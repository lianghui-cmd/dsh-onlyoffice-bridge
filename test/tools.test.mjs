import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'
import { OnlyOfficeBridge } from '../index.js'
import { cmToPx, collectPlaceholders, escapeForTest, imageTypeOf, renderPlaceholders } from './doc-helpers.mjs'
import { createFakeContext, makeFileKey, startFakeDocumentServer } from './helpers.mjs'

const SECRET = 'test-secret'

/** 建一个孤立缓存目录 + 桥接服务（不起 http 服务）。 */
async function setup(overrides = {}) {
  const storageDir = fs.mkdtempSync(path.join(os.tmpdir(), 'oo-tools-'))
  const ctx = createFakeContext()
  const bridge = new OnlyOfficeBridge(ctx, {
    dsUrl: 'http://127.0.0.1:9',
    jwtSecret: SECRET,
    jwtHeader: 'Authorization',
    jwtHeaderPrefixed: false,
    publicUrl: 'http://127.0.0.1:11006',
    editorDsUrl: '',
    storageDir,
    cacheTtlHour: 24,
    cacheCleanIntervalMin: 360,
    maxFileMb: 8,
    allowedOrigins: [],
    ...overrides,
  })
  const tool = (name) => {
    const definition = ctx.tools.get(`onlyoffice_${name}`)
    if (!definition) throw new Error(`工具未注册：${name}`)
    return definition
  }
  return { ctx, bridge, storageDir, tool, call: (name, args) => tool(name).execute(args, {}) }
}

function seedDoc(env, name, content = 'hello') {
  const key = makeFileKey(path.extname(name).slice(1) || 'docx')
  fs.writeFileSync(path.join(env.storageDir, key), content)
  fs.writeFileSync(
    path.join(env.storageDir, `${key.replace(/\.[^.]+$/, '')}.meta.json`),
    JSON.stringify({ fileKey: key, fileName: name, lineage: { originalPath: null, rootKey: key.replace(/\.[^.]+$/, '') } }),
  )
  return key
}

test('锁：lock / who-lock / unlock，写操作被锁拦截', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seedDoc(env, '合同.docx')

  const locked = await env.call('lock', { fileKey: key, lockedBy: 'alice', note: '改第三章' })
  assert.equal(locked.ok, true)
  assert.equal(locked.lock.lockedBy, 'alice')

  const again = await env.call('lock', { fileKey: key, lockedBy: 'bob' })
  assert.equal(again.ok, false)
  assert.match(again.msg, /已被 alice 锁定/)

  const who = await env.call('who-lock', { fileKey: key })
  assert.equal(who.locked, true)
  assert.equal(who.lockInfo.lockedBy, 'alice')

  // 被锁后 open 会被拒绝
  await assert.rejects(env.call('open', { fileKey: key, lockedBy: 'bob' }), /已被 alice 锁定/)
  // 锁持有者自己可以打开
  const opened = await env.call('open', { fileKey: key, lockedBy: 'alice' })
  assert.ok(opened.editorUrl)

  // 非持有者不能解锁
  const denied = await env.call('unlock', { fileKey: key, lockedBy: 'bob' })
  assert.equal(denied.ok, false)
  const unlocked = await env.call('unlock', { fileKey: key, lockedBy: 'alice' })
  assert.equal(unlocked.ok, true)
  assert.equal((await env.call('who-lock', { fileKey: key })).locked, false)
})

test('rename / meta：修改显示名与备注，rename 自动补扩展名', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seedDoc(env, '旧名.docx')
  const renamed = await env.call('rename', { fileKey: key, name: '新版合同', remark: '2026 评审' })
  assert.equal(renamed.ok, true)
  assert.equal(renamed.meta.fileName, '新版合同.docx')
  assert.equal(renamed.meta.remark, '2026 评审')

  const meta = await env.call('meta', { fileKey: key })
  assert.equal(meta.ok, true)
  assert.equal(meta.meta.fileName, '新版合同.docx')
  assert.equal(meta.lock, null)
  assert.ok(meta.size > 0)
})

test('status：列出缓存文档并附带锁状态；purge 清理元信息与锁', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seedDoc(env, '报表.xlsx', 'x')
  await env.call('lock', { fileKey: key, lockedBy: 'alice' })

  const status = await env.call('status', {})
  assert.equal(status.total, 1)
  assert.equal(status.documents[0].fileKey, key)
  assert.equal(status.documents[0].lock.lockedBy, 'alice')

  const purged = await env.call('purge', { target: 'one', fileKey: key })
  assert.equal(purged.ok, true)
  assert.equal(fs.existsSync(path.join(env.storageDir, key)), false)
  assert.equal((await env.call('who-lock', { fileKey: key })).locked, false)
})

test('安全：fileKey 路径穿越被拒绝', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  await assert.rejects(env.call('meta', { fileKey: '../../package.json' }), /fileKey 非法/)
  await assert.rejects(env.call('read-doc', { fileKey: 'a/../../b.docx' }), /fileKey 非法/)
  await assert.rejects(env.call('lock', { fileKey: '..\\x.docx' }), /fileKey 非法/)
})

test('batch-open / batch-read：逐项返回成功与失败原因', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const good = path.join(env.storageDir, 'a.docx')
  fs.writeFileSync(good, 'PK\u0003\u0004x', 'latin1')
  const result = await env.call('batch-open', { filePathList: [good, path.join(env.storageDir, 'missing.docx')] })
  assert.equal(result.total, 2)
  assert.equal(result.succeeded, 1)
  assert.equal(result.result[0].ok, true)
  assert.equal(result.result[1].ok, false)
  assert.match(result.result[1].msg, /不存在/)
})

test('read-doc：不支持格式给出明确提示而不抛错', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seedDoc(env, '幻灯片.pptx', 'PK')
  const result = await env.call('read-doc', { fileKey: key })
  assert.equal(result.ok, false)
  assert.match(result.msg, /pptx/)
})

test('export-docx / export-xlsx 等转换工具：未安装可选依赖时给出安装提示', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seedDoc(env, 'a.docx')
  // 转换类工具依赖 DS，测试环境把 DS 指向空端口，应得到可读的失败信息
  await assert.rejects(env.call('export-html', { fileKey: key }), /ConvertService|fetch|下载失败|HTTP/i)
  await assert.rejects(env.call('export-xlsx', { fileKey: key, outputPath: path.join(env.storageDir, 'x.xlsx') }), /ConvertService|fetch|下载失败|HTTP/i)
})

test('insert-image / template-fill：仅支持 docx，参数缺失时给出可读错误', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const xlsxKey = seedDoc(env, '表.xlsx', 'x')
  await assert.rejects(env.call('insert-image', { fileKey: xlsxKey, imgPath: 'x.png' }), /仅支持 docx/)
  const docKey = seedDoc(env, '文.docx', 'x')
  await assert.rejects(env.call('insert-image', { fileKey: docKey }), /imgPath/)
  await assert.rejects(env.call('template-fill', { templateKey: xlsxKey, vars: {} }), /仅支持 docx/)
  await assert.rejects(env.call('template-fill', { templateKey: docKey, vars: 'not-json' }), /JSON/)
})

test('docx 工具链纯函数：占位符收集/渲染、图片类型、厘米换算', () => {
  const xml =
    '<w:p><w:r><w:t>甲方：${party}</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>金额 </w:t></w:r><w:r><w:t>${amount}</w:t></w:r><w:r><w:t> 元</w:t></w:r></w:p>'
  const names = collectPlaceholders(xml)
  assert.deepEqual(names.sort(), ['amount', 'party'])

  const rendered = renderPlaceholders(xml, { party: '甲方<公司>', amount: 100 })
  assert.deepEqual(rendered.replaced.sort(), ['amount', 'party'])
  assert.equal(rendered.missing.length, 0)
  assert.match(rendered.xml, /甲方&lt;公司&gt;/)

  const missing = renderPlaceholders(xml, { party: 'x' })
  assert.deepEqual(missing.missing, ['amount'])

  const png = Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex')
  const jpg = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0x00, 0x10])
  assert.equal(imageTypeOf(png), 'png')
  assert.equal(imageTypeOf(jpg), 'jpg')
  assert.throws(() => imageTypeOf(Buffer.from('not an image')), /不支持的图片格式/)

  assert.equal(cmToPx(2.54), 96)
  assert.ok(cmToPx(14) > cmToPx(10))
  assert.equal(escapeForTest('<a&b>'), '&lt;a&amp;b&gt;')
})

test('export-pdf：只接受 docx/xlsx/pptx', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const txt = seedDoc(env, 'a.txt')
  await assert.rejects(env.call('export-pdf', { fileKey: txt }), /不支持导出 PDF/)
})

test('purge：target=all 跳过被锁文档', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const a = seedDoc(env, 'a.docx')
  const b = seedDoc(env, 'b.docx')
  await env.call('lock', { fileKey: a, lockedBy: 'alice' })
  const result = await env.call('purge', { target: 'all' })
  assert.equal(result.removed, 1)
  assert.equal(fs.existsSync(path.join(env.storageDir, a)), true)
  assert.equal(fs.existsSync(path.join(env.storageDir, b)), false)
})

test('导出状态：status 附带最近一次回调', async (t) => {
  const ds = await startFakeDocumentServer(Buffer.from('saved'))
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await ds.close()
    await env.ctx.dispose()
  })
  assert.equal(env.bridge.lastCallback, null)
  const status = await env.call('status', {})
  assert.equal(status.lastCallback, null)
  assert.equal(status.dsUrl, ds.origin)
})
