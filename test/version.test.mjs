import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'
import { OnlyOfficeBridge } from '../index.js'
import { createFakeContext, makeFileKey } from './helpers.mjs'

/** 建一个孤立缓存目录 + 桥接服务（不起 http 服务）。 */
async function setup(overrides = {}) {
  const storageDir = fs.mkdtempSync(path.join(os.tmpdir(), 'oo-ver-'))
  const ctx = createFakeContext()
  const bridge = new OnlyOfficeBridge(ctx, {
    dsUrl: 'http://127.0.0.1:9',
    jwtSecret: 'test-secret',
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
  const call = (name, args) => ctx.tools.get(`onlyoffice_${name}`).execute(args, {})
  return { ctx, bridge, storageDir, call }
}

function seed(env, name, content = 'x') {
  const key = makeFileKey(path.extname(name).slice(1) || 'docx')
  fs.writeFileSync(path.join(env.storageDir, key), content)
  fs.writeFileSync(
    path.join(env.storageDir, `${key.replace(/\.[^.]+$/, '')}.meta.json`),
    JSON.stringify({ fileKey: key, fileName: name }),
  )
  return key
}

test('snapshot / snapshot-list / restore-snapshot：完整版本流转', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seed(env, '合同.docx', 'v1')

  const s1 = await env.call('snapshot', { fileKey: key, label: '初稿' })
  assert.equal(s1.ok, true)
  // snapId = 毫秒时间戳(17位) + 同毫秒自增序号(3位) + 随机串
  assert.match(s1.snapId, /^\d{17}\d{3}-[0-9a-f]{6}$/)
  assert.equal(s1.label, '初稿')

  fs.writeFileSync(path.join(env.storageDir, key), 'v2')
  const s2 = await env.call('snapshot', { fileKey: key, label: '审核版' })

  const list = await env.call('snapshot-list', { fileKey: key })
  assert.equal(list.total, 2)
  assert.deepEqual(list.snapshots.map((s) => s.label), ['初稿', '审核版'])

  const restored = await env.call('restore-snapshot', { fileKey: key, snapId: s1.snapId })
  assert.equal(restored.ok, true)
  assert.equal(fs.readFileSync(path.join(env.storageDir, key), 'utf8'), 'v1')
  // 回滚前的版本被自动备份
  const backupPath = path.join(env.storageDir, 'snapshots', key.replace(/\.[^.]+$/, ''), `${restored.backupSnapId}.docx`)
  assert.equal(fs.readFileSync(backupPath, 'utf8'), 'v2')

  const after = await env.call('snapshot-list', { fileKey: key })
  assert.equal(after.total, 3)
  void s2
})

test('snapshot：fileKey 不存在 / 快照不存在时给出可读错误', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  await assert.rejects(env.call('snapshot', { fileKey: 'missing.docx' }), /fileKey 不存在/)
  const key = seed(env, 'a.docx')
  await assert.rejects(env.call('restore-snapshot', { fileKey: key, snapId: 'nope' }), /快照不存在/)
  await assert.rejects(env.call('restore-snapshot', { fileKey: key }), /snapId 不能为空/)
})

test('快照路径穿越被拒绝', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  await assert.rejects(env.call('snapshot', { fileKey: '../../x.docx' }), /fileKey 非法/)
  await assert.rejects(env.call('snapshot-list', { fileKey: 'a/b.docx' }), /fileKey 非法/)
})

test('merge-doc / split-doc / toc-regen / header-footer-set：参数校验', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())

  await assert.rejects(env.call('merge-doc', { fileKeyList: [] }), /不能为空/)
  await assert.rejects(env.call('merge-doc', { fileKeyList: ['only-one.docx'] }), /至少需要 2 个/)
  await assert.rejects(env.call('split-doc', { fileKey: 'missing.docx' }), /fileKey 不存在/)

  // 非 Office 扩展名：按扩展名提前拦截，错误信息要明确
  const txtKey = seed(env, 'a.txt', 'not office')
  await assert.rejects(env.call('toc-regen', { fileKey: txtKey }), /仅支持 docx \/ pptx/)
  await assert.rejects(env.call('header-footer-set', { fileKey: txtKey }), /仅支持 docx/)

  // 后缀是 .docx 但内容损坏：应给出可读的「不是有效的 Office 文档」，而不是 jszip 原始报错
  const brokenKey = seed(env, 'broken.docx', 'this is not a zip')
  await assert.rejects(env.call('split-doc', { fileKey: brokenKey }), /不是有效的 Office 文档/)
  await assert.rejects(env.call('toc-regen', { fileKey: brokenKey }), /不是有效的 Office 文档/)
  await assert.rejects(
    env.call('header-footer-set', { fileKey: brokenKey, headerText: 'h' }),
    /不是有效的 Office 文档/,
  )
})

test('header-footer-set：至少要设置一项', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  // 需要一个真实可解析的 docx 才能走到参数校验
  const { loadDocx } = await import('../lib/optional-deps.js')
  const { markdownToDocxBuffer } = await import('../lib/document.js')
  const docx = await loadDocx()
  const buffer = await markdownToDocxBuffer('# 标题\n\n正文', docx)
  const key = makeFileKey('docx')
  fs.writeFileSync(path.join(env.storageDir, key), buffer)
  await assert.rejects(env.call('header-footer-set', { fileKey: key }), /至少要设置/)
})

test('lock 会阻止 snapshot 之外的写操作', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seed(env, 'c.docx')
  await env.call('lock', { fileKey: key, lockedBy: 'alice' })
  await assert.rejects(env.call('restore-snapshot', { fileKey: key, snapId: 'x', lockedBy: 'bob' }), /已被 alice 锁定/)
  await assert.rejects(env.call('header-footer-set', { fileKey: key, headerText: 'h', lockedBy: 'bob' }), /已被 alice 锁定/)
})

test('purge 时快照一起清理，不留孤儿目录', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const key = seed(env, 'd.docx')
  await env.call('snapshot', { fileKey: key, label: 'x' })
  const snapDir = path.join(env.storageDir, 'snapshots', key.replace(/\.[^.]+$/, ''))
  assert.equal(fs.existsSync(snapDir), true)

  // 直接走 bridge.purge（工具层的 target=one 会连带清理快照）
  const result = env.bridge.purge(key)
  assert.equal(result.ok, true)
  assert.equal(fs.existsSync(snapDir), false, '快照目录应随文档一起删除')
})

test('新增的 7 个工具都注册且参数结构合法', async (t) => {
  const env = await setup()
  t.after(() => env.ctx.dispose())
  const names = [
    'onlyoffice_toc-regen',
    'onlyoffice_header-footer-set',
    'onlyoffice_merge-doc',
    'onlyoffice_split-doc',
    'onlyoffice_snapshot',
    'onlyoffice_restore-snapshot',
    'onlyoffice_snapshot-list',
  ]
  for (const name of names) {
    const tool = env.ctx.tools.get(name)
    assert.ok(tool, `${name} 未注册`)
    assert.equal(typeof tool.description, 'string')
    assert.ok(tool.description.length > 15, `${name} 描述过短`)
    assert.equal(tool.parameters.type, 'object')
    assert.ok(Array.isArray(tool.parameters.required))
    assert.equal(typeof tool.execute, 'function')
  }
  // split-doc 的 mode 必须是枚举
  const split = env.ctx.tools.get('onlyoffice_split-doc')
  assert.deepEqual(split.parameters.properties.mode.enum, ['section', 'slide'])
})
