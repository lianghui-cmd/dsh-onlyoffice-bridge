import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import test from 'node:test'
import { OnlyOfficeBridge, signJwt } from '../index.js'
import { createFakeContext, makeFileKey, startFakeDocumentServer, startRouteServer } from './helpers.mjs'

const SECRET = 'test-secret'

/** 建一个独立临时目录 + 已挂好路由的桥接服务。 */
async function setup(overrides = {}) {
  const storageDir = fs.mkdtempSync(path.join(os.tmpdir(), 'oo-bridge-'))
  const ctx = createFakeContext()
  const config = {
    dsUrl: 'http://127.0.0.1:9',
    jwtSecret: SECRET,
    jwtHeader: 'Authorization',
    jwtHeaderPrefixed: false,
    publicUrl: 'http://127.0.0.1:11006',
    editorDsUrl: '',
    storageDir,
    cacheTtlHour: 24,
    cacheCleanIntervalMin: 360,
    maxFileMb: 1,
    allowedOrigins: [],
    ...overrides,
  }
  // 插件在构造函数里通过 ctx 注册路由与工具；DS 地址在用例里可以覆盖
  const bridge = new OnlyOfficeBridge(ctx, { ...config, dsUrl: overrides.dsUrl ?? config.dsUrl })
  const server = await startRouteServer(ctx.routes)
  return { ctx, bridge, config, storageDir, server, routes: ctx.routes }
}

function jsonRequest(url, body, headers = {}) {
  return fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', ...headers },
    body: JSON.stringify(body),
  })
}

test('open：生成带 token 的 api.js 链接，并把文件写进缓存', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'src.docx')
  fs.writeFileSync(source, 'docx-bytes')
  const result = await env.bridge.open({ filePath: source })
  assert.match(result.fileKey, /^[a-z0-9]{32}-[0-9a-f]{16}\.docx$/)
  assert.ok(result.editorUrl.includes('/web-apps/apps/api/documents/api.js?token='))
  assert.equal(result.callbackUrl, 'http://127.0.0.1:11006/onlyoffice/callback')
  assert.equal(result.downloadUrl, `http://127.0.0.1:11006/onlyoffice/static-file/${result.fileKey}`)
  assert.equal(result.config.documentType, 'word')
  assert.equal(result.config.document.fileType, 'docx')
  assert.equal(result.config.editorConfig.callbackUrl, result.callbackUrl)
  assert.equal(result.config.document.url, `http://127.0.0.1:11006/onlyoffice/static-file/${result.fileKey}`)
  const cached = path.join(env.storageDir, result.fileKey)
  assert.ok(fs.existsSync(cached))
  assert.equal(fs.readFileSync(cached, 'utf8'), 'docx-bytes')
})

test('open：每次打开都换新 key（DS 编辑会话不复用旧缓存），同时记录原文件路径', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'again.docx')
  fs.writeFileSync(source, 'v1')
  const first = await env.bridge.open({ filePath: source })
  const second = await env.bridge.open({ fileKey: first.fileKey })
  assert.notEqual(first.fileKey, second.fileKey)
  const meta = JSON.parse(fs.readFileSync(path.join(env.storageDir, `${second.fileKey.split('.')[0]}.meta.json`), 'utf8'))
  assert.equal(meta.lineage.originalPath, source)
  assert.match(meta.lineage.rootKey, /^[a-z0-9]{32}-[0-9a-f]{16}$/)
  // 上一代编辑缓存被回收，只保留当前这一代
  assert.equal(fs.existsSync(path.join(env.storageDir, first.fileKey)), false)
  assert.equal(fs.existsSync(path.join(env.storageDir, second.fileKey)), true)
})

test('open：浏览器上传后可以用 fileKey 续开', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const response = await fetch(`${env.server.origin}/onlyoffice/upload?name=web.pptx`, {
    method: 'POST',
    body: Buffer.from('ppt'),
  })
  const uploaded = await response.json()
  const opened = await env.bridge.open({ fileKey: uploaded.fileKey })
  assert.match(opened.fileKey, /\.pptx$/)
  assert.equal(opened.config.documentType, 'slide')
  assert.equal(opened.config.document.title, 'web.pptx')
})

test('open：拒绝不存在的文件与不支持的扩展名', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  await assert.rejects(env.bridge.open({ filePath: path.join(env.storageDir, 'missing.docx') }), /文件不存在/)
  await assert.rejects(env.bridge.open({ fileName: 'evil.exe', contentBase64: Buffer.from('x').toString('base64') }), /不支持/)
})

test('static-file：返回文件内容，拒绝目录穿越与非法 key', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const key = makeFileKey('docx')
  fs.writeFileSync(path.join(env.storageDir, key), 'hello')
  const ok = await fetch(`${env.server.origin}/onlyoffice/static-file/${key}`)
  assert.equal(ok.status, 200)
  assert.equal(await ok.text(), 'hello')
  assert.match(ok.headers.get('content-type') ?? '', /wordprocessingml/)
  const trav = await fetch(`${env.server.origin}/onlyoffice/static-file/..%2f..%2fpackage.json`)
  assert.equal(trav.status, 404)
  const bad = await fetch(`${env.server.origin}/onlyoffice/static-file/${'z'.repeat(32)}.docx`)
  assert.equal(bad.status, 404)
})

test('callback：status=2 下载并回写缓存与原文件，响应 {"error":0}', async (t) => {
  const ds = await startFakeDocumentServer(Buffer.from('PK\u0003\u0004 updated'))
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await env.server.close()
    await ds.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'report.docx')
  fs.writeFileSync(source, 'old')
  const opened = await env.bridge.open({ filePath: source })
  const response = await jsonRequest(
    `${env.server.origin}/onlyoffice/callback`,
    { token: signJwt({ key: opened.fileKey, status: 2, url: `${ds.origin}/file` }, SECRET) },
  )
  assert.equal(response.status, 200)
  assert.deepEqual(await response.json(), { error: 0 })
  assert.equal(fs.readFileSync(source, 'utf8'), ds.payload.toString())
  assert.equal(fs.readFileSync(path.join(env.storageDir, opened.fileKey), 'utf8'), ds.payload.toString())
  assert.equal(env.bridge.lastCallback.status, 2)
  assert.equal(env.bridge.lastCallback.error, null)
})

test('callback：status=6（自动保存）同样落盘', async (t) => {
  const ds = await startFakeDocumentServer(Buffer.from('autosaved'))
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await env.server.close()
    await ds.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'auto.docx')
  fs.writeFileSync(source, 'old')
  const opened = await env.bridge.open({ filePath: source })
  const response = await jsonRequest(
    `${env.server.origin}/onlyoffice/callback`,
    { token: signJwt({ key: opened.fileKey, status: 6, url: `${ds.origin}/file` }, SECRET) },
  )
  assert.deepEqual(await response.json(), { error: 0 })
  assert.equal(fs.readFileSync(source, 'utf8'), 'autosaved')
})

test('callback：错误密钥 / 缺少 token 一律 403', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const forged = await jsonRequest(
    `${env.server.origin}/onlyoffice/callback`,
    { token: signJwt({ key: 'x', status: 2 }, 'wrong-secret') },
  )
  assert.equal(forged.status, 403)
  const missing = await jsonRequest(`${env.server.origin}/onlyoffice/callback`, { key: 'x', status: 2 })
  assert.equal(missing.status, 403)
})

test('callback：请求头携带 token 也可通过', async (t) => {
  const ds = await startFakeDocumentServer(Buffer.from('headered'))
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await env.server.close()
    await ds.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'header.docx')
  fs.writeFileSync(source, 'old')
  const opened = await env.bridge.open({ filePath: source })
  const token = signJwt({ payload: { key: opened.fileKey, status: 2, url: `${ds.origin}/file` } }, SECRET)
  const response = await jsonRequest(`${env.server.origin}/onlyoffice/callback`, {}, { Authorization: token })
  assert.equal(response.status, 200)
  assert.equal(fs.readFileSync(source, 'utf8'), 'headered')
})

test('upload：接收浏览器上传，返回 fileKey；非法来源被拒绝', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const ok = await fetch(`${env.server.origin}/onlyoffice/upload?name=demo.xlsx`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/octet-stream' },
    body: Buffer.from('sheet'),
  })
  assert.equal(ok.status, 200)
  const body = await ok.json()
  assert.match(body.fileKey, /\.xlsx$/)
  assert.ok(fs.existsSync(path.join(env.storageDir, body.fileKey)))
  const untrusted = await fetch(`${env.server.origin}/onlyoffice/upload?name=demo.docx`, {
    method: 'POST',
    headers: { Origin: 'https://evil.example' },
    body: Buffer.from('x'),
  })
  assert.equal(untrusted.status, 403)
})

test('health：报告 DS 地址、存储目录与最近回调', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const response = await fetch(`${env.server.origin}/onlyoffice/health`)
  const body = await response.json()
  assert.equal(body.ok, true)
  assert.equal(body.storageDir, env.storageDir)
  assert.equal(body.storageWritable, true)
})

test('exportPdf：向 DS 发 forcesave 命令，只接受 docx/xlsx/pptx', async (t) => {
  const ds = await startFakeDocumentServer()
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await env.server.close()
    await ds.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'export.docx')
  fs.writeFileSync(source, 'x')
  const opened = await env.bridge.open({ filePath: source })
  const result = await env.bridge.exportPdf({ fileKey: opened.fileKey, outputPath: path.join(env.storageDir, 'out.pdf') })
  assert.equal(result.ok, true)
  assert.equal(result.outputPath, path.join(env.storageDir, 'out.pdf'))
  assert.equal(ds.commands.length, 1)
  assert.equal(ds.commands[0].body.c, 'forcesave')
  assert.ok(ds.commands[0].headers.authorization, '命令请求必须带 JWT')

  const txtKey = makeFileKey('txt')
  fs.writeFileSync(path.join(env.storageDir, txtKey), 'x')
  await assert.rejects(env.bridge.exportPdf({ fileKey: txtKey }), /不支持导出 PDF/)
})

test('exportPdf：保存回调返回 PDF 时写入 outputPath', async (t) => {
  const pdf = Buffer.from('%PDF-1.7 fake')
  const ds = await startFakeDocumentServer(pdf)
  const env = await setup({ dsUrl: ds.origin })
  t.after(async () => {
    await env.server.close()
    await ds.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'conv.docx')
  fs.writeFileSync(source, 'x')
  const opened = await env.bridge.open({ filePath: source })
  const output = path.join(env.storageDir, 'nested', 'converted.pdf')
  const exported = await env.bridge.exportPdf({ fileKey: opened.fileKey, outputPath: output })
  await jsonRequest(
    `${env.server.origin}/onlyoffice/callback`,
    { token: signJwt({ key: exported.fileKey, status: 2, url: `${ds.origin}/file` }, SECRET) },
  )
  assert.ok(fs.existsSync(output))
  assert.equal(fs.readFileSync(output).subarray(0, 5).toString(), '%PDF-')
})

test('cleanExpired：只删除超过 TTL 的缓存；purge 删除单个 key', async (t) => {
  const env = await setup({ cacheTtlHour: 1 })
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const oldKey = makeFileKey('docx')
  const newKey = makeFileKey('docx')
  const oldPath = path.join(env.storageDir, oldKey)
  const newPath = path.join(env.storageDir, newKey)
  fs.writeFileSync(oldPath, 'old')
  fs.writeFileSync(newPath, 'new')
  const twoHoursAgo = new Date(Date.now() - 2 * 60 * 60 * 1000)
  fs.utimesSync(oldPath, twoHoursAgo, twoHoursAgo)

  const removed = await env.bridge.cleanExpired()
  assert.deepEqual(removed, [oldKey])
  assert.equal(fs.existsSync(oldPath), false)
  assert.equal(fs.existsSync(newPath), true)

  const purged = env.bridge.purge(newKey)
  assert.equal(purged.ok, true)
  assert.equal(fs.existsSync(newPath), false)
  assert.throws(() => env.bridge.purge('../escape'), /非法/)
})

test('status 工具列出缓存文档与最近回调', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const source = path.join(env.storageDir, 'listed.pptx')
  fs.writeFileSync(source, 'x')
  const opened = await env.bridge.open({ filePath: source })
  const tool = env.ctx.tools.get('onlyoffice_status')
  const value = await tool.execute({}, {})
  assert.equal(value.documents.length, 1)
  assert.equal(value.documents[0].fileKey, opened.fileKey)
  assert.equal(value.documents[0].originalPath, source)
  const content = tool.output.render({}, value)
  assert.match(content[0].text, /listed\.pptx/)
})

test('注册了完整工具清单（25 个），命名与参数结构符合 DSH 规范', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const names = [...env.ctx.tools.keys()].sort()
  const expected = [
    'onlyoffice_batch-export-pdf',
    'onlyoffice_batch-open',
    'onlyoffice_batch-read',
    'onlyoffice_diff',
    'onlyoffice_export-docx',
    'onlyoffice_export-html',
    'onlyoffice_export-markdown',
    'onlyoffice_export-pdf',
    'onlyoffice_export-pptx',
    'onlyoffice_export-xlsx',
    'onlyoffice_insert-image',
    'onlyoffice_lock',
    'onlyoffice_meta',
    'onlyoffice_open',
    'onlyoffice_purge',
    'onlyoffice_read-doc',
    'onlyoffice_rename',
    'onlyoffice_status',
    'onlyoffice_template-fill',
    'onlyoffice_unlock',
    'onlyoffice_who-lock',
    'onlyoffice_write-doc',
  ].sort()
  assert.deepEqual(names, expected)
  for (const name of names) {
    const tool = env.ctx.tools.get(name)
    assert.equal(typeof tool.description, 'string')
    assert.ok(tool.description.length > 10, `${name} 描述过短`)
    assert.equal(tool.parameters.type, 'object')
    assert.ok(tool.output && typeof tool.output.render === 'function')
    assert.equal(typeof tool.execute, 'function')
  }
})

test('purge 工具：target=one 需要 fileKey，target=expired 返回删除列表', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
    await env.ctx.dispose()
  })
  const purge = env.ctx.tools.get('onlyoffice_purge')
  await assert.rejects(purge.execute({ target: 'one' }, {}), /fileKey/)
  const expired = await purge.execute({ target: 'expired' }, {})
  assert.equal(expired.target, 'expired')
  assert.ok(Array.isArray(expired.removed))
})

test('卸载时释放路由、工具与定时器', async (t) => {
  const env = await setup()
  t.after(async () => {
    await env.server.close()
  })
  assert.equal(new Set(env.routes.map((route) => route.path)).size, 4)
  assert.equal(env.ctx.tools.size, 22)
  env.bridge.unregisterTools()
  assert.equal(env.ctx.tools.size, 0)
  env.bridge.registerTools()
  assert.equal(env.ctx.tools.size, 22)
  const disposeRoutes = env.ctx.routes.splice(0).length
  assert.equal(disposeRoutes, 4)
  // 重复注册不会产生重复工具，也不会换掉已有的注册
  const before = [...env.ctx.tools.keys()].sort()
  env.bridge.registerTools()
  env.bridge.registerTools()
  assert.deepEqual([...env.ctx.tools.keys()].sort(), before)
  const toolDisposers = env.bridge.toolDisposers.flat().slice()
  assert.equal(toolDisposers.length, 22)
  for (const dispose of toolDisposers) dispose()
  assert.equal(env.ctx.tools.size, 0)
  await env.ctx.dispose()
})
