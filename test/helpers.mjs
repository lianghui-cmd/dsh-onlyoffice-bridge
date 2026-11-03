// 测试替身：真实的 DSH 由 @deepseek-ai/cordis 提供 Service，
// 仓库里独立运行测试时把 @deepseek-ai/cordis 映射到 test/shims/cordis.mjs（见 register.mjs）。
import crypto from 'node:crypto'

/** 极简 Context：effect / interval / webServer 路由表 / tools 注册表。 */
export function createFakeContext() {
  const effects = []
  const routes = []
  const tools = new Map()
  const logs = []
  const toolRegistry = {
    register(definition) {
      tools.set(definition.name, definition)
      return () => tools.delete(definition.name)
    },
    get: (name) => tools.get(name),
    has: (name) => tools.has(name),
    get size() {
      return tools.size
    },
    keys: () => tools.keys(),
  }
  const ctx = {
    logger: {
      info: (...args) => logs.push(['info', args]),
      warn: (...args) => logs.push(['warn', args]),
      error: (...args) => logs.push(['error', args]),
    },
    logs,
    effect(fn) {
      const dispose = fn()
      if (typeof dispose === 'function') effects.push(dispose)
      return dispose
    },
    interval(callback, delay) {
      const timer = setInterval(callback, delay)
      timer.unref?.()
      const dispose = () => clearInterval(timer)
      effects.push(dispose)
      return dispose
    },
    plugin(plugin, config) {
      return plugin(ctx, config)
    },
    webServer: {
      getOrigin: () => 'http://127.0.0.1:11006',
      register(route) {
        routes.push(route)
        return () => {
          const index = routes.indexOf(route)
          if (index >= 0) routes.splice(index, 1)
        }
      },
    },
    tools: toolRegistry,
    routes,
    async dispose() {
      while (effects.length) await effects.pop()()
    },
  }
  return ctx
}

/** 起一个真实的 http 服务，把请求转成 plugin 的 route handler。 */
export async function startRouteServer(routes) {
  const { createServer } = await import('node:http')
  const server = createServer((req, res) => {
    const pathname = new URL(req.url ?? '/', 'http://x').pathname
    let matched = routes.find((route) => route.kind === 'exact' && route.path === pathname)
    if (!matched) {
      matched = routes
        .filter((route) => route.kind === 'prefix' && (pathname === route.path || pathname.startsWith(`${route.path}/`)))
        .sort((a, b) => b.path.length - a.path.length)[0]
    }
    if (!matched) {
      res.writeHead(404)
      res.end()
      return
    }
    Promise.resolve(matched.handler(req, res)).catch((error) => {
      if (!res.headersSent) res.writeHead(500)
      res.end(String(error))
    })
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  const address = server.address()
  const origin = `http://127.0.0.1:${address.port}`
  return {
    origin,
    close: () => new Promise((resolve) => server.close(resolve)),
  }
}

/**
 * 模拟 ONLYOFFICE Document Server：/file 返回固定内容，/commands 记录调用。
 */
export async function startFakeDocumentServer(payload = Buffer.from('PK\u0003\u0004 saved', 'latin1')) {
  const { createServer } = await import('node:http')
  const commands = []
  const server = createServer((req, res) => {
    const url = new URL(req.url ?? '/', 'http://x')
    if (url.pathname === '/file') {
      res.writeHead(200, { 'Content-Type': 'application/octet-stream' })
      res.end(payload)
      return
    }
    if (url.pathname === '/coauthoring/CommandService.ashx') {
      let raw = ''
      req.on('data', (chunk) => (raw += chunk))
      req.on('end', () => {
        commands.push({ headers: req.headers, body: raw ? JSON.parse(raw) : null })
        res.writeHead(200, { 'Content-Type': 'application/json' })
        res.end(JSON.stringify({ error: 0, key: 'ok' }))
      })
      return
    }
    res.writeHead(404)
    res.end()
  })
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
  const address = server.address()
  return {
    origin: `http://127.0.0.1:${address.port}`,
    commands,
    payload,
    close: () => new Promise((resolve) => server.close(resolve)),
  }
}

/** 生成一个固定格式的 fileKey。 */
export function makeFileKey(ext = 'docx') {
  return `${crypto.randomBytes(16).toString('hex')}-${crypto.randomBytes(8).toString('hex')}.${ext}`
}
