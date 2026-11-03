// @ts-check
/**
 * dsh-onlyoffice-bridge
 *
 * DeepSeek Harness（DSH）Cordis 插件：把 ONLYOFFICE Document Server 桥接给 Agent。
 * 纯 Node 实现（node:http / node:crypto / node:fs），不依赖 express。
 *
 * 对外能力：
 *   GET  /onlyoffice/static-file/<fileKey>   DS 拉取源文件（每次编辑生成新 key）
 *   POST /onlyoffice/callback                编辑保存回调（status=2/6 下载并落回原文件）
 *   POST /onlyoffice/upload                  浏览器上传本地文件到缓存
 *   GET  /onlyoffice/health                  健康检查 + 最近一次回调状态
 *
 * Agent 工具：
 *   onlyoffice_open         生成在线编辑链接
 *   onlyoffice_status       列出 / 查看缓存文档与最近回调
 *   onlyoffice_export_pdf   导出 PDF（本地落盘）
 *   onlyoffice_purge        清理缓存（单个 / 过期 / 全部）
 *
 * 配置：进程环境变量优先于插件目录下的 .env；也可以在 cordis.patch.yml 的
 * config 里用同名 camelCase 字段覆盖（见 export const Config）。
 */

import { Service } from '@deepseek-ai/cordis'
import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { Readable } from 'node:stream'
import { pipeline } from 'node:stream/promises'
import { fileURLToPath } from 'node:url'
import { TOOL_DEFINITIONS } from './lib/definitions.js'
import { createToolkit } from './lib/tools.js'

/** @typedef {import('@deepseek-ai/cordis').Context} Context */
/** @typedef {import('node:http').IncomingMessage} IncomingMessage */
/** @typedef {import('node:http').ServerResponse} ServerResponse */

export const name = 'dsh-onlyoffice-bridge'
export const inject = ['webServer', 'tools']

/** 插件目录（用于定位 .env、内置 tool 定义等）。 */
const PLUGIN_DIR = path.dirname(fileURLToPath(import.meta.url))

const PLUGIN_TAG = 'dsh-onlyoffice-bridge'
const ROUTE_PREFIX = '/onlyoffice'
const STATIC_PREFIX = `${ROUTE_PREFIX}/static-file`
const CALLBACK_PATH = `${ROUTE_PREFIX}/callback`
const UPLOAD_PATH = `${ROUTE_PREFIX}/upload`
const HEALTH_PATH = `${ROUTE_PREFIX}/health`

/** 支持的 Office 格式：本地保存用的是同名替换，键名保持原样。 */
const OFFICE_EXTENSIONS = ['docx', 'doc', 'odt', 'rtf', 'txt', 'xlsx', 'xls', 'ods', 'csv', 'pptx', 'ppt', 'odp']
const CONVERTIBLE_EXTENSIONS = ['docx', 'xlsx', 'pptx']
/** ConvertService 可导出的目标格式。 */
const CONVERT_FORMATS = ['pdf', 'docx', 'xlsx', 'pptx', 'html', 'odt', 'ods', 'odp', 'rtf', 'txt', 'csv']
const FILE_KEY_STEM = /^[a-z0-9]{32}-[0-9a-f]{16}$/
const FILE_KEY_EXT = /^[a-z]{1,5}$/
const VERSION_ID = /^[a-z0-9]{32}$/

/** 回调 status 语义（ONLYOFFICE Callback handler 文档）。 */
const CALLBACK_STATUS = {
  1: '文档被打开（仅通知）',
  2: '文档已就绪保存，可以下载',
  3: '保存出错',
  4: '文档被无改动关闭',
  6: '强制保存（force save）',
  7: '强制保存出错',
}

/** 默认配置。环境变量 → 内置默认值。 */
function defaults() {
  return {
    dsUrl: 'http://127.0.0.1:8080',
    jwtSecret: 'replace-me-secret-key',
    jwtHeader: 'Authorization',
    jwtHeaderPrefixed: false,
    publicUrl: '',
    editorDsUrl: '',
    storageDir: '',
    cacheTtlHour: 24,
    cacheCleanIntervalMin: 360,
    maxFileMb: 128,
    allowedOrigins: [],
  }
}

/** 读取插件目录下的 .env（不覆盖进程里已经存在的环境变量）。 */
function loadDotEnv() {
  const file = path.join(PLUGIN_DIR, '.env')
  let text
  try {
    text = fs.readFileSync(file, 'utf8')
  } catch {
    return {}
  }
  /** @type {Record<string, string>} */
  const parsed = {}
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim()
    if (!line || line.startsWith('#')) continue
    const eq = line.indexOf('=')
    if (eq <= 0) continue
    const key = line.slice(0, eq).trim()
    let value = line.slice(eq + 1).trim()
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1)
    }
    parsed[key] = value
  }
  return parsed
}

/** 环境变量（进程优先，回退 .env）读取一个键。 */
function createEnvReader(dotenv) {
  /**
   * @param {string} key
   * @returns {string | undefined}
   */
  return (key) => {
    const fromProcess = process.env[key]
    if (fromProcess !== undefined && fromProcess !== '') return fromProcess
    const fromFile = dotenv[key]
    if (fromFile !== undefined && fromFile !== '') return fromFile
    return undefined
  }
}

/**
 * @param {string} value
 * @param {string | undefined} fallback
 */
function cleanUrl(value, fallback = '') {
  const source = value.trim()
  if (!source) return fallback
  // 允许省略协议：192.168.1.20:11006 → http://192.168.1.20:11006
  const withScheme = /^[a-z][a-z0-9+.-]*:\/\//i.test(source) ? source : `http://${source}`
  return withScheme.replace(/\/+$/, '')
}

/**
 * @param {string | undefined} value
 * @param {number} fallback
 */
function positiveInt(value, fallback) {
  const n = Number(value)
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : fallback
}

/**
 * @param {string | undefined} value
 */
function splitList(value) {
  if (!value) return []
  return value
    .split(',')
    .map((item) => item.trim().replace(/\/+$/, ''))
    .filter(Boolean)
}

/**
 * 读取环境配置。
 * @param {{ dotenv?: Record<string, string>, cwd?: string }} [options]
 */
export function readEnvConfig(options = {}) {
  const env = createEnvReader(options.dotenv ?? {})
  const cwd = options.cwd ?? process.cwd()
  const home = process.env.DSH_HOME || path.join(process.env.USERPROFILE || process.env.HOME || cwd, '.dsh')
  const storage = env('STORAGE_DIR')
  const base = defaults()
  return {
    dsUrl: cleanUrl(env('ONLYOFFICE_DS_URL') ?? base.dsUrl, base.dsUrl),
    jwtSecret: env('ONLYOFFICE_SECRET') ?? env('ONLYOFFICE_SECRETS') ?? base.jwtSecret,
    jwtHeader: env('ONLYOFFICE_JWT_HEADER') ?? base.jwtHeader,
    jwtHeaderPrefixed: (env('ONLYOFFICE_JWT_HEADER_PREFIXED') ?? 'false').toLowerCase() === 'true',
    publicUrl: cleanUrl(env('DSH_PUBLIC_URL') ?? base.publicUrl),
    editorDsUrl: cleanUrl(env('EDITOR_DS_URL') ?? base.editorDsUrl),
    storageDir: storage ? path.resolve(cwd, storage) : path.join(home, 'onlyoffice'),
    cacheTtlHour: positiveInt(env('CACHE_TTL_HOUR'), base.cacheTtlHour),
    cacheCleanIntervalMin: positiveInt(env('CACHE_CLEAN_INTERVAL_MIN'), base.cacheCleanIntervalMin),
    maxFileMb: positiveInt(env('MAX_FILE_MB'), base.maxFileMb) || base.maxFileMb,
    allowedOrigins: splitList(env('ALLOWED_ORIGINS')),
  }
}

/**
 * 环境变量优先，其次插件 config，最后内置默认值。
 * @param {{ dotenv?: Record<string, string>, cwd?: string }} [envOptions]
 * @param {Record<string, unknown>} [config]
 */
export function resolveConfig(envOptions = {}, config = {}) {
  const defaultsConfig = defaults()
  const env = createEnvReader(envOptions.dotenv ?? {})
  const pick = (key) => {
    const value = config[key]
    return value === undefined || value === '' ? undefined : value
  }
  const pickString = (key) => {
    const value = pick(key)
    return typeof value === 'string' ? value : undefined
  }
  const pickNumber = (key) => {
    const value = pick(key)
    return typeof value === 'number' && Number.isFinite(value) ? positiveInt(String(value), -1) : undefined
  }
  const cwd = envOptions.cwd ?? process.cwd()
  const home = process.env.DSH_HOME || path.join(process.env.USERPROFILE || process.env.HOME || cwd, '.dsh')
  const storageFromEnv = env('STORAGE_DIR')
  const storageFromConfig = pickString('storageDir')
  const allowedFromConfig = Array.isArray(config.allowedOrigins)
    ? config.allowedOrigins.map((v) => String(v).replace(/\/+$/, '')).filter(Boolean)
    : undefined

  return {
    dsUrl: cleanUrl(env('ONLYOFFICE_DS_URL') ?? pickString('dsUrl') ?? '', defaultsConfig.dsUrl),
    jwtSecret: env('ONLYOFFICE_SECRET') ?? env('ONLYOFFICE_SECRETS') ?? pickString('jwtSecret') ?? defaultsConfig.jwtSecret,
    jwtHeader: env('ONLYOFFICE_JWT_HEADER') ?? pickString('jwtHeader') ?? defaultsConfig.jwtHeader,
    jwtHeaderPrefixed: (env('ONLYOFFICE_JWT_HEADER_PREFIXED') ?? String(config.jwtHeaderPrefixed ?? defaultsConfig.jwtHeaderPrefixed)).toLowerCase() === 'true',
    publicUrl: cleanUrl(env('DSH_PUBLIC_URL') ?? pickString('publicUrl') ?? '', defaultsConfig.publicUrl),
    editorDsUrl: cleanUrl(env('EDITOR_DS_URL') ?? pickString('editorDsUrl') ?? '', defaultsConfig.editorDsUrl),
    storageDir: storageFromEnv
      ? path.resolve(cwd, storageFromEnv)
      : storageFromConfig
        ? path.resolve(cwd, storageFromConfig)
        : path.join(home, 'onlyoffice'),
    cacheTtlHour: pickNumber('cacheTtlHour') ?? positiveInt(env('CACHE_TTL_HOUR'), defaultsConfig.cacheTtlHour),
    cacheCleanIntervalMin: pickNumber('cacheCleanIntervalMin') ?? positiveInt(env('CACHE_CLEAN_INTERVAL_MIN'), defaultsConfig.cacheCleanIntervalMin),
    maxFileMb:
      pickNumber('maxFileMb') || positiveInt(env('MAX_FILE_MB'), defaultsConfig.maxFileMb) || defaultsConfig.maxFileMb,
    allowedOrigins: allowedFromConfig ?? splitList(env('ALLOWED_ORIGINS')),
  }
}

/** @typedef {ReturnType<typeof readEnvConfig>} ResolvedConfig */

// ---------------------------------------------------------------------------
// JWT（ONLYOFFICE 约定的 HS256，payload 放在 body.token / Authorization 头）
// ---------------------------------------------------------------------------

/**
 * @param {string} part
 */
function base64url(part) {
  return Buffer.from(part, 'utf8').toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

/**
 * @param {string} part
 */
function fromBase64url(part) {
  const padded = part.replace(/-/g, '+').replace(/_/g, '/') + '='.repeat((4 - (part.length % 4)) % 4)
  return Buffer.from(padded, 'base64').toString('utf8')
}

/**
 * ONLYOFFICE 标准 HS256 签名。
 * @param {Record<string, unknown>} payload
 * @param {string} secret
 * @returns {string}
 */
export function signJwt(payload, secret) {
  const header = base64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }))
  const body = base64url(JSON.stringify(payload))
  const data = `${header}.${body}`
  const signature = crypto.createHmac('sha256', secret).update(data).digest('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
  return `${data}.${signature}`
}

/**
 * 校验并解出 payload；失败返回 undefined（不会抛异常）。
 * @param {string} token
 * @param {string} secret
 * @returns {Record<string, any> | undefined}
 */
export function verifyJwt(token, secret) {
  if (typeof token !== 'string') return undefined
  const parts = token.split('.')
  if (parts.length !== 3) return undefined
  const [header, body, signature] = parts
  const expected = crypto.createHmac('sha256', secret).update(`${header}.${body}`).digest('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
  const given = Buffer.from(signature)
  const want = Buffer.from(expected)
  if (given.length !== want.length || !crypto.timingSafeEqual(given, want)) return undefined
  try {
    const parsed = JSON.parse(fromBase64url(body))
    return parsed && typeof parsed === 'object' ? parsed : undefined
  } catch {
    return undefined
  }
}

/**
 * 旧版 ONLYOFFICE 使用简单 md5 签名，部分私有化部署仍开启。这里只做校验兜底。
 * @param {string} token
 * @param {string} secret
 */
function verifyLegacyToken(token, secret) {
  if (typeof token !== 'string' || !token.startsWith('#')) return undefined
  const body = token.slice(1)
  const [payload, signature] = body.split('.')
  if (!payload || !signature) return undefined
  const expected = crypto.createHash('md5').update(`${secret}${payload}`).digest('hex')
  if (signature !== expected) return undefined
  try {
    const parsed = JSON.parse(Buffer.from(payload, 'base64').toString('utf8'))
    return parsed && typeof parsed === 'object' ? parsed : undefined
  } catch {
    return undefined
  }
}

/**
 * 校验请求方带来的 token：请求头优先，其次 body.token（新版 DS 默认放 body.token）。
 * 也支持 query.token（重定向 / 下载场景）。
 * @param {IncomingMessage} req
 * @param {Record<string, any>} body
 * @param {{ jwtHeader: string, jwtHeaderPrefixed: boolean, jwtSecret: string }} config
 */
export function extractRequestToken(req, body, config) {
  const headerName = config.jwtHeader.trim().toLowerCase()
  const headerValue = req.headers[headerName]
  if (typeof headerValue === 'string' && headerValue) {
    const raw = headerValue.replace(/^\s*Bearer\s+/i, '')
    const payload = verifyJwt(raw, config.jwtSecret) ?? verifyLegacyToken(raw, config.jwtSecret)
    if (payload) return payload
  }
  if (body && typeof body.token === 'string') {
    const payload = verifyJwt(body.token, config.jwtSecret) ?? verifyLegacyToken(body.token, config.jwtSecret)
    if (payload) return payload
  }
  const url = new URL(req.url ?? '/', 'http://localhost')
  const queryToken = url.searchParams.get('token')
  if (queryToken) {
    const payload = verifyJwt(queryToken, config.jwtSecret) ?? verifyLegacyToken(queryToken, config.jwtSecret)
    if (payload) return payload
  }
  return undefined
}

/**
 * 签名并返回带 token 的请求头对象。
 * @param {Record<string, unknown>} payload
 * @param {{ jwtHeader: string, jwtHeaderPrefixed: boolean, jwtSecret: string }} config
 */
export function authHeaders(payload, config) {
  const token = signJwt(payload, config.jwtSecret)
  const headers = { Authorization: config.jwtHeaderPrefixed ? `Bearer ${token}` : token }
  if (config.jwtHeader.trim().toLowerCase() !== 'authorization') {
    headers[config.jwtHeader] = config.jwtHeaderPrefixed ? `Bearer ${token}` : token
  }
  return { token, headers }
}

// ---------------------------------------------------------------------------
// 小工具
// ---------------------------------------------------------------------------

/** @param {unknown} value */
function asRecord(value) {
  return value && typeof value === 'object' ? /** @type {Record<string, any>} */ (value) : {}
}

/** @param {string | string[] | undefined} value */
function firstHeader(value) {
  return Array.isArray(value) ? value[0] : value
}

/**
 * 解析并限制请求体大小。
 * @param {IncomingMessage} req
 * @param {number} limitBytes
 */
function readBody(req, limitBytes) {
  return new Promise((resolve, reject) => {
    /** @type {Buffer[]} */
    const chunks = []
    let size = 0
    req.on('data', (chunk) => {
      size += chunk.length
      if (size > limitBytes) {
        const error = new Error(`请求体超过上限 ${Math.round(limitBytes / 1024 / 1024)} MiB`)
        /** @type {any} */ (error).code = 'BODY_TOO_LARGE'
        reject(error)
        req.destroy()
        return
      }
      chunks.push(chunk)
    })
    req.on('end', () => resolve(Buffer.concat(chunks)))
    req.on('error', reject)
  })
}

/**
 * 解析一个 fileKey（不含扩展名）。fileKey 已包含随机串，扩展名只用于还原原文件。
 * @param {string} raw
 * @returns {{ stem: string, ext: string, name: string } | undefined}
 */
export function parseFileKey(raw) {
  if (typeof raw !== 'string' || !raw) return undefined
  const dot = raw.lastIndexOf('.')
  if (dot < 0) return undefined
  const stem = raw.slice(0, dot)
  const ext = raw.slice(dot + 1)
  if (!FILE_KEY_STEM.test(stem) || !FILE_KEY_EXT.test(ext)) return undefined
  return { stem, ext, name: raw }
}

/**
 * @param {string} candidate
 * @param {string} root
 */
function isInside(candidate, root) {
  const relative = path.relative(root, candidate)
  return relative !== '' && !relative.startsWith('..') && !path.isAbsolute(relative)
}

/**
 * 同源判断：用于拦截「任意网页通过 hidden form POST 触发保存」的 CSRF 面。
 * DS 服务端拉取 / 回调不带 Origin、Referer，因此天然通过。
 * @param {IncomingMessage} req
 * @param {ResolvedConfig} config
 * @param {boolean} allowSameOrigin
 */
export function isTrustedRequest(req, config, allowSameOrigin) {
  const origin = firstHeader(req.headers.origin)
  if (origin) {
    if (config.allowedOrigins.includes(origin.replace(/\/+$/, ''))) return true
    if (allowSameOrigin) {
      const host = firstHeader(req.headers.host)
      if (host) {
        try {
          const parsed = new URL(origin)
          if (parsed.host === host) return true
        } catch {
          return false
        }
      }
    }
    return false
  }
  const referer = firstHeader(req.headers.referer)
  if (referer) {
    try {
      const parsed = new URL(referer)
      const host = firstHeader(req.headers.host)
      if (host && parsed.host === host) return true
      return config.allowedOrigins.includes(parsed.origin)
    } catch {
      return false
    }
  }
  // 无 Origin / Referer：DS 服务端到服务端请求、本地 curl 等。
  return true
}

// ---------------------------------------------------------------------------
// 服务
// ---------------------------------------------------------------------------

export class OnlyOfficeBridge extends Service {
  static inject = ['webServer', 'tools']

  /**
   * @param {Context} ctx
   * @param {ResolvedConfig} config
   */
  constructor(ctx, config) {
    super(ctx, 'onlyoffice')
    this.config = config
    this.storageDir = config.storageDir
    /** 已注册工具的注销器（重复注册时先清掉旧的）。 */
    this.toolDisposers = []
    /** 定时清理句柄。 */
    this.cleanTimer = null
    /** 最近一次回调（同时持久化到 storage/last-callback.json）。 */
    this.lastCallback = this.loadLastCallback()

    ctx.effect(() => this.mkdirStorage())

    // 定时清理：用原生 timer 而不是 ctx.interval —— 后者在 DSH 里由可选的 timer 服务提供，
    // 引用它会把本插件绑死在那个服务上。卸载时由下面的 effect 与 dispose() 清掉。
    if (config.cacheTtlHour > 0 || config.cacheCleanIntervalMin > 0) {
      const intervalMinutes = config.cacheCleanIntervalMin > 0 ? config.cacheCleanIntervalMin : 60
      const timer = setInterval(() => {
        this.cleanExpired().then(
          (removed) => {
            if (removed.length) this.ctx.logger.info('[onlyoffice] 自动清理 %d 个过期缓存', removed.length)
          },
          (error) => this.ctx.logger.warn('[onlyoffice] 自动清理失败', error),
        )
      }, Math.max(1, intervalMinutes) * 60 * 1000)
      timer.unref?.()
      this.cleanTimer = timer
      ctx.effect(() => () => this.stop())
    }

    ctx.effect(() => this.registerRoutes())
    ctx.effect(() => [() => this.unregisterTools()])
    this.registerTools()
  }

  /** 停止定时清理并注销工具（服务被替换或进程退出时调用）。 */
  stop() {
    if (this.cleanTimer) {
      clearInterval(this.cleanTimer)
      this.cleanTimer = null
    }
    this.unregisterTools()
  }

  /** @deprecated 保留别名，等价于 stop()。 */
  dispose() {
    this.stop()
  }

  registerRoutes() {
    const disposers = [
      this.ctx.webServer.register({
        kind: 'prefix',
        path: STATIC_PREFIX,
        handler: (req, res) => this.handleStatic(req, res),
      }),
      this.ctx.webServer.register({
        kind: 'exact',
        path: CALLBACK_PATH,
        handler: (req, res) => this.handleCallback(req, res),
      }),
      this.ctx.webServer.register({
        kind: 'exact',
        path: UPLOAD_PATH,
        handler: (req, res) => this.handleUpload(req, res),
      }),
      this.ctx.webServer.register({
        kind: 'exact',
        path: HEALTH_PATH,
        handler: (req, res) => this.handleHealth(req, res),
      }),
    ]
    this.ctx.logger.info('[onlyoffice] 已挂载路由 %s', ROUTE_PREFIX)
    return () => {
      for (const dispose of disposers) dispose()
    }
  }

  /** DSH 对外地址：配置优先，否则用 webServer 探测；探测值只对运行在宿主机上的 DS 有效。 */
  publicBaseUrl() {
    if (this.config.publicUrl) return this.config.publicUrl
    try {
      return this.ctx.webServer.getOrigin()
    } catch {
      return `http://127.0.0.1:${process.env.PORT ?? 11006}`
    }
  }

  /** 浏览器打开编辑器使用的 DS 地址。 */
  editorBaseUrl() {
    return this.config.editorDsUrl || this.config.dsUrl
  }

  staticUrl(fileKey) {
    return `${this.publicBaseUrl()}${STATIC_PREFIX}/${encodeURIComponent(fileKey)}`
  }

  callbackUrl() {
    return `${this.publicBaseUrl()}${CALLBACK_PATH}`
  }

  // -- 存储 ---------------------------------------------------------------

  mkdirStorage() {
    fs.mkdirSync(this.storageDir, { recursive: true })
  }

  storagePath(fileKey) {
    const parsed = parseFileKey(fileKey)
    if (!parsed) return undefined
    const candidate = path.join(this.storageDir, parsed.name)
    return isInside(candidate, this.storageDir) ? candidate : undefined
  }

  /** @param {string} ext */
  assertSupported(ext) {
    const normalized = ext.toLowerCase()
    if (!OFFICE_EXTENSIONS.includes(normalized)) {
      throw new Error(`不支持的格式 .${normalized}；支持：${OFFICE_EXTENSIONS.join(' / ')}`)
    }
    return normalized
  }

  /**
   * 写入一份新缓存，返回 fileKey。
   * @param {string} originalName
   * @param {Buffer} data
   */
  writeCache(originalName, data) {
    this.mkdirStorage()
    const ext = this.assertSupported(path.extname(originalName).slice(1) || 'txt')
    const stem = `${crypto.randomBytes(16).toString('hex')}-${crypto.randomBytes(8).toString('hex')}`
    const fileKey = `${stem}.${ext}`
    const target = /** @type {string} */ (this.storagePath(fileKey))
    fs.writeFileSync(target, data)
    const metaPath = path.join(this.storageDir, `${stem}.meta.json`)
    const meta = {
      fileKey,
      fileName: path.basename(originalName) || `document.${ext}`,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      lastStatus: /** @type {number | null} */ (null),
      editorVersion: /** @type {number | null} */ (null),
    }
    fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2), 'utf8')
    return { fileKey, meta, target }
  }

  readMeta(fileKey) {
    const parsed = parseFileKey(fileKey)
    if (!parsed) return undefined
    const metaPath = path.join(this.storageDir, `${parsed.stem}.meta.json`)
    try {
      const meta = JSON.parse(fs.readFileSync(metaPath, 'utf8'))
      return { metaPath, meta: asRecord(meta) }
    } catch {
      return undefined
    }
  }

  loadLastCallback() {
    try {
      const raw = fs.readFileSync(path.join(this.storageDir, 'last-callback.json'), 'utf8')
      return JSON.parse(raw)
    } catch {
      return null
    }
  }

  // -- DS 交互 ------------------------------------------------------------

  /**
   * 调用 DS CommandService.ashx（forcesave 等）。
   * @param {Record<string, unknown>} payload
   */
  async commandService(payload) {
    const { headers } = authHeaders(payload, this.config)
    const response = await fetch(`${this.config.dsUrl}/coauthoring/CommandService.ashx`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...headers },
      body: JSON.stringify(payload),
    })
    const text = await response.text()
    let json
    try {
      json = JSON.parse(text)
    } catch {
      json = { raw: text }
    }
    return { ok: response.ok, status: response.status, body: json }
  }

  /**
   * 下载 DS 上的保存结果（单次，带 token）。
   * @param {string} url
   */
  async downloadSavedFile(url) {
    const { headers } = authHeaders({ url }, this.config)
    const response = await fetch(url, { headers })
    if (!response.ok) throw new Error(`下载保存内容失败 HTTP ${response.status}`)
    const buffer = Buffer.from(await response.arrayBuffer())
    const limit = this.config.maxFileMb * 1024 * 1024
    if (buffer.length > limit) throw new Error(`保存内容超过 ${this.config.maxFileMb} MiB 上限`)
    return buffer
  }

  /**
   * 下载任意 http(s) 文件（转换产物、模板源文件等）。
   * @param {string} url
   * @param {{ jwt?: boolean }} [options]
   */
  async downloadUrl(url, options = {}) {
    const headers = options.jwt ? authHeaders({ url }, this.config).headers : {}
    const response = await fetch(url, { headers })
    if (!response.ok) throw new Error(`下载失败 HTTP ${response.status}：${url}`)
    const buffer = Buffer.from(await response.arrayBuffer())
    const limit = this.config.maxFileMb * 1024 * 1024
    if (buffer.length > limit) throw new Error(`文件超过 ${this.config.maxFileMb} MiB 上限`)
    return buffer
  }

  /**
   * ONLYOFFICE ConvertService：把 DS 编辑缓存里的文档转成另一种格式。
   *
   * 转换接口按 key 从 DS 自己的缓存里取文档，所以 key 必须是 DS 当前持有编辑会话的那个；
   * 先对同一 key 发一次 forcesave 把最新内容刷进缓存，再请求转换。
   * @param {string} fileKey
   * @param {{ outputType: string, title?: string, region?: string }} options
   */
  async convert(fileKey, options) {
    try {
      await this.commandService({ c: 'forcesave', key: fileKey })
    } catch {
      // 该 key 还没有编辑会话时 forcesave 会失败，此时 DS 直接转换源文件即可
    }
    /** @type {Record<string, any>} */
    const payload = {
      async: false,
      filetype: /** @type {{ ext: string }} */ (parseFileKey(fileKey)).ext,
      key: fileKey,
      outputtype: options.outputType,
      title: options.title,
      url: this.staticUrl(fileKey),
    }
    if (options.region) payload.region = options.region
    const { headers } = authHeaders(payload, this.config)
    const response = await fetch(`${this.config.dsUrl}/ConvertService.ashx`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json', ...headers },
      body: JSON.stringify(payload),
    })
    const text = await response.text()
    let json
    try {
      json = JSON.parse(text)
    } catch {
      json = { raw: text }
    }
    if (!response.ok) throw new Error(`ConvertService HTTP ${response.status}：${String(text).slice(0, 300)}`)
    const errorCode = Number(json?.error ?? 0)
    if (errorCode !== 0) throw new Error(`ConvertService 返回 error=${errorCode}：${String(text).slice(0, 300)}`)
    return { url: typeof json?.fileUrl === 'string' ? json.fileUrl : undefined, percent: json?.percent, raw: json }
  }

  /**
   * 通用格式导出：转换后直接把产物下载到本地。
   * @param {{ fileKey: string, format: string, outputPath?: string }} args
   */
  async exportAs(args) {
    const fileKey = typeof args.fileKey === 'string' ? args.fileKey : ''
    const format = String(args.format ?? '').toLowerCase()
    const parsed = parseFileKey(fileKey)
    if (!parsed) throw new Error(`fileKey 非法：${fileKey}`)
    if (!CONVERT_FORMATS.includes(format)) {
      throw new Error(`不支持的导出格式 .${format}；支持：${CONVERT_FORMATS.join(' / ')}`)
    }
    const source = this.storagePath(fileKey)
    if (!source || !fs.existsSync(source)) throw new Error(`fileKey 不存在：${fileKey}`)
    const record = this.readMeta(fileKey)
    const meta = record?.meta ?? {}
    const originalName = String(meta.fileName ?? `document.${parsed.ext}`)
    const baseName = path.basename(originalName, path.extname(originalName)) || parsed.stem
    const target = args.outputPath
      ? path.resolve(args.outputPath)
      : path.join(path.dirname(source), `${baseName}.${format}`)

    const converted = await this.convert(fileKey, { outputType: format, title: originalName })
    if (!converted.url) throw new Error('ConvertService 没有返回 fileUrl')
    const buffer = await this.downloadUrl(converted.url, { jwt: true })
    fs.mkdirSync(path.dirname(target), { recursive: true })
    fs.writeFileSync(target, buffer)
    return {
      ok: true,
      fileKey,
      format,
      outputPath: target,
      size: buffer.length,
      convertUrl: converted.url,
      hint: '转换产物已下载到 outputPath',
    }
  }

  /**
   * 编辑前给文档取一个全新的 key。
   *
   * DS 的 key 决定编辑会话缓存：同一个 key 再打开会复用旧缓存（旧版本内容、过期链接），
   * 所以每次 open 都复制到一个新 key。文档级元数据（原始路径等）放在
   * meta.json 的 lineage 里，保存时据此写回原文件。
   * @param {string} fileKey
   */
  rekeyForEdit(fileKey) {
    const parsed = parseFileKey(fileKey)
    const source = this.storagePath(fileKey)
    if (!parsed || !source || !fs.existsSync(source)) throw new Error(`fileKey 不存在：${fileKey}`)
    const record = this.readMeta(fileKey)
    const meta = record?.meta ?? {}
    const stem = `${crypto.randomBytes(16).toString('hex')}-${crypto.randomBytes(8).toString('hex')}`
    const nextKey = `${stem}.${parsed.ext}`
    const nextPath = /** @type {string} */ (this.storagePath(nextKey))
      fs.copyFileSync(source, nextPath)
      const lineage = asRecord(meta.lineage)
      const rootKey = typeof lineage.rootKey === 'string' ? lineage.rootKey : parsed.stem
      const nextMeta = {
        ...meta,
        fileKey: nextKey,
        lineage: {
          originalPath: lineage.originalPath ?? null,
          rootKey,
          parentKey: this.generationKey(fileKey),
          depth: (Number(lineage.depth) || 0) + 1,
        },
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    }
    fs.writeFileSync(path.join(this.storageDir, `${stem}.meta.json`), JSON.stringify(nextMeta, null, 2), 'utf8')
    return { fileKey: nextKey, meta: nextMeta, path: nextPath }
  }

  /**
   * 当前这一代编辑会话的标识（key 由 <stem>.<ext> 组成，stem 才是文档 key 的一部分）。
   * @param {string} fileKey
   */
  generationKey(fileKey) {
    return /** @type {{ stem: string, ext: string, name: string }} */ (parseFileKey(fileKey)).stem
  }

  /**
   * 清理同一份原文件的历史编辑世代（保留 keepKey 当前这一代）。
   * 同一原文件的世代有相同的 lineage.rootKey。
   * @param {string} keepKey
   */
  retireFamily(keepKey) {
    const keep = parseFileKey(keepKey)
    const record = this.readMeta(keepKey)
    const rootKey = /** @type {string | undefined} */ (asRecord(asRecord(record?.meta).lineage).rootKey) ?? keep?.stem
    if (!rootKey) return
    let files = []
    try {
      files = fs.readdirSync(this.storageDir)
    } catch {
      return
    }
    for (const file of files) {
      const parsed = parseFileKey(file)
      if (!parsed || parsed.stem === keep?.stem) continue
      const other = this.readMeta(file)
      const otherRoot = asRecord(asRecord(other?.meta).lineage).rootKey
      if (otherRoot !== rootKey) continue
      fs.rmSync(path.join(this.storageDir, file), { force: true })
      fs.rmSync(path.join(this.storageDir, `${parsed.stem}.meta.json`), { force: true })
    }
  }

  /**
   * 打开（或创建）一个文档并返回编辑链接。
   * @param {{ filePath?: string, fileKey?: string, fileName?: string, contentBase64?: string, mode?: 'edit' | 'view' }} args
   */
  async open(args) {
    let fileKey
    let fileName = args.fileName
    if (typeof args.contentBase64 === 'string' && args.contentBase64) {
      const data = Buffer.from(args.contentBase64, 'base64')
      if (!data.length) throw new Error('contentBase64 解码为空')
      const limit = this.config.maxFileMb * 1024 * 1024
      if (data.length > limit) throw new Error(`文件超过 ${this.config.maxFileMb} MiB 上限`)
      const name = fileName || 'document.docx'
      const written = this.writeCache(name, data)
      fileKey = written.fileKey
      fileName = name
      const record = this.readMeta(fileKey)
      const rootKey = String(written.meta.fileKey).split('.')[0]
      const meta = { ...(record?.meta ?? {}), lineage: { originalPath: null, rootKey, parentKey: null, depth: 0 } }
      if (record) fs.writeFileSync(record.metaPath, JSON.stringify(meta, null, 2), 'utf8')
    } else if (typeof args.filePath === 'string' && args.filePath) {
      const absolute = path.resolve(args.filePath)
      const stat = fs.existsSync(absolute) ? fs.statSync(absolute) : undefined
      if (!stat) throw new Error(`文件不存在：${absolute}`)
      if (stat.isDirectory()) throw new Error(`这是一个目录：${absolute}`)
      const limit = this.config.maxFileMb * 1024 * 1024
      if (stat.size > limit) throw new Error(`文件超过 ${this.config.maxFileMb} MiB 上限`)
      const buffer = fs.readFileSync(absolute)
      const written = this.writeCache(path.basename(absolute), buffer)
      fileKey = written.fileKey
      fileName = path.basename(absolute)
      const record = this.readMeta(fileKey)
      const rootKey = String(written.meta.fileKey).split('.')[0]
      const meta = { ...(record?.meta ?? {}), lineage: { originalPath: absolute, rootKey, parentKey: null, depth: 0 } }
      if (record) fs.writeFileSync(record.metaPath, JSON.stringify(meta, null, 2), 'utf8')
    } else if (typeof args.fileKey === 'string' && args.fileKey) {
      // 直接续开一个已有缓存（浏览器端上传后回传 fileKey 的场景）。
      const existing = this.storagePath(args.fileKey)
      if (!existing || !fs.existsSync(existing)) throw new Error(`fileKey 不存在：${args.fileKey}`)
      fileKey = args.fileKey
      fileName = fileName || String(this.readMeta(args.fileKey)?.meta.fileName ?? args.fileKey)
    } else {
      throw new Error('需要 filePath、contentBase64 或 fileKey 之一')
    }

    const fresh = this.rekeyForEdit(/** @type {string} */ (fileKey))
    const parsed = /** @type {{ ext: string, name: string }} */ (parseFileKey(fresh.fileKey))
    const fileType = parsed.ext
    const mode = args.mode === 'view' ? 'view' : 'edit'
    // 同一份原文件只保留最新一代编辑缓存：上一代已不可能再收到保存回调（key 已更换）。
    this.retireFamily(fresh.fileKey)

    const document = {
      fileType,
      key: fresh.fileKey,
      title: fileName || `document.${fileType}`,
      url: this.staticUrl(fresh.fileKey),
    }
    /** @type {Record<string, any>} */
    const permissions = {
      edit: mode === 'edit',
      download: true,
      print: true,
      review: mode === 'edit',
      comment: mode === 'edit',
    }
    const editorConfig = {
      mode,
      lang: 'zh-CN',
      callbackUrl: this.callbackUrl(),
      user: { id: 'dsh-agent', name: 'DSH Agent' },
      customization: {
        forcesave: true,
        // 自动保存（键值保存）交给保存回调 status=6 落盘，避免只有手动保存才回写。
        autosave: true,
        feedback: false,
        compactHeader: true,
        help: false,
      },
    }
    const payload = { document, documentType: documentTypeOf(fileType), editorConfig: { ...editorConfig, permissions } }
    const token = signJwt(payload, this.config.jwtSecret)
    const query = new URLSearchParams({ token })
    if (this.config.jwtHeader.trim().toLowerCase() !== 'authorization') {
      query.set('customHeader', this.config.jwtHeader)
    }
    const editorUrl = `${this.editorBaseUrl()}/web-apps/apps/api/documents/api.js?${query.toString()}`
    return {
      fileKey: fresh.fileKey,
      editorUrl,
      apiJs: true,
      downloadUrl: this.staticUrl(fresh.fileKey),
      callbackUrl: this.callbackUrl(),
      docType: documentTypeOf(fileType),
      mode,
      token,
      config: payload,
    }
  }

  /**
   * 导出 PDF。
   *
   * ONLYOFFICE 的转换服务是按 key 从 DS 自己的编辑缓存里取文档，再按 outputtype 转换。
   * 所以流程是：先把当前缓存内容挂到一个新 key 上、作为该 key 的编辑会话交给 DS 打开
   * （这里把保存格式设为 pdf），然后对同一 key 发 forcesave；DS 的保存回调带回来的就是
   * PDF，由 handleCallback 写到 outputPath。
   * 只有 docx / xlsx / pptx 支持直接转换。
   * @param {{ fileKey: string, outputPath?: string }} args
   */
  async exportPdf(args) {
    const fileKey = typeof args.fileKey === 'string' ? args.fileKey : ''
    const parsed = parseFileKey(fileKey)
    if (!parsed) throw new Error(`fileKey 非法：${fileKey}`)
    if (!CONVERTIBLE_EXTENSIONS.includes(parsed.ext)) {
      throw new Error(`.${parsed.ext} 不支持导出 PDF，仅支持 ${CONVERTIBLE_EXTENSIONS.join(' / ')}`)
    }
    const source = this.storagePath(fileKey)
    if (!source || !fs.existsSync(source)) throw new Error(`fileKey 不存在：${fileKey}`)

    // 同一份原文件只保留最新一代（与 open 一致），避免历史缓存堆积。
    const fresh = this.rekeyForEdit(fileKey)
    this.retireFamily(fresh.fileKey)
    const outputPath = args.outputPath
      ? path.resolve(args.outputPath)
      : path.join(path.dirname(source), `${parsed.stem}.pdf`)
    const record = this.readMeta(fresh.fileKey)
    const originalName = String(record?.meta.fileName ?? `document.${parsed.ext}`)
    const pdfName = `${path.basename(originalName, path.extname(originalName)) || path.basename(outputPath, '.pdf')}.pdf`
    const meta = { ...(record?.meta ?? {}), exportPdfTo: outputPath }
    fs.writeFileSync(path.join(this.storageDir, `${this.generationKey(fresh.fileKey)}.meta.json`), JSON.stringify(meta, null, 2), 'utf8')

    const document = {
      fileType: parsed.ext,
      key: fresh.fileKey,
      title: originalName,
      url: this.staticUrl(fresh.fileKey),
      outputtype: 'pdf',
    }
    const editorConfig = {
      mode: 'edit',
      callbackUrl: this.callbackUrl(),
      user: { id: 'dsh-agent', name: 'DSH Agent' },
      customization: { forcesave: true, autosave: false, help: false },
    }
    const payload = { document, documentType: documentTypeOf(parsed.ext), editorConfig }
    const token = signJwt(payload, this.config.jwtSecret)
    // 打开参数里带上 DS 的 filename 覆盖，让 DS 把转换结果命名为 xxx.pdf。
    const query = new URLSearchParams({ token, filename: pdfName })
    const editorUrl = `${this.editorBaseUrl()}/web-apps/apps/api/documents/api.js?${query.toString()}`

    let command
    let commandError = null
    try {
      command = await this.commandService({ c: 'forcesave', key: fresh.fileKey, userdata: 'dsh-export-pdf' })
    } catch (error) {
      commandError = error instanceof Error ? error.message : String(error)
      command = null
    }
    return {
      ok: true,
      fileKey: fresh.fileKey,
      outputPath,
      editorUrl,
      async: true,
      command,
      commandError,
      convertHint:
        'PDF 由 DS 异步生成：打开 editorUrl（或让用户打开）并保存一次后，保存回调会把 PDF 写到 outputPath。' +
        (commandError ? ' 本次 forcesave 请求失败，可能是该 key 尚未在 DS 中建立编辑会话。' : ''),
    }
  }

  /**
   * 清理过期缓存（按 TTL）。PDF 导出产物、last-callback.json 不受影响。
   * @param {number} [ttlHours]
   */
  async cleanExpired(ttlHours = this.config.cacheTtlHour) {
    if (!(ttlHours > 0)) return []
    const removed = []
    const cutoff = Date.now() - ttlHours * 60 * 60 * 1000
    const files = await fs.promises.readdir(this.storageDir)
    /** @type {Map<string, string[]>} */
    const byStem = new Map()
    for (const file of files) {
      const dot = file.lastIndexOf('.')
      if (dot < 0) continue
      const stem = file.slice(0, dot)
      const list = byStem.get(stem) ?? []
      list.push(file)
      byStem.set(stem, list)
    }
    for (const [stem, group] of byStem) {
      if (!FILE_KEY_STEM.test(stem)) continue
      const primary = group.filter((f) => !f.endsWith('.pdf'))
      let expired = true
      for (const file of primary) {
        try {
          const stat = await fs.promises.stat(path.join(this.storageDir, file))
          if (stat.mtimeMs >= cutoff) {
            expired = false
            break
          }
        } catch {
          // 读取失败的文件交给下一轮或人工处理
          expired = false
        }
      }
      if (!expired) continue
      for (const file of group) {
        try {
          await fs.promises.unlink(path.join(this.storageDir, file))
          removed.push(file)
        } catch (error) {
          this.ctx.logger.warn('[onlyoffice] 清理文件失败 %s: %s', file, error instanceof Error ? error.message : String(error))
        }
      }
    }
    return removed
  }

  /**
   * 列出缓存文档。
   * @returns {Promise<Array<Record<string, unknown>>>}
   */
  async list() {
    /** @type {Array<Record<string, unknown>>} */
    const out = []
    const files = await fs.promises.readdir(this.storageDir)
    for (const file of files) {
      const dot = file.lastIndexOf('.')
      if (dot < 0) continue
      const stem = file.slice(0, dot)
      const ext = file.slice(dot + 1)
      if (!FILE_KEY_STEM.test(stem) || !OFFICE_EXTENSIONS.includes(ext)) continue
      const record = this.readMeta(file)
      const meta = record?.meta ?? {}
      // 每次编辑都会生成新的 key（见 rekeyForEdit）；这里只列出当前这一代，避免重复项刷屏。
      if (meta.retired === true) continue
      const lineage = asRecord(meta.lineage)
      const currentStem = typeof lineage.rootKey === 'string' ? lineage.rootKey : stem
      if (currentStem !== stem && fs.existsSync(path.join(this.storageDir, `${currentStem}.meta.json`))) continue
      let stat
      try {
        stat = await fs.promises.stat(path.join(this.storageDir, file))
      } catch {
        continue
      }
      out.push({
        fileKey: file,
        fileName: meta.fileName ?? null,
        originalPath: asRecord(meta.lineage).originalPath ?? null,
        size: stat.size,
        updatedAt: new Date(stat.mtimeMs).toISOString(),
        lastStatus: meta.lastStatus ?? null,
        edited: meta.lastStatus === 2 || meta.lastStatus === 6,
      })
    }
    return out.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)))
  }

  /**
   * 删除一个缓存（含 meta 与旁路 PDF）。
   * @param {string} fileKey
   */
  purge(fileKey) {
    const parsed = parseFileKey(fileKey)
    const target = this.storagePath(fileKey)
    if (!parsed || !target) throw new Error(`fileKey 非法：${fileKey}`)
    const removed = []
    for (const candidate of [
      target,
      path.join(this.storageDir, `${parsed.stem}.meta.json`),
      path.join(this.storageDir, `${parsed.stem}.pdf`),
    ]) {
      if (fs.existsSync(candidate)) {
        fs.rmSync(candidate, { force: true })
        removed.push(path.basename(candidate))
      }
    }
    return { ok: removed.length > 0, removed, fileKey }
  }

  // -- HTTP ---------------------------------------------------------------

  /**
   * @param {IncomingMessage} req
   * @param {ServerResponse} res
   */
  async handleStatic(req, res) {
    if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, { error: 'method not allowed' })
    const url = new URL(req.url ?? '/', 'http://localhost')
    if (!url.pathname.startsWith(`${STATIC_PREFIX}/`)) return send(res, 404, { error: 'not found' })
    const rawKey = decodeURIComponent(url.pathname.slice(STATIC_PREFIX.length + 1))
    const parsed = parseFileKey(rawKey)
    const filePath = parsed ? this.storagePath(parsed.name) : undefined
    if (!filePath || !fs.existsSync(filePath)) return send(res, 404, { error: 'file not found' })
    if (!isTrustedRequest(req, this.config, true)) return send(res, 403, { error: 'untrusted origin' })
    const stat = await fs.promises.stat(filePath)
    res.writeHead(200, {
      'Content-Type': contentTypeOf(parsed.ext),
      'Content-Length': String(stat.size),
      'Cache-Control': 'no-store',
    })
    if (req.method === 'HEAD') {
      res.end()
      return
    }
    await pipeline(Readable.from(fs.createReadStream(filePath)), res)
  }

  /**
   * @param {IncomingMessage} req
   * @param {ServerResponse} res
   */
  async handleCallback(req, res) {
    if (req.method !== 'POST') return send(res, 405, { error: 'method not allowed' })
    const limit = this.config.maxFileMb * 1024 * 1024
    let raw
    try {
      raw = await readBody(req, Math.min(limit, 16 * 1024 * 1024))
    } catch (error) {
      return send(res, 413, { error: error instanceof Error ? error.message : String(error) })
    }
    /** @type {Record<string, any>} */
    let body = {}
    try {
      body = raw.length ? JSON.parse(raw.toString('utf8')) : {}
    } catch {
      return send(res, 400, { error: 'invalid json' })
    }
    const claims = extractRequestToken(req, body, this.config)
    if (!claims) {
      this.ctx.logger.warn('[onlyoffice] 回调 JWT 校验失败（检查 ONLYOFFICE_SECRET / secret 配置是否一致）')
      return send(res, 403, { error: 1, message: 'invalid token' }, 'application/json')
    }
    const data = asRecord(claims.payload ?? claims)
    const status = Number(data.status ?? body.status ?? 0)
    const key = String(data.key ?? body.key ?? '')
    this.ctx.logger.info('[onlyoffice] 回调 key=%s status=%s', key, status)
    const callback = {
      at: new Date().toISOString(),
      key,
      status,
      statusText: CALLBACK_STATUS[status] ?? `未知状态 ${status}`,
      userdata: data.userdata ?? null,
      url: data.url ?? null,
      error: null,
    }
    try {
      if (status === 2 || status === 6) {
        const saveUrl = data.url ?? body.url
        if (!saveUrl) throw new Error('回调缺少 content.url')
        const buffer = await this.downloadSavedFile(String(saveUrl))
        await this.applySavedContent(key, buffer)
        callback.error = null
      }
    } catch (error) {
      callback.error = error instanceof Error ? error.message : String(error)
      this.ctx.logger.error('[onlyoffice] 保存回调失败 key=%s: %s', key, callback.error)
    }
    this.lastCallback = callback
    try {
      fs.writeFileSync(path.join(this.storageDir, 'last-callback.json'), JSON.stringify(callback, null, 2), 'utf8')
    } catch {
      // 持久化失败不影响回调响应
    }
    // ONLYOFFICE 要求回调响应体为 {"error":0}
    return send(res, 200, { error: callback.error && status !== 4 ? 1 : 0 }, 'application/json')
  }

  /**
   * 把保存结果落回缓存文件与原文件。
   * @param {string} fileKey
   * @param {Buffer} buffer
   */
  async applySavedContent(fileKey, buffer) {
    const parsed = parseFileKey(fileKey)
    const target = this.storagePath(fileKey)
    if (!parsed || !target) throw new Error(`回调 key 非法：${fileKey}`)
    fs.writeFileSync(target, buffer)
    const record = this.readMeta(fileKey)
    const meta = record?.meta ?? {}
    const lineage = asRecord(meta.lineage)
    const isPdf = buffer.subarray(0, 5).toString('latin1') === '%PDF-'
    if (meta.exportPdfTo && isPdf) {
      const out = String(meta.exportPdfTo)
      fs.mkdirSync(path.dirname(out), { recursive: true })
      fs.writeFileSync(out, buffer)
      this.ctx.logger.info('[onlyoffice] PDF 已导出：%s', out)
    }
    const originalPath = typeof lineage.originalPath === 'string' ? lineage.originalPath : undefined
    if (originalPath && !isPdf) {
      try {
        fs.writeFileSync(originalPath, buffer)
        this.ctx.logger.info('[onlyoffice] 已回写原文件：%s', originalPath)
      } catch (error) {
        this.ctx.logger.warn(
          '[onlyoffice] 回写原文件失败（%s）：%s；内容已保存在 %s',
          originalPath,
          error instanceof Error ? error.message : String(error),
          target,
        )
      }
    }
    if (record) {
      const next = { ...meta, updatedAt: new Date().toISOString(), lastStatus: null }
      fs.writeFileSync(record.metaPath, JSON.stringify(next, null, 2), 'utf8')
    }
    return { originalPath, isPdf }
  }

  /**
   * @param {IncomingMessage} req
   * @param {ServerResponse} res
   */
  async handleUpload(req, res) {
    if (req.method !== 'POST' && req.method !== 'PUT') return send(res, 405, { error: 'method not allowed' })
    if (!isTrustedRequest(req, this.config, true)) return send(res, 403, { error: 'untrusted origin' })
    const url = new URL(req.url ?? '/', 'http://localhost')
    const name = url.searchParams.get('name') ?? 'upload.docx'
    const limit = this.config.maxFileMb * 1024 * 1024
    let raw
    try {
      raw = await readBody(req, limit)
    } catch (error) {
      return send(res, 413, { error: error instanceof Error ? error.message : String(error) })
    }
    if (!raw.length) return send(res, 400, { error: 'empty body' })
    try {
      const written = this.writeCache(path.basename(name), raw)
      this.ctx.logger.info('[onlyoffice] 已接收上传 %s → %s', written.meta.fileName, written.fileKey)
      return send(res, 200, { fileKey: written.fileKey, fileName: written.meta.fileName })
    } catch (error) {
      return send(res, 400, { error: error instanceof Error ? error.message : String(error) })
    }
  }

  /**
   * @param {IncomingMessage} req
   * @param {ServerResponse} res
   */
  async handleHealth(req, res) {
    if (req.method !== 'GET') return send(res, 405, { error: 'method not allowed' })
    let storageWritable = true
    try {
      fs.accessSync(this.storageDir, fs.constants.W_OK)
    } catch {
      storageWritable = false
    }
    return send(res, 200, {
      ok: true,
      dsUrl: this.config.dsUrl,
      publicUrl: this.publicBaseUrl(),
      storageDir: this.storageDir,
      storageWritable,
      cacheTtlHour: this.config.cacheTtlHour,
      lastCallback: this.lastCallback,
    })
  }

  // -- Agent 工具 ---------------------------------------------------------

  registerTools() {
    // 重复调用时先注销旧注册，避免同一服务出现两份工具。
    this.unregisterTools()
    const toolkit = createToolkit(this, { logger: this.ctx.logger })
    const disposers = []
    this.toolDisposers.push(disposers)

    for (const definition of TOOL_DEFINITIONS) {
      const handler = toolkit.handlers[definition.key]
      if (typeof handler !== 'function') continue
      disposers.push(
        this.ctx.tools.register({
          name: `onlyoffice_${definition.name}`,
          description: definition.description,
          parameters: definition.parameters,
          output: { schema: { type: 'object' }, render: renderJson },
          execute: async (args) => handler(/** @type {any} */ (args ?? {})),
        }),
      )
    }
    this.ctx.logger.info('[onlyoffice] 已注册 %d 个 Agent 工具', disposers.length)
  }

  /** 幂等注销全部 Agent 工具（服务卸载或 ctx 销毁时调用）。 */
  unregisterTools() {
    const batches = this.toolDisposers ?? []
    this.toolDisposers = []
    for (const batch of batches) {
      if (typeof batch === 'function') {
        batch()
        continue
      }
      while (batch.length) {
        const dispose = batch.pop()
        if (typeof dispose === 'function') dispose()
      }
    }
  }
}

// ---------------------------------------------------------------------------
// 纯函数辅助
// ---------------------------------------------------------------------------

/** @param {string} ext */
function documentTypeOf(ext) {
  if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) return 'cell'
  if (['pptx', 'ppt', 'odp'].includes(ext)) return 'slide'
  return 'word'
}

/** @param {string} ext */
function contentTypeOf(ext) {
  switch (ext) {
    case 'docx':
      return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    case 'doc':
      return 'application/msword'
    case 'xlsx':
      return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    case 'xls':
      return 'application/vnd.ms-excel'
    case 'pptx':
      return 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    case 'ppt':
      return 'application/vnd.ms-powerpoint'
    case 'odt':
      return 'application/vnd.oasis.opendocument.text'
    case 'ods':
      return 'application/vnd.oasis.opendocument.spreadsheet'
    case 'odp':
      return 'application/vnd.oasis.opendocument.presentation'
    case 'csv':
      return 'text/csv; charset=utf-8'
    case 'rtf':
      return 'application/rtf'
    case 'pdf':
      return 'application/pdf'
    default:
      return 'text/plain; charset=utf-8'
  }
}

/**
 * @param {ServerResponse} res
 * @param {number} status
 * @param {unknown} body
 * @param {string} [contentType]
 */
function send(res, status, body, contentType = 'application/json; charset=utf-8') {
  const payload = Buffer.from(JSON.stringify(body), 'utf8')
  res.writeHead(status, { 'Content-Type': contentType, 'Content-Length': String(payload.length) })
  res.end(payload)
}

/**
 * @param {unknown} _args
 * @param {unknown} value
 * @returns {Array<{ type: 'text', text: string }>}
 */
function renderJson(_args, value) {
  return [{ type: 'text', text: JSON.stringify(value, null, 2) }]
}

// ---------------------------------------------------------------------------
// 插件入口
// ---------------------------------------------------------------------------
//
// 这里刻意不导出 schema 形式的 `Config`：包内只依赖 DSH 必然提供的
// @deepseek-ai/cordis，schema 库（schemastery 等）只对 dsh-tools 是必需依赖，
// 依赖它会让插件在缺少该包的 profile 里导入失败。配置校验由 resolveConfig
// 自己完成（非法值回退默认值并在日志里提示），对外契约一致。

/**
 * @param {Context} ctx
 * @param {Partial<ResolvedConfig>} [config]
 */
export function apply(ctx, config = {}) {
  const resolved = resolveConfig({ dotenv: loadDotEnv() }, /** @type {Record<string, unknown>} */ (config))
  if (resolved.jwtSecret === defaults().jwtSecret) {
    ctx.logger.warn('[onlyoffice] 仍在使用默认 JWT 密钥，生产环境请设置 ONLYOFFICE_SECRET / ONLYOFFICE_SECRETS')
  }
  if (!resolved.publicUrl) {
    ctx.logger.info('[onlyoffice] 未配置 DSH_PUBLIC_URL，使用自动探测地址（DS 在 docker 内时需显式配置）')
  }
  ctx.plugin(OnlyOfficeBridge, resolved)
}

export default { name, inject, apply }
