import type { Context } from '@deepseek-ai/cordis'

/** `cordis.patch.yml` 的 `config` 接受的字段（全部可选，环境变量优先）。 */
export interface OnlyOfficeBridgeConfig {
  /** ONLYOFFICE Document Server 地址，如 http://127.0.0.1:8080 */
  dsUrl?: string
  /** 与 DS local.json 一致的 JWT 密钥 */
  jwtSecret?: string
  /** JWT 请求头名（默认 Authorization） */
  jwtHeader?: string
  /** 请求头值是否带 "Bearer " 前缀 */
  jwtHeaderPrefixed?: boolean
  /** DS 能访问到的 DSH 地址；留空用 webServer 自动探测 */
  publicUrl?: string
  /** 浏览器打开编辑器使用的 DS 地址；留空用 dsUrl */
  editorDsUrl?: string
  /** 缓存目录；默认 <DSH_HOME>/onlyoffice */
  storageDir?: string
  /** 缓存过期时间（小时） */
  cacheTtlHour?: number
  /** 自动清理间隔（分钟） */
  cacheCleanIntervalMin?: number
  /** 单文件大小上限（MiB） */
  maxFileMb?: number
  /** 除同源外额外允许的网页来源 */
  allowedOrigins?: string[]
}

/** 已解析的完整配置（环境变量 > config > 默认值）。 */
export interface ResolvedConfig extends Required<OnlyOfficeBridgeConfig> {}

export declare const name: 'dsh-onlyoffice-bridge'
export declare const inject: ['webServer', 'tools']
export declare const Config: unknown

/** Cordis 插件入口：读取 env / .env / config 后挂载 OnlyOfficeBridge 服务。 */
export declare function apply(ctx: Context, config?: OnlyOfficeBridgeConfig): void

/** HS256 JWT 签名（ONLYOFFICE 格式）。 */
export declare function signJwt(payload: Record<string, unknown>, secret: string): string
/** 校验 HS256 JWT；失败返回 undefined。 */
export declare function verifyJwt(token: string, secret: string): Record<string, any> | undefined
/** 解析请求携带的 token（请求头 / body.token / query）。 */
export declare function extractRequestToken(
  req: import('node:http').IncomingMessage,
  body: Record<string, any>,
  config: Pick<ResolvedConfig, 'jwtHeader' | 'jwtHeaderPrefixed' | 'jwtSecret'>,
): Record<string, any> | undefined
/** 生成带 token 的请求头。 */
export declare function authHeaders(
  payload: Record<string, unknown>,
  config: Pick<ResolvedConfig, 'jwtHeader' | 'jwtHeaderPrefixed' | 'jwtSecret'>,
): { token: string; headers: Record<string, string> }
/** 校验 fileKey（<32hex>-<16hex>.<ext>）。 */
export declare function parseFileKey(raw: string): { stem: string; ext: string; name: string } | undefined
/** 同源 / 白名单校验。 */
export declare function isTrustedRequest(
  req: import('node:http').IncomingMessage,
  config: Pick<ResolvedConfig, 'allowedOrigins'>,
  allowSameOrigin: boolean,
): boolean
/** 环境变量 + .env 读取。 */
export declare function readEnvConfig(options?: { dotenv?: Record<string, string>; cwd?: string }): ResolvedConfig
/** 环境变量与 config 合并。 */
export declare function resolveConfig(
  envOptions?: { dotenv?: Record<string, string>; cwd?: string },
  config?: Record<string, unknown>,
): ResolvedConfig

/** `ctx.onlyoffice` 服务。 */
export declare class OnlyOfficeBridge {
  constructor(ctx: Context, config: ResolvedConfig)
  readonly config: ResolvedConfig
  readonly storageDir: string
  /** 生成编辑链接。 */
  open(args: { filePath?: string; fileName?: string; contentBase64?: string; mode?: 'edit' | 'view' }): Promise<Record<string, any>>
  /** 导出 PDF（异步，保存回调落盘）。 */
  exportPdf(args: { fileKey: string; outputPath?: string }): Promise<Record<string, any>>
  /** 列出缓存文档。 */
  list(): Promise<Array<Record<string, unknown>>>
  /** 清理过期缓存，返回被删除的文件名。 */
  cleanExpired(ttlHours?: number): Promise<string[]>
  /** 删除单个缓存。 */
  purge(fileKey: string): { ok: boolean; removed: string[]; fileKey: string }
}

declare const plugin: {
  name: typeof name
  inject: typeof inject
  apply: typeof apply
  Config: typeof Config
}
export default plugin
