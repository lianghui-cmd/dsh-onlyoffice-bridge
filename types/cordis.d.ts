// 仓库内的最小类型桩：只声明本插件实际用到的类型。
// 已安装 DSH profile / 安装包时，可把 tsconfig 的 paths 指到 node_modules 里的真实 .d.ts。
import type { IncomingMessage, ServerResponse } from 'node:http'

export interface RouteRequest extends IncomingMessage {}

export interface WebRoute {
  kind: 'exact' | 'prefix'
  path: string
  handler: (req: IncomingMessage, res: ServerResponse) => void | Promise<void>
}

export interface ToolDefinitionLike {
  name: string
  description: string
  parameters: Record<string, unknown>
  output: { schema: Record<string, unknown>; render(args: unknown, value: unknown): unknown }
  execute(args: unknown, exec: unknown): Promise<unknown>
}

export interface LoggerLike {
  info(...args: unknown[]): void
  warn(...args: unknown[]): void
  error(...args: unknown[]): void
}

export interface Context {
  readonly logger: LoggerLike
  effect<T>(execute: () => T): T
  plugin(plugin: unknown, config?: unknown): unknown
  readonly webServer: {
    getOrigin(): string
    register(route: WebRoute): () => void
  }
  readonly tools: {
    register(definition: ToolDefinitionLike): () => void
  }
}

export declare class Service {
  static init: symbol
  static check: symbol
  static config: symbol
  static invoke: symbol
  static extend: symbol
  static tracker: symbol
  constructor(ctx: Context, name?: string)
  readonly ctx: Context
}

declare const cordis: { Service: typeof Service }
export default cordis
