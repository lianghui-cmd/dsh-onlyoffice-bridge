// @ts-check
// docx / jszip / mammoth / xlsx / diff / marked 的懒加载封装。
//
// 解析顺序（保证「插件自包含」）：
//   1. vendor/<pkg>  —— 随仓库打包的内置副本（scripts/vendor-deps.mjs 生成）
//   2. 裸包名        —— 外部 node_modules（本地开发 / 已 npm install 的情况）
//
// 为什么这样排：plugin_manager 重装 bundle 会清理 profile 依赖，插件目录的 node_modules
// 也可能被清空；内置副本不依赖外部安装，重装后插件依旧可用。
//
// 依赖图里没有任何可选包是静态 import，缺包只影响用到它的那几个工具，
// 插件启动与其余工具不受影响。
//
// 注意：@deepseek-ai/cordis 这类 DSH 运行时包必须保持静态 import——
// 它们由 DSH 注入，不能也不应该 vendor。

import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { installVendorResolver } from './vendor-resolver.js'

const VENDOR_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'vendor')

/** 内置副本是否存在。 */
export function hasVendor() {
  return fs.existsSync(VENDOR_DIR)
}

/**
 * 解析一个包的入口 URL：优先 vendor，其次外部 node_modules。
 * @param {string} name
 * @returns {string | undefined} 可直接 import 的 URL；都找不到时返回 undefined
 */
function resolveEntry(name) {
  const vendorDir = path.join(VENDOR_DIR, name)
  const manifestPath = path.join(vendorDir, 'package.json')
  if (fs.existsSync(manifestPath)) {
    /** @type {any} */
    let manifest = {}
    try {
      manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
    } catch {
      manifest = {}
    }
    const entry = pickEntry(manifest, vendorDir)
    if (entry) return pathToFileURL(entry).href
  }
  return undefined
}

/**
 * 按 exports → module → main 的顺序找入口文件。
 * @param {any} manifest
 * @param {string} dir
 * @returns {string | undefined}
 */
function pickEntry(manifest, dir) {
  /** @type {string[]} */
  const candidates = []
  const dot = manifest?.exports?.['.']
  if (typeof dot === 'string') candidates.push(dot)
  else if (dot && typeof dot === 'object') {
    for (const key of ['import', 'require', 'default']) {
      const value = dot[key]
      if (typeof value === 'string') candidates.push(value)
      else if (value && typeof value === 'object' && typeof value.default === 'string') candidates.push(value.default)
    }
  }
  for (const key of ['module', 'main']) {
    if (typeof manifest?.[key] === 'string') candidates.push(manifest[key])
  }
  for (const candidate of candidates) {
    const absolute = path.resolve(dir, candidate)
    if (fs.existsSync(absolute)) return absolute
    // 省略扩展名时补 index.js
    for (const suffix of ['.js', '.cjs', '.mjs']) {
      if (fs.existsSync(`${absolute}${suffix}`)) return `${absolute}${suffix}`
    }
  }
  return undefined
}

/**
 * 加载一个可选依赖：vendor 优先，外部兜底。
 * @param {string} name
 */
async function loadOptional(name) {
  const vendorEntry = resolveEntry(name)
  if (vendorEntry) installVendorResolver()
  /** @type {unknown[]} */
  const attempts = []
  const specifiers = vendorEntry ? [vendorEntry, name] : [name]
  for (const specifier of specifiers) {
    try {
      return await import(specifier)
    } catch (error) {
      attempts.push(error)
    }
  }
  const details = attempts
    .map((error) => (error instanceof Error ? error.message : String(error)))
    .join(' | ')
  const hint = `缺少可选依赖 ${JSON.stringify(name)}：内置副本 ${vendorEntry ? '加载失败' : '不在 vendor/ 中'}，外部也没有安装`
  /** @type {Error & { code?: string }} */
  const wrapped = new Error(
    `${hint}。请重新生成内置副本（npm run vendor）或在插件目录执行 npm install。原始错误：${details}`,
  )
  wrapped.code = 'OPTIONAL_DEPENDENCY_MISSING'
  throw wrapped
}

/**
 * 兼容 ESM / CJS 双形态：CJS 包经 import() 进来时，真正的导出可能是
 * `mod.default`（module.exports 被包装）或 `mod` 本身。
 * 注意导出也可能是**函数**（如 jszip 的构造器，loadAsync 挂在函数对象上），
 * 所以要同时接受 function 与 object。
 * @param {unknown} mod
 */
function interop(mod) {
  /** @type {any} */
  let value = mod ?? {}
  for (let depth = 0; depth < 3; depth += 1) {
    const inner = value?.default
    if (inner && (typeof inner === 'object' || typeof inner === 'function')) {
      value = inner
      continue
    }
    break
  }
  return value
}

/** 从一个模块命名空间里找可用的 API 对象（用于 docx 这类既有具名导出又有 default 的包）。 */
function apiOf(mod) {
  /** @type {Record<string, any>} */
  const namespace = (mod ?? {}) ?? {}
  const inner = interop(mod)
  const spreadable = inner && (typeof inner === 'object' || typeof inner === 'function') ? { ...inner } : {}
  return { ...spreadable, ...namespace }
}

/** DOCX 相关 API（docx 库）。 */
export async function loadDocx() {
  const merged = apiOf(await loadOptional('docx'))
  if (typeof merged.Document !== 'function' || !merged.Packer) {
    throw new Error('docx 模块缺少 Document / Packer 导出，请升级 docx 到 9.x')
  }
  return merged
}

/** Word → 纯文本（mammoth）。 */
export async function loadMammoth() {
  const api = interop(await loadOptional('mammoth'))
  if (typeof api.extractRawText !== 'function') throw new Error('mammoth 模块缺少 extractRawText 导出')
  return /** @type {Record<string, any>} */ (api)
}

/** 电子表格读写（xlsx）。 */
export async function loadXlsx() {
  const api = interop(await loadOptional('xlsx'))
  if (!api.utils || typeof api.read !== 'function') throw new Error('xlsx 模块缺少 utils / read 导出')
  return /** @type {Record<string, any>} */ (api)
}

/** 文本差异（diff）。 */
export async function loadDiff() {
  const api = interop(await loadOptional('diff'))
  if (typeof api.diffLines !== 'function') throw new Error('diff 模块缺少 diffLines 导出')
  return /** @type {Record<string, any>} */ (api)
}

/** Markdown → HTML（marked，write-doc 的 md 渲染用）。 */
export async function loadMarked() {
  const api = interop(await loadOptional('marked'))
  const parse =
    typeof api.parse === 'function' ? api.parse : typeof api.marked === 'function' ? api.marked : undefined
  if (typeof parse !== 'function') throw new Error('marked 模块缺少 parse / marked 导出')
  return { parse: /** @type {(md: string) => string} */ (parse) }
}

/** ZIP 读写（docx 解包修补用）。 */
export async function loadZip() {
  const api = interop(await loadOptional('jszip'))
  if (typeof api.loadAsync !== 'function') throw new Error('jszip 模块缺少 loadAsync 导出')
  return /** @type {Record<string, any>} */ (api)
}
