// @ts-check
// 让 vendor/ 里的包也能解析彼此的裸名依赖（require('underscore') / import 'x' 等）。
//
// 问题：vendor 下的包按 `<plugin>/vendor/<pkg>` 平铺。Node 从
// `<plugin>/vendor/mammoth/lib/xml/xmldom.js` 找 `underscore` 时只会逐级向上找
// node_modules，不会命中兄弟目录 `<plugin>/vendor/underscore`，于是报 Cannot find module。
//
// 方案：<plugin>/vendor 本身就是一个 node_modules 形状的目录（里面直接放包）。
// 所以只要把 `<plugin>/vendor` 这条路径补进查找链即可，两条链路各补一次：
//   - ESM：resolve 钩子（registerHooks），只在 vendor 内模块发起、且默认解析失败时兜底；
//   - CJS：模块编译时用 Module._resolveFilename 的 paths 参数补齐，
//          通过包装 Module.prototype.require 实现，只在 vendor 内模块生效。
//
// 只作用于 vendor 目录内的模块，插件其他部分解析行为不变。

import fs from 'node:fs'
import Module from 'node:module'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const PLUGIN_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const VENDOR_DIR = path.join(PLUGIN_ROOT, 'vendor')
const VENDOR_PREFIX = VENDOR_DIR.replace(/\\/g, '/')

let installed = false

/** vendor 下是否存在任何包。 */
function hasVendorPackages() {
  try {
    return fs
      .readdirSync(VENDOR_DIR, { withFileTypes: true })
      .some((entry) => entry.isDirectory() && fs.existsSync(path.join(VENDOR_DIR, entry.name, 'package.json'))
        || (entry.isDirectory() && entry.name.startsWith('@')))
  } catch {
    return false
  }
}

/** @param {string} from */
function isInVendor(from) {
  return String(from).replace(/\\/g, '/').startsWith(VENDOR_PREFIX)
}

/**
 * 在 vendor 内解析一个裸名（main/exports 语义）。
 * @param {string} specifier
 * @returns {string | undefined}
 */
function resolveInVendor(specifier) {
  if (!specifier || specifier.startsWith('.') || specifier.startsWith('/') || /^[a-z]+:/.test(specifier)) return undefined
  const parts = specifier.split('/')
  const pkgName = specifier.startsWith('@') ? parts.slice(0, 2).join('/') : parts[0]
  const subpath = (specifier.startsWith('@') ? parts.slice(2) : parts.slice(1)).join('/')
  const pkgDir = path.join(VENDOR_DIR, pkgName)
  if (!fs.existsSync(path.join(pkgDir, 'package.json'))) return undefined

  /** @param {string} relative */
  const complete = (relative) => {
    const absolute = path.resolve(pkgDir, relative)
    for (const candidate of [absolute, `${absolute}.js`, `${absolute}.cjs`, `${absolute}.mjs`, path.join(absolute, 'index.js')]) {
      if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) return candidate
    }
    return undefined
  }

  /** @param {unknown} value @returns {string | undefined} */
  const firstString = (value) => {
    if (typeof value === 'string') return value
    if (value && typeof value === 'object' && !Array.isArray(value)) {
      for (const key of ['require', 'import', 'node', 'default']) {
        const nested = firstString(/** @type {any} */ (value)[key])
        if (nested) return nested
      }
    }
    return undefined
  }

  let manifest
  try {
    manifest = JSON.parse(fs.readFileSync(path.join(pkgDir, 'package.json'), 'utf8'))
  } catch {
    return undefined
  }

  if (subpath) {
    const exportsMap = manifest?.exports
    if (exportsMap && typeof exportsMap === 'object' && !Array.isArray(exportsMap)) {
      const mapped = firstString(/** @type {any} */ (exportsMap)[`./${subpath}`])
      if (mapped) {
        const done = complete(mapped)
        if (done) return done
      }
    }
    return complete(subpath)
  }
  const root = firstString(manifest?.exports?.['.'])
  if (root) {
    const done = complete(root)
    if (done) return done
  }
  return complete(manifest?.main ?? 'index.js')
}

/** CJS：vendor 内模块 require 失败时，按 vendor 内路径重试。 */
function patchCommonJs() {
  const proto = /** @type {any} */ (Module).prototype
  if (!proto || proto.__dshVendorRequirePatched) return
  const originalRequire = proto.require
  if (typeof originalRequire !== 'function') return
  proto.require = function patchedRequire(specifier) {
    try {
      return originalRequire.apply(this, arguments)
    } catch (error) {
      const from = typeof this?.filename === 'string' ? this.filename : ''
      if (!from || !isInVendor(from)) throw error
      const target = resolveInVendor(specifier)
      if (!target) throw error
      return originalRequire.call(this, target)
    }
  }
  proto.__dshVendorRequirePatched = true
}

/** ESM：解析失败时在 vendor 内兜底一次。 */
function registerEsmFallback() {
  if (typeof Module.registerHooks !== 'function') return
  Module.registerHooks({
    resolve(specifier, context, nextResolve) {
      try {
        return nextResolve(specifier, context)
      } catch (error) {
        const parent = context?.parentURL
        if (typeof parent !== 'string' || !parent.startsWith('file:')) throw error
        if (!isInVendor(fileURLToPath(parent))) throw error
        const target = resolveInVendor(specifier)
        if (!target) throw error
        return { url: pathToFileURL(target).href, shortCircuit: true }
      }
    },
  })
}

/** 安装（幂等）。vendor 不存在时什么也不做。 */
export function installVendorResolver() {
  if (installed) return
  installed = true
  if (!hasVendorPackages()) return
  patchCommonJs()
  registerEsmFallback()
}
