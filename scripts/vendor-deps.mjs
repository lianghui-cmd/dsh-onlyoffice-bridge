#!/usr/bin/env node
// 把文档处理依赖「真的打包进仓库」，让插件不再依赖外部 node_modules。
//
// 背景：plugin_manager 重装 bundle 会清理 profile 依赖，插件目录的 node_modules 也可能被清；
// 把运行时闭包 vendor 进来后，`lib/optional-deps.js` 会优先加载内置副本。
//
// 用法：
//   node scripts/vendor-deps.mjs            # 写入 vendor/（会先清空）
//   node scripts/vendor-deps.mjs --dry-run  # 只报告体积，不写文件
//
// 只搬运行时需要的文件：剔除 test/、maps、umd/iife 等无关构建产物。

import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const SOURCE_NODE_MODULES = path.join(ROOT, 'node_modules')
const TARGET = path.join(ROOT, 'vendor')
const ENTRY_PACKAGES = ['docx', 'jszip', 'mammoth', 'xlsx', 'diff', 'marked']

const dryRun = process.argv.includes('--dry-run')

/** 任何包都会剔除的文件/目录名（不参与运行）。 */
const DROP_ALWAYS = new Set([
  'test', 'tests', 'spec', '__tests__', 'benchmark', 'benchmarks', 'examples', 'example',
  'docs', 'doc', 'man', 'coverage', 'scripts', '.github', '.circleci', '.vscode',
  '.nyc_output', 'bower_components',
])
/** 剔除的扩展名。 */
const DROP_EXT = new Set(['.map', '.ts', '.md', '.markdown', '.flow', '.yml', '.yaml', '.txt', '.png', '.jpg', '.gif'])
/** 剔除的文件名。 */
const DROP_FILE = /^(readme|changelog|changes|news|history|contributing|code_of_conduct|security|governance|sponsors|license|license\.md|license\.markdown|tsconfig|\.eslintrc|\.editorconfig|\.travis|\.codeclimate|appveyor)/i

/**
 * 剔除的构建变体。
 * 注意：不能无差别删 *.umd.*——underscore 的 main 就是 underscore-umd.js。
 * 只在包没有被显式引用该文件时才按变体剔除，具体见 shouldKeep。
 */
const DROP_VARIANT = /(\.iife\.|\.browser\.|\.min\.js$)/

/** 上游 docx 依赖 @types/node，是类型包，运行时不需要。 */
const DROP_PACKAGES = new Set(['@types/node', 'undici-types'])

/** 包内精确路径的删除规则（包名 → 相对路径数组）。 */
const PACKAGE_OVERRIDES = {
  // docx：运行时只走 index.mjs + index.d.ts，layout/iife/umd 等旁支构建全部剔除
  docx: [
    'dist/index.umd.cjs', 'dist/index.iife.js', 'dist/index.cjs', 'dist/index.d.cts',
    'dist/layout.umd.cjs', 'dist/layout.iife.js', 'dist/layout.cjs', 'dist/layout.mjs',
    'dist/shapes.cjs', 'dist/shapes.mjs', 'dist/charts.cjs', 'dist/charts.mjs',
    'dist/math.cjs', 'dist/math.mjs', 'dist/watermarks.cjs', 'dist/watermarks.mjs',
    'dist/html.cjs', 'dist/html.mjs', 'dist/macros.cjs', 'dist/macros.mjs',
    'dist/embedded-files.cjs', 'dist/embedded-files.mjs', 'dist/bibliography.cjs', 'dist/bibliography.mjs',
  ],
  mammoth: ['mammoth.browser.js', 'mammoth.browser.min.js', 'browser'],
  // xlsx.js 第 13 行 require('./dist/cpexcel.js')：dist/cpexcel.js 必须保留，
  // 去掉它会让 GBK 等非 UTF-8 的旧 xls / csv 读写出现乱码。
  xlsx: [
    'formats.png', 'legend.png', 'xlsxworker.js', 'xlsx.mjs', 'dist/xlsx.full.min.js',
    'dist/xlsx.extendscript.js', 'dist/xlsx.full.min.js.map', 'dist/xlsx.core.min.map',
    'dist/cpexcel.full.mjs', 'dist/cpexcel.full.js', 'dist/xlsx.core.min.js',
  ],
  marked: ['marked.min.js', 'bin'],
  jszip: ['vendor', 'sponsors.md'],
  underscore: ['amd', 'modules'],
  // codepage 由 xlsx 的 dist/cpexcel.js 内联提供编码表；这里的 main 是 cputils.js，
  // 保留它自身 + bits/ 以备直接按包名加载的情况。
  codepage: ['dist'],
  pako: ['dist'],
  lie: ['polyfill.js'],
  'readable-stream': ['doc'],
}

/**
 * 判断文件是否保留。
 * @param {string} pkgName
 * @param {string} relativePath
 * @param {Set<string>} entryFiles 该包 package.json 里被显式引用的入口文件名
 */
function shouldKeep(pkgName, relativePath, entryFiles) {
  const segments = relativePath.split(/[\\/]/)
  const fileName = segments[segments.length - 1]
  const dirs = segments.slice(0, -1)
  if (dirs.some((dir) => DROP_ALWAYS.has(dir))) return false
  if (DROP_FILE.test(fileName)) return false
  const overrides = PACKAGE_OVERRIDES[pkgName] ?? []
  if (overrides.includes(relativePath.replace(/\\/g, '/'))) return false
  if (entryFiles.has(fileName)) return true
  if (DROP_EXT.has(path.extname(fileName).toLowerCase())) return false
  if (DROP_VARIANT.test(fileName)) return false
  return true
}

/** 收集 package.json 里显式引用的入口文件名（main/module/browser/exports/types）。 */
function entryFileNames(manifest) {
  const names = new Set()
  /** @param {unknown} value */
  const visit = (value) => {
    if (typeof value === 'string') {
      if (/\.(m?js|cjs)$/.test(value)) names.add(path.basename(value))
      return
    }
    if (Array.isArray(value)) {
      for (const item of value) visit(item)
      return
    }
    if (value && typeof value === 'object') {
      for (const item of Object.values(value)) visit(item)
    }
  }
  for (const key of ['main', 'module', 'browser', 'types', 'typings', 'exports']) visit(manifest?.[key])
  return names
}

/** 递归收集一个包内需要保留的文件。 */
function collect(pkgDir, pkgName, files, prefix = '', entryFiles = new Set()) {
  for (const entry of fs.readdirSync(pkgDir, { withFileTypes: true })) {
    const absolute = path.join(pkgDir, entry.name)
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name
    if (entry.isDirectory()) {
      if (DROP_ALWAYS.has(entry.name)) continue
      // 依赖自己的嵌套 node_modules 不复制：外层按闭包解析
      if (entry.name === 'node_modules') continue
      collect(absolute, pkgName, files, relative, entryFiles)
      continue
    }
    if (!shouldKeep(pkgName, relative, entryFiles)) continue
    files.push({ relative, absolute, size: fs.statSync(absolute).size })
  }
}

/** 解析运行时依赖闭包（只走 dependencies，不碰 devDependencies）。 */
function resolveClosure() {
  const closure = new Map()
  const queue = [...ENTRY_PACKAGES]
  while (queue.length) {
    const name = queue.shift()
    if (closure.has(name) || DROP_PACKAGES.has(name)) continue
    const dir = path.join(SOURCE_NODE_MODULES, name)
    const manifestPath = path.join(dir, 'package.json')
    if (!fs.existsSync(manifestPath)) {
      throw new Error(`缺少依赖 ${name}：请先在插件目录执行 npm install`)
    }
    const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
    closure.set(name, { dir, manifest })
    for (const dep of Object.keys(manifest.dependencies ?? {})) {
      if (!closure.has(dep) && !DROP_PACKAGES.has(dep)) queue.push(dep)
    }
  }
  return closure
}

function main() {
  const closure = resolveClosure()
  const report = []
  let totalBytes = 0
  let totalFiles = 0

  if (!dryRun) fs.rmSync(TARGET, { recursive: true, force: true })

  for (const [name, { dir, manifest }] of closure) {
    const files = []
    collect(dir, name, files, '', entryFileNames(manifest))
    const bytes = files.reduce((sum, file) => sum + file.size, 0)
    const before = duSizeOf(dir)
    totalBytes += bytes
    totalFiles += files.length
    report.push({ name, version: manifest.version, files: files.length, bytes, before, license: manifest.license ?? 'UNKNOWN' })

    if (dryRun) continue
    const outDir = path.join(TARGET, name)
    fs.mkdirSync(outDir, { recursive: true })
    // 保留一份精简的 package.json：Node 用 exports/main 解析入口，不能省
    const slim = {
      name: manifest.name,
      version: manifest.version,
      ...(manifest.license ? { license: manifest.license } : {}),
      ...(manifest.type ? { type: manifest.type } : {}),
      main: manifest.main,
      module: manifest.module,
      browser: manifest.browser,
      types: manifest.types,
      typings: manifest.typings,
      exports: manifest.exports,
      dependencies: manifest.dependencies,
      optionalDependencies: manifest.optionalDependencies,
    }
    fs.writeFileSync(path.join(outDir, 'package.json'), `${JSON.stringify(slim, null, 2)}\n`)
    for (const file of files) {
      const destination = path.join(outDir, file.relative)
      fs.mkdirSync(path.dirname(destination), { recursive: true })
      fs.copyFileSync(file.absolute, destination)
    }
    // 保留许可证原文（合规要求）
    for (const licenseFile of fs.readdirSync(dir)) {
      if (/^licen[cs]e/i.test(licenseFile) && fs.statSync(path.join(dir, licenseFile)).isFile()) {
        fs.copyFileSync(path.join(dir, licenseFile), path.join(outDir, licenseFile))
        break
      }
    }
  }

  report.sort((a, b) => b.bytes - a.bytes)
  console.log(`${dryRun ? '[dry-run] ' : ''}vendor 运行时闭包：${report.length} 个包，${totalFiles} 个文件`)
  console.log('包名'.padEnd(22), '版本'.padEnd(9), '文件'.padEnd(6), 'vendor', '原体积')
  for (const item of report) {
    console.log(
      item.name.padEnd(22),
      item.version.padEnd(9),
      String(item.files).padEnd(6),
      `${(item.bytes / 1048576).toFixed(2)}MB`.padEnd(8),
      `${(item.before / 1048576).toFixed(2)}MB`,
    )
  }
  console.log(`合计 vendor: ${(totalBytes / 1048576).toFixed(2)} MB`)
  if (!dryRun) {
    fs.writeFileSync(
      path.join(TARGET, 'MANIFEST.json'),
      `${JSON.stringify({ generatedAt: new Date().toISOString(), entries: ENTRY_PACKAGES, packages: report.map(({ name, version, files, license }) => ({ name, version, files, license })) }, null, 2)}\n`,
    )
    console.log(`已写入 ${path.relative(ROOT, TARGET)}/`)
  }
}

function duSizeOf(dir) {
  let size = 0
  const walk = (current) => {
    for (const entry of fs.readdirSync(current, { withFileTypes: true })) {
      const absolute = path.join(current, entry.name)
      if (entry.isDirectory()) walk(absolute)
      else size += fs.statSync(absolute).size
    }
  }
  walk(dir)
  return size
}

main()
