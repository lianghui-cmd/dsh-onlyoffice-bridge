// @ts-check
// 文档工具层：把桥接服务（存储 / DS 交互）和文档处理能力组合成 25 个 Agent 工具。
//
// 为什么单独一层：
//   - 每个工具都要做「参数校验 + 路径/锁检查 + 错误信息可读」这套胶水；
//   - 文档处理依赖是可选的（docx / mammoth / xlsx / diff），只有真正用到时才加载；
//   - 测试可以注入假 fs 根目录与假依赖，不必启动整个服务。

import crypto from 'node:crypto'
import fs from 'node:fs'
import path from 'node:path'
import { loadDiff, loadDocx, loadMammoth, loadXlsx, loadZip } from './optional-deps.js'
import {
  TEXT_LIMIT,
  cmToPx,
  extractText,
  insertImageIntoDocx,
  imageTypeOf,
  markdownToDocxBuffer,
  patchDocxXml,
  renderPlaceholders,
  rowsToXlsxBuffer,
} from './document.js'

/** 转 PDF 以外还可以导出的格式（走 DS ConvertService）。 */
export const CONVERT_FORMATS = ['pdf', 'docx', 'xlsx', 'pptx', 'html', 'odt', 'ods', 'odp', 'rtf', 'txt', 'csv']
/** 只读正文时可抽取的格式。 */
const READABLE = ['docx', 'xlsx', 'xls', 'ods', 'csv']

/**
 * 读出 docx 模板里的全部占位符。
 * @param {Buffer} buffer
 * @param {Record<string, any>} jszip
 */
export async function collectDocxPlaceholders(buffer, jszip) {
  const zip = await jszip.loadAsync(buffer)
  const entry = zip.file('word/document.xml')
  if (!entry) throw new Error('不是有效的 docx：缺少 word/document.xml')
  const xml = await entry.async('string')
  const pattern = /\$\{\s*([A-Za-z0-9_.\-]+)\s*\}/g
  const plain = String(xml)
    .replace(/<\/w:p>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
  const names = new Set()
  let match
  while ((match = pattern.exec(plain)) !== null) names.add(match[1])
  return [...names]
}

/**
 * @typedef {object} BridgeLike
 * @property {import('../index.js').ResolvedConfig} config
 * @property {string} storageDir
 * @property {any} lastCallback
 * @property {(args: any) => Promise<any>} open
 * @property {(args: any) => Promise<any>} exportPdf
 * @property {(args: any) => Promise<any>} exportAs
 * @property {(fileKey: string) => Promise<Array<Record<string, any>>> | Array<Record<string, any>>} list
 * @property {(fileKey: string) => any} purge
 * @property {(ttlHours?: number) => Promise<string[]>} cleanExpired
 * @property {(fileKey: string, buffer: Buffer) => Promise<any>} applySavedContent
 * @property {(url: string, options?: { jwt?: boolean }) => Promise<Buffer>} downloadUrl
 * @property {() => string} publicBaseUrl
 */

/**
 * @param {BridgeLike} bridge
 * @param {{ logger?: { info: Function, warn: Function, error: Function } }} [options]
 */
export function createToolkit(bridge, options = {}) {
  const logger = options.logger ?? console
  const deps = {
    docx: /** @type {Promise<any> | undefined} */ (undefined),
    mammoth: /** @type {Promise<any> | undefined} */ (undefined),
    xlsx: /** @type {Promise<any> | undefined} */ (undefined),
    diff: /** @type {Promise<any> | undefined} */ (undefined),
    jszip: /** @type {Promise<any> | undefined} */ (undefined),
  }
  const get = {
    docx: () => (deps.docx ??= loadDocx()),
    mammoth: () => (deps.mammoth ??= loadMammoth()),
    xlsx: () => (deps.xlsx ??= loadXlsx()),
    diff: () => (deps.diff ??= loadDiff()),
    jszip: () => (deps.jszip ??= loadZip()),
  }

  // -- 存储辅助 -----------------------------------------------------------

  /** @param {string} fileKey */
  const storagePath = (fileKey) => {
    const name = safeKey(fileKey)
    const candidate = path.join(bridge.storageDir, name)
    const relative = path.relative(bridge.storageDir, candidate)
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) {
      throw new Error(`fileKey 非法：${fileKey}`)
    }
    return candidate
  }

  /** @param {string} fileKey */
  const requirePath = (fileKey) => {
    const target = storagePath(fileKey)
    if (!fs.existsSync(target)) throw new Error(`fileKey 不存在：${fileKey}`)
    return target
  }

  /** @param {string} fileKey */
  const metaOf = (fileKey) => {
    const target = storagePath(fileKey)
    const stem = path.basename(target).replace(/\.[^.]+$/, '')
    const metaPath = path.join(bridge.storageDir, `${stem}.meta.json`)
    if (!fs.existsSync(metaPath)) return {}
    try {
      return JSON.parse(fs.readFileSync(metaPath, 'utf8'))
    } catch {
      return {}
    }
  }

  /** @param {string} fileKey @param {Record<string, any>} patch */
  const writeMeta = (fileKey, patch) => {
    const target = storagePath(fileKey)
    const stem = path.basename(target).replace(/\.[^.]+$/, '')
    const metaPath = path.join(bridge.storageDir, `${stem}.meta.json`)
    const meta = { ...metaOf(fileKey), ...patch, updatedAt: new Date().toISOString() }
    fs.writeFileSync(metaPath, JSON.stringify(meta, null, 2), 'utf8')
    return meta
  }

  // -- 锁（应用层软锁）----------------------------------------------------

  const lockPath = path.join(bridge.storageDir, 'onlyoffice-lock.json')
  const readLocks = () => {
    try {
      const parsed = JSON.parse(fs.readFileSync(lockPath, 'utf8'))
      return parsed && typeof parsed === 'object' ? parsed : {}
    } catch {
      return {}
    }
  }
  /** @param {Record<string, any>} locks */
  const writeLocks = (locks) => fs.writeFileSync(lockPath, JSON.stringify(locks, null, 2), 'utf8')

  /**
   * 写操作前检查软锁；被占用时给出明确提示（软锁只拦 Agent 的写操作，不阻止编辑器）。
   * @param {string} fileKey
   * @param {string | undefined} actor
   */
  const assertUnlocked = (fileKey, actor) => {
    const lock = readLocks()[fileKey]
    if (!lock) return
    if (actor && lock.lockedBy === actor) return
    throw new Error(`文档已被 ${lock.lockedBy} 锁定（${new Date(lock.lockedAt).toLocaleString('zh-CN')}），先 onlyoffice_unlock 再修改`)
  }

  // -- 文档生成 -----------------------------------------------------------

  /**
   * 生成一份新文档写入缓存。
   * @param {string} fileName
   * @param {Buffer} buffer
   * @param {Record<string, any>} [lineage]
   */
  const saveNew = (fileName, buffer, lineage = {}) => {
    const target = /** @type {any} */ (bridge).writeCache(fileName, buffer)
    const key = target.fileKey
    const stem = key.replace(/\.[^.]+$/, '')
    writeMeta(key, {
      fileName,
      lineage: { originalPath: null, rootKey: stem, parentKey: null, depth: 0, ...lineage },
      lastStatus: null,
    })
    return {
      fileKey: key,
      path: target.target,
      fileName,
      size: buffer.length,
    }
  }

  // -- 工具实现 -----------------------------------------------------------

  /** @type {Record<string, (args: any) => Promise<any> | any>} */
  const handlers = {
    // ========== 基础 ==========
    open: async (args) => {
      const input = args ?? {}
      if (input.fileKey) assertUnlocked(String(input.fileKey), input.lockedBy)
      return bridge.open(input)
    },

    'read-doc': async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const target = requirePath(fileKey)
      const ext = path.extname(target).slice(1).toLowerCase()
      if (!READABLE.includes(ext)) {
        return {
          ok: false,
          fileKey,
          msg: `.${ext} 暂不支持正文抽取（支持 ${READABLE.join(' / ')}）；pptx 请用 onlyoffice_export_pdf 或 export_html`,
        }
      }
      const result = await extractText(target, ext, {
        xlsx: await get.xlsx(),
        mammoth: await get.mammoth(),
      })
      return { fileKey, ...result }
    },

    status: async (args) => {
      const docs = await bridge.list()
      const locks = readLocks()
      const filter = typeof args?.fileKey === 'string' && args.fileKey ? args.fileKey : undefined
      const list = docs
        .filter((doc) => !filter || doc.fileKey === filter)
        .map((doc) => ({ ...doc, lock: locks[String(doc.fileKey)] ?? null }))
      return {
        storageDir: bridge.storageDir,
        dsUrl: bridge.config.dsUrl,
        publicUrl: bridge.publicBaseUrl(),
        lastCallback: bridge.lastCallback,
        total: list.length,
        documents: list,
      }
    },

    'export-pdf': async (args) => bridge.exportPdf({ fileKey: safeKey(args?.fileKey), outputPath: args?.outputPath }),

    purge: async (args) => {
      const target = args?.target === 'one' || args?.target === 'all' ? args.target : 'expired'
      if (target === 'one') {
        const fileKey = safeKey(args?.fileKey)
        const result = bridge.purge(fileKey)
        const locks = readLocks()
        delete locks[fileKey]
        writeLocks(locks)
        return { ...result, target }
      }
      if (target === 'all') {
        const docs = await bridge.list()
        const removed = docs
          .filter((doc) => !readLocks()[String(doc.fileKey)])
          .map((doc) => bridge.purge(String(doc.fileKey)))
          .filter((result) => result.ok)
        return { ok: true, target, removed: removed.length, skippedLocked: docs.length - removed.length }
      }
      const removed = await bridge.cleanExpired(bridge.config.cacheTtlHour)
      return { ok: true, target, removed }
    },

    // ========== 第一梯队：写 / 对比 / 元信息 / 锁 / 批量 ==========
    'write-doc': async (args) => {
      const mode = String(args?.mode ?? '').toLowerCase()
      const fileName = String(args?.fileName ?? (mode === 'xlsx' ? '工作簿.xlsx' : '文档.docx'))
      const content = String(args?.content ?? '')
      if (mode === 'docx') {
        const docx = await get.docx()
        const buffer = await markdownToDocxBuffer(content, docx)
        return { ok: true, ...saveNew(ensureExt(fileName, 'docx'), buffer), message: '已生成 docx（Markdown 渲染）' }
      }
      if (mode === 'xlsx') {
        let rows
        try {
          rows = JSON.parse(content)
        } catch (error) {
          throw new Error(`content 不是合法 JSON：${error instanceof Error ? error.message : String(error)}`)
        }
        const xlsx = await get.xlsx()
        const buffer = rowsToXlsxBuffer(rows, String(args?.sheetName ?? 'Sheet1'), xlsx)
        return { ok: true, ...saveNew(ensureExt(fileName, 'xlsx'), buffer) }
      }
      throw new Error('mode 仅支持 docx / xlsx')
    },

    diff: async (args) => {
      const oldKey = safeKey(args?.oldKey)
      const newKey = safeKey(args?.newKey)
      const oldPath = requirePath(oldKey)
      const newPath = requirePath(newKey)
      const extOf = (p) => path.extname(p).slice(1).toLowerCase()
      if (extOf(oldPath) !== extOf(newPath)) throw new Error('两个文档的格式不同，无法逐行对比')
      if (!READABLE.includes(extOf(oldPath))) throw new Error(`.${extOf(oldPath)} 暂不支持对比`)
      const extractor = { xlsx: await get.xlsx(), mammoth: await get.mammoth() }
      const [left, right] = await Promise.all([extractText(oldPath, extOf(oldPath), extractor), extractText(newPath, extOf(newPath), extractor)])
      if (!left.ok || !right.ok) throw new Error(`读取失败：${left.msg ?? ''}${right.msg ?? ''}`)
      const { diffLines } = await get.diff()
      const changes = diffLines(left.text, right.text)
      let added = 0
      let removed = 0
      for (const change of changes) {
        const lines = String(change.value ?? '').split('\n').filter((line) => line !== '').length
        if (change.added) added += lines
        else if (change.removed) removed += lines
      }
      return {
        ok: true,
        oldKey,
        newKey,
        addedLines: added,
        removedLines: removed,
        changed: added + removed > 0,
        hunks: changes
          .filter((change) => change.added || change.removed)
          .slice(0, 50)
          .map((change) => ({ type: change.added ? 'added' : 'removed', value: truncate(String(change.value ?? ''), 4000) })),
      }
    },

    rename: async (args) => {
      const fileKey = safeKey(args?.fileKey)
      requirePath(fileKey)
      const name = String(args?.name ?? '').trim()
      if (!name) throw new Error('name 不能为空')
      const current = metaOf(fileKey)
      const currentExt = String(current.fileName ?? fileKey).split('.').pop()
      const nextName = path.extname(name) ? name : `${name}.${currentExt}`
      const meta = writeMeta(fileKey, { fileName: nextName, remark: args?.remark ?? current.remark ?? '' })
      return { ok: true, fileKey, meta }
    },

    meta: async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const target = requirePath(fileKey)
      const stat = fs.statSync(target)
      return {
        ok: true,
        fileKey,
        meta: metaOf(fileKey),
        lock: readLocks()[fileKey] ?? null,
        size: stat.size,
        updatedAt: new Date(stat.mtimeMs).toISOString(),
        storagePath: target,
      }
    },

    'export-markdown': async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const target = requirePath(fileKey)
      const ext = path.extname(target).slice(1).toLowerCase()
      if (ext === 'docx') {
        // mammoth 的 markdown 输出保留标题/列表/表格语义，比纯文本更适合二次编辑
        const mammoth = await get.mammoth()
        const result = await mammoth.convertToMarkdown({ path: target })
        const markdown = truncate(String(result?.value ?? ''), TEXT_LIMIT * 2)
        return {
          ok: true,
          fileKey,
          markdown,
          truncated: String(result?.value ?? '').length > markdown.length,
          messages: (result?.messages ?? []).slice(0, 20),
        }
      }
      if (ext === 'xlsx' || ext === 'csv') {
        const xlsx = await get.xlsx()
        const workbook = xlsx.read(fs.readFileSync(target), { type: 'buffer' })
        const parts = []
        for (const sheetName of workbook.SheetNames ?? []) {
          const csv = xlsx.utils.sheet_to_csv(workbook.Sheets[sheetName])
          parts.push(`## ${sheetName}`, '', '```csv', csv.trim(), '```', '')
        }
        return { ok: true, fileKey, markdown: parts.join('\n') }
      }
      const fallback = await extractText(target, ext, { xlsx: await get.xlsx(), mammoth: await get.mammoth() })
      if (!fallback.ok) return { ok: false, fileKey, msg: fallback.msg }
      return { ok: true, fileKey, markdown: fallback.text, fallback: true }
    },

    lock: async (args) => {
      const fileKey = safeKey(args?.fileKey)
      requirePath(fileKey)
      const locks = readLocks()
      const existing = locks[fileKey]
      if (existing) {
        return { ok: false, fileKey, msg: `已被 ${existing.lockedBy} 锁定`, lock: existing }
      }
      const lock = { lockedBy: String(args?.lockedBy ?? 'agent'), lockedAt: Date.now(), note: args?.note ?? '' }
      locks[fileKey] = lock
      writeLocks(locks)
      return { ok: true, fileKey, lock }
    },

    unlock: async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const locks = readLocks()
      const existing = locks[fileKey]
      if (!existing) return { ok: true, fileKey, msg: '文档未被锁定' }
      if (args?.lockedBy && existing.lockedBy !== args.lockedBy) {
        return { ok: false, fileKey, msg: `只有 ${existing.lockedBy} 能解锁`, lock: existing }
      }
      delete locks[fileKey]
      writeLocks(locks)
      return { ok: true, fileKey, release: existing }
    },

    'who-lock': async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const lock = readLocks()[fileKey] ?? null
      return { ok: true, fileKey, locked: Boolean(lock), lockInfo: lock }
    },

    'batch-open': async (args) => {
      const list = Array.isArray(args?.filePathList) ? args.filePathList : []
      if (!list.length) throw new Error('filePathList 不能为空')
      const results = []
      for (const filePath of list) {
        try {
          const opened = await bridge.open({ filePath: String(filePath), mode: args?.mode })
          results.push({ ok: true, filePath, fileKey: opened.fileKey, editorUrl: opened.editorUrl })
        } catch (error) {
          results.push({ ok: false, filePath, msg: error instanceof Error ? error.message : String(error) })
        }
      }
      return { total: results.length, succeeded: results.filter((item) => item.ok).length, result: results }
    },

    'batch-read': async (args) => {
      const list = Array.isArray(args?.fileKeyList) ? args.fileKeyList : []
      if (!list.length) throw new Error('fileKeyList 不能为空')
      const xlsx = await get.xlsx()
      const mammoth = await get.mammoth()
      const results = []
      for (const rawKey of list) {
        try {
          const fileKey = safeKey(rawKey)
          const target = requirePath(fileKey)
          const ext = path.extname(target).slice(1).toLowerCase()
          if (!READABLE.includes(ext)) {
            results.push({ fileKey, ok: false, msg: `.${ext} 暂不支持正文抽取` })
            continue
          }
          results.push({ fileKey, ...(await extractText(target, ext, { xlsx, mammoth })) })
        } catch (error) {
          results.push({ fileKey: String(rawKey), ok: false, msg: error instanceof Error ? error.message : String(error) })
        }
      }
      return { total: results.length, result: results }
    },

    'batch-export-pdf': async (args) => {
      const list = Array.isArray(args?.fileKeyList) ? args.fileKeyList : []
      if (!list.length) throw new Error('fileKeyList 不能为空')
      const outDir = args?.outputDir ? path.resolve(String(args.outputDir)) : undefined
      const results = []
      for (const rawKey of list) {
        try {
          const fileKey = safeKey(rawKey)
          const outputPath = outDir
            ? path.join(outDir, `${path.basename(fileKey).replace(/\.[^.]+$/, '')}.pdf`)
            : undefined
          const exported = await bridge.exportPdf({ fileKey, outputPath })
          results.push({ ok: true, fileKey, outputPath: exported.outputPath, editorUrl: exported.editorUrl })
        } catch (error) {
          results.push({ ok: false, fileKey: String(rawKey), msg: error instanceof Error ? error.message : String(error) })
        }
      }
      return { total: results.length, succeeded: results.filter((item) => item.ok).length, result: results }
    },

    // ========== 第二梯队：多格式导出 ==========
    'export-docx': async (args) => bridge.exportAs({ fileKey: safeKey(args?.fileKey), format: 'docx', outputPath: args?.outputPath }),
    'export-xlsx': async (args) => bridge.exportAs({ fileKey: safeKey(args?.fileKey), format: 'xlsx', outputPath: args?.outputPath }),
    'export-pptx': async (args) => bridge.exportAs({ fileKey: safeKey(args?.fileKey), format: 'pptx', outputPath: args?.outputPath }),
    'export-html': async (args) => {
      const result = await bridge.exportAs({ fileKey: safeKey(args?.fileKey), format: 'html', outputPath: args?.outputPath })
      const html = fs.readFileSync(result.outputPath, 'utf8')
      return { ...result, preview: truncate(html, 2000) }
    },

    // ========== 第二梯队：模板填充（docx） / 插图（docx） ==========
    'template-fill': async (args) => {
      const templateKey = safeKey(args?.templateKey ?? args?.fileKey)
      const templatePath = requirePath(templateKey)
      if (!templatePath.toLowerCase().endsWith('.docx')) throw new Error('模板仅支持 docx')
      let vars = args?.vars ?? args?.variables
      if (typeof vars === 'string') {
        try {
          vars = JSON.parse(vars)
        } catch (error) {
          throw new Error(`vars 不是合法 JSON：${error instanceof Error ? error.message : String(error)}`)
        }
      }
      if (!vars || typeof vars !== 'object') throw new Error('vars 必须是对象或 JSON 字符串')

      const buffer = fs.readFileSync(templatePath)
      const placeholders = await collectDocxPlaceholders(buffer, await get.jszip())

      let replaced = []
      let missing = []
      const out = await patchDocxXml(
        buffer,
        (xml) => {
          const result = renderPlaceholders(xml, /** @type {Record<string, unknown>} */ (vars))
          replaced = result.replaced
          missing = result.missing
          return result.xml
        },
        await get.jszip(),
      )

      const newFileName = ensureExt(String(args?.newFileName ?? `模板填充-${Date.now()}.docx`), 'docx')
      const saved = saveNew(newFileName, out, { templateKey, remark: `来自模板 ${templateKey}` })
      writeMeta(saved.fileKey, { remark: `来自模板 ${templateKey}` })
      return {
        ok: true,
        newFileKey: saved.fileKey,
        fileName: newFileName,
        placeholders,
        replaced,
        missing,
        message: missing.length
          ? `已填充，但模板里这些占位符没有给值：${missing.join(', ')}`
          : '已按模板生成新文档',
      }
    },

    'insert-image': async (args) => {
      const fileKey = safeKey(args?.fileKey)
      const target = requirePath(fileKey)
      if (!target.toLowerCase().endsWith('.docx')) throw new Error('insert-image 仅支持 docx（xlsx / pptx 暂不支持）')
      assertUnlocked(fileKey, args?.lockedBy)
      const imagePath = path.resolve(String(args?.imgPath ?? ''))
      if (!args?.imgPath) throw new Error('imgPath 不能为空')
      if (!fs.existsSync(imagePath)) throw new Error(`图片不存在：${imagePath}`)
      const data = fs.readFileSync(imagePath)
      const type = imageTypeOf(data)
      const widthPx = cmToPx(Number(args?.widthCm ?? 14))
      const ratio = typeof args?.heightCm === 'number' ? cmToPx(args.heightCm) / widthPx : guessRatio(data)
      const heightPx = Math.round(widthPx * (Number.isFinite(ratio) && ratio > 0 ? ratio : 0.6))
      const next = await insertImageIntoDocx(
        fs.readFileSync(target),
        { data, type, widthPx, heightPx },
        await get.jszip(),
      )
      fs.writeFileSync(target, next)
      writeMeta(fileKey, { lastImage: imagePath })
      return {
        ok: true,
        fileKey,
        imagePath,
        widthCm: Number(args?.widthCm ?? 14),
        message: '图片已追加到文档末尾；已打开编辑器的会话需要重新打开链接才能看到新内容',
      }
    },
  }

  return { handlers, storagePath, readLocks, writeMeta, metaOf }
}

// ---------------------------------------------------------------------------
// 小工具
// ---------------------------------------------------------------------------

/**
 * 规范化并校验 fileKey / 文件名，拒绝路径穿越。
 * @param {unknown} raw
 */
export function safeKey(raw) {
  const value = String(raw ?? '').trim()
  if (!value) throw new Error('fileKey 不能为空')
  if (value.includes('/') || value.includes('\\') || value.includes('..') || value.includes('\0')) {
    throw new Error(`fileKey 非法：${value}`)
  }
  return value
}

/**
 * 确保文件名词带有指定扩展名。
 * @param {string} name
 * @param {string} ext
 */
export function ensureExt(name, ext) {
  const clean = path.basename(String(name ?? '').trim()) || `document.${ext}`
  return clean.toLowerCase().endsWith(`.${ext}`) ? clean : `${clean.replace(/\.[^.]+$/, '')}.${ext}`
}

/** @param {string} value @param {number} limit */
function truncate(value, limit) {
  return value.length > limit ? `${value.slice(0, limit)}\n…（已截断，原文 ${value.length} 字符）` : value
}

/**
 * 依据常见图片格式的特征字节粗略估算宽高比（png / jpg / gif）。
 * @param {Buffer} data
 */
function guessRatio(data) {
  try {
    if (data.subarray(0, 8).toString('hex') === '89504e470d0a1a0a' && data.length > 24) {
      const width = data.readUInt32BE(16)
      const height = data.readUInt32BE(20)
      if (width > 0 && height > 0) return height / width
    }
    if (data.subarray(0, 4).toString('latin1') === 'GIF8' && data.length > 10) {
      return data.readUInt16LE(8) / data.readUInt16LE(6)
    }
    for (let i = 2; i < data.length - 9; ) {
      if (data[i] !== 0xff) {
        i += 1
        continue
      }
      const marker = data[i + 1]
      const length = data.readUInt16BE(i + 2)
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        const height = data.readUInt16BE(i + 5)
        const width = data.readUInt16BE(i + 7)
        if (width > 0) return height / width
        break
      }
      i += 2 + Math.max(length, 2)
    }
  } catch {
    // 读不出来就用默认比例
  }
  return 0.6
}

/** 只读正文支持的扩展名（供文档/测试引用）。 */
export { READABLE }
