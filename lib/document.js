// @ts-check
// Word / Excel 文档的读取、生成与原地修补。
// 所有第三方库都由调用方通过参数注入（见 lib/optional-deps.js），本模块只做逻辑，
// 既不产生静态依赖，也方便单元测试替换实现。

import fs from 'node:fs'

/** 单次抽取正文的长度上限，避免把超大文档直接塞进模型上下文。 */
export const TEXT_LIMIT = 200_000

/**
 * @typedef {Record<string, any>} DocxApi
 * @typedef {Record<string, any>} XlsxApi
 * @typedef {Record<string, any>} MammothApi
 */

/**
 * 抽取文档纯文本。
 * @param {string} filePath
 * @param {string} ext 小写扩展名（不含点）
 * @param {{ xlsx: XlsxApi, mammoth: MammothApi }} deps
 * @returns {Promise<{ ok: boolean, text: string, msg?: string, truncated?: boolean }>}
 */
export async function extractText(filePath, ext, deps) {
  try {
    let text = ''
    if (ext === 'docx') {
      const result = await deps.mammoth.extractRawText({ path: filePath })
      text = String(result?.value ?? '')
    } else if (['xlsx', 'xls', 'ods', 'csv'].includes(ext)) {
      const workbook = deps.xlsx.read(fs.readFileSync(filePath), { type: 'buffer' })
      const parts = []
      for (const sheetName of workbook.SheetNames ?? []) {
        const sheet = workbook.Sheets[sheetName]
        const rows = deps.xlsx.utils.sheet_to_json(sheet, { header: 1, blankrows: false })
        parts.push(`===== 工作表：${sheetName} =====`)
        parts.push(JSON.stringify(rows, null, 2))
        if (parts.join('\n').length > TEXT_LIMIT) break
      }
      text = parts.join('\n')
    } else {
      return { ok: false, text: '', msg: `暂不支持抽取 .${ext} 的正文（支持 docx / xlsx / csv）` }
    }
    if (text.length > TEXT_LIMIT) {
      return { ok: true, text: text.slice(0, TEXT_LIMIT), truncated: true }
    }
    return { ok: true, text }
  } catch (error) {
    return { ok: false, text: '', msg: `解析异常：${error instanceof Error ? error.message : String(error)}` }
  }
}

/**
 * Markdown → docx Buffer。支持标题、段落、粗体/斜体/行内代码、列表、引用、代码块、表格、分隔线。
 * @param {string} markdown
 * @param {DocxApi} docx
 * @returns {Promise<Buffer>}
 */
export async function markdownToDocxBuffer(markdown, docx) {
  const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx
  /** @type {any[]} */
  const children = []
  let inCode = false
  /** @type {string[]} */
  let codeLines = []
  /** @type {string[][]} */
  let tableRows = []

  const flushTable = () => {
    if (!tableRows.length) return
    const rows = tableRows
    tableRows = []
    const width = Math.max(...rows.map((cells) => cells.length))
    children.push(
      new docx.Table({
        rows: rows.map(
          /** @param {string[]} cells */
          (cells) =>
            new docx.TableRow({
              children: Array.from(
                { length: width },
                (_, index) =>
                  new docx.TableCell({
                    children: [new Paragraph({ children: [new TextRun({ text: cells[index] ?? '', size: 20 })] })],
                  }),
              ),
            }),
        ),
      }),
    )
    children.push(new Paragraph({ text: '' }))
  }

  const flushCode = () => {
    if (!codeLines.length) return
    for (const line of codeLines) {
      children.push(
        new Paragraph({
          children: [new TextRun({ text: line || ' ', font: 'Consolas', size: 19 })],
          shading: { fill: 'F5F5F5' },
        }),
      )
    }
    codeLines = []
    children.push(new Paragraph({ text: '' }))
  }

  for (const rawLine of String(markdown ?? '').replace(/\r\n?/g, '\n').split('\n')) {
    const line = rawLine.trimEnd()
    if (/^\s*```/.test(line)) {
      if (inCode) {
        flushCode()
        inCode = false
      } else {
        flushTable()
        inCode = true
      }
      continue
    }
    if (inCode) {
      codeLines.push(rawLine)
      continue
    }
    if (!line.trim()) {
      flushTable()
      continue
    }
    if (/^\s*\|.*\|\s*$/.test(line)) {
      if (/^[\s|:-]+$/.test(line)) continue
      tableRows.push(
        line
          .trim()
          .replace(/^\|/, '')
          .replace(/\|$/, '')
          .split('|')
          .map((cell) => cell.trim()),
      )
      continue
    }
    flushTable()
    const heading = /^(#{1,6})\s+(.*)$/.exec(line)
    if (heading) {
      const level = Math.min(6, heading[1].length)
      children.push(
        new Paragraph({
          heading: HeadingLevel?.[`HEADING_${level}`],
          children: inlineRuns(heading[2], docx, level <= 2 ? 30 : 24),
          spacing: { after: 120 },
        }),
      )
      continue
    }
    if (/^\s*([-*_])\s*\1\s*\1/.test(line.trim())) {
      children.push(new Paragraph({ text: '', border: { bottom: { style: 'single', size: 6, color: '999999' } } }))
      continue
    }
    const unordered = /^\s*[-*+]\s+(.*)$/.exec(line)
    if (unordered) {
      children.push(new Paragraph({ bullet: { level: 0 }, children: inlineRuns(unordered[1], docx, 22) }))
      continue
    }
    const ordered = /^\s*(\d+)[.)]\s+(.*)$/.exec(line)
    if (ordered) {
      children.push(
        new Paragraph({
          numbering: { reference: 'dsh-ordered', level: 0 },
          children: inlineRuns(ordered[2], docx, 22),
        }),
      )
      continue
    }
    const quote = /^\s*>\s?(.*)$/.exec(line)
    if (quote) {
      children.push(
        new Paragraph({ children: inlineRuns(quote[1], docx, 22), indent: { left: 360 }, shading: { fill: 'F7F7F7' } }),
      )
      continue
    }
    children.push(new Paragraph({ children: inlineRuns(line, docx, 22), spacing: { after: 100 } }))
  }
  flushTable()
  if (inCode) flushCode()
  if (!children.length) children.push(new Paragraph({ alignment: AlignmentType?.LEFT, text: '' }))

  const document = new Document({
    numbering: {
      config: [
        { reference: 'dsh-ordered', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: 'left' }] },
      ],
    },
    sections: [{ properties: {}, children }],
  })
  return /** @type {Buffer} */ (await Packer.toBuffer(document))
}

/**
 * 解析 **粗体** / *斜体* / `代码`，返回 TextRun 数组。
 * @param {string} text
 * @param {DocxApi} docx
 * @param {number} size 半磅值（22 约等于 11pt）
 * @returns {any[]}
 */
function inlineRuns(text, docx, size) {
  const { TextRun } = docx
  /** @type {any[]} */
  const runs = []
  const pattern = /(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/g
  let last = 0
  let match
  while ((match = pattern.exec(text)) !== null) {
    if (match.index > last) runs.push(new TextRun({ text: text.slice(last, match.index), size }))
    const token = match[0]
    if (token.startsWith('**')) runs.push(new TextRun({ text: token.slice(2, -2), bold: true, size }))
    else if (token.startsWith('`')) runs.push(new TextRun({ text: token.slice(1, -1), font: 'Consolas', size }))
    else runs.push(new TextRun({ text: token.slice(1, -1), italics: true, size }))
    last = match.index + token.length
  }
  if (last < text.length) runs.push(new TextRun({ text: text.slice(last), size }))
  if (!runs.length) runs.push(new TextRun({ text: '', size }))
  return runs
}

/**
 * 二维数组或对象数组 → xlsx Buffer。
 * @param {unknown} rows
 * @param {string} sheetName
 * @param {XlsxApi} xlsx
 * @returns {Buffer}
 */
export function rowsToXlsxBuffer(rows, sheetName, xlsx) {
  if (!Array.isArray(rows)) throw new Error('xlsx 内容必须是数组：二维数组或对象数组')
  const isMatrix = rows.every((row) => Array.isArray(row))
  const sheet = isMatrix
    ? xlsx.utils.aoa_to_sheet(/** @type {any[][]} */ (rows))
    : xlsx.utils.json_to_sheet(/** @type {any[]} */ (rows))
  const workbook = xlsx.utils.book_new()
  const name = String(sheetName || 'Sheet1').replace(/[\\/?*[\]:]/g, '').slice(0, 31) || 'Sheet1'
  xlsx.utils.book_append_sheet(workbook, sheet, name)
  return /** @type {Buffer} */ (xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' }))
}

/**
 * 解包 docx → 改 word/document.xml → 重新打包，保留原有样式。
 * @param {Buffer} buffer
 * @param {(xml: string) => string} transform
 * @param {Record<string, any>} [jszip]
 * @returns {Promise<Buffer>}
 */
export async function patchDocxXml(buffer, transform, jszip) {
  const JSZip = jszip ?? (await import('jszip')).default
  const zip = await JSZip.loadAsync(buffer)
  const entry = zip.file('word/document.xml')
  if (!entry) throw new Error('不是有效的 docx：缺少 word/document.xml')
  const xml = await entry.async('string')
  zip.file('word/document.xml', transform(xml))
  return /** @type {Buffer} */ (await zip.generateAsync({ type: 'nodebuffer' }))
}

/**
 * 往 docx 末尾插入图片（保留原有内容），返回新的 docx Buffer。
 * @param {Buffer} buffer
 * @param {{ data: Buffer, type: string, widthPx: number, heightPx: number }} image
 * @param {Record<string, any>} jszip
 * @returns {Promise<Buffer>}
 */
export async function insertImageIntoDocx(buffer, image, jszip) {
  const zip = await jszip.loadAsync(buffer)
  const mediaName = `dsh-image-${Date.now()}.${image.type}`
  zip.file(`word/media/${mediaName}`, image.data)

  const relsEntry = zip.file('word/_rels/document.xml.rels')
  const relsXml = relsEntry ? await relsEntry.async('string') : ''
  /** @type {string[]} */
  const usedIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map((match) => Number(match[1]))
  const rId = `rId${(usedIds.length ? Math.max(...usedIds) : 0) + 1}`
  zip.file('word/_rels/document.xml.rels', appendImageRelationship(relsXml, rId, `media/${mediaName}`))

  const documentEntry = zip.file('word/document.xml')
  if (!documentEntry) throw new Error('不是有效的 docx：缺少 word/document.xml')
  const xml = await documentEntry.async('string')
  /** @type {string[]} */
  const docPrIds = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map((match) => Number(match[1]))
  const docPrId = (docPrIds.length ? Math.max(...docPrIds) : 0) + 1
  zip.file(
    'word/document.xml',
    appendImageParagraph(xml, { rId, widthPx: image.widthPx, heightPx: image.heightPx, docPrId }),
  )
  // 新增的媒体文件写进 [Content_Types].xml，避免 Office 报“内容类型缺失”
  const contentTypes = zip.file('[Content_Types].xml')
  if (contentTypes) {
    const typesXml = await contentTypes.async('string')
    const ext = image.type
    if (!new RegExp(`Extension="${ext}"`, 'i').test(typesXml)) {
      const mime = ext === 'jpg' ? 'image/jpeg' : `image/${ext}`
      zip.file(
        '[Content_Types].xml',
        typesXml.replace(/(<Types\b[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${mime}"/>`),
      )
    }
  }
  return /** @type {Buffer} */ (await zip.generateAsync({ type: 'nodebuffer' }))
}

/**
 * 收集 docx 正文里的 `${var}` 占位符。
 * @param {string} xml
 * @returns {string[]}
 */
export function collectPlaceholders(xml) {
  const names = new Set()
  const pattern = /\$\{\s*([A-Za-z0-9_.\-]+)\s*\}/g
  let match
  while ((match = pattern.exec(stripXmlTags(xml))) !== null) names.add(match[1])
  return [...names]
}

/**
 * 把 `${var}` 渲染进 document.xml。值按 XML 转义写入，避免破坏文档结构。
 * @param {string} xml
 * @param {Record<string, unknown>} vars
 * @returns {{ xml: string, replaced: string[], missing: string[] }}
 */
export function renderPlaceholders(xml, vars) {
  const present = collectPlaceholders(xml)
  const replaced = []
  const missing = present.filter((name) => vars[name] === undefined)
  let out = xml
  for (const [name, value] of Object.entries(vars)) {
    const pattern = new RegExp(`\\$\\{\\s*${escapeRegExp(name)}\\s*\\}`, 'g')
    if (!pattern.test(out)) continue
    replaced.push(name)
    out = out.replace(pattern, escapeXml(value === null || value === undefined ? '' : String(value)))
  }
  return { xml: out, replaced, missing }
}

/**
 * 往 docx 末尾追加一段图片。
 *
 * 做法是在 `<w:body><w:sectPr>` 之前插入一个居中的段落 + 内联图片，
 * 并同步升级 drawing 的命名空间版本（docx 9 生成的文档才认 wp14 的尺寸单位）。
 * @param {string} xml
 * @param {{ rId: string, widthPx: number, heightPx: number, docPrId: number }} image
 * @returns {string}
 */
export function appendImageParagraph(xml, image) {
  const docPrId = image.docPrId
  const paragraph =
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${Math.round(image.widthPx * 9525)}" cy="${Math.round(image.heightPx * 9525)}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="${docPrId}" name="DSH Image ${docPrId}"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${docPrId}" name="image"/>` +
    `<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${image.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/>` +
    `<a:ext cx="${Math.round(image.widthPx * 9525)}" cy="${Math.round(image.heightPx * 9525)}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    `</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`

  const ensureNamespaces = (source) => {
    let out = source
    if (!out.includes('xmlns:wp="')) {
      out = out.replace(
        /<w:document\b/,
        '<w:document xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
      )
    }
    for (const [prefix, uri] of [
      ['a', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
      ['pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture'],
      ['r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'],
    ]) {
      if (!out.includes(`xmlns:${prefix}="`)) out = out.replace(/<w:document\b/, `<w:document xmlns:${prefix}="${uri}"`)
    }
    return out
  }

  const withNamespaces = ensureNamespaces(xml)
  if (withNamespaces.includes('<w:sectPr')) {
    return withNamespaces.replace(/<w:sectPr/, `${paragraph}<w:sectPr`)
  }
  return withNamespaces.replace(/<\/w:body>/, `${paragraph}</w:body>`)
}

/**
 * 往 document.xml.rels 里追加一条图片关系。
 * @param {string} relsXml
 * @param {string} rId
 * @param {string} target 形如 media/image1.png
 */
export function appendImageRelationship(relsXml, rId, target) {
  const relationship =
    `<Relationship Id="${rId}" ` +
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ' +
    `Target="${target}"/>`
  if (relsXml.includes('</Relationships>')) return relsXml.replace('</Relationships>', `${relationship}</Relationships>`)
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `${relationship}</Relationships>`
  )
}

/** @param {string} xml */
function stripXmlTags(xml) {
  return String(xml)
    .replace(/<w:tab\/>/g, '\t')
    .replace(/<\/w:p>/g, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

/** @param {string} value */
function escapeXml(value) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/** @param {string} value */
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * 依据魔数判断 OOXML 需要的图片类型。
 * @param {Buffer} data
 */
export function imageTypeOf(data) {
  if (data.subarray(0, 3).toString('latin1') === '\xFF\xD8\xFF') return 'jpg'
  if (data.subarray(0, 8).toString('hex') === '89504e470d0a1a0a') return 'png'
  if (data.subarray(0, 4).toString('latin1') === 'GIF8') return 'gif'
  if (data.subarray(0, 4).toString('latin1') === 'RIFF' && data.subarray(8, 12).toString('latin1') === 'WEBP') return 'webp'
  throw new Error('不支持的图片格式（仅 png / jpg / gif / webp）')
}

/** 厘米 → 96dpi 像素。 */
export function cmToPx(cm) {
  return Math.max(16, Math.round((Number(cm) || 14) * (96 / 2.54)))
}
