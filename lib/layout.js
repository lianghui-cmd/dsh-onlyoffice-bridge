// @ts-check
// 排版与结构类能力：刷新目录、设置页眉页脚、docx 合并/拆分、pptx 合并/拆分。
//
// 全部基于 OOXML 部件级改写（见 lib/ooxml.js），不重建文档树，
// 因此原有的样式、分节、图片、关系都能保留。

import {
  CT,
  REL,
  allocPartName,
  appendRel,
  ensureDefaultExtension,
  ensureOverride,
  hasNoteReferences,
  nextRelId,
  openPackage,
  parseRels,
  readRels,
  relsPathFor,
  remapRelIds,
  splitTopLevelElements,
} from './ooxml.js'

// ---------------------------------------------------------------------------
// Word：目录（TOC）刷新
// ---------------------------------------------------------------------------

/**
 * 刷新 Word 的目录域。
 *
 * 做法：找到正文里第一个 TOC 域（`TOC \o ...`），把它的域结果清空并
 * 标记为「需要更新」（`w:dirty="true"`）。Word / ONLYOFFICE 打开文档时
 * 会依据标题样式重新生成条目与页码，所以这里不需要自己算页码。
 *
 * 文档里没有 TOC 域时，在正文开头插入一个可自动更新的 TOC 域。
 * @param {{ buffer: Buffer, jszip: any }} input
 * @returns {Promise<{ buffer: Buffer, action: 'refreshed' | 'inserted', entriesHint: number }>}
 */
export async function refreshToc(input) {
  const pkg = await openPackage(input.buffer, input.jszip)
  const path = 'word/document.xml'
  if (!pkg.has(path)) throw new Error('不是有效的 docx：缺少 word/document.xml')
  let xml = await pkg.read(path)

  // 统计可能成为目录条目的段落：优先按标题样式，其次按 outlineLvl（大纲级别）
  const headingCount = countHeadingParagraphs(xml)

  const existing = /<w:instrText[^>]*>\s*TOC\s+\\o[^<]*<\/w:instrText>/i.exec(xml)
  if (existing) {
    // 已有 TOC 域：把域结果区间清空并置 dirty，让打开时重算
    xml = markTocDirty(xml)
    pkg.write(path, xml)
    return { buffer: await pkg.save(), action: 'refreshed', entriesHint: headingCount }
  }

  // 没有 TOC：在 body 开头插入一个完整的 TOC 域
  xml = insertTocAtBodyStart(xml)
  pkg.write(path, xml)
  return { buffer: await pkg.save(), action: 'inserted', entriesHint: headingCount }
}

/**
 * 统计可能进目录的标题段落。
 *
 * 判定顺序：显式标题样式（Heading1-9 / 标题 1-9）→ 大纲级别（outlineLvl 0-8）。
 * 很多中文模板并不使用内置标题样式，而是直接给段落设大纲级别，
 * 这类段落同样会被 TOC 收录，所以两种都要算。
 * @param {string} xml
 */
function countHeadingParagraphs(xml) {
  const byStyle = (xml.match(/<w:pStyle w:val="(?:Heading|heading|标题)[\s]*[1-9][^"]*"/g) ?? []).length
  const byOutline = (xml.match(/<w:outlineLvl w:val="[0-8]"/g) ?? []).length
  return Math.max(byStyle, byOutline)
}

/**
 * 把已有的 TOC 域标记为 dirty，并清空上一次的条目缓存。
 * @param {string} xml
 */
function markTocDirty(xml) {
  // 1) 找到 TOC 域指令所在段落，标记 dirty
  let out = xml.replace(
    /(<w:fldChar\s+[^>]*w:fldCharType="begin")([^>]*>)/g,
    (whole, prefix, suffix) => (whole.includes('w:dirty=') ? whole : `${prefix} w:dirty="true"${suffix}`),
  )

  // 2) 清空 begin 与 separate 之间的域结果（那是上次生成的目录条目）
  out = out.replace(
    /(<w:instrText[^>]*>\s*TOC\s+\\o[^<]*<\/w:instrText>[\s\S]*?<w:fldChar\s+[^>]*w:fldCharType="separate"[^>]*\/>)([\s\S]*?)(<w:fldChar\s+[^>]*w:fldCharType="end")/,
    (whole, head, body, tail) => {
      if (!head) return whole
      // 只清掉域结果段落，保留分隔符结构
      const cleaned = body.replace(/<w:p\b[\s\S]*?<\/w:p>/g, '')
      return `${head}${cleaned}${tail}`
    },
  )
  return out
}

/**
 * 在正文最前面插入一个 TOC 域。
 * @param {string} xml
 */
function insertTocAtBodyStart(xml) {
  const tocParagraphs =
    '<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>目录</w:t></w:r></w:p>' +
    '<w:p>' +
    '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' +
    '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>' +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:t>右键此处选择「更新域」生成目录</w:t></w:r>' +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
    '</w:p>' +
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

  if (/<w:body[^>]*>/.test(xml)) {
    return xml.replace(/(<w:body[^>]*>)/, `$1${tocParagraphs}`)
  }
  throw new Error('document.xml 结构异常：找不到 w:body')
}

// ---------------------------------------------------------------------------
// Word：页眉页脚
// ---------------------------------------------------------------------------

/**
 * 设置 Word 的页眉页脚。
 *
 * - 支持固定文字、自动页码（PAGE 域）、首页独立
 * - 每节（w:sectPr）独立写入，因此多节文档每节都能生效
 * - 已有页眉/页脚时保留原文件、只替换引用
 * @param {{ buffer: Buffer, jszip: any, headerText?: string, footerText?: string, enablePageNumber?: boolean, differentFirstPage?: boolean }} input
 */
export async function setHeaderFooter(input) {
  const pkg = await openPackage(input.buffer, input.jszip)
  const docPath = 'word/document.xml'
  if (!pkg.has(docPath)) throw new Error('不是有效的 docx：缺少 word/document.xml')
  let xml = await pkg.read(docPath)

  const hasHeader = typeof input.headerText === 'string' && input.headerText.trim() !== ''
  const hasFooter = typeof input.footerText === 'string' && input.footerText.trim() !== ''
  const wantPageNumber = Boolean(input.enablePageNumber)
  if (!hasHeader && !hasFooter && !wantPageNumber) {
    throw new Error('至少要设置 headerText / footerText / enablePageNumber 其中一项')
  }

  /** @type {{ headerRid?: string, footerRid?: string, firstHeaderRid?: string, firstFooterRid?: string }} */
  const ids = {}
  const rels = await readRels(pkg, docPath)
  let relsXml = rels.xml
  let relsChanged = false

  if (hasHeader) {
    ids.headerRid = await installHeaderFooterPart(pkg, 'header', input.headerText, false)
    relsXml = appendRel(relsXml, { id: nextRelId(parseRels(relsXml)), type: REL.header, target: ids.headerRid.target })
    relsChanged = true
    if (input.differentFirstPage) {
      ids.firstHeaderRid = await installHeaderFooterPart(pkg, 'header', '', true)
      relsXml = appendRel(relsXml, { id: nextRelId(parseRels(relsXml)), type: REL.header, target: ids.firstHeaderRid.target })
    }
  }

  if (hasFooter || wantPageNumber) {
    const footerXml = buildFooterText(input.footerText, wantPageNumber)
    ids.footerRid = await installHeaderFooterPart(pkg, 'footer', footerXml, true)
    relsXml = appendRel(relsXml, { id: nextRelId(parseRels(relsXml)), type: REL.footer, target: ids.footerRid.target })
    relsChanged = true
    if (input.differentFirstPage) {
      ids.firstFooterRid = await installHeaderFooterPart(pkg, 'footer', '', true)
      relsXml = appendRel(relsXml, { id: nextRelId(parseRels(relsXml)), type: REL.footer, target: ids.firstFooterRid.target })
    }
  }

  if (relsChanged) pkg.write(rels.path, relsXml)

  // 在每个 sectPr 上挂引用
  const refs = []
  if (ids.headerRid) refs.push(`<w:headerReference w:type="default" r:id="${ids.headerRid.rid}"/>`)
  if (ids.firstHeaderRid) refs.push(`<w:headerReference w:type="first" r:id="${ids.firstHeaderRid.rid}"/>`)
  if (ids.footerRid) refs.push(`<w:footerReference w:type="default" r:id="${ids.footerRid.rid}"/>`)
  if (ids.firstFooterRid) refs.push(`<w:footerReference w:type="first" r:id="${ids.firstFooterRid.rid}"/>`)
  const titlePg = input.differentFirstPage ? '<w:titlePg/>' : ''
  const referenceXml = `${refs.join('')}${titlePg}`

  xml = applyToAllSections(xml, referenceXml, Boolean(input.differentFirstPage))
  pkg.write(docPath, xml)

  return {
    buffer: await pkg.save(),
    sections: (xml.match(/<w:sectPr/g) ?? []).length,
    hasHeader,
    hasFooter,
    pageNumber: wantPageNumber,
  }
}

/**
 * 把页眉/页脚引用写进每个 sectPr。
 * 已有同名引用时替换，避免出现两个 default 引用。
 * @param {string} xml
 * @param {string} referenceXml
 * @param {boolean} clearFirst
 */
function applyToAllSections(xml, referenceXml, clearFirst) {
  return xml.replace(/<w:sectPr\b([^>]*)>([\s\S]*?)<\/w:sectPr>/g, (whole, attrs, body) => {
    let inner = body
    // 去掉已有的默认页眉页脚引用（保留 first/even 的旧值，避免破坏原设计）
    inner = inner.replace(/<w:(header|footer)Reference\s+w:type="default"[^>]*\/>/g, '')
    if (clearFirst) inner = inner.replace(/<w:titlePg\/>/g, '')
    // 引用必须排在 sectPr 子元素最前面
    return `<w:sectPr${attrs}>${referenceXml}${inner}</w:sectPr>`
  })
}

/**
 * 生成页脚 XML：固定文字 + 可选自动页码。
 * @param {string | undefined} footerText
 * @param {boolean} pageNumber
 */
function buildFooterText(footerText, pageNumber) {
  const parts = []
  const text = (footerText ?? '').trim()
  if (text) parts.push(`<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`)
  if (pageNumber) {
    if (text) parts.push('<w:r><w:t xml:space="preserve">  </w:t></w:r>')
    parts.push('<w:r><w:t xml:space="preserve">第 </w:t></w:r>')
    parts.push('<w:r><w:fldChar w:fldCharType="begin"/></w:r>')
    parts.push('<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>')
    parts.push('<w:r><w:fldChar w:fldCharType="separate"/></w:r>')
    parts.push('<w:r><w:t>1</w:t></w:r>')
    parts.push('<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    parts.push('<w:r><w:t xml:space="preserve"> 页</w:t></w:r>')
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${parts.join('')}</w:p></w:ftr>`
  )
}

/**
 * 新建一个页眉/页脚部件并登记内容类型。
 * @param {any} pkg
 * @param {'header' | 'footer'} kind
 * @param {string} contentXml
 * @param {boolean} isRawXml 传入的已经是完整 XML（footer 走这个分支）
 */
async function installHeaderFooterPart(pkg, kind, contentXml, isRawXml) {
  const partName = allocPartName(pkg, 'word', kind, 'xml')
  const xml = isRawXml
    ? contentXml
    : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXmlText(contentXml)}</w:t></w:r></w:p></w:hdr>`
  pkg.write(partName, xml)
  await ensureOverride(pkg, `/${partName}`, kind === 'header' ? CT.header : CT.footer)
  return { rid: '', target: partName.replace(/^word\//, ''), partName }
}

/** @param {string} value */
function escapeXmlText(value) {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

// ---------------------------------------------------------------------------
// Word：合并
// ---------------------------------------------------------------------------

/**
 * 合并多个 docx。
 *
 * 保留每个源文档的正文（段落/表格/图片）、分节属性、页眉页脚；
 * 图片等关系会重新编号并复制到目标包，避免 rId 冲突。
 * 脚注/尾注引用无法跨文档搬运，遇到时给出提示。
 * @param {{ buffers: Buffer[], jszip: any }} input
 */
export async function mergeDocx(input) {
  if (input.buffers.length < 2) throw new Error('合并至少需要 2 个 docx 文档')
  const target = await openPackage(input.buffers[0], input.jszip)
  const docPath = 'word/document.xml'
  const baseXml = await target.read(docPath)
  const bodyMatch = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(baseXml)
  if (!bodyMatch) throw new Error('第一个文档结构异常：找不到 w:body')

  // 目标文档 body 内容（去掉结尾的 sectPr，它代表最后一节，要保留）
  const baseBody = bodyMatch[1]
  const trailingSectPr = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/g
  const sectionsInBase = baseBody.match(trailingSectPr) ?? []
  const lastSectPr = sectionsInBase.length ? sectionsInBase[sectionsInBase.length - 1] : ''
  const baseContent = lastSectPr ? baseBody.replace(trailingSectPr, '') : baseBody

  /** @type {string[]} */
  const chunks = [baseContent]
  /** @type {string[]} */
  const notes = []

  for (let index = 1; index < input.buffers.length; index += 1) {
    const source = await openPackage(input.buffers[index], input.jszip)
    if (!source.has(docPath)) {
      notes.push(`第 ${index + 1} 个文档缺少 word/document.xml，已跳过`)
      continue
    }
    const sourceXml = await source.read(docPath)
    const match = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(sourceXml)
    if (!match) {
      notes.push(`第 ${index + 1} 个文档结构异常，已跳过`)
      continue
    }
    if (hasNoteReferences(match[1])) {
      notes.push(`第 ${index + 1} 个文档含脚注/尾注引用，这些引用在合并后可能失效`)
    }

    // 搬运关系（图片/超链接等）
    const sourceRels = await readRels(source, docPath)
    const targetRels = await readRels(target, docPath)
    /** @type {Map<string, string>} */
    const ridMap = new Map()
    let targetRelsXml = targetRels.xml
    const usedTargets = new Set(targetRels.entries.map((rel) => rel.target))
    /** @type {Map<string, string>} */
    const copiedParts = new Map()

    for (const rel of sourceRels.entries) {
      if (rel.type === REL.header || rel.type === REL.footer) continue // 页眉页脚不跨文档搬运
      const newId = nextRelId(parseRels(targetRelsXml))
      if (rel.external) {
        ridMap.set(rel.id, newId)
        targetRelsXml = appendRel(targetRelsXml, { id: newId, type: rel.type, target: rel.target, external: true })
        continue
      }
      const sourcePart = `word/${rel.target}`.replace(/^word\/word\//, 'word/')
      if (!source.has(sourcePart)) continue

      let newTarget = copiedParts.get(sourcePart)
      if (!newTarget) {
        const ext = sourcePart.split('.').pop() ?? 'bin'
        const baseName = sourcePart.split('/').pop() ?? `part.${ext}`
        const stem = baseName.replace(/\.[^.]+$/, '').replace(/[^A-Za-z0-9_-]/g, '')
        const candidate = `media/merge-${index}-${stem}.${ext}`
        const finalTarget = usedTargets.has(candidate) ? `media/merge-${index}-${stem}-${Date.now()}.${ext}` : candidate
        usedTargets.add(finalTarget)
        const bytes = await source.readBinary(sourcePart)
        target.write(`word/${finalTarget}`, bytes)
        if (ext === 'png' || ext === 'jpg' || ext === 'jpeg' || ext === 'gif') {
          await ensureDefaultExtension(target, ext, `image/${ext === 'jpg' ? 'jpeg' : ext}`)
        }
        copiedParts.set(sourcePart, finalTarget)
        newTarget = finalTarget
      }
      ridMap.set(rel.id, newId)
      targetRelsXml = appendRel(targetRelsXml, { id: newId, type: rel.type, target: newTarget })
    }

    target.write(targetRels.path, targetRelsXml)

    // 分节：源文档的非末尾 sectPr 原样搬入，末尾 sectPr 丢弃（目标已有自己的结尾节）
    const sourceSections = match[1].match(trailingSectPr) ?? []
    let content = sourceSections.length
      ? match[1].replace(trailingSectPr, '')
      : match[1]
    if (sourceSections.length) {
      // 把源文档的中间分节属性搬进来，保证分节结构延续
      for (const section of sourceSections.slice(0, -1)) {
        content = `${content}${section}`
      }
    }
    chunks.push(remapRelIds(content, ridMap))
  }

  const mergedBody = `${chunks.join('')}${lastSectPr}`
  target.write(docPath, baseXml.replace(bodyMatch[0], `<w:body>${mergedBody}</w:body>`))

  return { buffer: await target.save(), notes }
}

// ---------------------------------------------------------------------------
// Word：按分节拆分
// ---------------------------------------------------------------------------

/**
 * 按分节符拆分 docx。
 * @param {{ buffer: Buffer, jszip: any }} input
 */
export async function splitDocxBySection(input) {
  const pkg = await openPackage(input.buffer, input.jszip)
  const docPath = 'word/document.xml'
  if (!pkg.has(docPath)) throw new Error('不是有效的 docx：缺少 word/document.xml')
  const xml = await pkg.read(docPath)
  const bodyMatch = /<w:body[^>]*>([\s\S]*)<\/w:body>/.exec(xml)
  if (!bodyMatch) throw new Error('document.xml 结构异常：找不到 w:body')
  const body = bodyMatch[1]
  const finalSect = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>$/g

  /** @type {Array<{ content: string, sectPr: string }>} */
  const sections = []
  /** @type {string[]} */
  let current = []
  // 逐段扫描：段落里带 sectPr 的，就是一个分节结束
  const elements = splitTopLevelElements(body.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>\s*$/g, ''))
  let pendingSect = ''
  for (const element of elements) {
    current.push(element.xml)
    const sectMatch = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(element.xml)
    if (sectMatch) {
      pendingSect = sectMatch[0]
      sections.push({ content: current.join(''), sectPr: pendingSect })
      current = []
      pendingSect = ''
    }
  }
  const tailSect = /<w:sectPr\b[^>]*\/>|<w:sectPr\b[\s\S]*?<\/w:sectPr>/.exec(body)
  if (current.length) {
    sections.push({ content: current.join(''), sectPr: tailSect ? tailSect[0] : '' })
  }
  if (!sections.length) throw new Error('文档没有可拆分的内容')
  // 只有一个分节时，按段落数量兜底拆分，避免「拆了但等于没拆」
  if (sections.length === 1) {
    const paras = splitTopLevelElements(body).filter((element) => element.tag === 'w:p')
    if (paras.length < 2) throw new Error('文档只有一个分节且段落不足，无法拆分')
    const half = Math.ceil(paras.length / 2)
    return {
      buffers: [
        await rebuildDocx(pkg, xml, bodyMatch[0], paras.slice(0, half).map((p) => p.xml).join(''), tailSect ? tailSect[0] : '', input.jszip),
        await rebuildDocx(pkg, xml, bodyMatch[0], paras.slice(half).map((p) => p.xml).join(''), tailSect ? tailSect[0] : '', input.jszip),
      ],
      mode: 'section-fallback',
      notes: ['文档只有一个分节，已按段落数量均分为 2 份'],
    }
  }

  /** @type {Buffer[]} */
  const buffers = []
  for (const section of sections) {
    buffers.push(await rebuildDocx(pkg, xml, bodyMatch[0], section.content, section.sectPr, input.jszip))
  }
  return { buffers, mode: 'section', notes: [] }
}

/**
 * 用给定 body 内容重建一份 docx（复用同一个包的其它部件）。
 * @param {any} pkg
 * @param {string} originalXml
 * @param {string} originalBody 原 body 整段（含标签）
 * @param {string} content 新 body 内部内容
 * @param {string} sectPr 结尾分节属性
 * @param {any} jszip 用于从字节重建包
 */
async function rebuildDocx(pkg, originalXml, originalBody, content, sectPr, jszip) {
  const rebuilt = originalXml.replace(originalBody, `<w:body>${content}${sectPr}</w:body>`)
  // 基于当前包的字节克隆一份新包，再替换主部件：这样其它部件（图片/样式/页眉）都保留
  const bytes = await pkg.save()
  const cloneZip = await jszip.loadAsync(bytes)
  cloneZip.file('word/document.xml', rebuilt)
  return /** @type {Buffer} */ (await cloneZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}

// ---------------------------------------------------------------------------
// PPT：按页拆分
// ---------------------------------------------------------------------------

/**
 * 按页拆分 pptx。
 * @param {{ buffer: Buffer, jszip: any }} input
 */
export async function splitPptxBySlide(input) {
  const pkg = await openPackage(input.buffer, input.jszip)
  const presPath = 'ppt/presentation.xml'
  if (!pkg.has(presPath)) throw new Error('不是有效的 pptx：缺少 ppt/presentation.xml')
  const presXml = await pkg.read(presPath)
  const rels = await readRels(pkg, presPath)
  const slideRels = rels.entries.filter((rel) => rel.type === REL.slide)
  if (!slideRels.length) throw new Error('pptx 里没有幻灯片')

  const slideParts = slideRels
    .map((rel) => ({ rid: rel.id, part: `ppt/${rel.target}`.replace(/^ppt\/ppt\//, 'ppt/') }))
    .filter((item) => pkg.has(item.part))
    .sort((a, b) => {
      const na = Number(/slide(\d+)\.xml$/.exec(a.part)?.[1] ?? 0)
      const nb = Number(/slide(\d+)\.xml$/.exec(b.part)?.[1] ?? 0)
      return na - nb
    })

  /** @type {Buffer[]} */
  const buffers = []
  const packageBytes = await pkg.save()

  // 先建立「备注页 → 它引用的幻灯片」映射，删除时才能连带清理，
  // 否则残留的备注页会引用已删除的幻灯片，PowerPoint 打开会报修复提示。
  /** @type {Map<string, string>} */
  const noteOwner = new Map()
  for (const name of pkg.names()) {
    if (!/^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)) continue
    const relsName = relsPathFor(name)
    if (!pkg.has(relsName)) continue
    const relsXml = await pkg.read(relsName)
    const slideRef = /Target="\.\.\/slides\/(slide\d+\.xml)"/.exec(relsXml)
    if (slideRef) noteOwner.set(name, `ppt/slides/${slideRef[1]}`)
  }

  for (const slide of slideParts) {
    const cloneZip = await input.jszip.loadAsync(packageBytes)
    // 删掉其它幻灯片部件及其 rels，以及它们的备注页（否则备注页会引用已删除的幻灯片）
    for (const other of slideParts) {
      if (other.part === slide.part) continue
      cloneZip.remove(other.part)
      cloneZip.remove(relsPathFor(other.part))
      // 该幻灯片对应的备注页：用预建的映射表精确删除，
      // 否则残留备注页会引用已删除的幻灯片，PowerPoint 会提示修复。
      for (const [notePart, ownerSlide] of noteOwner) {
        if (ownerSlide !== other.part) continue
        cloneZip.remove(notePart)
        cloneZip.remove(relsPathFor(notePart))
        if (cloneZip.file('[Content_Types].xml')) {
          const typesXml = await cloneZip.file('[Content_Types].xml').async('string')
          cloneZip.file(
            '[Content_Types].xml',
            typesXml.replace(new RegExp(`<Override\\s+PartName="/${notePart}"[^/>]*/>`, 'g'), ''),
          )
        }
      }
    }
    // presentation.xml 只保留这一页
    const onePagePres = presXml
      .replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst><p:sldId id="256" r:id="${slide.rid}"/></p:sldIdLst>`)
    cloneZip.file(presPath, onePagePres)
    // presentation 的 rels 只保留 slide 相关 + 其它必要关系
    const keptRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      rels.entries
        .filter((rel) => rel.type !== REL.slide || rel.id === slide.rid)
        .map(
          (rel) =>
            `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"${rel.external ? ' TargetMode="External"' : ''}/>`,
        )
        .join('') +
      '</Relationships>'
    cloneZip.file(relsPathFor(presPath), keptRels)
    // 内容类型里去掉被删部件的 Override
    if (cloneZip.file('[Content_Types].xml')) {
      const typesXml = await cloneZip.file('[Content_Types].xml').async('string')
      const removed = new Set(slideParts.filter((other) => other.part !== slide.part).map((other) => `/${other.part}`))
      const cleaned = typesXml.replace(/<Override\s+PartName="([^"]+)"[^/>]*\/>/g, (whole, partName) =>
        removed.has(partName) ? '' : whole,
      )
      cloneZip.file('[Content_Types].xml', cleaned)
    }
    buffers.push(/** @type {Buffer} */ (await cloneZip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })))
  }
  return { buffers, slides: slideParts.length, notes: [] }
}

// ---------------------------------------------------------------------------
// PPT：合并
// ---------------------------------------------------------------------------

/**
 * 合并多个 pptx：把后续演示文稿的幻灯片追加到第一个演示文稿末尾。
 *
 * 幻灯片可能引用各自的版式/母版/主题，这里把整条依赖链（slide → layout →
 * master → theme）连同图片等媒体一起复制进目标包并重编号关系。
 * @param {{ buffers: Buffer[], jszip: any }} input
 */
export async function mergePptx(input) {
  if (input.buffers.length < 2) throw new Error('合并至少需要 2 个 pptx 文档')
  const target = await openPackage(input.buffers[0], input.jszip)
  const presPath = 'ppt/presentation.xml'
  if (!target.has(presPath)) throw new Error('第一个文档不是有效的 pptx')
  const presXml = await target.read(presPath)
  const presRels = await readRels(target, presPath)
  let presRelsXml = presRels.xml

  const slideIdList = /<p:sldIdLst>([\s\S]*?)<\/p:sldIdLst>/.exec(presXml)
  /** @type {string[]} */
  const appendedIds = []
  let maxSlideId = 255
  if (slideIdList) {
    for (const match of slideIdList[1].matchAll(/<p:sldId\s+id="(\d+)"/g)) {
      maxSlideId = Math.max(maxSlideId, Number(match[1]))
    }
  }

  /** @type {string[]} */
  const notes = []
  /** @type {Set<string>} */
  const usedTargets = new Set(target.names())

  /** 递归复制一个部件及其关系闭包，返回新的部件路径。 */
  const copyPart = async (sourcePkg, sourcePart, depth = 0) => {
    if (depth > 6) return undefined
    const ext = sourcePart.split('.').pop() ?? 'xml'
    const baseName = (sourcePart.split('/').pop() ?? 'part').replace(/\.[^.]+$/, '')
    const dir = sourcePart.slice(0, sourcePart.lastIndexOf('/'))
    let candidate = `${dir}/${baseName}.${ext}`
    if (usedTargets.has(candidate)) {
      candidate = `${dir}/${baseName}-merged${maxSlideId}.${ext}`
      let n = 1
      while (usedTargets.has(candidate)) {
        candidate = `${dir}/${baseName}-merged${maxSlideId}-${n}.${ext}`
        n += 1
      }
    }
    usedTargets.add(candidate)
    target.write(candidate, await sourcePkg.readBinary(sourcePart))

    // 复制它的 rels，并把 target 路径按新路径重写
    const sourceRels = await readRels(sourcePkg, sourcePart)
    if (!sourceRels.entries.length) return candidate
    /** @type {Map<string, string>} */
    const ridMap = new Map()
    let targetRelsXml = ''
    for (const rel of sourceRels.entries) {
      if (rel.external) {
        const newId = `rId${ridMap.size + 1}`
        ridMap.set(rel.id, newId)
        targetRelsXml = appendRel(targetRelsXml || emptyRels(), { id: newId, type: rel.type, target: rel.target, external: true })
        continue
      }
      const childSource = `${dir}/${rel.target}`.replace(/^([^/]+)\/\.\.\//, '')
      if (!sourcePkg.has(childSource)) continue
      const childTarget = await copyPart(sourcePkg, childSource, depth + 1)
      if (!childTarget) continue
      const newId = `rId${ridMap.size + 1}`
      ridMap.set(rel.id, newId)
      const relative = childTarget.startsWith(`${dir}/`) ? childTarget.slice(dir.length + 1) : `../${childTarget}`
      targetRelsXml = appendRel(targetRelsXml || emptyRels(), { id: newId, type: rel.type, target: relative })
    }
    // 重写复制过来的部件 XML 里的 rId
    if (ridMap.size) {
      const xml = await target.read(candidate)
      target.write(candidate, remapRelIds(xml, ridMap))
      target.write(relsPathFor(candidate), targetRelsXml)
    }
    return candidate
  }

  for (let index = 1; index < input.buffers.length; index += 1) {
    const source = await openPackage(input.buffers[index], input.jszip)
    if (!source.has(presPath)) {
      notes.push(`第 ${index + 1} 个文档不是 pptx，已跳过`)
      continue
    }
    const sourcePres = await source.read(presPath)
    const sourcePresRels = await readRels(source, presPath)
    const sourceSlides = sourcePresRels.entries
      .filter((rel) => rel.type === REL.slide)
      .map((rel) => ({ rid: rel.id, part: `ppt/${rel.target}`.replace(/^ppt\/ppt\//, 'ppt/') }))
      .sort((a, b) => {
        const na = Number(/slide(\d+)\.xml$/.exec(a.part)?.[1] ?? 0)
        const nb = Number(/slide(\d+)\.xml$/.exec(b.part)?.[1] ?? 0)
        return na - nb
      })

    for (const slide of sourceSlides) {
      if (!source.has(slide.part)) continue
      const copied = await copyPart(source, slide.part)
      if (!copied) continue
      maxSlideId += 1
      const newRid = nextRelId(parseRels(presRelsXml))
      presRelsXml = appendRel(presRelsXml, { id: newRid, type: REL.slide, target: copied.replace(/^ppt\//, '') })
      appendedIds.push(`<p:sldId id="${maxSlideId}" r:id="${newRid}"/>`)
    }
    void sourcePres
  }

  if (!appendedIds.length) throw new Error('没有可追加的幻灯片')

  const newPresXml = slideIdList
    ? presXml.replace(/(<p:sldIdLst>)([\s\S]*?)(<\/p:sldIdLst>)/, (whole, open, body, close) => `${open}${body}${appendedIds.join('')}${close}`)
    : presXml.replace(/(<p:sldMasterIdLst[\s\S]*?<\/p:sldMasterIdLst>)/, `$1<p:sldIdLst>${appendedIds.join('')}</p:sldIdLst>`)
  target.write(presPath, newPresXml)
  target.write(presRels.path, presRelsXml)

  return { buffer: await target.save(), appended: appendedIds.length, notes }
}

/** 空 rels 骨架。 */
function emptyRels() {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  )
}

// ---------------------------------------------------------------------------
// PPT：页码刷新（PPT 没有 Word 那种 TOC 域）
// ---------------------------------------------------------------------------

/**
 * 刷新 pptx 的页码。
 *
 * PowerPoint 的页码是幻灯片里的 `<p:sldNum>` 占位符（由母版定义），
 * 这里做两件真实有效的事：
 *   1. 给每页的 sldNum 占位符写入当前页序（第 N 页）；
 *   2. 把幻灯片正文里形如「第 X 页 / X/Y 页」的文字替换为真实页序。
 * 没有占位符也没有页码文字时，只返回统计信息，不做破坏性改写。
 * @param {{ buffer: Buffer, jszip: any }} input
 */
export async function renumberSlides(input) {
  const pkg = await openPackage(input.buffer, input.jszip)
  const presPath = 'ppt/presentation.xml'
  if (!pkg.has(presPath)) throw new Error('不是有效的 pptx：缺少 ppt/presentation.xml')
  const rels = await readRels(pkg, presPath)
  const slideParts = rels.entries
    .filter((rel) => rel.type === REL.slide)
    .map((rel) => `ppt/${rel.target}`.replace(/^ppt\/ppt\//, 'ppt/'))
    .filter((part) => pkg.has(part))
    .sort((a, b) => {
      const na = Number(/slide(\d+)\.xml$/.exec(a)?.[1] ?? 0)
      const nb = Number(/slide(\d+)\.xml$/.exec(b)?.[1] ?? 0)
      return na - nb
    })

  let updated = 0
  const total = slideParts.length
  for (const [index, part] of slideParts.entries()) {
    const xml = await pkg.read(part)
    let next = xml

    // 1) 页码占位符：把 <p:sldNum> 里的文本节点换成当前页序
    next = next.replace(
      /(<p:ph[^>]*type="sldNum"[^>]*\/>[\s\S]*?<a:t>)([^<]*)(<\/a:t>)/g,
      (whole, head, _old, tail) => {
        updated += 1
        return `${head}${index + 1}${tail}`
      }
    )

    // 2) 正文里的页码文字占位符：「第 X 页」「X / N 页」「Page X」
    const patterns = [
      { re: /第\s*[XxＸ]\s*页/g, value: `第 ${index + 1} 页` },
      { re: /第\s*\d+\s*页\s*[／/]\s*共\s*\d+\s*页/g, value: `第 ${index + 1} 页 / 共 ${total} 页` },
      { re: /\bPage\s*[Xx]\s*(of\s*\d+)?/g, value: `Page ${index + 1} of ${total}` },
    ]
    for (const pattern of patterns) {
      if (pattern.re.test(next)) {
        updated += 1
        next = next.replace(pattern.re, pattern.value)
      }
    }

    if (next !== xml) pkg.write(part, next)
  }

  return { buffer: await pkg.save(), slides: total, updated }
}

// ---------------------------------------------------------------------------
// 辅助：判断包类型
// ---------------------------------------------------------------------------

/**
 * 判断 buffer 是 docx 还是 pptx（按主部件存在性）。
 * @param {Buffer} buffer
 * @param {any} jszip
 * @returns {Promise<'docx' | 'pptx' | 'unknown'>}
 */
export async function detectKind(buffer, jszip) {
  let zip
  try {
    zip = await jszip.loadAsync(buffer)
  } catch (error) {
    throw new Error(
      `文件不是有效的 Office 文档（zip 结构损坏或不是 OOXML）：${error instanceof Error ? error.message : String(error)}`,
    )
  }
  if (zip.file('word/document.xml')) return 'docx'
  if (zip.file('ppt/presentation.xml')) return 'pptx'
  return 'unknown'
}
