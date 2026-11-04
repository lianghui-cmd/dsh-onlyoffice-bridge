// @ts-check
// OOXML 包操作原语：zip 读写、关系（rels）增删、内容类型登记、部件名分配。
//
// docx / pptx 本质都是 zip 包 + 一批 XML 部件，所有排版类操作都建立在这层之上。
// 这里刻意不引入 XML 解析器：只做「定位 + 改写 + 校验」，避免把整棵树序列化后
// 丢掉 Word 自己写入的命名空间声明（那是导致文档打不开的常见原因）。

/** @typedef {import('jszip')} JSZip */

const CONTENT_TYPES = '[Content_Types].xml'

/** 关系类型常量。 */
export const REL = {
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  styles: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
  footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
  footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
  endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
  settings: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  fontTable: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable',
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
}

/** 各部件的内容类型。 */
export const CT = {
  header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
  footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
  document: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  presentation: 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
  slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
  slideLayout: 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml',
  slideMaster: 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
  notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
}

/**
 * 打开一个 OOXML 包。
 * @param {Buffer} buffer
 * @param {Record<string, any>} jszip
 */
export async function openPackage(buffer, jszip) {
  let zip
  try {
    zip = await jszip.loadAsync(buffer)
  } catch (error) {
    // 把 jszip 的底层报错翻译成用户能理解的提示：
    // 常见原因是文件损坏、只写了一半，或根本不是 docx/pptx。
    throw new Error(
      `文件不是有效的 Office 文档（zip 结构损坏或不是 OOXML 格式）：${error instanceof Error ? error.message : String(error)}`,
    )
  }
  return {
    zip,
    /** @type {string[]} */
    names() {
      return Object.keys(zip.files).filter((name) => !zip.files[name].dir)
    },
    /** @param {string} name */
    has(name) {
      return zip.file(name) !== null && zip.file(name) !== undefined
    },
    /** @param {string} name @returns {Promise<string>} */
    async read(name) {
      const entry = zip.file(name)
      if (!entry) throw new Error(`部件不存在：${name}`)
      return entry.async('string')
    },
    /** @param {string} name @returns {Promise<Buffer>} */
    async readBinary(name) {
      const entry = zip.file(name)
      if (!entry) throw new Error(`部件不存在：${name}`)
      return /** @type {Buffer} */ (await entry.async('nodebuffer'))
    },
    /** @param {string} name @param {string | Buffer} content */
    write(name, content) {
      zip.file(name, content)
    },
    /** @param {string} name */
    remove(name) {
      zip.remove(name)
    },
    /** @returns {Promise<Buffer>} */
    async save() {
      return /** @type {Buffer} */ (
        await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } })
      )
    },
  }
}

// ---------------------------------------------------------------------------
// 内容类型
// ---------------------------------------------------------------------------

/**
 * 确保 [Content_Types].xml 里登记了某个部件的 Override。
 * @param {any} pkg openPackage 的返回值
 * @param {string} partName 形如 /word/header1.xml
 * @param {string} contentType
 */
export async function ensureOverride(pkg, partName, contentType) {
  if (!pkg.has(CONTENT_TYPES)) return
  const xml = await pkg.read(CONTENT_TYPES)
  const normalized = partName.startsWith('/') ? partName : `/${partName}`
  if (new RegExp(`PartName="${escapeRe(normalized)}"`, 'i').test(xml)) return
  const override = `<Override PartName="${normalized}" ContentType="${contentType}"/>`
  pkg.write(CONTENT_TYPES, xml.replace('</Types>', `${override}</Types>`))
}

/**
 * 确保登记了某个扩展名的 Default 内容类型。
 * @param {any} pkg
 * @param {string} ext 不带点，如 png
 * @param {string} contentType
 */
export async function ensureDefaultExtension(pkg, ext, contentType) {
  if (!pkg.has(CONTENT_TYPES)) return
  const xml = await pkg.read(CONTENT_TYPES)
  if (new RegExp(`Extension="${escapeRe(ext)}"`, 'i').test(xml)) return
  pkg.write(
    CONTENT_TYPES,
    xml.replace(/(<Types[^>]*>)/, `$1<Default Extension="${ext}" ContentType="${contentType}"/>`),
  )
}

/**
 * 移除某个部件的 Override 登记。
 * @param {any} pkg
 * @param {string} partName
 */
export async function removeOverride(pkg, partName) {
  if (!pkg.has(CONTENT_TYPES)) return
  const xml = await pkg.read(CONTENT_TYPES)
  const normalized = partName.startsWith('/') ? partName : `/${partName}`
  pkg.write(
    CONTENT_TYPES,
    xml.replace(new RegExp(`<Override\\s+PartName="${escapeRe(normalized)}"[^/>]*/>`, 'gi'), ''),
  )
}

// ---------------------------------------------------------------------------
// 关系（.rels）
// ---------------------------------------------------------------------------

/**
 * 主部件的 rels 路径：word/document.xml → word/_rels/document.xml.rels
 * @param {string} partPath
 */
export function relsPathFor(partPath) {
  const slash = partPath.lastIndexOf('/')
  const dir = partPath.slice(0, slash)
  const file = partPath.slice(slash + 1)
  return `${dir}/_rels/${file}.rels`
}

/**
 * 解析 rels 里的条目。
 * @param {string} relsXml
 * @returns {Array<{ id: string, type: string, target: string, external: boolean }>}
 */
export function parseRels(relsXml) {
  /** @type {Array<{ id: string, type: string, target: string, external: boolean }>} */
  const out = []
  const pattern = /<Relationship\s+([^>]*?)\/>/g
  let match
  while ((match = pattern.exec(relsXml)) !== null) {
    const attrs = match[1]
    const id = attrOf(attrs, 'Id')
    const type = attrOf(attrs, 'Type')
    const target = attrOf(attrs, 'Target')
    const mode = attrOf(attrs, 'TargetMode')
    if (!id) continue
    out.push({ id, type, target, external: mode === 'External' })
  }
  return out
}

/**
 * 往 rels 里追加一条关系。
 * @param {string} relsXml
 * @param {{ id: string, type: string, target: string, external?: boolean }} rel
 */
export function appendRel(relsXml, rel) {
  const element =
    `<Relationship Id="${rel.id}" Type="${rel.type}" Target="${escapeAttr(rel.target)}"` +
    `${rel.external ? ' TargetMode="External"' : ''}/>`
  if (/<\/Relationships>/.test(relsXml)) {
    return relsXml.replace('</Relationships>', `${element}</Relationships>`)
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `${element}</Relationships>`
  )
}

/**
 * 分配一个未被占用的 rId。
 * @param {Array<{ id: string }>} rels
 */
export function nextRelId(rels) {
  let max = 0
  for (const rel of rels) {
    const match = /rId(\d+)/.exec(rel.id)
    if (match) max = Math.max(max, Number(match[1]))
  }
  return `rId${max + 1}`
}

/**
 * 读取部件对应的 rels（不存在时返回空骨架）。
 * @param {any} pkg
 * @param {string} partPath
 */
export async function readRels(pkg, partPath) {
  const path = relsPathFor(partPath)
  if (!pkg.has(path)) {
    return {
      path,
      xml:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
      entries: [],
    }
  }
  const xml = await pkg.read(path)
  return { path, xml, entries: parseRels(xml) }
}

// ---------------------------------------------------------------------------
// 部件名分配
// ---------------------------------------------------------------------------

/**
 * 按前缀分配一个唯一部件名，如 word/header3.xml。
 * @param {any} pkg
 * @param {string} dir 例如 word
 * @param {string} base 例如 header
 * @param {string} ext 例如 xml
 */
export function allocPartName(pkg, dir, base, ext) {
  const existing = new Set(pkg.names())
  for (let index = 1; index < 9999; index += 1) {
    const candidate = `${dir}/${base}${index}.${ext}`
    if (!existing.has(candidate)) return candidate
  }
  throw new Error(`无法分配部件名：${dir}/${base}*.${ext}`)
}

// ---------------------------------------------------------------------------
// 小工具
// ---------------------------------------------------------------------------

/**
 * 从属性串里取一个属性值。
 * @param {string} attrs
 * @param {string} name
 */
export function attrOf(attrs, name) {
  const pattern = new RegExp(`(?:^|\\s)${name}="([^"]*)"`)
  const match = pattern.exec(attrs)
  return match ? match[1] : undefined
}

/** @param {string} value */
export function escapeAttr(value) {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** @param {string} value */
export function escapeRe(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * 把 XML 文本片段按顶层元素切成数组（不解析树，只按标签深度切分）。
 * 用于从 document.xml 的 body 里取出段落/表格等顶层元素。
 * @param {string} bodyXml body 内部内容
 * @returns {Array<{ xml: string, tag: string }>}
 */
export function splitTopLevelElements(bodyXml) {
  /** @type {Array<{ xml: string, tag: string }>} */
  const out = []
  let index = 0
  while (index < bodyXml.length) {
    const open = /<([A-Za-z0-9:]+)[^>]*?(\/?)>/.exec(bodyXml.slice(index))
    if (!open) break
    const start = index + open.index
    const tag = open[1]
    if (open[2] === '/') {
      out.push({ xml: bodyXml.slice(start, start + open[0].length), tag })
      index = start + open[0].length
      continue
    }
    // 找到匹配的闭合标签（同名标签计数）
    const pattern = new RegExp(`<${escapeRe(tag)}(?:\\s[^>]*)?>|</${escapeRe(tag)}>`, 'g')
    pattern.lastIndex = start
    let depth = 0
    let end = -1
    let match
    while ((match = pattern.exec(bodyXml)) !== null) {
      if (match[0].startsWith('</')) depth -= 1
      else depth += 1
      if (depth === 0) {
        end = match.index + match[0].length
        break
      }
    }
    if (end < 0) {
      // 自闭合或被截断，保守收尾
      const selfClose = /\/>/.exec(open[0])
      end = start + (selfClose ? open[0].length : open[0].length)
    }
    out.push({ xml: bodyXml.slice(start, end), tag })
    index = end
  }
  return out
}

/**
 * 判断一段 XML 里是否引用了脚注/尾注（合并时需要提示这类内容无法搬运）。
 * @param {string} xml
 */
export function hasNoteReferences(xml) {
  return /<w:footnoteReference|<w:endnoteReference/.test(xml)
}

/**
 * 重写 XML 里引用的 rId 映射。
 * @param {string} xml
 * @param {Map<string, string>} map 旧 rId → 新 rId
 */
export function remapRelIds(xml, map) {
  if (!map.size) return xml
  return xml.replace(/(\s(?:r:id|r:embed|r:link)=")([^"]+)(")/g, (whole, prefix, id, suffix) => {
    const next = map.get(id)
    return next ? `${prefix}${next}${suffix}` : whole
  })
}
