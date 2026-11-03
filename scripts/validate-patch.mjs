#!/usr/bin/env node
// 校验 cordis.patch.yml 是否是 DSH 能接受的 bundle patch（不依赖第三方库）。
//
// 只做结构校验，不追求完整 YAML 规范。核心规则：
//   1. 去掉 BOM；忽略空行与 # 注释
//   2. 顶层必须是 "- " 开头的序列
//   3. 每项的第一行必须是 "key:"（通常是 insert:）
//   4. insert 列表项必须同时有 id 与 name，且 name 与 package.json 一致
//
// 专门拦截历史事故：行首裸 "-" 后换行会被 YAML 当作文档分隔符，
// 安装时报 "end of the stream or a document separator is expected"。
//
// 用法：node scripts/validate-patch.mjs

import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const PATCH = path.join(ROOT, 'cordis.patch.yml')
const MANIFEST = path.join(ROOT, 'package.json')

/** @type {string[]} */
const errors = []

/** 收集有效行（去 BOM、去注释、去空行），保留原始行号与缩进。 */
function tokenize(rawText) {
  const text = rawText.charCodeAt(0) === 0xfeff ? rawText.slice(1) : rawText
  /** @type {Array<{ indent: number, text: string, line: number }>} */
  const tokens = []
  const lines = text.split(/\r?\n/)
  for (let index = 0; index < lines.length; index += 1) {
    const raw = lines[index]
    if (!raw.trim()) continue
    // YAML 注释：# 在行首或前面是空白
    if (/^\s*#/.test(raw)) continue
    const indent = (raw.match(/^ */) ?? [''])[0].length
    let content = raw.slice(indent).replace(/\s+$/, '')
    content = stripInlineComment(content)
    if (!content.trim()) continue
    if (content === '-') {
      errors.push(
        `第 ${index + 1} 行：行首裸 "-" 后直接换行，YAML 会把它当作本文档结束符。` +
          '安装会报 "end of the stream or a document separator is expected"，请写成 "- insert:"。',
      )
      continue
    }
    tokens.push({ indent, text: content, line: index + 1 })
  }
  return tokens
}

/** 去掉行内注释（尊重引号）。 */
function stripInlineComment(text) {
  let inSingle = false
  let inDouble = false
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i]
    if (char === "'" && !inDouble) inSingle = !inSingle
    else if (char === '"' && !inSingle) inDouble = !inDouble
    else if (char === '#' && !inSingle && !inDouble && (i === 0 || /\s/.test(text[i - 1]))) {
      return text.slice(0, i).trimEnd()
    }
  }
  return text
}

/** @param {string} value */
function scalar(value) {
  const text = value.trim()
  if ((text.startsWith("'") && text.endsWith("'")) || (text.startsWith('"') && text.endsWith('"'))) {
    return text.slice(1, -1)
  }
  return text
}

function main() {
  let manifest
  try {
    manifest = JSON.parse(fs.readFileSync(MANIFEST, 'utf8'))
  } catch (error) {
    errors.push(`package.json 无法解析：${error instanceof Error ? error.message : String(error)}`)
  }

  let patchText
  try {
    patchText = fs.readFileSync(PATCH, 'utf8')
  } catch {
    patchText = undefined
    errors.push('缺少 cordis.patch.yml')
  }

  const tokens = typeof patchText === 'string' ? tokenize(patchText) : []
  if (typeof patchText === 'string') {
    if (!tokens.length) {
      errors.push('cordis.patch.yml 没有任何有效内容（不能为空或只有注释）')
    } else {
      if (!/^-\s/.test(tokens[0].text) || tokens[0].indent !== 0) {
        errors.push(`第 ${tokens[0].line} 行：顶层必须是序列（每项以 "- " 开头）`)
      }
      if (!/^-\s*insert:\s*$/.test(tokens[0].text)) {
        errors.push(`第 ${tokens[0].line} 行：第一项必须是 "- insert:"（DSH 靠它插入插件行）`)
      }

      // 收集 insert 下的条目：形如 "- id: xxx"，随后缩进的 "name: xxx"
      /** @type {Array<{ id?: string, name?: string, line: number }>} */
      const entries = []
      /** @type {any} */
      let current
      for (const token of tokens.slice(1)) {
        const itemMatch = token.text.match(/^-\s+([\w.$-]+):\s*(.*)$/)
        if (itemMatch) {
          current = { line: token.line }
          entries.push(current)
          current[itemMatch[1]] = scalar(itemMatch[2])
          continue
        }
        const pair = token.text.match(/^([\w.$-]+):\s*(.*)$/)
        if (pair && current) current[pair[1]] = scalar(pair[2])
      }

      if (!entries.length) errors.push('insert 下没有任何条目')
      for (const [index, entry] of entries.entries()) {
        if (!entry.id) errors.push(`insert[${index}]（第 ${entry.line} 行）缺少 id`)
        if (!entry.name) errors.push(`insert[${index}]（第 ${entry.line} 行）缺少 name`)
        if (manifest?.name && entry.name && entry.name !== manifest.name) {
          errors.push(
            `insert[${index}].name (${entry.name}) 与 package.json name (${manifest.name}) 不一致，DSH 会解析不到这个包`,
          )
        }
      }
    }
  }

  if (!manifest?.dsh?.bundle?.patch) {
    errors.push('package.json 缺少 dsh.bundle.patch，DSH 不会把它当作 bundle')
  } else if (manifest.dsh.bundle.patch !== './cordis.patch.yml') {
    errors.push(`dsh.bundle.patch 应指向 ./cordis.patch.yml，当前为 ${manifest.dsh.bundle.patch}`)
  }

  for (const file of ['index.js', 'lib/definitions.js', 'locale/zh.json', 'locale/en.json', 'icon.svg']) {
    if (!fs.existsSync(path.join(ROOT, file))) errors.push(`缺少文件 ${file}`)
  }

  if (errors.length) {
    console.error('cordis.patch.yml 校验失败：')
    for (const error of errors) console.error(`  - ${error}`)
    process.exit(1)
  }
  console.log('cordis.patch.yml 校验通过：insert 结构与 package.json 一致')
}

main()
