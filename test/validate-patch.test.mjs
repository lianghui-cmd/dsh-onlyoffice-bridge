// validate-patch.mjs 的用例矩阵（用 node 写，避开 PowerShell 编码干扰）。
import { execFileSync } from 'node:child_process'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const PATCH = path.join(ROOT, 'cordis.patch.yml')
const VALIDATOR = path.join(ROOT, 'scripts', 'validate-patch.mjs')

const original = fs.readFileSync(PATCH, 'utf8')
const QUOTE = String.fromCharCode(39)

/** @param {...string} lines */
const lines = (...lines) => lines.join('\n')
/** @param {...string} parts */
const join = (...parts) => parts.join('\n') + '\n'

const CASES = [
  { name: '当前文件（合法）', content: original, expect: 'pass' },
  {
    name: '嵌套缩进写法（合法）',
    content: join('- insert:', '    - id: onlyoffice-bridge', `      name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE}`),
    expect: 'pass',
  },
  {
    name: '同缩进写法（合法）',
    content: join('- insert:', '  - id: onlyoffice-bridge', `    name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE}`),
    expect: 'pass',
  },
  {
    name: '同缩进 + BOM（合法）',
    content: '\uFEFF' + join('- insert:', '- id: onlyoffice-bridge', `  name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE}`),
    expect: 'pass',
  },
  {
    name: '带 # 注释（合法）',
    content: join(
      '# 顶部注释',
      '- insert:',
      '    # 行内注释',
      '    - id: onlyoffice-bridge',
      `      name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE} # 尾注释`,
    ),
    expect: 'pass',
  },
  {
    name: '事故写法：裸 - 换行（必须拦）',
    content: join(
      '# 注释',
      '-',
      '  insert:',
      '    - id: onlyoffice-bridge',
      `      name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE}`,
    ),
    expect: 'fail',
  },
  { name: '空文件（必须拦）', content: '', expect: 'fail' },
  {
    name: '顶层不是序列（必须拦）',
    content: join('insert:', '  - id: onlyoffice-bridge', `    name: ${QUOTE}dsh-onlyoffice-bridge${QUOTE}`),
    expect: 'fail',
  },
  {
    name: '缺少 name（必须拦）',
    content: join('- insert:', '    - id: onlyoffice-bridge'),
    expect: 'fail',
  },
  {
    name: 'name 与 package.json 不一致（必须拦）',
    content: join('- insert:', '    - id: someone-else', `      name: ${QUOTE}other-package-name${QUOTE}`),
    expect: 'fail',
  },
]
void lines

let failed = 0
try {
  for (const testCase of CASES) {
    fs.writeFileSync(PATCH, testCase.content, 'utf8')
    let outcome
    try {
      execFileSync(process.execPath, [VALIDATOR], { stdio: 'pipe' })
      outcome = 'pass'
    } catch {
      outcome = 'fail'
    }
    const ok = outcome === testCase.expect
    if (!ok) failed += 1
    console.log(`${ok ? 'ok  ' : 'FAIL'}  ${testCase.name} => 期望 ${testCase.expect}，实际 ${outcome}`)
  }
} finally {
  fs.writeFileSync(PATCH, original, 'utf8')
  console.log('--- cordis.patch.yml 已还原 ---')
}

console.log(failed ? `${failed} 个用例失败` : '全部用例通过')
process.exit(failed ? 1 : 0)
