// 单进程测试入口：顺序加载 node:test 用例并汇总结果。
// 沙箱里 node --test 的文件级子进程可能被拦截（spawn EPERM），所以在同一进程内加载。
//
// 只加载 test/*.test.mjs 里用 node:test 的报告型用例；
// scripts 自检类用例（validate-patch.test.mjs 会 spawn 子进程并自己打印结果）
// 由 npm test 的第二条命令单独执行。
import { register } from 'node:module'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

register('./resolve.mjs', import.meta.url)

/** 这类用例自带独立运行入口，不参与 node:test 汇总。 */
const STANDALONE = new Set(['validate-patch.test.mjs'])

const here = path.dirname(fileURLToPath(import.meta.url))
const files = fs
  .readdirSync(here)
  .filter((file) => file.endsWith('.test.mjs') && !STANDALONE.has(file))
  .sort()
for (const file of files) {
  await import(pathToFileURL(path.join(here, file)).href)
}
