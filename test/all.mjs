// 单进程测试入口：顺序加载全部 *.test.mjs，node:test 收集并汇总结果。
// 沙箱里 node --test 的文件级子进程可能被拦截（spawn EPERM），所以在同一进程内加载。
import { register } from 'node:module'
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

register('./resolve.mjs', import.meta.url)

const here = path.dirname(fileURLToPath(import.meta.url))
const files = fs.readdirSync(here).filter((file) => file.endsWith('.test.mjs')).sort()
for (const file of files) {
  await import(pathToFileURL(path.join(here, file)).href)
}
