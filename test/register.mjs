// 独立运行测试时的解析钩子：把 DSH 提供的 @deepseek-ai/* 运行时包映射到本地轻量实现。
// 在真实 DSH 中由 profile 的 node_modules 提供，无需这个钩子。
import { register } from 'node:module'

register('./resolve.mjs', import.meta.url)
