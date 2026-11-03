// 见 register.mjs：只重写 cordis 的解析结果，其余交给默认解析。
const CORDIS = new URL('./shims/cordis.mjs', import.meta.url).href

export async function resolve(specifier, context, nextResolve) {
  if (specifier === '@deepseek-ai/cordis') return { url: CORDIS, shortCircuit: true }
  return nextResolve(specifier, context)
}
