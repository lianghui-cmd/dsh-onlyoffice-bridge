// 把 lib/document.js 里的纯函数重新导出给测试用（避免测试直接依赖内部实现细节）。
export {
  TEXT_LIMIT,
  cmToPx,
  collectPlaceholders,
  imageTypeOf,
  renderPlaceholders,
} from '../lib/document.js'

/** 测试里需要检查 XML 转义结果时用：把常见字符转义成 XML 实体。 */
export const escapeForTest = (value) =>
  String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
