// @deepseek-ai/cordis 的测试替身（只含插件实际使用的 Service 基类）。
// 与真实 cordis 一致：Service 是带抽象 init 的基类。
export class Service {
  static init = Symbol.for('cordis.init')
  static check = Symbol.for('cordis.check')
  static config = Symbol.for('cordis.config')
  static invoke = Symbol.for('cordis.invoke')
  static extend = Symbol.for('cordis.extend')
  static tracker = Symbol.for('cordis.tracker')

  /**
   * @param {unknown} ctx
   * @param {string} [name]
   */
  constructor(ctx, name) {
    this.ctx = ctx
    this.name = name ?? this.constructor.name
  }
}

export default { Service }
