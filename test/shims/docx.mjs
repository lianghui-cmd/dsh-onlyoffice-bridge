// @ts-check
// 测试替身：真实安装时会解析到 node_modules 里的正式包；仓库里未安装时由
// test/resolve.mjs 映射到这里，保证 `npm test` 在无依赖环境也能跑。
//
// 注意：不用 require()——DSH 插件在 ESM 运行时里加载，用 createRequire 反而会
// 绕过加载器的 resolve 钩子，导致测试环境与真实环境行为不一致。

const notInstalled = (name) => {
  throw new Error(`测试替身：${name} 未安装（npm install ${name} 后可跑真实用例）`)
}

export const Document = /** @type {any} */ (class {
  constructor() {
    notInstalled('docx')
  }
})

export const Packer = { async toBuffer() {
  notInstalled('docx')
} }

export const Paragraph = class {
  constructor() {
    notInstalled('docx')
  }
}

export const TextRun = class {
  constructor() {
    notInstalled('docx')
  }
}

export const ImageRun = /** @type {any} */ (class {})
export const AlignmentType = { CENTER: 'center', LEFT: 'left', RIGHT: 'right' }
export const patchDocument = async () => {
  notInstalled('docx')
}
export const PatchType = { PARAGRAPH: 1, DOCUMENT: 2, IMAGE: 3 }

export const mammoth = {
  async extractRawText() {
    notInstalled('mammoth')
  },
}
