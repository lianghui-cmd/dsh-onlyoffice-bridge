// @ts-check
// 全部 Agent 工具的定义表（名称 / 描述 / 参数 JSON Schema）。
// 实际执行落在 lib/tools.js 的同名处理函数上（key 与 name 一一对应）。

/** 复用的小片段 */
const fileKey = { type: 'string', description: '文档 fileKey（onlyoffice_open / onlyoffice_status 返回）' }
const outputPath = { type: 'string', description: '可选，输出文件的绝对路径；默认为缓存目录下同名文件' }
const lockedBy = { type: 'string', description: '操作者标识（默认 agent）；软锁用于防止并发修改' }

/**
 * @typedef {object} ToolDefinition
 * @property {string} key 处理函数名
 * @property {string} name 工具名（注册时会加 onlyoffice_ 前缀）
 * @property {string} description
 * @property {Record<string, unknown>} parameters
 */

/** @type {ToolDefinition[]} */
export const TOOL_DEFINITIONS = [
  // ===================== 基础 =====================
  {
    key: 'open',
    name: 'open',
    description:
      '在 ONLYOFFICE 中打开一个 Office 文档供人工编辑，返回 editorUrl 与 fileKey；保存后自动回写原文件。' +
      'filePath（本地文件）/ contentBase64（内容）/ fileKey（已有缓存）三选一。',
    parameters: {
      type: 'object',
      properties: {
        filePath: { type: 'string', description: '本地文件绝对路径（推荐）' },
        contentBase64: { type: 'string', description: '文件内容 Base64；与 filePath 二选一' },
        fileKey: { type: 'string', description: '已有缓存文档的 fileKey；与 filePath 二选一' },
        fileName: { type: 'string', description: '编辑器标题里的文件名，如 报告.docx' },
        mode: { type: 'string', enum: ['edit', 'view'], description: 'edit=可编辑（默认），view=只读' },
        lockedBy: lockedBy,
      },
      required: [],
      additionalProperties: false,
    },
  },
  {
    key: 'read-doc',
    name: 'read-doc',
    description: '读取缓存文档的正文文本。支持 docx（mammoth）与 xlsx/csv（按工作表输出二维数组）；pptx 不支持正文抽取。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'status',
    name: 'status',
    description: '查看桥接状态：缓存文档列表（fileKey、文件名、原文件路径、大小、最近编辑状态、软锁）、最近一次保存回调。',
    parameters: {
      type: 'object',
      properties: { fileKey: { type: 'string', description: '可选，只查看该文档' } },
      required: [],
      additionalProperties: false,
    },
  },
  {
    key: 'export-pdf',
    name: 'export-pdf',
    description:
      '把 docx/xlsx/pptx 导出为 PDF（仅这三种格式）。会建立一次编辑会话并 forcesave，PDF 在该文档下一次保存时写入 outputPath。',
    parameters: {
      type: 'object',
      properties: { fileKey, outputPath },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'purge',
    name: 'purge',
    description: '清理缓存：target=one 删除单个 fileKey（连同元信息与锁），target=expired 清理超过 TTL 的缓存，target=all 清空未加锁的文档。',
    parameters: {
      type: 'object',
      properties: {
        fileKey: { type: 'string', description: 'target=one 时必填' },
        target: { type: 'string', enum: ['one', 'expired', 'all'], description: '清理范围，默认 expired' },
      },
      required: [],
      additionalProperties: false,
    },
  },

  // ===================== 第一梯队 =====================
  {
    key: 'write-doc',
    name: 'write-doc',
    description:
      '用 Markdown 生成 docx（标题/列表/表格/代码块/粗体斜体），或用 JSON 生成 xlsx（二维数组或对象数组），写入缓存并返回 fileKey。',
    parameters: {
      type: 'object',
      properties: {
        mode: { type: 'string', enum: ['docx', 'xlsx'], description: 'docx=Markdown 渲染，xlsx=JSON 表格' },
        content: { type: 'string', description: 'mode=docx 时为 Markdown 文本；mode=xlsx 时为 JSON 数组字符串' },
        fileName: { type: 'string', description: '文件名，如 周报.docx / 台账.xlsx' },
        sheetName: { type: 'string', description: 'mode=xlsx 时的工作表名，默认 Sheet1' },
      },
      required: ['mode', 'content'],
      additionalProperties: false,
    },
  },
  {
    key: 'diff',
    name: 'diff',
    description: '对比两个同格式缓存文档的正文差异，返回增删行数与变更片段（最多 50 段）。',
    parameters: {
      type: 'object',
      properties: {
        oldKey: { type: 'string', description: '旧版本文档 fileKey' },
        newKey: { type: 'string', description: '新版本文档 fileKey' },
      },
      required: ['oldKey', 'newKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'rename',
    name: 'rename',
    description: '修改缓存文档的显示名称与备注（不改文件内容；建议保留原扩展名）。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        name: { type: 'string', description: '新文件名，如 新版合同.docx' },
        remark: { type: 'string', description: '可选备注，如 2026 评审' },
      },
      required: ['fileKey', 'name'],
      additionalProperties: false,
    },
  },
  {
    key: 'meta',
    name: 'meta',
    description: '查询单个文档的元信息：文件名、原文件路径、大小、更新时间、软锁状态与缓存路径。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'export-markdown',
    name: 'export-markdown',
    description: '把缓存文档转成 Markdown 文本（docx 用 mammoth 转 Markdown，xlsx 每张表输出 ```csv 代码块）。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'lock',
    name: 'lock',
    description: '给文档加应用层软锁，防止多个 Agent 并发改写；被锁后写操作（open/insert-image）会被拦截并提示。',
    parameters: {
      type: 'object',
      properties: { fileKey, lockedBy, note: { type: 'string', description: '可选说明，如 修改第三章' } },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'unlock',
    name: 'unlock',
    description: '解除文档软锁；传 lockedBy 时只有锁的持有者才能解除。',
    parameters: {
      type: 'object',
      properties: { fileKey, lockedBy: { type: 'string', description: '锁持有者；传了会校验身份' } },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'who-lock',
    name: 'who-lock',
    description: '查询文档当前被谁锁定以及锁定时间。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'batch-open',
    name: 'batch-open',
    description: '批量打开多个本地文件，逐个返回 fileKey 与 editorUrl；单个失败不影响其余文件。',
    parameters: {
      type: 'object',
      properties: {
        filePathList: { type: 'array', items: { type: 'string' }, description: '本地文件绝对路径列表' },
        mode: { type: 'string', enum: ['edit', 'view'], description: '统一使用 edit 或 view 模式，默认 edit' },
      },
      required: ['filePathList'],
      additionalProperties: false,
    },
  },
  {
    key: 'batch-read',
    name: 'batch-read',
    description: '批量读取多个文档正文，逐个返回结果与失败原因。',
    parameters: {
      type: 'object',
      properties: { fileKeyList: { type: 'array', items: { type: 'string' }, description: 'fileKey 列表' } },
      required: ['fileKeyList'],
      additionalProperties: false,
    },
  },
  {
    key: 'batch-export-pdf',
    name: 'batch-export-pdf',
    description: '批量导出 PDF：逐个建立转换会话并返回各自的 outputPath 与 editorUrl。',
    parameters: {
      type: 'object',
      properties: {
        fileKeyList: { type: 'array', items: { type: 'string' }, description: 'fileKey 列表' },
        outputDir: { type: 'string', description: '可选，PDF 统一输出目录' },
      },
      required: ['fileKeyList'],
      additionalProperties: false,
    },
  },

  // ===================== 第二梯队：多格式导出（DS ConvertService） =====================
  {
    key: 'export-docx',
    name: 'export-docx',
    description: '通过 ConvertService 把文档转成 docx 并下载到本地（支持 odt/rtf/txt/html 等源格式）。',
    parameters: { type: 'object', properties: { fileKey, outputPath }, required: ['fileKey'], additionalProperties: false },
  },
  {
    key: 'export-xlsx',
    name: 'export-xlsx',
    description: '通过 ConvertService 把文档转成 xlsx 并下载到本地（支持 ods/csv/html 等源格式）。',
    parameters: { type: 'object', properties: { fileKey, outputPath }, required: ['fileKey'], additionalProperties: false },
  },
  {
    key: 'export-pptx',
    name: 'export-pptx',
    description: '通过 ConvertService 把文档转成 pptx 并下载到本地（支持 odp 等源格式）。',
    parameters: { type: 'object', properties: { fileKey, outputPath }, required: ['fileKey'], additionalProperties: false },
  },
  {
    key: 'export-html',
    name: 'export-html',
    description: '通过 ConvertService 把文档转成单个 HTML 并下载到本地，返回前 2000 字符预览。',
    parameters: { type: 'object', properties: { fileKey, outputPath }, required: ['fileKey'], additionalProperties: false },
  },

  // ===================== 第二梯队：模板 / 图片 =====================
  {
    key: 'template-fill',
    name: 'template-fill',
    description:
      'docx 模板填充：模板正文中用 ${varName} 占位，传入变量生成新文档（保留原样式）。' +
      '返回新 fileKey、已替换与缺失的变量名。仅支持 docx。',
    parameters: {
      type: 'object',
      properties: {
        templateKey: { type: 'string', description: '模板 docx 的 fileKey' },
        vars: {
          type: ['object', 'string'],
          description: '变量对象，或 JSON 字符串，如 {"name":"张三","date":"2026-11-03"}',
        },
        newFileName: { type: 'string', description: '生成后的文件名，如 合同-张三.docx' },
      },
      required: ['templateKey', 'vars'],
      additionalProperties: false,
    },
  },
  {
    key: 'insert-image',
    name: 'insert-image',
    description:
      '往 docx 末尾插入一张图片（png/jpg/gif/webp），保留原有内容。仅支持 docx；xlsx / pptx 暂不支持。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        imgPath: { type: 'string', description: '图片本地绝对路径' },
        widthCm: { type: 'number', description: '图片宽度（厘米），默认 14；高度按原图比例' },
        heightCm: { type: 'number', description: '可选，显式指定高度（厘米）' },
        lockedBy: lockedBy,
      },
      required: ['fileKey', 'imgPath'],
      additionalProperties: false,
    },
  },

  // ===================== 排版：目录 / 页眉页脚 =====================
  {
    key: 'toc-regen',
    name: 'toc-regen',
    description:
      '自动刷新 Word 文档目录（TOC 域），PPT 同样适用；文档里没有目录时会在正文开头插入一个。' +
      '刷新后的页码与条目由编辑器在打开时按标题样式重算。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'header-footer-set',
    name: 'header-footer-set',
    description:
      '批量设置 Word 页眉、页脚：固定文字、公司名称、自动页码、版本号；支持多节独立页眉页脚与首页独立。仅 docx。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        headerText: { type: 'string', description: '页眉文字，可选' },
        footerText: { type: 'string', description: '页脚固定文字 / 版本号，可选' },
        enablePageNumber: { type: 'boolean', description: '是否开启自动页码，默认 false' },
        differentFirstPage: { type: 'boolean', description: '首页是否独立页眉页脚，默认 false' },
        lockedBy,
      },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },

  // ===================== 结构：合并 / 拆分 =====================
  {
    key: 'merge-doc',
    name: 'merge-doc',
    description:
      '按顺序合并多个 docx / pptx，保留原始格式、分节、页眉页脚与图片；输出一份新文档 fileKey。' +
      '合并 pptx 时会连同版式、母版、主题一起搬运。',
    parameters: {
      type: 'object',
      properties: {
        fileKeyList: { type: 'array', items: { type: 'string' }, description: '待合并文档 fileKey 数组，按传入顺序拼接' },
        outFileName: { type: 'string', description: '输出文档名称，可选' },
      },
      required: ['fileKeyList'],
      additionalProperties: false,
    },
  },
  {
    key: 'split-doc',
    name: 'split-doc',
    description:
      '拆分 docx / pptx：section=Word 按分节符（只有一个分节时按段落均分）；slide=PPT 按单页。' +
      '输出一批独立文档 fileKey 列表。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        mode: {
          type: 'string',
          enum: ['section', 'slide'],
          description: '拆分模式：section=Word 按分节符；slide=PPT 按单页',
        },
      },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },

  // ===================== 版本：快照 / 回滚 =====================
  {
    key: 'snapshot',
    name: 'snapshot',
    description:
      '给当前文档打一个不可变版本快照，附带标签，用于后续 diff 对比或回滚；快照本身不会被修改。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        label: { type: 'string', description: '快照备注标签，例如 初稿 / 合同审核版，可选' },
      },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'snapshot-list',
    name: 'snapshot-list',
    description: '列出一个文档的全部历史快照，拿到 snapId 后可以 diff 对比或 restore 回滚。',
    parameters: {
      type: 'object',
      properties: { fileKey },
      required: ['fileKey'],
      additionalProperties: false,
    },
  },
  {
    key: 'restore-snapshot',
    name: 'restore-snapshot',
    description:
      '把文档回滚到指定 snapId 快照；回滚前会自动备份当前版本（返回 backupSnapId），不会丢内容。',
    parameters: {
      type: 'object',
      properties: {
        fileKey,
        snapId: { type: 'string', description: '快照 ID，来自 snapshot 或 snapshot-list 的返回' },
        lockedBy,
      },
      required: ['fileKey', 'snapId'],
      additionalProperties: false,
    },
  },
]
