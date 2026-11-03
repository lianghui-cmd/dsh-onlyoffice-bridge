# 我把 ONLYOFFICE 塞进了 DeepSeek Harness：让 Agent 直接改 Word、Excel、PPT

> 一个 Cordis 插件，22 个 Agent 工具，自包含依赖，CI 全绿。
> 源码：<https://github.com/lianghui-cmd/dsh-onlyoffice-bridge>

---

## 一、AI Agent 都能跑浏览器了，为什么还改不了一个 .docx？

试过让 Agent 改过 Word 文档的人，大概都遇过下面这些场景：

- 「帮我把这份报价单里的单价改一下」→ Agent 只能告诉你改哪个文件，然后你自己开 Word
- 「把这三份 Excel 合并到一起」→ Agent 写了个脚本，但数据类型乱码、公式丢失
- 「让 Agent 看看这份合同哪里写得不对」→ 抽取出纯文本，样式、表格全没了

我自己最近在做一个教育融媒平台的项目，业务侧的同学每天都要改文档、改课件、改讲稿。**他们的诉求很朴素：让 AI 帮我干活，但我还得在编辑器里看一眼**。

于是就有了这个插件：**让 DeepSeek Harness 能调起一个真正的 Office 编辑器，改完自动回写原文件**。

## 二、这套东西怎么工作的

简单说就是「桥接」，不绕路：

```
DeepSeek Harness                          ONLYOFFICE Document Server
      │                                              │
      │  1. Agent 调 onlyoffice_open                │
      │     └─ 生成一个签名过的编辑链接              │
      │                                              │
      │  2. 你打开链接（浏览器）                     │
      │     └─ Document Server 拉取文档渲染成编辑器  │
      │                                              │
      │  3. 你编辑、保存                             │
      │     └─ Document Server 回调保存结果          │
      │  4. 自动写回缓存目录 + 原文件                │
```

关键点：

- **不内嵌 iframe**，跨域最稳，避免浏览器同源策略一堆麻烦
- **缓存目录 + 原文件双回写**：不会让你改了就丢
- **JWT（HS256）双向验签**：DS 调过来、DS 拉文件都带 token，伪造不了

## 三、22 个 Agent 工具，覆盖「编辑、生成、转换、对比、锁定」

分三组给你看：

**基础：**
- `onlyoffice_open` — 打开 / 编辑 / 新建，一个命令拿编辑链接
- `onlyoffice_read-doc` — 抽正文（docx、xlsx、csv；pptx 不支持）
- `onlyoffice_status` — 看缓存文档列表、最近保存回调
- `onlyoffice_export-pdf` — 导出 PDF（docx / xlsx / pptx）
- `onlyoffice_purge` — 清缓存

**第一梯队（批量与元信息）：**
- `onlyoffice_write-doc` — 用 **Markdown 生成 docx**、用 JSON 生成 xlsx
- `onlyoffice_diff` — 两个文档逐行对比（增删行数 + 变更片段）
- `onlyoffice_rename` / `onlyoffice_meta` / `onlyoffice_export-markdown`
- `onlyoffice_lock` / `onlyoffice_unlock` / `onlyoffice_who-lock` — 软锁，防多个 Agent 并发改
- `onlyoffice_batch-open` / `onlyoffice_batch-read` / `onlyoffice_batch-export-pdf`

**第二梯队（格式转换与富文本）：**
- `onlyoffice_export-docx` / `export-xlsx` / `export-pptx` / `export-html` — 走 DS ConvertService
- `onlyoffice_template-fill` — docx 模板 `${var}` 填充，**保留原样式**
- `onlyoffice_insert-image` — 往 docx 末尾追加图片

举个完整场景：

> **你：** 把 `D:\docs\报价单.xlsx` 打开给我改一下  
> **Agent →** `onlyoffice_open { filePath: "D:\\docs\\报价单.xlsx" }`  
> **你：** 浏览器打开编辑链接，改完保存  
> **Agent →** `onlyoffice_status`  ← 看到 `lastCallback.status=2`，已经落盘

> **你：** 用「合同模板.docx」生成一份给张三的合同，金额 100 万  
> **Agent →** `onlyoffice_template-fill`  ← 一份新 docx，占位符全部填好

## 四、安装方式（3 行）

```bash
# 1. 把插件目录放进 DeepSeek-Harness/plugins/
# 2. 复制 .env.example 为 .env，填上 DS 地址、JWT 密钥、DSH_PUBLIC_URL
# 3. 用 plugin_manager 安装并启用，或手工把 insert 段加进 cordis.patch.yml
```

⚠️ 一个关键配置（Docker 部署时最容易踩的坑）：

`DSH_PUBLIC_URL` 是**Document Server 能访问到的地址**，不是你浏览器的地址。DS 跑在 docker 里时，必须写宿主机局域网 IP，不能写 `127.0.0.1`。插件的 README 里有完整故障排查表。

## 五、为什么值得 star 这个仓库

这个项目有一个我觉得挺重要的设计：**依赖是内置的**。

很多 DSH 插件仓库会写「npm install 然后重启」，但实际上：

- 插件包依赖本身可能被 `plugin_manager` 重装时清空
- 用户的环境五花八门，npm 装依赖经常各种网络 / 权限问题
- 离线环境装不了

我用了 `vendor/` 目录，把 38 个运行时依赖（docx、mammoth、xlsx、diff、marked、jszip + 传递依赖）**真的打包进仓库**，剪枝后 14 MB，但作用极大：

> **CI 上专门跑一个「不装 node_modules」的测试套件**，36 条用例照样全绿。  
> 这就是「自包含」的证据。

另外：

- **36 条 node:test + 10 条配置校验矩阵**，覆盖 JWT、路径穿越、回调验签、软锁、模板填充、插图
- **CI 在 Node 22/24 × Linux/Windows 四个矩阵全绿**
- **MIT 许可**，可商用
- 插件启动失败也会给出**可读错误**（缺依赖时会告诉你缺哪个、怎么补）

## 六、踩过的坑（真·实战记录）

这个插件写下来踩的坑，比文章长，简单说几个：

1. **`editor.html?token=` 是错的**。ONLYOFFICE 的编辑器页面不直接认 token 参数，必须用 `api.js?token=...`
2. **每次编辑要换新 key**。DS 用 key 缓存编辑会话，同一个 key 会拿到旧缓存
3. **回调验签是必须的**，且要兼容 `Authorization` 头 / `body.token` / 旧版 md5 token 三种形态
4. **路径穿越防护**：fileKey 必须白名单校验 + 路径归属校验，不能直接 `path.join(STORAGE_DIR, relPath)`
5. **软锁 vs 真锁**：只做应用层软锁拦截 Agent 的写操作，不阻止编辑器（防止互锁死局）

## 七、后续计划

- [ ] pptx 的正文抽取（现在不支持）
- [ ] 模板填充支持表格、页眉页脚
- [ ] 编辑器内嵌方案（跨域稳定后）
- [ ] 多语言界面（现在中英混排，README 中文、代码注释中文、commit 英文）

## 八、点个 Star 吧

仓库：<https://github.com/lianghui-cmd/dsh-onlyoffice-bridge>

如果你正在做类似的「AI + Office 文档」的活，欢迎点个 ⭐、提个 issue、或者 fork 改造。
也欢迎在评论区聊聊：你是怎么让 AI 帮你改 Office 文档的？

---

**标签：** #DeepSeek #AI-Agent #Office #ONLYOFFICE #自动化 #文档处理 #开发者工具
