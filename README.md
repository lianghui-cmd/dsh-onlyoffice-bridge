# dsh-onlyoffice-bridge

[![CI](https://github.com/lianghui-cmd/dsh-onlyoffice-bridge/actions/workflows/ci.yml/badge.svg)](https://github.com/lianghui-cmd/dsh-onlyoffice-bridge/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Node](https://img.shields.io/badge/node-%3E%3D22.19-brightgreen)](https://nodejs.org)

DeepSeek Harness（DSH）Cordis 插件：把 **ONLYOFFICE Document Server** 桥接给 Agent，
让 Agent 能打开、人工在线编辑、再自动回写 docx / xlsx / pptx 文件，并能生成、转换、
对比、锁定 Office 文档。

> 模式：桥接模式。DSH 生成编辑链接 → 浏览器打开 ONLYOFFICE 编辑器 → 保存回调把内容写回
> DSH 缓存目录**和原始文件**。不内嵌 iframe，跨域最稳。

- **29 个 Agent 工具**：打开 / 读取 / 生成 / 转换 / 对比 / 软锁 / 批量 / 模板填充 / 插图 /
  排版（目录·页眉页脚）/ 合并拆分 / 版本快照
- 纯 Node 实现：`node:http` / `node:crypto` / `node:fs`，**不依赖 express**
- **依赖内置**：`vendor/` 打包了文档处理依赖，重装插件不会掉依赖，离线可用
- JWT（HS256）签名 + 回调验签，兼容 `Authorization` 头 / `body.token` / 旧版 md5 token
- 缓存 TTL 自动清理（随插件卸载停止，不残留定时器）
- 配置既可以来自环境变量 / `.env`，也可以来自 `cordis.patch.yml` 的 `config`
- **44 条自动化用例 + 10 条配置校验用例**，CI 在 Node 22/24 × Linux/Windows 全绿

---

## Contents

```
onlyoffice-bridge/
├── index.js            # 插件主代码（唯一运行时入口）
├── index.d.ts          # 类型声明
├── package.json        # 插件清单（dsh.bundle.patch）
├── cordis.patch.yml    # DSH 加载补丁（DSH 靠它识别并插入插件行）
├── .env.example        # 环境配置模板
├── locale/{zh,en}.json # 插件列表里的标题与描述
├── icon.svg            # 插件图标
├── types/cordis.d.ts   # 仓库内类型桩（离线 typecheck 用）
├── tsconfig.json       # JS + JSDoc 严格类型检查
├── test/               # node:test 用例（24 条，覆盖 JWT / 回调 / 静态文件 / 工具）
└── README.md
```

只依赖 DSH 必然提供的 `@deepseek-ai/cordis`；不依赖 express，也不依赖 schema 库
（配置校验由插件自己完成，非法值回退默认值）。

## Prerequisites

1. **ONLYOFFICE Document Server**（推荐 docker）：

   ```bash
   docker run -d --name onlyoffice -p 8080:80 \
     -e JWT_ENABLED=true \
     -e JWT_SECRET=replace-me-secret-key \
     onlyoffice/documentserver
   ```

   生产环境建议写在 `local.json` 里并重启 DS：

   ```json
   { "services": { "CoAuthoring": { "secret": { "inbox": { "string": "SAME_SECRET" }, "outbox": { "string": "SAME_SECRET" }, "session": { "string": "SAME_SECRET" } } } } }
   ```

   容器内路径：`/etc/onlyoffice/documentserver/local.json`，改完 `supervisorctl restart all`。

2. DSH 能访问 DS（`ONLYOFFICE_DS_URL`）。
3. **DS 能访问 DSH**（`DSH_PUBLIC_URL`）：DS 要用它来拉源文件、回调保存结果。
   - DS 和 DSH 在同一台机器且 DS 不是容器：`http://127.0.0.1:11006` 即可
   - DS 跑在 docker 里：必须写宿主机局域网 IP（如 `http://192.168.1.20:11006`）
     或 `http://host.docker.internal:11006`；**不能写 `127.0.0.1`**
   - DSH 默认只监听 `127.0.0.1`，要让 DS 容器访问到，需要让 DSH 监听局域网地址
     （webserver 的 `host` 配置，具体字段见你所用 DSH 版本的 `dsh-host-webserver` 文档）

## Installation

```text
DeepSeek-Harness/
└── plugins/
    └── onlyoffice-bridge/   ← 把本目录整个放进来
```

1. 复制配置模板并填写：`cp .env.example .env`
2. 让 DSH 加载插件，二选一：
   - **推荐**：在 DSH 里用插件管理器安装这个目录（`plugin_manager install_bundle`，
     target 指向本目录的绝对路径）。它会读取 `package.json` 的 `dsh.bundle.patch`
     并把 `cordis.patch.yml` 里的行插入当前 profile。
   - **手写**：把下面三行合并进 profile 的 `cordis.patch.yml`：

     ```yaml
     - insert:
         - id: onlyoffice-bridge
           name: 'dsh-onlyoffice-bridge'
     ```

     并让 profile 能解析到本包（例如用工作区目录安装 / `file:` 依赖）。
3. 重启 DSH（新安装的 bundle 在重启后稳定生效），或按插件管理器的提示热加载。
4. 验证：DSH 插件列表里应出现 **OnlyOffice Bridge**；Agent 里调用 `onlyoffice_status`，
   返回里应带 `storageDir`、`dsUrl`、`publicUrl`。

## Configuration

优先级：**进程环境变量 > 本目录 `.env` > `cordis.patch.yml` 的 `config` > 内置默认值**。

| 变量 | 默认 | 说明 |
| --- | --- | --- |
| `ONLYOFFICE_DS_URL` | `http://127.0.0.1:8080` | DS 地址，用于生成编辑链接、调用 CommandService |
| `ONLYOFFICE_SECRET` / `ONLYOFFICE_SECRETS` | `replace-me-secret-key` | 与 DS 一致的 JWT 密钥；前者优先 |
| `ONLYOFFICE_JWT_HEADER` | `Authorization` | DS 的 JWT 请求头名 |
| `ONLYOFFICE_JWT_HEADER_PREFIXED` | `false` | 请求头是否带 `Bearer ` |
| `DSH_PUBLIC_URL` | 自动探测 | **DS 能访问到的 DSH 地址**（最关键的一项） |
| `EDITOR_DS_URL` | 同 `dsUrl` | 浏览器打开编辑器用的 DS 地址（多网卡 / 反向代理时用） |
| `STORAGE_DIR` | `<DSH_HOME>/onlyoffice` | 缓存目录，相对路径按进程工作目录解析 |
| `CACHE_TTL_HOUR` | `24` | 缓存过期时间；保存 / 重新打开会刷新 |
| `CACHE_CLEAN_INTERVAL_MIN` | `360` | 自动清理间隔（分钟） |
| `MAX_FILE_MB` | `128` | 单文件上限 |
| `ALLOWED_ORIGINS` | 空 | 额外允许的网页来源，逗号分隔（默认同源即可） |

也可以在 `cordis.patch.yml` 里写同样的语义（camelCase）：

```yaml
- insert:
    - id: onlyoffice-bridge
      name: 'dsh-onlyoffice-bridge'
      config:
        dsUrl: http://192.168.1.20:8080
        jwtSecret: SAME_SECRET
        publicUrl: http://192.168.1.20:11006
        cacheTtlHour: 48
```

数值字段写错（例如把 `cacheTtlHour` 写成非数字）时会回退到默认值，不会让插件加载失败。

## All Tools (v0.5.0, 29 total)

Tool names registered into DSH are `onlyoffice_<name>`. The table below
grouped by capability.

### Basic
- `onlyoffice_open` — `filePath` / `contentBase64` / `fileKey` 三选一，返回 `editorUrl` 与 `fileKey`，保存后回写原文件
- `onlyoffice_read-doc` — 读正文（docx / xlsx / csv；pptx 不支持正文抽取）
- `onlyoffice_status` — 缓存文档列表 + 软锁 + 最近一次保存回调
- `onlyoffice_export-pdf` — 转 PDF（仅 docx / xlsx / pptx）
- `onlyoffice_purge` — `target=one|expired|all`

### Tier 1
- `onlyoffice_write-doc` — `mode=docx` 用 Markdown 生成；`mode=xlsx` 用 JSON 二维数组或对象数组生成
- `onlyoffice_diff` — 两个同格式文档逐行对比，返回增删行数 + 最多 50 段变更
- `onlyoffice_rename` — 改显示名与备注（缺扩展名会自动补）
- `onlyoffice_meta` — 单文档元信息、锁、大小、缓存路径
- `onlyoffice_export-markdown` — docx 转 Markdown（mammoth）；xlsx 每张表输出 `csv` 代码块
- `onlyoffice_lock` / `onlyoffice_unlock` / `onlyoffice_who-lock` — 应用层软锁
- `onlyoffice_batch-open` / `onlyoffice_batch-read` / `onlyoffice_batch-export-pdf` — 批量，逐项返回成功与失败原因

### Tier 2
- `onlyoffice_export-docx` / `onlyoffice_export-xlsx` / `onlyoffice_export-pptx` / `onlyoffice_export-html` — 走 DS ConvertService 转换并下载到本地
- `onlyoffice_template-fill` — docx 模板 `${var}` 填充，保留原样式，返回已替换 / 缺失变量
- `onlyoffice_insert-image` — docx 末尾追加图片（仅 docx；xlsx / pptx 暂不支持）

> 早期设计稿把 `template-fill`、`insert-image` 既算进梯队又单列，合计 25；加上 v0.5.0 的
> 7 个新工具后，实际注册数是 **29 个**，以本表为准。

### Typesetting（v0.5.0）

- `onlyoffice_toc-regen` — 刷新 Word 目录（TOC 域）；文档没有目录时在正文开头插入一个。
  刷新后由编辑器按标题样式重算条目与页码。PPT 则刷新页码占位符
- `onlyoffice_header-footer-set` — 设置页眉、页脚、自动页码、版本号；支持多节独立与首页独立（仅 docx）

### Structure（v0.5.0）

- `onlyoffice_merge-doc` — 按顺序合并多个 docx / pptx，保留分节、页眉页脚、图片；
  合并 pptx 时连版式、母版、主题一起搬运
- `onlyoffice_split-doc` — `section` 按分节符拆 Word（只有一个分节时按段落均分）；
  `slide` 按页拆 PPT

### Version（v0.5.0）

- `onlyoffice_snapshot` — 给文档打不可变版本快照，可带标签（如「初稿」「合同审核版」）
- `onlyoffice_snapshot-list` — 列出全部历史快照，拿到 `snapId`
- `onlyoffice_restore-snapshot` — 回滚到指定快照；**回滚前自动备份当前版本**，返回 `backupSnapId`

> 版本快照存放在缓存目录的 `snapshots/<fileKey>/` 下，随缓存 TTL 一起清理；
> 删除文档（`purge`）时其快照目录也会一并删除，不留孤儿文件。

典型对话：

```text
用户：把 D:\docs\报价单.xlsx 打开给我改一下
Agent → onlyoffice_open { "filePath": "D:\\docs\\报价单.xlsx" }
       ← editorUrl: http://192.168.1.20:8080/web-apps/apps/api/documents/api.js?token=...
用户在浏览器打开 editorUrl 编辑并保存
       → DS 回调 /onlyoffice/callback（status=2/6）→ 下载内容 → 写回 报价单.xlsx
Agent → onlyoffice_status   ← 看到 lastCallback.status=2、edited=true
```

### Why the editor link is `api.js`, not `editor.html`

`/web-apps/apps/api/documents/api.js?...` 是 ONLYOFFICE 官方推荐的接入方式：
浏览器打开它时，文档服务器会返回一个已经注入好配置的完整编辑器页面。
例如：

```html
<!-- 自己的页面里也可以这样嵌入（可选，不是必须） -->
<script src="http://ds:8080/web-apps/apps/api/documents/api.js?token=JWT"></script>
```

`/web-apps/apps/api/documents/editor.html` 是编辑器组件本身，**不直接接受** token 参数，
除非自己再套一层页面去加载 `api.js`。

## Caching Strategy

- 每次 `onlyoffice_open` 都会给文档分配一个**新的 key**（`<32hex>-<16hex>.<ext>`）。
  DS 用 key 缓存编辑会话，复用旧 key 会拿到过期会话，所以每次编辑都换新 key；
  同一份原文件只保留最新一代缓存，上一代在下次打开时回收。
- 缓存目录里每个文档有 `<key>.<ext>` 与 `<key>.meta.json`（文件名、原文件路径、导出目标等）。
- 超过 `CACHE_TTL_HOUR` 的缓存由定时任务清理，也可以让 Agent 调 `onlyoffice_purge`。
- 定时器在插件卸载时清除，不会残留。
- 版本快照存放在 `snapshots/<fileKey>/`，随 TTL 一起清理；`purge` 会连带删除该文档的快照目录。

## HTTP Routes

除了 DS 需要的三条（静态文件 / 保存回调 / 上传）和健康检查，v0.5.0 还开放了与
Agent 工具**同一份实现**的操作接口，方便前端或脚本直接调用：

| 路由 | Method | 对应工具 |
| --- | --- | --- |
| `/onlyoffice/toc` | POST | `toc-regen` |
| `/onlyoffice/header-footer` | POST | `header-footer-set` |
| `/onlyoffice/merge` | POST | `merge-doc` |
| `/onlyoffice/split` | POST | `split-doc` |
| `/onlyoffice/snapshot` | POST | `snapshot` |
| `/onlyoffice/restore` | POST | `restore-snapshot` |
| `/onlyoffice/snapshot-list` | GET | `snapshot-list`（query 传 `fileKey`） |

POST 的 body 就是工具参数（JSON），返回结构与工具返回值一致：

```bash
curl -X POST http://127.0.0.1:11006/onlyoffice/snapshot \
  -H 'Content-Type: application/json' \
  -d '{"fileKey":"<key>.docx","label":"审核前"}'
```

这些接口与 Agent 工具共享 `lib/tools.js` 的实现，因此行为、错误信息、软锁规则完全一致。

## Networking & Security

- **回调验签**：`/onlyoffice/callback` 必须携带有效 JWT（请求头或 `body.token`），
  密钥不一致会返回 403 并在日志里提示（这是最常见的故障）。
- **目录穿越防护**：静态文件路由只接受 `<32hex>-<16hex>.<ext>` 形式的 key，
  并校验解析后的路径仍在缓存目录内。
- **CSRF 防护**：`/onlyoffice/upload` 与静态文件路由限制同源（或 `ALLOWED_ORIGINS`）；
  DS 的服务端拉取 / 回调不带 `Origin` / `Referer`，因此不受影响。
- **原文件回写**：只有通过 `filePath` 打过的文档会回写原文件；回写失败会在日志里保留
  缓存路径，内容不会丢。

## Common Failures

| 现象 | 原因 | 处理 |
| --- | --- | --- |
| 编辑器提示 “The document security token is not correctly formed” | 插件与 DS 的 JWT 密钥不一致 | 对齐 `ONLYOFFICE_SECRET` 与 DS `local.json` 的 secret |
| 编辑器一直转圈 / “下载失败” | DS 拉不到 `DSH_PUBLIC_URL/onlyoffice/static-file/...` | DS 容器里 `curl` 一下这个地址；docker 场景改成宿主机 IP |
| 保存了但没有落盘 | 回调地址 DS 访问不到，或回调 403 | 看 `onlyoffice_status` 的 `lastCallback`，以及 DSH 日志里的 `[onlyoffice] 回调` |
| 打开时提示缓存过期 / 内容不对 | 复用了旧 key | 本插件每次打开都换新 key；手工拼链接时要重新调 `onlyoffice_open` |
| 导出 PDF 后没有文件 | PDF 是异步生成的 | 导出后需要该文档保存一次（打开 `editorUrl` 保存，或用户保存） |
| 端口/地址没生效 | DSH 桌面端不会自动读取 `.env` | 把变量注入 DSH 进程环境，或用 `cordis.patch.yml` 的 `config` |
| 工具报“缺少可选依赖 docx / mammoth / xlsx” | 插件目录没装依赖（重装 bundle 可能清空 `node_modules`） | 在插件目录执行 `npm install` 或 `pnpm install`，然后重载插件 |

## Development & Testing

```bash
# 36 cases: JWT sign/verify, path traversal, save callback (status 2/6),
# upload, export, soft lock, batch, template-fill / insert-image arg checks,
# cache cleanup, and registration of 22 tools
npm test

# Runtime shims needed to run the tests standalone:
# @deepseek-ai/cordis and the document dependencies
npm install
```

测试在单进程里加载全部 `*.test.mjs`，并用 `test/shims/cordis.mjs` + `test/resolve.mjs`
把 `@deepseek-ai/cordis` 换成轻量替身，因此不需要安装 DSH 本体也能跑。
文档依赖（docx / mammoth / xlsx / diff / marked / jszip）是**懒加载**的：缺包时只有用到它们的工具报错，
不影响插件启动与其他工具。

## Installation Notes (important)

`plugin_manager` 重装 bundle 会清理 profile 依赖，**插件目录的 `node_modules` 也可能被清**。
因此本插件把文档处理依赖**内置进仓库**（`vendor/`），不依赖任何外部安装：

```
vendor/            # 38 个运行时包，约 14 MB，随 git 一起提交
├── docx/          # 只保留 dist/index.mjs + 类型，剔掉 layout/iife/umd 等旁支构建
├── xlsx/          # 保留 dist/cpexcel.js（非 UTF-8 旧表格必需）
├── mammoth/ jszip/ marked/ diff/ + 传递依赖
└── MANIFEST.json  # 版本清单与许可证
```

加载顺序是 **`vendor/` 优先 → 外部 `node_modules` 兜底**，所以：

- 重装 bundle、清空依赖、换机器、离线环境，插件都照常工作；
- 想升级依赖：改 `package.json` 的版本 → `npm install` → `npm run vendor` → 提交 `vendor/`。

三个必须记住的操作规矩：

1. **先修好包再装**（`cordis.patch.yml` 能解析、`index.js` 能导入），确认无误只装一次；
2. 遇到 `installation rejected` **不要连续重试**——它会再次清空依赖；
3. 不要指望 profile 的 `node_modules` 提供文档依赖，它们只认 `vendor/`。

`vendor/` 内的包互相 `require` 兄弟包时，由 `lib/vendor-resolver.js` 补齐查找路径
（Node 默认不会从 `vendor/mammoth` 找到 `vendor/underscore`）。

## Development & Testing

```bash
# 36 cases: JWT sign/verify, path traversal, save callback (status 2/6),
# upload, export, soft lock, batch, template-fill / insert-image arg checks,
# cache cleanup, and registration of 22 tools
npm test

# Regenerate the vendored dependency copies (run after upgrading deps)
npm run vendor

# Preview vendor size without writing anything
npm run vendor:check
```

测试在单进程里加载全部 `*.test.mjs`，并用 `test/shims/cordis.mjs` + `test/resolve.mjs`
把 `@deepseek-ai/cordis` 换成轻量替身，因此不需要安装 DSH 本体。
**把 `node_modules` 整个删掉后，44 条用例依然全部通过**——这就是内置依赖的意义。

## License

MIT

## Third-party Licenses

`vendor/` 内各包均为 MIT 等宽松许可，原始版本与许可证文件已随包保留，
汇总见 `vendor/MANIFEST.json`。

## Links

- Repository: <https://github.com/lianghui-cmd/dsh-onlyoffice-bridge>
- Issues & feature requests: <https://github.com/lianghui-cmd/dsh-onlyoffice-bridge/issues>
- Contributing guide: [CONTRIBUTING.md](CONTRIBUTING.md)
- ONLYOFFICE Document Server: <https://github.com/ONLYOFFICE/DocumentServer>

如果这个插件帮你把 Office 文档接进了 Agent，欢迎点个 ⭐ Star。
