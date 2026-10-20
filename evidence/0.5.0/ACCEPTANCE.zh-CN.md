# 0.5.0 Desktop Agent 联调验收

日期：2026-10-20。实际应用为 `lib/main.dart` 构建的完整 macOS Release，非独立浏览器演示。

## 生产发布

- Registry：`https://plugins.openmuseai.com`
- 插件：`io.openmuse/genoffice-pptx`，0.5.0
- 审核单：`0b8e53f7-24f4-4c7d-a5ec-4666c50a240b`，approved
- MCPB SHA-256：`b37feb0ec3b422e00c77130333c6b1b509ceb4d5eb09725d7770a7e3e9bd67fb`
- 已审核 HTML SHA-256：`819293ab9fb7ac887e033f895913251a61f8efc4678735c073c69ad90fd10b88`
- 生产 Node 服务从已审核 MCPB 提取 HTML 部署。Desktop 经真实在线 Registry 下载、验包、发现并打开插件；在线安装测试通过，用时约 40 秒。

## 多设备连接

生产设备目录服务的 DSN 固化了旧 Docker IP，Postgres 重建后从 `172.18.0.3` 变成 `172.18.0.10`，导致注册/列表 HTTP 500。已改成按配置的容器与网络解析当前地址，每次连接重新获取，并设置连接超时。客户端的底层 Socket/HTTP 错误统一进入重试逻辑，避免 loading 快照卡住。

实际账号设置页显示：macOS 与 Android 在线、实时通道已连接、本地传输已就绪、公网通道已连接。见 `desktop-devices-online.png`。

## DSH 白屏

实际 WKWebView 记录了 **HTTP 431**，DOM 中不存在应用 root。原因是 `.default()` 共享持久 Cookie 仓库：DSH 为每个随机端口创建独立 Cookie 名，但 HTTP Cookie 按域名而不按端口发送，重启累积的 `127.0.0.1` Cookie 超过 Node 请求头限制。

修复：每个 DSH WebView 使用 `.nonPersistent()` 独立 Cookie 仓库，正常认证后加载页面；HTTP 错误和页面进程退出显示重试入口。同时修复启动探针的认证 Cookie 交换、账号切换取消旧启动任务、失效挂载目录拖累整个 Workspace 注册的问题。

修复后的实际嵌入页面返回 HTTP 200，rootChildren=1；同屏可操作 DSH 与 PPTX，见 `desktop-dsh-restored.png`。未提高 Node 请求头上限，也未禁用认证。

## 真实 DSH 对话修改

1. 在完整 Desktop 的本地资源搜索中打开 `PPTX-Agent-acceptance.pptx`（冻结样本的 Workspace 副本）。
2. 在现有 DSH 对话中通过其正式 Native Gateway 投递验收指令，实际配置的模型执行。
3. 模型仅调用六次 `mcp__resources__*` 工具：发现活动资源、发现插件能力、读取第一页、`pptx_edit_text`、`pptx_save`、再次读取。该验收回合没有 shell 或直接文件写入调用。
4. `sp_5` 从 `20XX` 改为 `2026 DSH`，revision 1→2；保存后 dirty=false。
5. 从 Host 的实际 WKWebView 工具通道读取并确认新文字，读取磁盘 ZIP 的 `ppt/slides/slide1.xml` 确认保存内容。

证据：`dsh-conversation.json`、`dsh-tool-trace.json`、`desktop-dsh-edited-saved.png`。对话投递使用实际 DSH Gateway；原生 WebView 内的 Agent 按钮鼠标自动化受 macOS AXError.notImplemented 限制，没有把自动点击该按钮宣称为通过。按钮的标准 ui/message 行为在独立 MCP Apps TCK 中通过。

## 标签与保存持久性

在同一真实 Desktop 编辑器内通过标准工具创建未保存备注草稿（revision 3、dirty=true），点击切换到已打开的 `test.txt`，再点击回 PPTX：仍为 revision 3、dirty=true，证明没有重载。随后撤销临时草稿并保存。

关闭该 PPTX 标签后，旧 URI 的工具调用被拒绝（-32602）。再次通过资源搜索打开，得到新会话 revision 1、dirty=false，但保存过的 `2026 DSH` 仍存在。见 `tab-retention.json`。

## 原版渲染回归

独立标准 MCP Apps 测试 Host 对已审核 HTML 验证：23/23 页、原版画布对照、手动编辑保存重开、View 工具发现、修改/保存/撤销、旧 revision 拒绝、ui/message、原版聊天面板移除均通过。

冻结样本 SHA：`d712b4672db3634457e1c3ee3d9d8485a0cfa3298edcacffbc4fb568c1b36afb`。23 页最大差异像素比例约 2.746%，沿用 0.4.0 的字体/浏览器差异；这不是逐像素完全一致或全部 Electron 功能的声明。原始样本仍可能被用户自己的 GenOffice/DSH 对话修改，本验收回合使用上述副本。

## 应用包与限制

完整应用 ZIP：`Muse-Client/dist/OpenMuse-macos.zip`。SHA：`0a577933313d07d1aebd23ae94b3021cb3e95380560cf970d4d21ce65390105d`。

macOS 本机构建为 ad hoc 签名，当前机器没有有效的代码签名身份；Keychain 数据保护权限仍返回 -34018，登录可在内存使用，但重启需要重新登录。没有降级到明文保存凭证。

本阶段验证 macOS 的 Agent 联动；Android 的同账号在线已核对，未把它表述为 Android Agent 编辑链路验收。
