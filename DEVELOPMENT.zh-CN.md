# GenOffice PPTX 开发、发布和验收

## 交付结构

- `src/`：Node 22 MCP Server，`inspect_pptx` 接收 JSON Schema `contentEncoding: base64`、标准 PPTX MIME 类型，返回页数等结构信息；不接收客户端内部路径。
- `vendor/`：从 genoffice 提取的 PPTX 引擎及原许可、来源记录。
- `flutter_renderer/`：独立 Flutter 包，只依赖 Flutter、`archive` 和 `xml`。OOXML 文件在应用内解包，绘制基本形状、文本与位图。它不导入 Muse-Client 代码。Flutter AOT 不支持安装后加载新 Dart 代码，所以兼容版客户端需预先编入此预览器；市场安装状态决定是否启用。
- `scripts/pack_desktop.py`：生成可复现的标准 MCPB，包含 MCP Server 与渲染器源码。`scripts/publish.py`：通过 MCP Registry v0.1 上传、提交元数据。`scripts/release.py`：开发者一键检查、打包、提交。`scripts/review.py`：审核人员操作独立市场的审核接口。

本版本保留单一插件标识 `io.openmuse/genoffice-pptx`。Desktop 从 MCPB 安装 stdio 服务；Android 下载并核验同一 MCPB 摘要，启用发行版内的原生 Flutter 预览器。普通文件点击只经 Host 的统一资源协议传递字节，不调用远端 MCP，也不打开系统浏览器。插件与 Host 之间使用 MCP、MCPB、MCP Registry、JSON Schema 的公开字段；客户端将公开插件名称与其 AOT 渲染器关联，这是发行版适配关系。

## 开发与样本验证

要求 Node 22+、Python 3、Flutter 3.44。插件目录执行：

```sh
npm ci
npm test
cd flutter_renderer
flutter pub get
flutter analyze lib
OPENMUSE_PPTX_SAMPLE_PATH=/Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx flutter test
cd ..
npm run inspect -- /Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx
npm run pack:desktop
shasum -a 256 dist/genoffice-pptx-0.3.1.mcpb
```

预期样本 23 页，首屏含“年终汇报”和图片。自动化测试只检查可打开、文字、图片和翻页；它不能证明版式正确。0.3.1 的实机画面已经观察到与 PPTX 缩略图明显不符：封面的曲线裁剪和底色没有还原，第二页文字发生重叠。复杂母版、占位符继承、裁剪路径、段落布局、动画、SmartArt、嵌入媒体、SVG/WDP 仍需完善。以原稿或可信渲染结果逐页视觉比对，是宣称“正常渲染”的发布门槛。修改后必须提升 `package.json` 与 Flutter 包版本，重新打包，不能覆盖已审核版本的字节。

## 推送更新

生产公开目录为 `https://plugins.openmuseai.com`。上传和审核接口仅在生产服务器回环地址监听，先由运维建立 SSH 隧道，切勿把令牌提交到 Git 或写入命令历史：

```sh
ssh -N -L 18795:127.0.0.1:18794 <user>@<market-host>
```

在第二个终端运行；脚本会从环境变量读取令牌，若未设置则安全交互输入：

```sh
python3 scripts/release.py \
  --registry http://127.0.0.1:18795 \
  --public-origin https://plugins.openmuseai.com
```

`release.py` 依次执行 Node 测试、Flutter 检查、打包、`PUT /packages/{sha256}.mcpb` 和 `POST /v0.1/publish`，输出 submission ID 与 `pending`。只想验证包可用时加 `--build-only`。开发者也可单独调用 `scripts/publish.py`。

审核员检查包摘要、许可、测试和两端画面后，在管理后台批准；或在私有隧道上执行。尤其不能只依据页数和 smoke test 批准复杂演示文稿的视觉适配：

```sh
python3 scripts/review.py <submission-id> approve \
  --registry http://127.0.0.1:18795 \
  --note 'macOS/Android 23-page preview checked'
```

`GET https://plugins.openmuseai.com/v0.1/servers` 仅展示审核通过的版本。客户端从公开目录获取 `server.json`，下载 MCPB 并核对 `fileSha256`。下架用 `review.py <submission-id> delist --note ...`；已安装用户的本地包不会被远程删除。

## 两端人工验收

1. 安装包含 `flutter_renderer/` 的最新 OpenMuse macOS 和 Android 整包；先核对应用包 SHA-256。
2. 在两个客户端的在线插件市场刷新、安装 `GenOffice PPTX 0.3.1`；检查已安装版本和服务端 MCPB SHA-256 一致。审核前应不可安装。
3. macOS：在工作区选中 `Sample_ppt.pptx`，使用 GenOffice PPTX 打开，检查原生 Flutter 画布、23 页计数、文字、图片、翻页。打开路径不创建 WebView。
4. Android：将同一样本放入可访问的工作区，点击 `.pptx`；Host 将资源字节交给已安装的预览器，在应用内显示并可翻页。也可以从在线市场点击“使用”，经系统文件选择器选择样本。普通文件点击不上传内容到远端服务。
5. 后台下架后刷新市场，新设备不可安装；已安装包仍可使用。重新上架再确认新安装可用，并核对审计记录。

macOS 可用 Host 的线上集成测试复现正式安装与同一样本的 MCP 调用。仅设置 Registry 地址时安装到临时目录；设置安装目录时会写入当前用户的 OpenMuse 应用支持目录，退出后不会清理：

```sh
cd ../../../Muse-Client/app/openmuse_host
OPENMUSE_MCP_REGISTRY_URL=https://plugins.openmuseai.com \
OPENMUSE_MCP_INSTALL_ROOT="$HOME/Library/Application Support/com.openmuseai.office/OpenMuse/mcpb" \
OPENMUSE_PPTX_SAMPLE_PATH=/Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx \
flutter test test/mcp_market_e2e_test.dart
```

## 故障定位

- 市场无卡片：核对审核状态、公开目录版本、客户端的 Registry 提供商地址。
- Android 无法安装：确认包从 HTTPS 下载、摘要匹配、当前 App 包含 0.3.1 AOT 适配；重新安装 0.3.1 包。
- 文件点击回退系统应用：核对在线市场安装回执、PPTX MIME 类型（`application/vnd.openxmlformats-officedocument.presentationml.presentation` 或 `application/pptx`）及工作区资源读取授权。
- 画面缺图：检查图片类型。当前 Flutter 解码器支持普通 PNG/JPEG；SVG/WDP、SmartArt 等需要后续适配。
