# 0.3.1 实机验收记录（2026-10-19）

测试文件：`/Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx`，SHA-256 `d712b4672db3634457e1c3ee3d9d8485a0cfa3298edcacffbc4fb568c1b36afb`。

- `android-market-0.3.1.png`：Android 正式商店显示已安装并可使用 0.3.1。
- `source-thumbnail.jpeg`：原 PPTX 自带的封面缩略图，用于比对。
- `android-sample-page-1.png`：在 OpenMuse 内打开第 1/23 页。图像与文字出现，但版式、裁剪、底色与原稿缩略图明显不符。
- `android-sample-page-2.png`：可翻至第 2/23 页，但文字重叠。

因此“安装、打开、翻页”的功能验收通过；“正确还原指定 PPTX 的视觉渲染”未通过。此证据不能用于宣称 PowerPoint 兼容性完成。

macOS 的正式 Registry 下载、MCPB 安装和 `inspect_pptx` 工具调用已由 `Muse-Client/app/openmuse_host/test/mcp_market_e2e_test.dart` 在本机对生产目录执行并通过，识别到服务端版本 0.3.1。测试通过 `OPENMUSE_MCP_INSTALL_ROOT` 安装到当前用户 OpenMuse 的应用支持目录，并以同一个样本调用工具。原生 Flutter viewer 的 1/23、2/23 页切换也在 macOS Flutter 测试运行器中通过。完整 App 已构建；前台打开样本与视觉渲染尚无有效截图证据。
