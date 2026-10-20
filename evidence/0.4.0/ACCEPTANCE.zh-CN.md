# 0.4.0 验收记录（2026-10-20）

## 发布

- 插件：`io.openmuse/genoffice-pptx@0.4.0`，生产目录 `https://plugins.openmuseai.com/v0.1/servers` 已展示。
- 审核提交：`98150fa4-d35e-48b5-a34d-97356337e38e`。
- MCPB SHA-256：`060aa1d49febc2d5c79b00360f448117b4d6a185677816f4ce9687d874e1985a`，14,000,049 字节。
- 包内 UI SHA-256：`15a58b7eebde5e3a3490a2a1ba8b7518fa8f07f36eac90f6cbe58d14ef175a2e`，18,344,029 字节。标准 MCP 客户端从生产资源读取后已逐字节核对。
- 原版源码：GenOffice submodule `1e643ac74416dbff8a3e75e283bd88b0c1cba5f9`，无上游源码修改。

## 已执行

| 项目 | 结果 | 证据/范围 |
| --- | --- | --- |
| 原版屏幕管线逐页比对 | 23/23 | `render-comparison.json`；同一原版 Node session builder、同一原版 Konva 屏幕画布；阈值 0.15 的不同像素比例最大 2.746% |
| 原版文本框编辑 → 标准 download-file → 重开 | 通过 | 独立 MCP Apps Host，实际操作原版 contentEditable；检查保存的 OOXML 和重开的 RenderSlide |
| macOS 原生文件桥保存 | 通过 | macOS smoke 应用里修改备注、点击保存，界面显示“已保存”，磁盘 PPTX notes XML 确认改动 |
| 完整 macOS 应用从生产市场安装 | 通过 | 安装回执摘要等于上面的 MCPB；Workspace 内打开样本，第 4/23 页截图分别见 `macos-fullapp-page04.png`、`macos-fullapp-page23.png` |
| Android 原生 WebView 从生产市场安装并渲染 | 通过 | 模拟器 Android 34，使用正式 Mobile 市场安装器和正式 Offline App View；`android-production-page01.png`，安装记录摘要一致 |
| Android 离线冷启动 | 通过 | `TEST_OFFLINE=true` 禁止 Dart HttpClient 创建；从安装记录、审核后的本地 HTML 和本地 PPTX 重开；原生层拦截外部网络，见 `android-offline-page01.png` |
| Workspace 资源分发回归 | 通过 | Desktop 输出真实 PPTX MIME；Mobile 将非 Markdown 资源交给 Host，并传递大于 2 MiB 的资源；授权句柄不传给插件 |
| 完整 macOS/Android 包 | 通过 | 正式构建脚本、静态分析和签名/资源门禁通过，使用 lib/main.dart 主入口；不是 smoke runner |
| Android 物理设备应用更新 | 通过 | 以 `adb install -r` 覆盖安装，未清除应用数据；真机的插件点击与手势编辑不作为本次自动化验收声明 |

## 发布过程中修复的问题

1. macOS 空白：WKWebView 无安全来源时没有 `crypto.randomUUID`，React 启动失败。Host 改为安全的合成本地文档来源并限制导航；插件提供 sandbox Web Storage 适配。
2. Android 内存：不将 20+ MiB HTML/base64 大字符串反复送进 loadDataWithBaseURL，改为 AssetLoader 的本地资源和流式工具输入。
3. MIME/路由：删除 GenOffice ID → Flutter renderer 的专用映射，按标准 JSON Schema 的 MIME 声明发现安装能力。Workspace 不再只分发 Markdown，也不再被 2 MiB 预览上限挡住。
4. 发布字节不一致：后续构建产物与包内 HTML 只差 13 字节的 `use strict`，Mobile 校验正确拒绝。部署现从审核 MCPB 提取 UI；`stage_service.py` 避免再次误部署后续构建结果。

## 仍需明确的边界

- 同一渲染代码不等于任意字体逐像素一致。样本已没有先前几何、图片裁剪和透明背景错乱；字体替代和抗锯齿仍有差异，不能宣称任意 PPTX 完美还原。
- 本次不是所有 GenOffice 菜单的功能认证。AI、系统剪贴板以及部分依赖 Electron 的菜单尚未完整接通。
- Mobile 远端 Workspace 的保存目前导出至 App Documents；未宣称能把编辑写回 Desktop 远端工作区。用户授予的本地文件路径可以回写。
- 验收期间原始工作区样本发生过修改（第 1/23 页 XML）。逐页比对使用先前冻结的样本副本，其 SHA 在 JSON 中记录；没有用比较脚本覆盖用户源文件。
