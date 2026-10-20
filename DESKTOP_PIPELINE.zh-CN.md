# Desktop：复用 GenOffice 原版 PPTX 管线

## 结论和已验证事实

Desktop 版应复用 GenOffice 的 `pptx-engine → pptx-render → React/Konva`，以及原编辑器的交互和保存代码。插件负责打包、启动、授权文件、把修改后的文件交还 Host，不再维护另一套 Flutter PPTX 绘制规则。

当前 `flutter_renderer/lib/src/deck.dart` 和 `viewer.dart` 是独立实现的简化解析与画布。它们没有使用 `@genoffice/pptx-render` 的 RenderTree，也没有使用原版 `SlideCanvas`、`SlideThumb` 和 `StaticNode`。因此文字排版、主题、母版、图片效果、裁剪和对象变换会与 GenOffice 分叉。这个差异来自迁移/抽取范围，并非原版引擎对样本文件的解析错误。

在当前 Mac 上，可运行：

```sh
python3 scripts/verify_upstream_desktop.py \
  /Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx \
  --out /tmp/openmuse-genoffice-upstream-check
```

原版 CLI 通过应用自己的 Electron 导出管线生成 23 张 960×540 PNG；`setText` 后另存 PPTX、重新打开，文本仍存在，页数仍是 23。该脚本是原版基准测试，**不表示 OpenMuse 当前插件已经完成嵌入或编辑**。

## Desktop 目标结构

```text
OpenMuse Workspace 文件资源
  → Host 选中已安装的 PPTX 能力
  → 标准 MCP 工具传入 PPTX 字节/资源句柄
  → 插件本地进程持有 GenOffice 原版编辑会话
  → 原版 React/Konva 编辑器在 OpenMuse 内的隔离 Web 视图显示
  → 原版保存管线产生 PPTX 字节
  → Host 写回用户授权的 Workspace 资源
```

插件与 Host 的契约限于标准 MCP 工具/资源及 MCP Apps UI；它不能依赖 `Muse-Client` 源码、插件 ID 特判或 Host 私有字段。Electron IPC 可以继续作为**插件内部**的实现细节，但不能穿过 Host 边界。文件访问须由 Host 对用户选定的资源授权；Web 视图没有任意本机文件访问能力。

原版 `apps/slides/src/renderer/App.tsx` 大量调用 Electron preload 的 `window.slidesApi`。要完整复用编辑器，需将原版 Electron 主进程作为插件 Desktop sidecar 保留，并给 Web 视图提供透明的 IPC 传输适配层，使已有 `slidesApi` 方法和事件仍由原版后端处理。只复用 `pptx-engine` 的 `openPptx` 或仅将 `RenderSlide` 转成 Flutter Widget，均不能实现相同的编辑效果。

macOS Host 现有 `McpAppWebView.swift` 只接受同源 HTTPS/localhost 页面，且没有 MCP Apps 消息桥。它只能显示页面，尚不能直接承载完整原版编辑器。实现时需补齐**标准** MCP Apps 的工具调用/资源消息桥，并确保插件页面限源、限权限。允许本地插件进程向 Web 视图提供静态资源与内部 IPC 服务；不能把临时 localhost 地址做成用户需手工配置的服务，也不能把这条 Desktop 方式套用到 Mobile。

当前市场 MCPB 上限为 32 MiB；已安装 GenOffice 的 CLI、`app.asar`、Electron Frameworks 分别约为 39、24、276 MiB（未压缩）。因此不能把整套 GenOffice.app 塞入现有 MCPB。应保留小型 MCPB 作为标准能力入口，另交付经过签名与审核的 macOS 平台产物，并在商店安装事务中校验其摘要与版本。这仍是同一个插件 ID 的 Desktop 分发物。

## 实施顺序与验收门槛

1. 固定 GenOffice 源码版本，打包原版 `pptx-engine`、`pptx-render`、`apps/slides` 编辑器、字体和所需 Electron 后端；不改写绘图规则。
2. 实现插件内部 IPC 适配层、标准 MCP/MCP Apps 边界、签名平台包安装；让 Host 在应用窗口内打开编辑器。
3. 使用 `Sample_ppt.pptx` 在 OpenMuse 内对照原版首屏、23 页缩略图、顶部工具栏和编辑画布。对同一缩放与字体环境的截图做视觉差异检查。
4. 在 OpenMuse 内修改第一页文字，保存、关闭、重新打开；确认修改保留、23 页和媒体资源完整，并验证撤销/重做及至少一种对象编辑。
5. 上述 Desktop 门槛通过后，再设计 Mobile/Web 的受限实现；不把 Electron 或本机 localhost 运行模型移植到 Mobile。

在完成第 2–4 步前，不应将当前 Flutter 预览称为“与 GenOffice 一致的 Desktop 编辑器”，也不应发布为满足高保真编辑的版本。
