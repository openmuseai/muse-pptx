# GenOffice PPTX 外部插件（纵向样例）

插件 ID：`com.openmuse.pptx`。当前 Desktop 包含从 GenOffice 上游提交 `5eab98adc8d0ebf6ecea687e42400a791ef58d64` 提取的 `pptx-engine`、`zip-gate` 和唯一必需的 `agent-core/redact-range`。上游 [GenOffice](https://github.com/genspark-ai/genoffice) 的根许可证是 Apache-2.0；完整许可证和 NOTICE 保留在本目录，嵌套许可仍须在公开发行前逐项核对。

## 已实现

```sh
npm ci --ignore-scripts
npm test
npm run pack:desktop
```

`dist/worker.mjs` 是独立的 Node 22+ ESM bundle，支持 `inspect <file.pptx>` 与 `roundtrip <file.pptx> <out.pptx>`；输入上限 64 MiB。`dist/catalog.json` 和 `.omplugin` 是 **本地未签名开发夹具**。可由 `Muse-Client/app/openmuse_host/bin/openmuse_plugin.dart` 的本地目录路径安装；安装后可手动运行工作区 `PPTX/worker.mjs`。

Desktop 应用的开发版可设置 `OPENMUSE_PPTX_DEV_CATALOG=/绝对路径/dist/catalog.json` 后启动，市场将出现 PPTX 开发预览卡片。安装后 Host 注册只读结构预览适配器，工作区的 `.pptx` 文件可由它打开。当前没有图形编辑能力。

Desktop 样例只证明 GenOffice 核心包能被拆出、构建、安装和调用。原应用的 Electron/React/Konva 编辑器、Host 对外部 worker 的自动激活、权限 Broker、商场远端签名目录尚未接入。不要将当前包上传公开商店。

Android 预览另由 `mobile/pptx.js` 生成 `pptx-mobile.json`；它提供 A2UI 入口，PPTX 文件解析调用随 App 发货的 Rust FFI。`Muse-Client` 的开发版内置市场夹具从该源码同步，并在安装后开启 Cloud Workspace 的只读 PPTX 入口。iOS 尚未标记支持。

## 源码边界

`vendor/pptx-engine` 与 `vendor/zip-gate` 保持上游包布局；`vendor/agent-core` 仅包含引擎直接依赖的 `redact-range`。升级时先比对固定提交及许可证，再重放本目录改动并跑上游测试。业务适配只放在 `src/` 和 `scripts/`，不在 vendor 内直接实现 OpenMuse 逻辑。
