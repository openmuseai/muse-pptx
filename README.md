# GenOffice PPTX 0.3.1

独立第三方 PPTX 插件。标准 MCPB 内含 GenOffice 的 `inspect_pptx` MCP 工具以及原生 Flutter 只读预览器源码。插件不引用客户端源码或私有协议字段。兼容的 macOS/Android 客户端在构建时编入 Flutter AOT 预览器，市场安装并校验 MCPB 后启用它；文件经 Host 的统一资源读取流程在应用内展示，不发送到远端 MCP 服务。

```sh
npm ci
npm test
cd flutter_renderer && flutter pub get && flutter test && cd ..
npm run pack:desktop
```

指定样本 `/Users/mac/src/ioffice/samples/ppt-render/flutter/assets/Sample_ppt.pptx` 有 23 页；`flutter_renderer/test/deck_test.dart` 会在样本存在时验证页数、文字与图片。预览当前支持基本图形、文字、位图和上一页/下一页；复杂动画、SmartArt、SVG/WDP 媒体、精确母版效果及编辑保存尚未覆盖。发布与两端验收步骤见 [DEVELOPMENT.zh-CN.md](DEVELOPMENT.zh-CN.md)。
