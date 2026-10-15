# GenOffice PPTX MCP Plugin

独立 Node MCP Server。使用公开 MCP、MCPB 与 JSON Schema，不导入任何客户端代码。

```sh
npm ci
npm test
npm run pack:desktop
```

`dist/genoffice-pptx-0.2.0.mcpb` 可交给支持 MCPB 的 Desktop Host。`inspect_pptx` 处理调用方显式选择的 PPTX 内容，结构预览已通过标准 MCP 客户端与 Desktop Host 验证。`npm run serve:http` 在 `/mcp` 提供 Streamable HTTP，MCP 资源 `a2ui://pptx-preview` 提供 A2UI v0.9 消息。`roundtrip` 是引擎内部测试，不是市场工具。Node 22+；完整图形编辑及 Mobile 本地执行未实现。发布流程见[开发者指南](../../docs/DEVELOPER-GUIDE.zh-CN.md)。
