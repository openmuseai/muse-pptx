import { McpServer } from '@modelcontextprotocol/server'
import { z } from 'zod'
import { inspectBytes } from './worker.js'

const uiResource = 'a2ui://pptx-preview'
const mimeType = 'application/a2ui+json'
const catalogId = 'https://a2ui.org/specification/v0_9/catalogs/basic/catalog.json'
const template = [
  { version: 'v0.9', createSurface: { surfaceId: 'pptx-preview', catalogId } },
  { version: 'v0.9', updateComponents: { surfaceId: 'pptx-preview', components: [
    { id: 'root', component: 'Column', children: ['title', 'summary'] },
    { id: 'title', component: 'Text', text: 'PPTX 结构预览', variant: 'h2' },
    { id: 'summary', component: 'Text', text: { path: '/summary' } },
  ] } },
]

export function buildServer() {
  const server = new McpServer(
    { name: 'io.openmuse/genoffice-pptx', version: '0.3.1' },
    { capabilities: { tools: {}, resources: {} } },
  )
  server.registerResource(
    'PPTX A2UI preview', uiResource,
    { title: 'PPTX A2UI preview', mimeType },
    async () => ({ contents: [{ uri: uiResource, mimeType, text: JSON.stringify(template) }] }),
  )
  server.registerTool(
    'inspect_pptx',
    {
      title: 'Inspect PowerPoint slides',
      description: 'Inspect a user-selected PPTX document and return its slide structure.',
      inputSchema: z.object({ content: z.base64().meta({
        contentMediaType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }) }),
      outputSchema: z.object({
        slides: z.number(),
        size: z.unknown(),
        sha256: z.string(),
        pages: z.array(z.object({ index: z.number(), elements: z.number() })),
      }),
      _meta: { ui: { resourceUri: uiResource, mimeType } },
    },
    async ({ content }) => {
      const output = await inspectBytes(Buffer.from(content, 'base64'))
      const update = [{ version: 'v0.9', updateDataModel: {
        surfaceId: 'pptx-preview', path: '/', value: { summary: `${output.slides} slides` },
      } }]
      return {
        content: [
          { type: 'text' as const, text: JSON.stringify(output) },
          { type: 'resource' as const, resource: {
            uri: `${uiResource}/data`, mimeType, text: JSON.stringify(update),
          } },
        ],
        structuredContent: output,
        _meta: { ui: { resourceUri: uiResource, mimeType } },
      }
    },
  )
  return server
}
