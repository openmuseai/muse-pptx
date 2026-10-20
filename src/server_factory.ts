import { readFile } from 'node:fs/promises'
import { McpServer } from '@modelcontextprotocol/server'
import { z } from 'zod'
import { inspectBytes } from './worker.js'

const uiResource = 'ui://genoffice-pptx/editor'
const mimeType = 'text/html;profile=mcp-app'

export function buildServer() {
  const server = new McpServer(
    { name: 'io.openmuse/genoffice-pptx', version: '0.5.0' },
    { capabilities: { tools: {}, resources: {} } },
  )
  server.registerResource(
    'GenOffice PowerPoint editor', uiResource,
    { title: 'GenOffice PowerPoint editor', mimeType },
    async () => {
      let text: string
      try {text=await readFile(new URL('./web/index.html',import.meta.url),'utf8')}
      catch {text=await readFile(new URL('../dist/web/index.html',import.meta.url),'utf8')}
      return {contents:[{uri:uiResource,mimeType,text}]}
    },
  )
  server.registerTool(
    'inspect_pptx',
    {
      title: 'Inspect PowerPoint slides',
      description: 'Inspect a user-selected PPTX document and return its slide structure.',
      inputSchema: z.object({ content: z.base64().meta({
        contentMediaType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      }), name: z.string().optional() }),
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
      return {
        content: [
          { type: 'text' as const, text: JSON.stringify(output) },

        ],
        structuredContent: output,
        _meta: { ui: { resourceUri: uiResource, mimeType } },
      }
    },
  )
  return server
}
