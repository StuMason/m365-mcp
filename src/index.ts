#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';

const server = new Server({ name: 'm365-mcp', version: '0.1.0' }, { capabilities: { tools: {} } });

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => ({
  content: [{ type: 'text', text: `Unknown tool: ${request.params.name}` }],
  isError: true,
}));

const transport = new StdioServerTransport();
await server.connect(transport);
