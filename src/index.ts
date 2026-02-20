#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { loadAuthConfig, getAccessToken } from './lib/auth.js';
import { profileToolDefinition, executeProfile } from './lib/tools/profile.js';
import { calendarToolDefinition, executeCalendar } from './lib/tools/calendar.js';

const server = new Server({ name: 'm365-mcp', version: '0.1.0' }, { capabilities: { tools: {} } });

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [profileToolDefinition, calendarToolDefinition],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const config = loadAuthConfig();
    const token = await getAccessToken(config);

    let result: string;

    switch (request.params.name) {
      case 'ms_profile':
        result = await executeProfile(token);
        break;
      case 'ms_calendar':
        result = await executeCalendar(
          token,
          (request.params.arguments ?? {}) as {
            date?: string;
            start?: string;
            end?: string;
          },
        );
        break;
      default:
        return {
          content: [{ type: 'text', text: `Unknown tool: ${request.params.name}` }],
          isError: true,
        };
    }

    return { content: [{ type: 'text', text: result }] };
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    return {
      content: [{ type: 'text', text: `Error: ${message}` }],
      isError: true,
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
