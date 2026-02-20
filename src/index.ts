#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { loadAuthConfig, getAccessToken } from './lib/auth.js';
import { authStatusToolDefinition, executeAuthStatus } from './lib/tools/auth-status.js';
import { profileToolDefinition, executeProfile } from './lib/tools/profile.js';
import { calendarToolDefinition, executeCalendar } from './lib/tools/calendar.js';
import { mailToolDefinition, executeMail } from './lib/tools/mail.js';
import { chatToolDefinition, executeChat } from './lib/tools/chat.js';
import { filesToolDefinition, executeFiles } from './lib/tools/files.js';
import { transcriptsToolDefinition, executeTranscripts } from './lib/tools/transcripts.js';

// Validate env vars at startup
try {
  loadAuthConfig();
} catch (error) {
  process.stderr.write(
    `Configuration error: ${error instanceof Error ? error.message : String(error)}\n`,
  );
  process.stderr.write('\nRequired environment variables:\n');
  process.stderr.write('  MS365_MCP_CLIENT_ID      - Azure AD application (client) ID\n');
  process.stderr.write('  MS365_MCP_TENANT_ID       - Azure AD tenant ID\n');
  process.stderr.write('\nOptional:\n');
  process.stderr.write(
    '  MS365_MCP_CLIENT_SECRET   - Azure AD client secret (confidential clients only)\n',
  );
  process.exit(1);
}

const server = new Server({ name: 'm365-mcp', version: '0.2.0' }, { capabilities: { tools: {} } });

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    authStatusToolDefinition,
    profileToolDefinition,
    calendarToolDefinition,
    mailToolDefinition,
    chatToolDefinition,
    filesToolDefinition,
    transcriptsToolDefinition,
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  try {
    const config = loadAuthConfig();

    // auth_status handles its own auth
    if (name === 'ms_auth_status') {
      const result = await executeAuthStatus(config);
      return { content: [{ type: 'text', text: result }] };
    }

    // All other tools need a valid token
    const token = await getAccessToken(config);

    let result: string;
    switch (name) {
      case 'ms_profile':
        result = await executeProfile(token);
        break;
      case 'ms_calendar':
        result = await executeCalendar(
          token,
          args as {
            date?: string;
            start?: string;
            end?: string;
          },
        );
        break;
      case 'ms_mail':
        result = await executeMail(
          token,
          args as {
            search?: string;
            count?: number;
          },
        );
        break;
      case 'ms_chat':
        result = await executeChat(
          token,
          args as {
            chat_id?: string;
            count?: number;
          },
        );
        break;
      case 'ms_files':
        result = await executeFiles(
          token,
          args as {
            path?: string;
            search?: string;
            count?: number;
          },
        );
        break;
      case 'ms_transcripts':
        result = await executeTranscripts(
          token,
          args as {
            date?: string;
            start?: string;
            end?: string;
            transcript_id?: string;
          },
        );
        break;
      default:
        return {
          content: [{ type: 'text', text: `Unknown tool: ${name}` }],
          isError: true,
        };
    }

    return { content: [{ type: 'text', text: result }] };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      content: [
        {
          type: 'text',
          text: `Error: ${message}\n\nTip: Use ms_auth_status to check or fix your connection.`,
        },
      ],
      isError: true,
    };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
