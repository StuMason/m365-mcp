#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { ListToolsRequestSchema, CallToolRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { loadAuthConfig, getAccessToken } from './lib/auth.js';
import { getVersion } from './lib/version.js';
import { authStatusToolDefinition, executeAuthStatus } from './lib/tools/auth-status.js';
import { profileToolDefinition, executeProfile } from './lib/tools/profile.js';
import { calendarToolDefinition, executeCalendar } from './lib/tools/calendar.js';
import { mailToolDefinition, executeMail } from './lib/tools/mail.js';
import { chatToolDefinition, executeChat } from './lib/tools/chat.js';
import { filesToolDefinition, executeFiles } from './lib/tools/files.js';
import { transcriptsToolDefinition, executeTranscripts } from './lib/tools/transcripts.js';
import { serverInfoToolDefinition, executeServerInfo } from './lib/tools/server-info.js';
import { scheduleToolDefinition, executeSchedule } from './lib/tools/schedule.js';
import { sharepointToolDefinition, executeSharepoint } from './lib/tools/sharepoint.js';
import { teamsToolDefinition, executeTeams } from './lib/tools/teams.js';
import { tasksToolDefinition, executeTasks } from './lib/tools/tasks.js';
import { peopleToolDefinition, executePeople } from './lib/tools/people.js';

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

const server = new Server(
  { name: 'm365-mcp', title: 'Microsoft 365', version: getVersion() },
  {
    capabilities: { tools: {} },
    instructions:
      'Read-only access to the signed-in user\u2019s own Microsoft 365 data via the Graph API. ' +
      'Every tool acts as that user and cannot reach anyone else\u2019s mailbox, chats or files. ' +
      'Start with ms_auth_status if a call reports an authentication problem. ' +
      'Many tools are progressive: called with no arguments they list items with IDs, ' +
      'and those IDs are passed back to drill into detail.',
  },
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    authStatusToolDefinition,
    profileToolDefinition,
    calendarToolDefinition,
    mailToolDefinition,
    chatToolDefinition,
    filesToolDefinition,
    transcriptsToolDefinition,
    scheduleToolDefinition,
    sharepointToolDefinition,
    teamsToolDefinition,
    tasksToolDefinition,
    peopleToolDefinition,
    serverInfoToolDefinition,
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args = {} } = request.params;

  try {
    const config = loadAuthConfig();

    // Tools that don't require a valid token
    if (name === 'ms_auth_status') {
      const result = await executeAuthStatus(config);
      return { content: [{ type: 'text', text: result }] };
    }

    if (name === 'ms_server_info') {
      const result = executeServerInfo();
      return { content: [{ type: 'text', text: result }] };
    }

    // All other tools need a valid token
    const token = await getAccessToken(config);

    let result: string;
    switch (name) {
      case 'ms_profile':
        result = await executeProfile(
          token,
          args as {
            include?: string[];
          },
        );
        break;
      case 'ms_calendar':
        result = await executeCalendar(
          token,
          args as {
            date?: string;
            start?: string;
            end?: string;
            event_id?: string;
            calendars?: boolean;
          },
        );
        break;
      case 'ms_mail':
        result = await executeMail(
          token,
          args as {
            search?: string;
            count?: number;
            message_id?: string;
            folder?: string;
            folders?: boolean;
            attachments?: boolean;
            filter?: string;
          },
        );
        break;
      case 'ms_chat':
        result = await executeChat(
          token,
          args as {
            chat_id?: string;
            count?: number;
            members?: boolean;
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
            item_id?: string;
            shared?: boolean;
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
            offset?: number;
            length?: number;
          },
        );
        break;
      case 'ms_schedule':
        result = await executeSchedule(
          token,
          args as {
            emails: string[];
            date?: string;
            start?: string;
            end?: string;
            interval?: number;
          },
        );
        break;
      case 'ms_sharepoint':
        result = await executeSharepoint(
          token,
          args as {
            search?: string;
            site_id?: string;
            list_id?: string;
            count?: number;
          },
        );
        break;
      case 'ms_teams':
        result = await executeTeams(
          token,
          args as {
            team_id?: string;
            channel_id?: string;
            message_id?: string;
            count?: number;
          },
        );
        break;
      case 'ms_tasks':
        result = await executeTasks(
          token,
          args as {
            list_id?: string;
            planner?: boolean;
            include_completed?: boolean;
            count?: number;
          },
        );
        break;
      case 'ms_people':
        result = await executePeople(
          token,
          args as {
            search?: string;
            user?: string;
            groups?: boolean;
            count?: number;
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
