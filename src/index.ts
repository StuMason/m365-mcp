#!/usr/bin/env node

import { McpServer } from '@modelcontextprotocol/server';
import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import { loadAuthConfig, getAccessToken } from './lib/auth.js';
import { getVersion } from './lib/version.js';
import { authStatusToolDefinition, executeAuthStatus } from './lib/tools/auth-status.js';
import { profileToolDefinition, executeProfile } from './lib/tools/profile.js';
import { calendarToolDefinition, executeCalendar } from './lib/tools/calendar.js';
import { mailToolDefinition, executeMail } from './lib/tools/mail.js';
import { chatToolDefinition, executeChat } from './lib/tools/chat.js';
import { filesToolDefinition, executeFiles } from './lib/tools/files.js';
import { transcriptsToolDefinition, executeTranscripts } from './lib/tools/transcripts.js';
import { scheduleToolDefinition, executeSchedule } from './lib/tools/schedule.js';
import { sharepointToolDefinition, executeSharepoint } from './lib/tools/sharepoint.js';
import { teamsToolDefinition, executeTeams } from './lib/tools/teams.js';
import { tasksToolDefinition, executeTasks } from './lib/tools/tasks.js';
import { peopleToolDefinition, executePeople } from './lib/tools/people.js';
import { serverInfoToolDefinition, executeServerInfo } from './lib/tools/server-info.js';
import { toolNames } from './lib/tools/index.js';

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

const server = new McpServer(
  { name: 'm365-mcp', title: 'Microsoft 365', version: getVersion() },
  {
    instructions:
      'Read-only access to the signed-in user’s own Microsoft 365 data via the Graph API. ' +
      'Every tool acts as that user and cannot reach anyone else’s mailbox, chats or files. ' +
      'Start with ms_auth_status if a call reports an authentication problem. ' +
      'Many tools are progressive: called with no arguments they list items with IDs, ' +
      'and those IDs are passed back to drill into detail.',
  },
);

type ToolResult = { content: [{ type: 'text'; text: string }]; isError?: true };

function ok(text: string): ToolResult {
  return { content: [{ type: 'text', text }] };
}

function failed(error: unknown): ToolResult {
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

/**
 * Wraps a tool that needs a Graph token.
 *
 * Acquires or refreshes the access token first — starting the browser sign-in
 * flow when there is none — then runs the tool and maps any throw onto the same
 * error text the v1 dispatch switch produced, so the failure UX is unchanged.
 */
function withToken<A>(run: (token: string, args: A) => Promise<string>) {
  return async (args: A): Promise<ToolResult> => {
    try {
      const token = await getAccessToken(loadAuthConfig());
      return ok(await run(token, args));
    } catch (error) {
      return failed(error);
    }
  };
}

// Registered in TOOL_DEFINITIONS order, so tools/list and ms_server_info agree.

// ms_auth_status manages its own auth lifecycle: it has to run when there is no
// valid token, which is the whole point of it.
server.registerTool(authStatusToolDefinition.name, authStatusToolDefinition, async () => {
  try {
    return ok(await executeAuthStatus(loadAuthConfig()));
  } catch (error) {
    return failed(error);
  }
});

server.registerTool(profileToolDefinition.name, profileToolDefinition, withToken(executeProfile));
server.registerTool(
  calendarToolDefinition.name,
  calendarToolDefinition,
  withToken(executeCalendar),
);
server.registerTool(mailToolDefinition.name, mailToolDefinition, withToken(executeMail));
server.registerTool(chatToolDefinition.name, chatToolDefinition, withToken(executeChat));
server.registerTool(filesToolDefinition.name, filesToolDefinition, withToken(executeFiles));
server.registerTool(
  transcriptsToolDefinition.name,
  transcriptsToolDefinition,
  withToken(executeTranscripts),
);
server.registerTool(
  scheduleToolDefinition.name,
  scheduleToolDefinition,
  withToken(executeSchedule),
);
server.registerTool(
  sharepointToolDefinition.name,
  sharepointToolDefinition,
  withToken(executeSharepoint),
);
server.registerTool(teamsToolDefinition.name, teamsToolDefinition, withToken(executeTeams));
server.registerTool(tasksToolDefinition.name, tasksToolDefinition, withToken(executeTasks));
server.registerTool(peopleToolDefinition.name, peopleToolDefinition, withToken(executePeople));

// ms_server_info touches nothing outside the process.
server.registerTool(serverInfoToolDefinition.name, serverInfoToolDefinition, () =>
  ok(executeServerInfo(toolNames())),
);

const transport = new StdioServerTransport();
await server.connect(transport);
