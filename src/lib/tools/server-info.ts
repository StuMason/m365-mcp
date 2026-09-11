import { getVersion } from '../version.js';

export const serverInfoToolDefinition = {
  name: 'ms_server_info',
  title: 'Server Info',
  description:
    'Returns m365-mcp server metadata: version, available tools, and runtime info. Useful for debugging.',
  inputSchema: {
    type: 'object' as const,
    properties: {},
  },
  annotations: {
    title: 'Server Info',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },
};

const TOOL_NAMES = [
  'ms_auth_status',
  'ms_profile',
  'ms_calendar',
  'ms_mail',
  'ms_chat',
  'ms_files',
  'ms_transcripts',
  'ms_schedule',
  'ms_sharepoint',
  'ms_teams',
  'ms_tasks',
  'ms_people',
  'ms_server_info',
];

export function executeServerInfo(): string {
  const version = getVersion();
  const lines: string[] = [];

  lines.push(`# m365-mcp v${version}`);
  lines.push('');
  lines.push(`Node: ${process.version}`);
  lines.push(`Platform: ${process.platform} ${process.arch}`);
  lines.push('');
  lines.push(`## Tools (${TOOL_NAMES.length})`);
  for (const name of TOOL_NAMES) {
    lines.push(`- ${name}`);
  }
  lines.push('');
  lines.push('## Environment');
  lines.push(`MS365_MCP_CLIENT_ID: ${process.env['MS365_MCP_CLIENT_ID'] ? 'set' : 'not set'}`);
  lines.push(
    `MS365_MCP_CLIENT_SECRET: ${process.env['MS365_MCP_CLIENT_SECRET'] ? 'set' : 'not set'}`,
  );
  lines.push(`MS365_MCP_TENANT_ID: ${process.env['MS365_MCP_TENANT_ID'] ? 'set' : 'not set'}`);
  lines.push(
    `MS365_MCP_REDIRECT_URL: ${process.env['MS365_MCP_REDIRECT_URL'] || 'default (dynamic port)'}`,
  );
  lines.push(`MS365_MCP_TIMEZONE: ${process.env['MS365_MCP_TIMEZONE'] || 'auto'}`);

  return lines.join('\n');
}
