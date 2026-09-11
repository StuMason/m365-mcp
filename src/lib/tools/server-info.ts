import { z } from 'zod';
import { getVersion, getBuildInfo } from '../version.js';

export const serverInfoToolDefinition = {
  name: 'ms_server_info',
  title: 'Server Info',
  description:
    'Returns m365-mcp server metadata: version, available tools, and runtime info. Useful for debugging.',
  inputSchema: z.object({}),
  annotations: {
    title: 'Server Info',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  },
};

export function executeServerInfo(names: string[]): string {
  const version = getVersion();
  const lines: string[] = [];

  lines.push(`# m365-mcp v${version}`);
  const build = getBuildInfo();
  if (build) {
    // Two builds of the same unreleased version are otherwise indistinguishable,
    // which makes "am I testing the fix?" unanswerable during a review cycle.
    const dirty = build.dirty ? ' (uncommitted changes)' : '';
    lines.push(`Build: ${build.commit ?? 'unknown'} on ${build.branch ?? 'unknown'}${dirty}`);
    lines.push(`Built: ${build.builtAt}`);
  }
  lines.push('');
  lines.push(`Node: ${process.version}`);
  lines.push(`Platform: ${process.platform} ${process.arch}`);
  lines.push('');
  lines.push(`## Tools (${names.length})`);
  for (const name of names) {
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
