#!/usr/bin/env node

import { StdioServerTransport } from '@modelcontextprotocol/server/stdio';
import { loadAuthConfig } from './lib/auth.js';
import { buildServer } from './lib/server.js';

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

const transport = new StdioServerTransport();
await buildServer().connect(transport);
