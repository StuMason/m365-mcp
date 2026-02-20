import { executeServerInfo } from '../../lib/tools/server-info.js';

describe('executeServerInfo', () => {
  it('returns server version and metadata', () => {
    const result = executeServerInfo();

    expect(result).toContain('# m365-mcp v');
    expect(result).toContain(`Node: ${process.version}`);
    expect(result).toContain(`Platform: ${process.platform} ${process.arch}`);
  });

  it('lists all available tools', () => {
    const result = executeServerInfo();

    expect(result).toContain('ms_auth_status');
    expect(result).toContain('ms_profile');
    expect(result).toContain('ms_calendar');
    expect(result).toContain('ms_mail');
    expect(result).toContain('ms_chat');
    expect(result).toContain('ms_files');
    expect(result).toContain('ms_transcripts');
    expect(result).toContain('ms_server_info');
    expect(result).toContain('Tools (8)');
  });

  it('shows environment variable status without exposing values', () => {
    const result = executeServerInfo();

    expect(result).toContain('MS365_MCP_CLIENT_ID:');
    expect(result).toContain('MS365_MCP_TENANT_ID:');
    // Should show 'set' or 'not set', never the actual value
    expect(result).not.toMatch(/MS365_MCP_CLIENT_ID: [a-f0-9-]{10,}/);
  });

  it('reports env var as set when present', () => {
    const original = process.env['MS365_MCP_TIMEZONE'];
    try {
      process.env['MS365_MCP_TIMEZONE'] = 'Europe/London';
      const result = executeServerInfo();
      expect(result).toContain('MS365_MCP_TIMEZONE: Europe/London');
    } finally {
      if (original === undefined) {
        delete process.env['MS365_MCP_TIMEZONE'];
      } else {
        process.env['MS365_MCP_TIMEZONE'] = original;
      }
    }
  });

  it('reports default for MS365_MCP_REDIRECT_URL when not set', () => {
    const original = process.env['MS365_MCP_REDIRECT_URL'];
    try {
      delete process.env['MS365_MCP_REDIRECT_URL'];
      const result = executeServerInfo();
      expect(result).toContain('MS365_MCP_REDIRECT_URL: default (dynamic port)');
    } finally {
      if (original === undefined) {
        delete process.env['MS365_MCP_REDIRECT_URL'];
      } else {
        process.env['MS365_MCP_REDIRECT_URL'] = original;
      }
    }
  });
});
