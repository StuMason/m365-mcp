import { mkdtempSync, rmSync, readFileSync, writeFileSync, statSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import type { TokenData } from '../types/tokens.js';
import {
  getConfigDir,
  loadTokens,
  saveTokens,
  deleteTokens,
  isTokenExpired,
  loadAuthConfig,
} from '../lib/auth.js';

const sampleTokens: TokenData = {
  access_token: 'access-abc-123',
  refresh_token: 'refresh-xyz-789',
  expires_at: new Date(Date.now() + 3_600_000).toISOString(), // 1 hour from now
  scopes: 'User.Read Mail.Read',
};

describe('getConfigDir', () => {
  it('creates directory if missing', () => {
    const tmp = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
    const configPath = join(tmp, 'nested', 'config');
    const originalEnv = process.env['XDG_CONFIG_HOME'];
    try {
      process.env['XDG_CONFIG_HOME'] = join(tmp, 'nested', 'config');
      const dir = getConfigDir();
      expect(dir).toBe(join(configPath, 'm365-mcp'));
      expect(existsSync(dir)).toBe(true);
    } finally {
      if (originalEnv === undefined) {
        delete process.env['XDG_CONFIG_HOME'];
      } else {
        process.env['XDG_CONFIG_HOME'] = originalEnv;
      }
      rmSync(tmp, { recursive: true, force: true });
    }
  });

  it('falls back to ~/.config when XDG_CONFIG_HOME is not set', () => {
    const originalEnv = process.env['XDG_CONFIG_HOME'];
    try {
      delete process.env['XDG_CONFIG_HOME'];
      const dir = getConfigDir();
      expect(dir).toContain('m365-mcp');
      expect(existsSync(dir)).toBe(true);
    } finally {
      if (originalEnv === undefined) {
        delete process.env['XDG_CONFIG_HOME'];
      } else {
        process.env['XDG_CONFIG_HOME'] = originalEnv;
      }
    }
  });
});

describe('saveTokens', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
  });

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('writes tokens.json', () => {
    saveTokens(sampleTokens, tmpDir);
    const filePath = join(tmpDir, 'tokens.json');
    expect(existsSync(filePath)).toBe(true);
    const raw = readFileSync(filePath, 'utf-8');
    const parsed = JSON.parse(raw) as TokenData;
    expect(parsed.access_token).toBe('access-abc-123');
  });

  it('sets chmod 600 on tokens.json', () => {
    saveTokens(sampleTokens, tmpDir);
    const filePath = join(tmpDir, 'tokens.json');
    const stats = statSync(filePath);
    const mode = stats.mode & 0o777;
    expect(mode).toBe(0o600);
  });
});

describe('loadTokens', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
  });

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('returns null for missing file', () => {
    const result = loadTokens(tmpDir);
    expect(result).toBeNull();
  });

  it('returns null for invalid JSON', () => {
    const filePath = join(tmpDir, 'tokens.json');
    writeFileSync(filePath, 'not-json!!!', 'utf-8');
    const result = loadTokens(tmpDir);
    expect(result).toBeNull();
  });

  it('round-trips with saveTokens', () => {
    saveTokens(sampleTokens, tmpDir);
    const loaded = loadTokens(tmpDir);
    expect(loaded).toEqual(sampleTokens);
  });
});

describe('deleteTokens', () => {
  let tmpDir: string;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
  });

  afterEach(() => {
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('removes the file', () => {
    saveTokens(sampleTokens, tmpDir);
    const filePath = join(tmpDir, 'tokens.json');
    expect(existsSync(filePath)).toBe(true);
    deleteTokens(tmpDir);
    expect(existsSync(filePath)).toBe(false);
  });

  it('does not throw when file does not exist', () => {
    expect(() => deleteTokens(tmpDir)).not.toThrow();
  });
});

describe('isTokenExpired', () => {
  it('returns false for token expiring in 1 hour', () => {
    const tokens: TokenData = {
      ...sampleTokens,
      expires_at: new Date(Date.now() + 3_600_000).toISOString(),
    };
    expect(isTokenExpired(tokens)).toBe(false);
  });

  it('returns true for token expiring in 90 seconds (within 2-min buffer)', () => {
    const tokens: TokenData = {
      ...sampleTokens,
      expires_at: new Date(Date.now() + 90_000).toISOString(),
    };
    expect(isTokenExpired(tokens)).toBe(true);
  });

  it('returns true for already expired token', () => {
    const tokens: TokenData = {
      ...sampleTokens,
      expires_at: new Date(Date.now() - 60_000).toISOString(),
    };
    expect(isTokenExpired(tokens)).toBe(true);
  });
});

describe('default configDir (no argument)', () => {
  let tmpDir: string;
  let originalXdg: string | undefined;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
    originalXdg = process.env['XDG_CONFIG_HOME'];
    process.env['XDG_CONFIG_HOME'] = tmpDir;
  });

  afterEach(() => {
    if (originalXdg === undefined) {
      delete process.env['XDG_CONFIG_HOME'];
    } else {
      process.env['XDG_CONFIG_HOME'] = originalXdg;
    }
    rmSync(tmpDir, { recursive: true, force: true });
  });

  it('saveTokens uses default configDir when none provided', () => {
    saveTokens(sampleTokens);
    const filePath = join(tmpDir, 'm365-mcp', 'tokens.json');
    expect(existsSync(filePath)).toBe(true);
  });

  it('loadTokens uses default configDir when none provided', () => {
    saveTokens(sampleTokens);
    const loaded = loadTokens();
    expect(loaded).toEqual(sampleTokens);
  });

  it('deleteTokens uses default configDir when none provided', () => {
    saveTokens(sampleTokens);
    deleteTokens();
    const filePath = join(tmpDir, 'm365-mcp', 'tokens.json');
    expect(existsSync(filePath)).toBe(false);
  });
});

describe('loadAuthConfig', () => {
  const originalEnv = { ...process.env };

  afterEach(() => {
    process.env = { ...originalEnv };
  });

  it('returns config when all env vars are set', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'test-client-id';
    process.env['MS365_MCP_CLIENT_SECRET'] = 'test-client-secret';
    process.env['MS365_MCP_TENANT_ID'] = 'test-tenant-id';

    const config = loadAuthConfig();
    expect(config).toEqual({
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      tenantId: 'test-tenant-id',
    });
  });

  it('throws when all env vars are missing', () => {
    delete process.env['MS365_MCP_CLIENT_ID'];
    delete process.env['MS365_MCP_CLIENT_SECRET'];
    delete process.env['MS365_MCP_TENANT_ID'];

    expect(() => loadAuthConfig()).toThrow(
      'Missing required environment variables: MS365_MCP_CLIENT_ID, MS365_MCP_CLIENT_SECRET, MS365_MCP_TENANT_ID',
    );
  });

  it('throws when only CLIENT_ID is missing', () => {
    delete process.env['MS365_MCP_CLIENT_ID'];
    process.env['MS365_MCP_CLIENT_SECRET'] = 'secret';
    process.env['MS365_MCP_TENANT_ID'] = 'tenant';

    expect(() => loadAuthConfig()).toThrow('MS365_MCP_CLIENT_ID');
  });

  it('throws when only CLIENT_SECRET is missing', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'client';
    delete process.env['MS365_MCP_CLIENT_SECRET'];
    process.env['MS365_MCP_TENANT_ID'] = 'tenant';

    expect(() => loadAuthConfig()).toThrow('MS365_MCP_CLIENT_SECRET');
  });

  it('throws when only TENANT_ID is missing', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'client';
    process.env['MS365_MCP_CLIENT_SECRET'] = 'secret';
    delete process.env['MS365_MCP_TENANT_ID'];

    expect(() => loadAuthConfig()).toThrow('MS365_MCP_TENANT_ID');
  });
});
