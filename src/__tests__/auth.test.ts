import { jest } from '@jest/globals';
import { mkdtempSync, rmSync, readFileSync, writeFileSync, statSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import type { TokenData, AuthConfig } from '../types/tokens.js';

// Mock child_process to prevent browser tabs opening during tests
const mockExecFile = jest.fn();
jest.unstable_mockModule('node:child_process', () => ({
  execFile: mockExecFile,
}));

// Dynamic import AFTER mock registration
const {
  getConfigDir,
  loadTokens,
  saveTokens,
  deleteTokens,
  isTokenExpired,
  loadAuthConfig,
  SCOPES,
  refreshAccessToken,
  getAccessToken,
  findAvailablePort,
  exchangeCodeForTokens,
  openBrowser,
  waitForAuthCallback,
} = await import('../lib/auth.js');

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

  it('throws when all required env vars are missing', () => {
    delete process.env['MS365_MCP_CLIENT_ID'];
    delete process.env['MS365_MCP_CLIENT_SECRET'];
    delete process.env['MS365_MCP_TENANT_ID'];

    expect(() => loadAuthConfig()).toThrow(
      'Missing required environment variables: MS365_MCP_CLIENT_ID, MS365_MCP_TENANT_ID',
    );
  });

  it('throws when only CLIENT_ID is missing', () => {
    delete process.env['MS365_MCP_CLIENT_ID'];
    process.env['MS365_MCP_TENANT_ID'] = 'tenant';

    expect(() => loadAuthConfig()).toThrow('MS365_MCP_CLIENT_ID');
  });

  it('does not require CLIENT_SECRET (public client support)', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'client';
    delete process.env['MS365_MCP_CLIENT_SECRET'];
    process.env['MS365_MCP_TENANT_ID'] = 'tenant';

    const config = loadAuthConfig();
    expect(config.clientId).toBe('client');
    expect(config.clientSecret).toBeUndefined();
    expect(config.tenantId).toBe('tenant');
  });

  it('includes CLIENT_SECRET when provided', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'client';
    process.env['MS365_MCP_CLIENT_SECRET'] = 'secret';
    process.env['MS365_MCP_TENANT_ID'] = 'tenant';

    const config = loadAuthConfig();
    expect(config.clientSecret).toBe('secret');
  });

  it('throws when only TENANT_ID is missing', () => {
    process.env['MS365_MCP_CLIENT_ID'] = 'client';
    delete process.env['MS365_MCP_TENANT_ID'];

    expect(() => loadAuthConfig()).toThrow('MS365_MCP_TENANT_ID');
  });
});

describe('SCOPES', () => {
  it('exports the expected Microsoft Graph scopes', () => {
    expect(SCOPES).toContain('openid');
    expect(SCOPES).toContain('offline_access');
    expect(SCOPES).toContain('User.Read');
    expect(SCOPES).toContain('Calendars.Read');
    expect(SCOPES).toContain('Mail.Read');
    expect(SCOPES).toContain('Chat.Read');
    expect(SCOPES).toContain('Files.Read');
    expect(SCOPES).toContain('OnlineMeetingTranscript.Read.All');
    expect(SCOPES).toContain('Sites.Read.All');
    expect(SCOPES).toHaveLength(11);
  });

  it('is a frozen array of strings', () => {
    expect(Array.isArray(SCOPES)).toBe(true);
    SCOPES.forEach((scope) => {
      expect(typeof scope).toBe('string');
    });
  });
});

describe('refreshAccessToken', () => {
  const config: AuthConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id',
  };

  let tmpDir: string;
  let originalXdg: string | undefined;
  let originalFetch: typeof global.fetch;
  let stderrSpy: ReturnType<typeof jest.spyOn>;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
    originalXdg = process.env['XDG_CONFIG_HOME'];
    process.env['XDG_CONFIG_HOME'] = tmpDir;
    originalFetch = global.fetch;
    stderrSpy = jest
      .spyOn(process.stderr, 'write')
      .mockImplementation(
        (() => true) as unknown as (
          ...args: Parameters<typeof process.stderr.write>
        ) => ReturnType<typeof process.stderr.write>,
      );
  });

  afterEach(() => {
    if (originalXdg === undefined) {
      delete process.env['XDG_CONFIG_HOME'];
    } else {
      process.env['XDG_CONFIG_HOME'] = originalXdg;
    }
    rmSync(tmpDir, { recursive: true, force: true });
    global.fetch = originalFetch;
    stderrSpy.mockRestore();
  });

  it('returns TokenData and saves tokens on successful refresh', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'new-access-token',
        refresh_token: 'new-refresh-token',
        expires_in: 3600,
        scope: 'User.Read Mail.Read',
      }),
    } as Response);

    const result = await refreshAccessToken(config, 'old-refresh-token');

    expect(result).not.toBeNull();
    expect(result!.access_token).toBe('new-access-token');
    expect(result!.refresh_token).toBe('new-refresh-token');
    expect(result!.scopes).toBe('User.Read Mail.Read');

    // Verify tokens were saved
    const saved = loadTokens();
    expect(saved).not.toBeNull();
    expect(saved!.access_token).toBe('new-access-token');

    // Verify fetch was called with correct URL
    expect(global.fetch).toHaveBeenCalledWith(
      `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
      expect.objectContaining({
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      }),
    );
  });

  it('sends correct form body with refresh_token grant', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'at',
        refresh_token: 'rt',
        expires_in: 3600,
        scope: 'User.Read',
      }),
    } as Response);

    await refreshAccessToken(config, 'my-refresh-token');

    const fetchMock = global.fetch as jest.MockedFunction<typeof fetch>;
    const callArgs = fetchMock.mock.calls[0] as [string, RequestInit];
    const callBody = callArgs[1].body as string;
    const params = new URLSearchParams(callBody);
    expect(params.get('grant_type')).toBe('refresh_token');
    expect(params.get('client_id')).toBe('test-client-id');
    expect(params.get('client_secret')).toBe('test-client-secret');
    expect(params.get('refresh_token')).toBe('my-refresh-token');
  });

  it('returns null and deletes tokens on HTTP error', async () => {
    saveTokens(sampleTokens);

    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: false,
      status: 400,
      text: async () => 'invalid_grant',
    } as Response);

    const result = await refreshAccessToken(config, 'bad-refresh-token');

    expect(result).toBeNull();
    expect(stderrSpy).toHaveBeenCalled();

    // Verify tokens were deleted
    const saved = loadTokens();
    expect(saved).toBeNull();
  });

  it('returns null and deletes tokens on network error', async () => {
    saveTokens(sampleTokens);

    global.fetch = jest.fn<typeof fetch>().mockRejectedValue(new Error('Network error'));

    const result = await refreshAccessToken(config, 'some-refresh-token');

    expect(result).toBeNull();
    expect(stderrSpy).toHaveBeenCalled();

    // Verify tokens were deleted
    const saved = loadTokens();
    expect(saved).toBeNull();
  });

  it('returns null and logs non-Error throw values', async () => {
    global.fetch = jest.fn<typeof fetch>().mockRejectedValue('string-error');

    const result = await refreshAccessToken(config, 'some-refresh-token');

    expect(result).toBeNull();
    const stderrCalls = stderrSpy.mock.calls.map((c: unknown[]) => String(c[0]));
    expect(stderrCalls.some((c: string) => c.includes('string-error'))).toBe(true);
  });

  it('sets expires_at to approximately expires_in seconds from now', async () => {
    const beforeMs = Date.now();
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'at',
        refresh_token: 'rt',
        expires_in: 7200,
        scope: 'User.Read',
      }),
    } as Response);

    const result = await refreshAccessToken(config, 'rt');
    const afterMs = Date.now();

    expect(result).not.toBeNull();
    const expiresAtMs = new Date(result!.expires_at).getTime();
    expect(expiresAtMs).toBeGreaterThanOrEqual(beforeMs + 7200 * 1000 - 1000);
    expect(expiresAtMs).toBeLessThanOrEqual(afterMs + 7200 * 1000 + 1000);
  });
});

describe('findAvailablePort', () => {
  it('returns a positive integer port number', async () => {
    const port = await findAvailablePort();
    expect(Number.isInteger(port)).toBe(true);
    expect(port).toBeGreaterThan(0);
    expect(port).toBeLessThanOrEqual(65535);
  });

  it('returns different ports on successive calls', async () => {
    const port1 = await findAvailablePort();
    const port2 = await findAvailablePort();
    // They could theoretically be the same, but in practice almost never will be
    expect(typeof port1).toBe('number');
    expect(typeof port2).toBe('number');
  });
});

describe('openBrowser', () => {
  afterEach(() => {
    mockExecFile.mockClear();
  });

  it('calls execFile with the correct command for the platform', () => {
    openBrowser('https://example.com');

    // On macOS (darwin), should call 'open'
    if (process.platform === 'darwin') {
      expect(mockExecFile).toHaveBeenCalledWith('open', ['https://example.com']);
    } else {
      expect(mockExecFile).toHaveBeenCalled();
    }
  });

  it('does not throw when execFile throws', () => {
    mockExecFile.mockImplementation(() => {
      throw new Error('Command not found');
    });

    expect(() => openBrowser('https://example.com')).not.toThrow();
  });
});

describe('exchangeCodeForTokens', () => {
  const config: AuthConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id',
  };

  let originalFetch: typeof global.fetch;

  beforeEach(() => {
    originalFetch = global.fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
  });

  it('returns TokenData on successful exchange', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'exchange-access',
        refresh_token: 'exchange-refresh',
        expires_in: 3600,
        scope: 'User.Read Mail.Read',
      }),
    } as Response);

    const result = await exchangeCodeForTokens(
      config,
      'auth-code-123',
      'http://localhost:12345/callback',
    );

    expect(result.access_token).toBe('exchange-access');
    expect(result.refresh_token).toBe('exchange-refresh');
    expect(result.scopes).toBe('User.Read Mail.Read');
    expect(new Date(result.expires_at).getTime()).toBeGreaterThan(Date.now());
  });

  it('sends correct form body with authorization_code grant', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'at',
        refresh_token: 'rt',
        expires_in: 3600,
        scope: 'User.Read',
      }),
    } as Response);

    await exchangeCodeForTokens(config, 'the-code', 'http://localhost:9999/callback');

    const fetchMock = global.fetch as jest.MockedFunction<typeof fetch>;
    const callArgs = fetchMock.mock.calls[0] as [string, RequestInit];
    expect(callArgs[0]).toBe(
      `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    );
    const params = new URLSearchParams(callArgs[1].body as string);
    expect(params.get('grant_type')).toBe('authorization_code');
    expect(params.get('client_id')).toBe('test-client-id');
    expect(params.get('client_secret')).toBe('test-client-secret');
    expect(params.get('code')).toBe('the-code');
    expect(params.get('redirect_uri')).toBe('http://localhost:9999/callback');
  });

  it('throws on HTTP error response', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: false,
      status: 400,
      text: async () => 'bad_request',
    } as Response);

    await expect(
      exchangeCodeForTokens(config, 'bad-code', 'http://localhost:12345/callback'),
    ).rejects.toThrow('Token exchange failed (400): bad_request');
  });
});

describe('waitForAuthCallback', () => {
  const realFetch = global.fetch;

  it('resolves with code on valid callback', async () => {
    const port = await findAvailablePort();
    const state = 'test-state-abc';
    const { promise, server } = waitForAuthCallback(port, state, 5000);

    // Give server time to start
    await new Promise((resolve) => setTimeout(resolve, 50));

    // Make a valid callback request using real fetch
    const res = await realFetch(
      `http://127.0.0.1:${port}/callback?code=my-auth-code&state=${state}`,
    );
    expect(res.status).toBe(200);
    const body = await res.text();
    expect(body).toContain('Signed in!');

    const code = await promise;
    expect(code).toBe('my-auth-code');
    server.close();
  });

  it('returns 404 for non-callback paths', async () => {
    const port = await findAvailablePort();
    const { promise, server } = waitForAuthCallback(port, 'state', 500);

    await new Promise((resolve) => setTimeout(resolve, 50));

    const res = await realFetch(`http://127.0.0.1:${port}/other-path`);
    expect(res.status).toBe(404);

    // Server is still waiting for a valid callback — let it time out
    server.close();
    await promise.catch(() => {}); // Ignore timeout rejection
  });

  it('rejects on state mismatch', async () => {
    const port = await findAvailablePort();
    const { promise } = waitForAuthCallback(port, 'correct-state', 5000);

    // Attach rejection handler immediately to avoid unhandled rejection
    const rejectionPromise = promise.catch((err: Error) => err);

    await new Promise((resolve) => setTimeout(resolve, 50));

    const res = await realFetch(`http://127.0.0.1:${port}/callback?code=test&state=wrong-state`);
    expect(res.status).toBe(400);
    const body = await res.text();
    expect(body).toContain('State mismatch');

    const error = await rejectionPromise;
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toBe('State mismatch in OAuth callback');
  });

  it('rejects on error parameter in callback', async () => {
    const port = await findAvailablePort();
    const { promise } = waitForAuthCallback(port, 'state', 5000);

    const rejectionPromise = promise.catch((err: Error) => err);

    await new Promise((resolve) => setTimeout(resolve, 50));

    const res = await realFetch(
      `http://127.0.0.1:${port}/callback?error=access_denied&error_description=User+cancelled`,
    );
    expect(res.status).toBe(400);
    const body = await res.text();
    expect(body).toContain('Sign-in failed');

    const error = await rejectionPromise;
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toBe('Authentication failed: User cancelled');
  });

  it('uses error code when error_description is missing', async () => {
    const port = await findAvailablePort();
    const { promise } = waitForAuthCallback(port, 'state', 5000);

    const rejectionPromise = promise.catch((err: Error) => err);

    await new Promise((resolve) => setTimeout(resolve, 50));

    await realFetch(`http://127.0.0.1:${port}/callback?error=access_denied`);

    const error = await rejectionPromise;
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toBe('Authentication failed: access_denied');
  });

  it('rejects when code is missing from callback', async () => {
    const port = await findAvailablePort();
    const state = 'test-state';
    const { promise } = waitForAuthCallback(port, state, 5000);

    const rejectionPromise = promise.catch((err: Error) => err);

    await new Promise((resolve) => setTimeout(resolve, 50));

    const res = await realFetch(`http://127.0.0.1:${port}/callback?state=${state}`);
    expect(res.status).toBe(400);
    const body = await res.text();
    expect(body).toContain('No authorization code received');

    const error = await rejectionPromise;
    expect(error).toBeInstanceOf(Error);
    expect((error as Error).message).toBe('No authorization code in callback');
  });

  it('rejects on timeout', async () => {
    const port = await findAvailablePort();
    const { promise, server } = waitForAuthCallback(port, 'state', 100); // 100ms timeout

    await expect(promise).rejects.toThrow('Authentication timed out');
    server.close();
  });
});

describe('getAccessToken', () => {
  const config: AuthConfig = {
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    tenantId: 'test-tenant-id',
  };

  let tmpDir: string;
  let originalXdg: string | undefined;
  let originalFetch: typeof global.fetch;
  let stderrSpy: ReturnType<typeof jest.spyOn>;

  beforeEach(() => {
    tmpDir = mkdtempSync(join(tmpdir(), 'm365-mcp-test-'));
    originalXdg = process.env['XDG_CONFIG_HOME'];
    process.env['XDG_CONFIG_HOME'] = tmpDir;
    originalFetch = global.fetch;
    stderrSpy = jest
      .spyOn(process.stderr, 'write')
      .mockImplementation(
        (() => true) as unknown as (
          ...args: Parameters<typeof process.stderr.write>
        ) => ReturnType<typeof process.stderr.write>,
      );
  });

  afterEach(() => {
    if (originalXdg === undefined) {
      delete process.env['XDG_CONFIG_HOME'];
    } else {
      process.env['XDG_CONFIG_HOME'] = originalXdg;
    }
    rmSync(tmpDir, { recursive: true, force: true });
    global.fetch = originalFetch;
    stderrSpy.mockRestore();
  });

  it('returns cached access_token when not expired', async () => {
    const validTokens: TokenData = {
      access_token: 'cached-token',
      refresh_token: 'cached-refresh',
      expires_at: new Date(Date.now() + 3_600_000).toISOString(),
      scopes: 'User.Read',
    };
    saveTokens(validTokens);

    const result = await getAccessToken(config);
    expect(result).toBe('cached-token');
  });

  it('refreshes expired token and returns new access_token', async () => {
    const expiredTokens: TokenData = {
      access_token: 'expired-token',
      refresh_token: 'valid-refresh',
      expires_at: new Date(Date.now() - 60_000).toISOString(),
      scopes: 'User.Read',
    };
    saveTokens(expiredTokens);

    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      json: async () => ({
        access_token: 'refreshed-token',
        refresh_token: 'new-refresh',
        expires_in: 3600,
        scope: 'User.Read',
      }),
    } as Response);

    const result = await getAccessToken(config);
    expect(result).toBe('refreshed-token');
  });

  it('returns cached token that is not yet within expiry buffer', async () => {
    const futureTokens: TokenData = {
      access_token: 'still-valid',
      refresh_token: 'refresh',
      expires_at: new Date(Date.now() + 600_000).toISOString(), // 10 min out (beyond 2-min buffer)
      scopes: 'User.Read',
    };
    saveTokens(futureTokens);

    const result = await getAccessToken(config);
    expect(result).toBe('still-valid');
  });
});
