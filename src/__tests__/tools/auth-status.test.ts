import { jest } from '@jest/globals';
import type { TokenData, AuthConfig } from '../../types/tokens.js';
import type { GraphResult } from '../../lib/graph.js';

const mockLoadTokens = jest.fn<() => TokenData | null>();
const mockIsTokenExpired = jest.fn<(tokens: TokenData) => boolean>();
const mockStartAuthFlow = jest.fn<(config: AuthConfig) => Promise<TokenData>>();
const mockRefreshAccessToken =
  jest.fn<(config: AuthConfig, refreshToken: string) => Promise<TokenData | null>>();

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/auth.js', () => ({
  loadTokens: mockLoadTokens,
  isTokenExpired: mockIsTokenExpired,
  startAuthFlow: mockStartAuthFlow,
  refreshAccessToken: mockRefreshAccessToken,
  SCOPES: ['User.Read', 'Calendars.Read'],
}));

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
}));

// Dynamic import AFTER mocks are registered
const { executeAuthStatus } = await import('../../lib/tools/auth-status.js');

const TEST_CONFIG: AuthConfig = {
  clientId: 'test-client-id',
  clientSecret: 'test-client-secret',
  tenantId: 'test-tenant-id',
};

const VALID_TOKENS: TokenData = {
  access_token: 'valid-token',
  refresh_token: 'refresh-token',
  expires_at: new Date(Date.now() + 3600_000).toISOString(),
  scopes: 'User.Read Calendars.Read',
};

describe('executeAuthStatus', () => {
  let stderrSpy: ReturnType<typeof jest.spyOn>;

  beforeEach(() => {
    stderrSpy = jest.spyOn(process.stderr, 'write').mockImplementation(() => true);
  });

  afterEach(() => {
    mockLoadTokens.mockReset();
    mockIsTokenExpired.mockReset();
    mockStartAuthFlow.mockReset();
    mockRefreshAccessToken.mockReset();
    mockGraphFetch.mockReset();
    stderrSpy.mockRestore();
  });

  it('returns connected status with profile info when tokens valid', async () => {
    mockLoadTokens.mockReturnValue(VALID_TOKENS);
    mockIsTokenExpired.mockReturnValue(false);
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Connected \u2713');
    expect(result).toContain('User: Stuart Mason (stuart@example.com)');
    expect(result).toContain('Token expires:');
    expect(result).toContain('Scopes: User.Read Calendars.Read');
    expect(mockGraphFetch).toHaveBeenCalledWith('/me', 'valid-token');
  });

  it('handles expired token with successful refresh', async () => {
    const refreshedTokens: TokenData = {
      access_token: 'refreshed-token',
      refresh_token: 'new-refresh',
      expires_at: new Date(Date.now() + 3600_000).toISOString(),
      scopes: 'User.Read Calendars.Read',
    };

    mockLoadTokens.mockReturnValue(VALID_TOKENS);
    mockIsTokenExpired.mockReturnValue(true);
    mockRefreshAccessToken.mockResolvedValue(refreshedTokens);
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Connected \u2713');
    expect(result).toContain('User: Stuart Mason (stuart@example.com)');
    expect(mockRefreshAccessToken).toHaveBeenCalledWith(TEST_CONFIG, 'refresh-token');
    expect(mockGraphFetch).toHaveBeenCalledWith('/me', 'refreshed-token');
  });

  it('handles expired token with failed refresh', async () => {
    mockLoadTokens.mockReturnValue(VALID_TOKENS);
    mockIsTokenExpired.mockReturnValue(true);
    mockRefreshAccessToken.mockResolvedValue(null);

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Not connected');
    expect(result).toContain('Token expired and refresh failed');
    expect(result).toContain('Action: Run ms_auth_status again to sign in.');
  });

  it('handles no tokens by triggering auth flow', async () => {
    const newTokens: TokenData = {
      access_token: 'new-token',
      refresh_token: 'new-refresh',
      expires_at: new Date(Date.now() + 3600_000).toISOString(),
      scopes: 'User.Read Calendars.Read',
    };

    mockLoadTokens.mockReturnValue(null);
    mockStartAuthFlow.mockResolvedValue(newTokens);
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(stderrSpy).toHaveBeenCalledWith('Not connected. Starting auth flow...\n');
    expect(mockStartAuthFlow).toHaveBeenCalledWith(TEST_CONFIG);
    expect(result).toContain('Status: Connected \u2713 (just signed in)');
    expect(result).toContain('User: Stuart Mason (stuart@example.com)');
  });

  it('handles profile fetch failure gracefully with valid tokens', async () => {
    mockLoadTokens.mockReturnValue(VALID_TOKENS);
    mockIsTokenExpired.mockReturnValue(false);
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 401, message: 'Graph token expired' },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Connected \u2713');
    expect(result).toContain('Token expires:');
    expect(result).toContain('Scopes:');
    expect(result).not.toContain('User:');
  });

  it('handles profile fetch failure after auth flow', async () => {
    const newTokens: TokenData = {
      access_token: 'new-token',
      refresh_token: 'new-refresh',
      expires_at: new Date(Date.now() + 3600_000).toISOString(),
      scopes: 'User.Read',
    };

    mockLoadTokens.mockReturnValue(null);
    mockStartAuthFlow.mockResolvedValue(newTokens);
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 0, message: 'Network error' },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Connected \u2713 (just signed in)');
    expect(result).not.toContain('User:');
  });

  it('catches auth flow errors and returns error status', async () => {
    mockLoadTokens.mockReturnValue(null);
    mockStartAuthFlow.mockRejectedValue(new Error('Authentication timed out'));

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Status: Not connected');
    expect(result).toContain('Error: Authentication timed out');
    expect(result).toContain('Action: Run ms_auth_status again to sign in.');
  });

  it('uses userPrincipalName when mail is missing', async () => {
    mockLoadTokens.mockReturnValue(VALID_TOKENS);
    mockIsTokenExpired.mockReturnValue(false);
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        userPrincipalName: 'stuart@example.onmicrosoft.com',
      },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('User: Stuart Mason (stuart@example.onmicrosoft.com)');
  });

  it('falls back to SCOPES constant when tokens have no scopes', async () => {
    const tokensNoScopes: TokenData = {
      ...VALID_TOKENS,
      scopes: '',
    };

    mockLoadTokens.mockReturnValue(tokensNoScopes);
    mockIsTokenExpired.mockReturnValue(false);
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 0, message: 'Network error' },
    });

    const result = await executeAuthStatus(TEST_CONFIG);

    expect(result).toContain('Scopes: User.Read Calendars.Read');
  });
});
