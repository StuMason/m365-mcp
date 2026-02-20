import { jest } from '@jest/globals';
import { graphFetch } from '../lib/graph.js';

// Save original fetch
const originalFetch = globalThis.fetch;

afterEach(() => {
  globalThis.fetch = originalFetch;
});

function mockFetch(response: Partial<Response>): jest.Mock<typeof fetch> {
  const mock = jest.fn<typeof fetch>().mockResolvedValue(response as Response);
  globalThis.fetch = mock;
  return mock;
}

const systemTimezone = Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';

describe('graphFetch', () => {
  it('returns ok with data on successful response', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({ displayName: 'Test User' }),
    } as Partial<Response>);

    const result = await graphFetch<{ displayName: string }>('/me', 'test-token');

    expect(result).toEqual({ ok: true, data: { displayName: 'Test User' } });
    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/me',
      expect.objectContaining({
        headers: expect.objectContaining({
          Authorization: 'Bearer test-token',
          'Content-Type': 'application/json',
          Prefer: `outlook.timezone="${systemTimezone}"`,
        }),
      }),
    );
  });

  it('returns 401 error with reconnect message', async () => {
    mockFetch({
      ok: false,
      status: 401,
    } as Partial<Response>);

    const result = await graphFetch('/me', 'expired-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 401,
        message: 'Graph token expired. Use ms_auth_status to reconnect.',
      },
    });
  });

  it('returns 403 error with permissions message', async () => {
    mockFetch({
      ok: false,
      status: 403,
    } as Partial<Response>);

    const result = await graphFetch('/me', 'limited-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });
  });

  it('returns 404 error with license message', async () => {
    mockFetch({
      ok: false,
      status: 404,
    } as Partial<Response>);

    const result = await graphFetch('/me/calendarView', 'test-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 404,
        message: 'Resource not found. Your account may not have an Exchange Online license.',
      },
    });
  });

  it('returns generic error for other status codes', async () => {
    mockFetch({
      ok: false,
      status: 500,
      text: () => Promise.resolve('Internal Server Error'),
    } as Partial<Response>);

    const result = await graphFetch('/me', 'test-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 500,
        message: 'Graph API error (500): Internal Server Error',
      },
    });
  });

  it('uses beta URL when beta option is true', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({ value: [] }),
    } as Partial<Response>);

    await graphFetch('/me/chats', 'test-token', { beta: true });

    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/beta/me/chats',
      expect.any(Object),
    );
  });

  it('includes timezone header by default', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({}),
    } as Partial<Response>);

    await graphFetch('/me', 'test-token');

    expect(mock).toHaveBeenCalledWith(
      expect.any(String),
      expect.objectContaining({
        headers: expect.objectContaining({
          Prefer: `outlook.timezone="${systemTimezone}"`,
        }),
      }),
    );
  });

  it('uses MS365_MCP_TIMEZONE env var for timezone header', async () => {
    const original = process.env['MS365_MCP_TIMEZONE'];
    process.env['MS365_MCP_TIMEZONE'] = 'America/New_York';

    try {
      const mock = mockFetch({
        ok: true,
        json: () => Promise.resolve({}),
      } as Partial<Response>);

      await graphFetch('/me', 'test-token');

      expect(mock).toHaveBeenCalledWith(
        expect.any(String),
        expect.objectContaining({
          headers: expect.objectContaining({
            Prefer: 'outlook.timezone="America/New_York"',
          }),
        }),
      );
    } finally {
      if (original === undefined) {
        delete process.env['MS365_MCP_TIMEZONE'];
      } else {
        process.env['MS365_MCP_TIMEZONE'] = original;
      }
    }
  });

  it('excludes timezone header when timezone is false', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({}),
    } as Partial<Response>);

    await graphFetch('/me', 'test-token', { timezone: false });

    const callHeaders = mock.mock.calls[0][1]!.headers as Record<string, string>;
    expect(callHeaders).not.toHaveProperty('Prefer');
  });

  it('returns network error when fetch throws', async () => {
    globalThis.fetch = jest.fn<typeof fetch>().mockRejectedValue(new Error('Network failure'));

    const result = await graphFetch('/me', 'test-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 0,
        message: 'Network error: Network failure',
      },
    });
  });

  it('handles non-Error thrown values in network catch', async () => {
    globalThis.fetch = jest.fn<typeof fetch>().mockRejectedValue('string error');

    const result = await graphFetch('/me', 'test-token');

    expect(result).toEqual({
      ok: false,
      error: {
        status: 0,
        message: 'Network error: string error',
      },
    });
  });
});
