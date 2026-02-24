import { jest } from '@jest/globals';
import { graphFetch, graphPost } from '../lib/graph.js';

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
        message: 'Resource not found. The item may not exist or you may lack access.',
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

  it('merges custom headers into request', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({ value: [] }),
    } as Partial<Response>);

    await graphFetch('/me/messages', 'test-token', {
      timezone: false,
      headers: { ConsistencyLevel: 'eventual' },
    });

    const callHeaders = mock.mock.calls[0][1]!.headers as Record<string, string>;
    expect(callHeaders).toHaveProperty('ConsistencyLevel', 'eventual');
    expect(callHeaders).toHaveProperty('Authorization', 'Bearer test-token');
  });

  it('returns error when success response has invalid JSON', async () => {
    mockFetch({
      ok: true,
      status: 200,
      json: () => Promise.reject(new SyntaxError('Unexpected end of JSON input')),
    } as Partial<Response>);

    const result = await graphFetch('/me', 'test-token');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.status).toBe(200);
      expect(result.error.message).toContain('not valid JSON');
    }
  });

  it('handles response.text() failure on error path', async () => {
    mockFetch({
      ok: false,
      status: 500,
      text: () => Promise.reject(new Error('stream error')),
    } as Partial<Response>);

    const result = await graphFetch('/me', 'test-token');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.status).toBe(500);
      expect(result.error.message).toContain('unable to read error response body');
    }
  });
});

describe('graphPost', () => {
  it('sends POST request with JSON body', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({ value: [{ scheduleId: 'user@example.com' }] }),
    } as Partial<Response>);

    const result = await graphPost<{ schedules: string[] }, { value: unknown[] }>(
      '/me/calendar/getSchedule',
      'test-token',
      { schedules: ['user@example.com'] },
    );

    expect(result).toEqual({
      ok: true,
      data: { value: [{ scheduleId: 'user@example.com' }] },
    });
    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/me/calendar/getSchedule',
      expect.objectContaining({
        method: 'POST',
        body: JSON.stringify({ schedules: ['user@example.com'] }),
        headers: expect.objectContaining({
          Authorization: 'Bearer test-token',
          'Content-Type': 'application/json',
        }),
      }),
    );
  });

  it('returns error on failed POST', async () => {
    mockFetch({ ok: false, status: 403 } as Partial<Response>);
    const result = await graphPost('/me/calendar/getSchedule', 'test-token', {});
    expect(result).toEqual({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });
  });

  it('handles network error on POST', async () => {
    globalThis.fetch = jest.fn<typeof fetch>().mockRejectedValue(new Error('Network failure'));
    const result = await graphPost('/me/calendar/getSchedule', 'test-token', {});
    expect(result).toEqual({
      ok: false,
      error: { status: 0, message: 'Network error: Network failure' },
    });
  });

  it('uses beta URL when beta option is true', async () => {
    const mock = mockFetch({ ok: true, json: () => Promise.resolve({}) } as Partial<Response>);
    await graphPost('/me/calendar/getSchedule', 'test-token', {}, { beta: true });
    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/beta/me/calendar/getSchedule',
      expect.any(Object),
    );
  });

  it('merges custom headers into POST request', async () => {
    const mock = mockFetch({ ok: true, json: () => Promise.resolve({}) } as Partial<Response>);
    await graphPost(
      '/test',
      'test-token',
      {},
      { timezone: false, headers: { 'X-Custom': 'value' } },
    );
    const callHeaders = mock.mock.calls[0][1]!.headers as Record<string, string>;
    expect(callHeaders).toHaveProperty('X-Custom', 'value');
  });
});
