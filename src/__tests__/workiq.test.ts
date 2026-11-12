import { jest } from '@jest/globals';
import type { TokenData } from '../types/tokens.js';

const mockLoadTokens = jest.fn<() => TokenData | null>();
const mockSaveTokens = jest.fn<(t: TokenData) => void>();

jest.unstable_mockModule('../lib/auth.js', () => ({
  loadTokens: mockLoadTokens,
  saveTokens: mockSaveTokens,
}));

const {
  getWorkIqToken,
  workIqScope,
  listWorkIqTools,
  callWorkIqTool,
  resetWorkIqToken,
  resetWorkIqSession,
  READ_ONLY_TOOLS,
} = await import('../lib/workiq.js');

const config = { clientId: 'client-1', tenantId: 'tenant-1' };
const SCOPE = 'api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask';

const tokens: TokenData = {
  access_token: 'graph-at',
  refresh_token: 'rt-original',
  expires_at: new Date(Date.now() + 3_600_000).toISOString(),
  scopes: 'User.Read',
};

/** A token response body. */
function tokenBody(over: Partial<Record<string, unknown>> = {}): string {
  return JSON.stringify({ access_token: 'wq-at', expires_in: 3600, ...over });
}

beforeEach(() => {
  process.env['MS365_MCP_WORKIQ_SCOPE'] = SCOPE;
  mockLoadTokens.mockReturnValue({ ...tokens });
  mockSaveTokens.mockReset();
  resetWorkIqToken();
  resetWorkIqSession();
});

afterEach(() => {
  delete process.env['MS365_MCP_WORKIQ_SCOPE'];
});

describe('workIqScope', () => {
  it('returns null when unset, so the feature stays off by default', () => {
    delete process.env['MS365_MCP_WORKIQ_SCOPE'];
    expect(workIqScope()).toBeNull();
  });

  it('ignores whitespace-only values', () => {
    process.env['MS365_MCP_WORKIQ_SCOPE'] = '   ';
    expect(workIqScope()).toBeNull();
  });
});

describe('getWorkIqToken', () => {
  it('redeems the stored refresh token against the Work IQ resource', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      text: async () => tokenBody(),
    } as Response);

    expect(await getWorkIqToken(config)).toBe('wq-at');

    const [, init] = (global.fetch as jest.MockedFunction<typeof fetch>).mock.calls[0] as [
      string,
      RequestInit,
    ];
    const body = new URLSearchParams(init.body as string);
    expect(body.get('grant_type')).toBe('refresh_token');
    expect(body.get('scope')).toBe(SCOPE);
    expect(body.get('refresh_token')).toBe('rt-original');
    // Origin is SPA-only and this is a token request like any other.
    expect((init.headers as Record<string, string>)['Origin']).toBeUndefined();
  });

  // The refresh grant rotates the refresh token, and Graph must use the new one
  // from then on. Dropping it strands the Graph session on a retired token.
  it('persists a rotated refresh token', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      text: async () => tokenBody({ refresh_token: 'rt-rotated' }),
    } as Response);

    await getWorkIqToken(config);

    expect(mockSaveTokens).toHaveBeenCalledWith(
      expect.objectContaining({ refresh_token: 'rt-rotated', access_token: 'graph-at' }),
    );
  });

  it('does not rewrite the token file when the refresh token is unchanged', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      text: async () => tokenBody({ refresh_token: 'rt-original' }),
    } as Response);

    await getWorkIqToken(config);
    expect(mockSaveTokens).not.toHaveBeenCalled();
  });

  it('caches the token instead of redeeming on every call', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      text: async () => tokenBody(),
    } as Response);

    await getWorkIqToken(config);
    await getWorkIqToken(config);

    expect(global.fetch).toHaveBeenCalledTimes(1);
  });

  it('explains an unconsented scope rather than surfacing the raw Azure error', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: false,
      status: 400,
      text: async () => JSON.stringify({ error_description: 'AADSTS65001: no consent' }),
    } as Response);

    await expect(getWorkIqToken(config)).rejects.toThrow(/not been granted the Work IQ scope/);
    // The Graph tools keep working, and the message has to say so.
    await expect(getWorkIqToken(config)).rejects.toThrow(/Graph tools are unaffected/);
  });

  it('does not leak the Azure body on other failures', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: false,
      status: 500,
      text: async () => 'internal server trace 12345',
    } as Response);

    await expect(getWorkIqToken(config)).rejects.toThrow(/Could not get a Work IQ token \(500\)/);
    await expect(getWorkIqToken(config)).rejects.not.toThrow(/12345/);
  });

  it('asks the caller to sign in when there is no refresh token', async () => {
    mockLoadTokens.mockReturnValue(null);
    await expect(getWorkIqToken(config)).rejects.toThrow(/Not signed in/);
  });

  it('refuses when Work IQ is not enabled', async () => {
    delete process.env['MS365_MCP_WORKIQ_SCOPE'];
    await expect(getWorkIqToken(config)).rejects.toThrow(/not enabled/);
  });

  it('sends the client secret for a confidential client', async () => {
    global.fetch = jest.fn<typeof fetch>().mockResolvedValue({
      ok: true,
      text: async () => tokenBody(),
    } as Response);

    await getWorkIqToken({ ...config, clientSecret: 'shh' });

    const [, init] = (global.fetch as jest.MockedFunction<typeof fetch>).mock.calls[0] as [
      string,
      RequestInit,
    ];
    expect(new URLSearchParams(init.body as string).get('client_secret')).toBe('shh');
  });
});

describe('MCP transport', () => {
  /** Queues a token response followed by the given MCP replies. */
  function mockMcp(
    ...replies: Array<{ body: string; ok?: boolean; status?: number; session?: string }>
  ): void {
    const fn = jest.fn<typeof fetch>();
    fn.mockResolvedValueOnce({ ok: true, text: async () => tokenBody() } as Response);
    for (const reply of replies) {
      fn.mockResolvedValueOnce({
        ok: reply.ok ?? true,
        status: reply.status ?? 200,
        text: async () => reply.body,
        headers: new Headers(reply.session ? { 'mcp-session-id': reply.session } : {}),
      } as Response);
    }
    global.fetch = fn;
  }

  const initReply = {
    body: JSON.stringify({ result: { serverInfo: { name: 'WorkIQ' } } }),
    session: 'sess-1',
  };

  it('lists tools after the handshake', async () => {
    mockMcp(
      initReply,
      { body: '' }, // notifications/initialized
      {
        body: JSON.stringify({ result: { tools: [{ name: 'fetch' }, { name: 'create_entity' }] } }),
      },
    );

    const tools = await listWorkIqTools(config);
    expect(tools.map((t) => t.name)).toEqual(['fetch', 'create_entity']);
  });

  // Streamable HTTP may answer with an SSE stream rather than plain JSON.
  it('reassembles an SSE reply', async () => {
    const sse = 'event: message\ndata: {"result":{"tools":[{"name":"ask"}]}}\n\n';
    mockMcp(initReply, { body: '' }, { body: sse });

    const tools = await listWorkIqTools(config);
    expect(tools.map((t) => t.name)).toEqual(['ask']);
  });

  it('carries the session id returned by the server', async () => {
    mockMcp(initReply, { body: '' }, { body: JSON.stringify({ result: { tools: [] } }) });
    await listWorkIqTools(config);

    const calls = (global.fetch as jest.MockedFunction<typeof fetch>).mock.calls;
    const last = calls[calls.length - 1] as [string, RequestInit];
    expect((last[1].headers as Record<string, string>)['Mcp-Session-Id']).toBe('sess-1');
  });

  it('performs the handshake once across calls', async () => {
    mockMcp(
      initReply,
      { body: '' },
      { body: JSON.stringify({ result: { tools: [] } }) },
      { body: JSON.stringify({ result: { content: [] } }) },
    );

    await listWorkIqTools(config);
    await callWorkIqTool(config, 'fetch', { entityUrls: ['/me'] });

    const methods = (global.fetch as jest.MockedFunction<typeof fetch>).mock.calls
      .slice(1)
      .map((c) => JSON.parse((c[1] as RequestInit).body as string).method);
    expect(methods.filter((m) => m === 'initialize')).toHaveLength(1);
  });

  it('surfaces a JSON-RPC error as a thrown message', async () => {
    mockMcp(
      initReply,
      { body: '' },
      { body: JSON.stringify({ error: { code: -32602, message: 'bad args' } }) },
    );
    await expect(listWorkIqTools(config)).rejects.toThrow('bad args');
  });

  it('reports an HTTP failure without echoing the body', async () => {
    mockMcp(initReply, { body: '' }, { body: 'upstream trace 999', ok: false, status: 503 });
    await expect(listWorkIqTools(config)).rejects.toThrow(/Work IQ returned HTTP 503/);
  });

  it('passes tool arguments through unchanged', async () => {
    mockMcp(initReply, { body: '' }, { body: JSON.stringify({ result: { content: [] } }) });
    await callWorkIqTool(config, 'fetch', { entityUrls: ['/me/manager'] });

    const calls = (global.fetch as jest.MockedFunction<typeof fetch>).mock.calls;
    const last = calls[calls.length - 1] as [string, RequestInit];
    const sent = JSON.parse(last[1].body as string);
    expect(sent.params).toEqual({ name: 'fetch', arguments: { entityUrls: ['/me/manager'] } });
  });
});

describe('READ_ONLY_TOOLS', () => {
  // ask delegates to an agent and call_function invokes arbitrary functions, so
  // neither can be promised read-only however harmless a given call looks.
  it.each(['ask', 'call_function', 'create_entity', 'update_entity', 'delete_entity', 'do_action'])(
    'does not claim %s is read-only',
    (name) => {
      expect(READ_ONLY_TOOLS).not.toContain(name);
    },
  );

  it.each(['fetch', 'fetch_blob', 'get_schema', 'search_paths', 'list_agents'])(
    'includes %s',
    (name) => {
      expect(READ_ONLY_TOOLS).toContain(name);
    },
  );
});
