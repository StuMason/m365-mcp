import { jest } from '@jest/globals';
import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { InMemoryTransport } from '@modelcontextprotocol/server';
import { Client } from '@modelcontextprotocol/client';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..', '..', '..');

// getAccessToken is the seam every data tool goes through. Stubbing it lets the
// wire tests exercise registration and the error path without touching Graph.
const mockGetAccessToken = jest.fn<() => Promise<string>>();

jest.unstable_mockModule('../../lib/auth.js', () => ({
  getAccessToken: mockGetAccessToken,
  loadAuthConfig: (): { clientId: string; tenantId: string } => ({
    clientId: 'test-id',
    tenantId: 'test-tenant',
  }),
  // ms_auth_status pulls these in; none are reached by these tests.
  loadTokens: (): null => null,
  isTokenExpired: (): boolean => false,
  startAuthFlow: async (): Promise<never> => {
    throw new Error('auth flow not expected in roster tests');
  },
  refreshAccessToken: async (): Promise<null> => null,
  SCOPES: [],
}));

const { buildServer } = await import('../../lib/server.js');
const { TOOL_DEFINITIONS, toolNames } = await import('../../lib/tools/index.js');

/**
 * 0.7.0 shipped three hand-maintained lists of tools that disagreed with each
 * other — index.ts registered ten, server-info claimed ten, the README documented
 * seven. These tests make that class of drift a build failure.
 *
 * They assert against what a real client sees over a real transport, not against
 * the source text, so a refactor that changes how tools are registered cannot
 * quietly break them.
 */
async function connectClient(): Promise<Client> {
  const client = new Client({ name: 'roster-test', version: '1.0.0' }, {});
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  await Promise.all([buildServer().connect(serverTransport), client.connect(clientTransport)]);
  return client;
}

describe('tool roster', () => {
  afterEach(() => {
    mockGetAccessToken.mockReset();
  });

  it('every definition is well formed', () => {
    for (const def of TOOL_DEFINITIONS) {
      expect(def.name).toMatch(/^ms_[a-z_]+$/);
      expect(def.title).toBeTruthy();
      expect(def.description.length).toBeGreaterThan(20);
      expect(def.inputSchema).toBeDefined();
      expect(def.annotations.title).toBe(def.title);
    }
  });

  it('has no duplicate names', () => {
    const names = toolNames();
    expect(new Set(names).size).toBe(names.length);
  });

  it('declares every data tool read-only and non-destructive', () => {
    for (const def of TOOL_DEFINITIONS) {
      // ms_auth_status is the deliberate exception: it writes tokens.json and can
      // open a browser, so it is not read-only and not idempotent.
      const expectedReadOnly = def.name !== 'ms_auth_status';
      expect({ name: def.name, readOnly: def.annotations.readOnlyHint }).toEqual({
        name: def.name,
        readOnly: expectedReadOnly,
      });
      expect(def.annotations.destructiveHint).toBe(false);
    }
  });

  it('advertises exactly the roster, in order, over the wire', async () => {
    const client = await connectClient();
    const { tools } = await client.listTools();

    expect(tools.map((t) => t.name)).toEqual(toolNames());
    expect(tools.map((t) => t.title)).toEqual(TOOL_DEFINITIONS.map((d) => d.title));
  });

  it('carries annotations and input schemas onto the wire', async () => {
    const client = await connectClient();
    const { tools } = await client.listTools();

    for (const def of TOOL_DEFINITIONS) {
      const wire = tools.find((t) => t.name === def.name);
      expect(wire?.annotations).toEqual(def.annotations);
      expect(wire?.inputSchema).toMatchObject({ type: 'object' });
    }
  });

  it('publishes real numeric bounds, not the safe-integer range', async () => {
    // A bare z.int() emits minimum: -9007199254740991, which is noise in every
    // schema and tells a caller nothing about the real limit.
    const client = await connectClient();
    const { tools } = await client.listTools();

    const people = tools.find((t) => t.name === 'ms_people');
    expect(people?.inputSchema.properties?.['count']).toMatchObject({
      type: 'integer',
      minimum: 1,
      maximum: 50,
    });
  });

  it('reports an auth failure as a tool error, not a protocol error', async () => {
    mockGetAccessToken.mockRejectedValue(new Error('Token refresh failed'));
    const client = await connectClient();

    const result = await client.callTool({ name: 'ms_profile', arguments: {} });

    expect(result.isError).toBe(true);
    const text = (result.content as Array<{ text: string }>)[0]!.text;
    expect(text).toContain('Token refresh failed');
    expect(text).toContain('Tip: Use ms_auth_status');
  });

  it('rejects out-of-range arguments before the handler runs', async () => {
    mockGetAccessToken.mockResolvedValue('unused-token');
    const client = await connectClient();

    const result = await client.callTool({ name: 'ms_people', arguments: { count: 999 } });

    expect(result.isError).toBe(true);
    // The handler must not have been reached — no token was ever requested.
    expect(mockGetAccessToken).not.toHaveBeenCalled();
  });

  it('matches the tool count claimed in the README and CLAUDE.md', () => {
    const expected = TOOL_DEFINITIONS.length;
    for (const file of ['README.md', 'CLAUDE.md']) {
      const text = readFileSync(join(repoRoot, file), 'utf-8');
      for (const [claim, n] of text.matchAll(/(\d+) tools/g)) {
        expect({ file, claim, n: Number(n) }).toEqual({ file, claim, n: expected });
      }
    }
  });
});
