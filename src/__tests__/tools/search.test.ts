import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphPost =
  jest.fn<
    <TBody, TResult>(
      path: string,
      token: string,
      body: TBody,
      options?: { beta?: boolean; timezone?: boolean },
    ) => Promise<GraphResult<TResult>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphPost: mockGraphPost,
  graphFetch: jest.fn(),
}));

const { executeSearch, stripHighlights, searchToolDefinition } =
  await import('../../lib/tools/search.js');

/** Builds a /search/query response with the given hits. */
function response(hits: unknown[], total?: number): { ok: true; data: unknown } {
  return {
    ok: true,
    data: { value: [{ hitsContainers: [{ total: total ?? hits.length, hits }] }] },
  };
}

const EMPTY = response([]);

describe('stripHighlights', () => {
  it('removes the <c0> hit markers Graph wraps around matches', () => {
    expect(stripHighlights('a <c0>match</c0> here')).toBe('a match here');
    expect(stripHighlights('<c1>x</c1> <c2>y</c2>')).toBe('x y');
  });

  it('collapses whitespace', () => {
    expect(stripHighlights('  lots   of\n\nspace  ')).toBe('lots of space');
  });
});

describe('searchToolDefinition', () => {
  it('is declared read-only', () => {
    expect(searchToolDefinition.name).toBe('ms_search');
    expect(searchToolDefinition.annotations.readOnlyHint).toBe(true);
  });
});

describe('executeSearch', () => {
  afterEach(() => {
    mockGraphPost.mockReset();
  });

  it('queries each entity-type group as a separate call', async () => {
    mockGraphPost.mockResolvedValue(EMPTY);

    await executeSearch('test-token', { query: 'budget' });

    // Graph accepts only one entityRequest per call and rejects arbitrary type
    // combinations, so a full search is three calls.
    expect(mockGraphPost).toHaveBeenCalledTimes(3);
    const sent = mockGraphPost.mock.calls.map(
      (c) => (c[2] as { requests: Array<{ entityTypes: string[] }> }).requests[0]!.entityTypes,
    );
    expect(sent).toEqual([
      ['message', 'chatMessage'],
      ['event'],
      ['driveItem', 'listItem', 'site'],
    ]);
  });

  it('never asks for the person entity type', async () => {
    // person needs People.Read, which is not among the consented scopes.
    mockGraphPost.mockResolvedValue(EMPTY);

    await executeSearch('test-token', { query: 'anything' });

    const all = mockGraphPost.mock.calls.flatMap(
      (c) => (c[2] as { requests: Array<{ entityTypes: string[] }> }).requests[0]!.entityTypes,
    );
    expect(all).not.toContain('person');
  });

  it('narrows to the requested areas', async () => {
    mockGraphPost.mockResolvedValue(EMPTY);

    await executeSearch('test-token', { query: 'q', types: ['mail'] });

    expect(mockGraphPost).toHaveBeenCalledTimes(1);
    const body = mockGraphPost.mock.calls[0]![2] as {
      requests: Array<{ entityTypes: string[] }>;
    };
    expect(body.requests[0]!.entityTypes).toEqual(['message']);
  });

  it('maps sharepoint onto both site and listItem', async () => {
    mockGraphPost.mockResolvedValue(EMPTY);

    await executeSearch('test-token', { query: 'q', types: ['sharepoint'] });

    const body = mockGraphPost.mock.calls[0]![2] as {
      requests: Array<{ entityTypes: string[] }>;
    };
    // Order follows the GROUPS declaration, not the argument order.
    expect(body.requests[0]!.entityTypes).toEqual(['listItem', 'site']);
  });

  it('formats a mail hit', async () => {
    mockGraphPost.mockResolvedValueOnce(
      response([
        {
          summary: 'the <c0>budget</c0> is fine',
          resource: {
            '@odata.type': '#microsoft.graph.message',
            subject: 'Q3 Budget',
            from: { emailAddress: { name: 'Jane Doe' } },
            receivedDateTime: '2026-09-10T10:00:00Z',
            webLink: 'https://outlook.example/1',
          },
        },
      ]),
    );
    mockGraphPost.mockResolvedValue(EMPTY);

    const result = await executeSearch('test-token', { query: 'budget' });

    expect(result).toContain('## Mail & Chat');
    expect(result).toContain('### Q3 Budget');
    expect(result).toContain('Type: message');
    expect(result).toContain('From: Jane Doe');
    expect(result).toContain('the budget is fine');
    expect(result).toContain('https://outlook.example/1');
  });

  it('titles chat messages by sender, since they have no subject', async () => {
    mockGraphPost.mockResolvedValueOnce(
      response([
        {
          summary: 'sounds good',
          resource: {
            '@odata.type': 'microsoft.graph.chatMessage',
            from: { user: { displayName: 'Jamie Hook' } },
          },
        },
      ]),
    );
    mockGraphPost.mockResolvedValue(EMPTY);

    const result = await executeSearch('test-token', { query: 'q' });

    expect(result).toContain('### Message from Jamie Hook');
    expect(result).not.toContain('untitled');
  });

  it('falls back to the type when there is no title or sender', async () => {
    mockGraphPost.mockResolvedValueOnce(
      response([{ resource: { '@odata.type': '#microsoft.graph.driveItem' } }]),
    );
    mockGraphPost.mockResolvedValue(EMPTY);

    expect(await executeSearch('test-token', { query: 'q' })).toContain('### (driveItem)');
  });

  it('reports how many results it is showing out of the total', async () => {
    mockGraphPost.mockResolvedValueOnce(
      response([{ resource: { '@odata.type': '#microsoft.graph.message', subject: 'One' } }], 615),
    );
    mockGraphPost.mockResolvedValue(EMPTY);

    expect(await executeSearch('test-token', { query: 'q' })).toContain('(showing 1 of 615)');
  });

  it('reports when nothing matched', async () => {
    mockGraphPost.mockResolvedValue(EMPTY);

    expect(await executeSearch('test-token', { query: 'nothing' })).toBe(
      'No results for "nothing".',
    );
  });

  it('returns the results it got when only some areas fail', async () => {
    mockGraphPost.mockResolvedValueOnce(
      response([{ resource: { '@odata.type': '#microsoft.graph.message', subject: 'Found' } }]),
    );
    mockGraphPost.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSearch('test-token', { query: 'q' });

    expect(result).toContain('Found');
    expect(result).toContain('Some areas failed:');
    expect(result).toContain('Insufficient permissions.');
  });

  it('reports the errors when every area fails', async () => {
    mockGraphPost.mockResolvedValue({
      ok: false,
      error: { status: 500, message: 'Graph exploded.' },
    });

    const result = await executeSearch('test-token', { query: 'q' });

    expect(result).toContain('No results');
    expect(result).toContain('Graph exploded.');
  });

  it('clamps the page size', async () => {
    mockGraphPost.mockResolvedValue(EMPTY);

    await executeSearch('test-token', { query: 'q', types: ['mail'], count: 25 });
    let body = mockGraphPost.mock.calls[0]![2] as { requests: Array<{ size: number }> };
    expect(body.requests[0]!.size).toBe(25);

    mockGraphPost.mockReset();
    mockGraphPost.mockResolvedValue(EMPTY);
    await executeSearch('test-token', { query: 'q', types: ['mail'] });
    body = mockGraphPost.mock.calls[0]![2] as { requests: Array<{ size: number }> };
    expect(body.requests[0]!.size).toBe(5);
  });
});
