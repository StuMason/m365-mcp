import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean; headers?: Record<string, string> },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
}));

const { executeSharepoint } = await import('../../lib/tools/sharepoint.js');

describe('executeSharepoint', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('searches sites with default wildcard', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            displayName: 'Engineering Hub',
            webUrl: 'https://example.sharepoint.com/sites/engineering',
            description: 'Engineering team site',
            id: 'site-1',
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('Engineering Hub');
    expect(result).toContain('https://example.sharepoint.com/sites/engineering');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites?search=*'),
      'test-token',
      expect.any(Object),
    );
  });

  it('searches sites with custom query', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSharepoint('test-token', { search: 'marketing' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites?search=marketing'),
      'test-token',
      expect.any(Object),
    );
  });

  it('lists site lists', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            displayName: 'Documents',
            name: 'documents',
            description: 'Shared documents',
            webUrl: 'https://example.sharepoint.com/sites/eng/Documents',
            lastModifiedDateTime: '2026-01-15T10:00:00Z',
            list: { template: 'documentLibrary' },
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', { site_id: 'site-1' });

    expect(result).toContain('Documents');
    expect(result).toContain('documentLibrary');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites/site-1/lists'),
      'test-token',
      expect.any(Object),
    );
  });

  it('lists items from a specific list', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'item-1',
            fields: { Title: 'Q1 Report', Status: 'Published' },
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', {
      site_id: 'site-1',
      list_id: 'list-1',
    });

    expect(result).toContain('Q1 Report');
    expect(result).toContain('Published');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites/site-1/lists/list-1/items'),
      'test-token',
      expect.any(Object),
    );
  });

  it('handles empty results', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('No sites found');
  });

  it('handles error from Graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('Error');
  });

  it('clamps count', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSharepoint('test-token', { count: 100 });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );
  });

  it('handles empty lists for a site', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    const result = await executeSharepoint('test-token', { site_id: 'site-1' });
    expect(result).toBe('No lists found for this site.');
  });

  it('handles empty items for a list', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    const result = await executeSharepoint('test-token', { site_id: 'site-1', list_id: 'list-1' });
    expect(result).toBe('No items found in this list.');
  });

  it('handles error fetching site lists', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSharepoint('test-token', { site_id: 'site-1' });
    expect(result).toBe('Error: Insufficient permissions.');
  });

  it('handles error fetching list items', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    const result = await executeSharepoint('test-token', { site_id: 'site-1', list_id: 'list-1' });
    expect(result).toBe('Error: Resource not found.');
  });
});
