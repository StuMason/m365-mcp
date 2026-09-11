import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
  graphPost: jest.fn(),
}));

const { executeInsights, insightsToolDefinition } = await import('../../lib/tools/insights.js');

describe('insightsToolDefinition', () => {
  it('is declared read-only', () => {
    expect(insightsToolDefinition.name).toBe('ms_insights');
    expect(insightsToolDefinition.annotations.readOnlyHint).toBe(true);
  });
});

describe('executeInsights', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('defaults to recently used documents', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'i-1',
            lastUsed: { lastAccessedDateTime: '2026-09-09T14:46:48Z' },
            resourceVisualization: {
              title: 'Roadmap.pptx',
              type: 'PowerPoint',
              containerDisplayName: 'Front Door',
            },
            resourceReference: { webUrl: 'https://example.sharepoint.com/roadmap.pptx' },
          },
        ],
      },
    });

    const result = await executeInsights('test-token', {});

    expect(result).toContain('Roadmap.pptx');
    expect(result).toContain('Type: PowerPoint');
    expect(result).toContain('Location: Front Door');
    expect(result).toContain('Last opened:');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/insights/used'),
      'test-token',
      expect.any(Object),
    );
  });

  it('reads the shared list when asked', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'i-2',
            lastShared: {
              sharedDateTime: '2026-09-08T09:00:00Z',
              sharingSubject: 'FYI',
              sharedBy: { displayName: 'Jane Doe' },
            },
            resourceVisualization: { title: 'Notes.docx' },
          },
        ],
      },
    });

    const result = await executeInsights('test-token', { kind: 'shared' });

    expect(result).toContain('Notes.docx');
    expect(result).toContain('Shared by Jane Doe:');
    expect(result).toContain('Context: FYI');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/insights/shared'),
      'test-token',
      expect.any(Object),
    );
  });

  it('names untitled documents', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'x' }] } });

    expect(await executeInsights('test-token', {})).toContain('Untitled');
  });

  it('explains a tenant-wide 403 rather than surfacing it raw', async () => {
    // Item insights are commonly switched off by policy; that is not a fault.
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'ItemInsightsDisabled' },
    });

    const result = await executeInsights('test-token', { kind: 'trending' });

    expect(result).toContain('turned off for this Microsoft 365 tenant');
    expect(result).toContain('trending');
    expect(result).not.toContain('Error:');
  });

  it('surfaces other errors normally', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 500, message: 'Graph exploded.' },
    });

    expect(await executeInsights('test-token', {})).toBe('Error: Graph exploded.');
  });

  it('reports an empty list', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executeInsights('test-token', { kind: 'used' })).toBe(
      'No "used" document insights found.',
    );
  });

  it('clamps count to the documented range', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    await executeInsights('test-token', { count: 50 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );

    await executeInsights('test-token', {});
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=15'),
      'test-token',
      expect.any(Object),
    );
  });
});
