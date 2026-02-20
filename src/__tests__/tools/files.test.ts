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
}));

// Dynamic import AFTER the mock is registered
const { executeFiles, formatFileSize } = await import('../../lib/tools/files.js');

describe('executeFiles', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('lists root files by default', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            name: 'Documents',
            size: 0,
            lastModifiedDateTime: '2025-06-15T10:00:00Z',
            webUrl: 'https://onedrive.example.com/Documents',
            folder: { childCount: 5 },
          },
          {
            name: 'report.docx',
            size: 25600,
            lastModifiedDateTime: '2025-06-14T09:00:00Z',
            webUrl: 'https://onedrive.example.com/report.docx',
            file: {
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            },
          },
        ],
      },
    });

    const result = await executeFiles('test-token', {});

    expect(result).toContain('\u{1F4C1} Documents');
    expect(result).toContain('\u{1F4C4} report.docx');
    expect(result).toContain('https://onedrive.example.com/Documents');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/root/children'),
      'test-token',
      { timezone: false },
    );
  });

  it('handles path param', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { path: '/Documents/Reports' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/root:%2FDocuments%2FReports:/children'),
      'test-token',
      { timezone: false },
    );
  });

  it('treats root path as default listing', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { path: '/' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/root/children'),
      'test-token',
      { timezone: false },
    );
  });

  it('handles search param', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { search: 'budget 2025' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining("/me/drive/root/search(q='budget%202025')"),
      'test-token',
      { timezone: false },
    );
  });

  it('handles empty results', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeFiles('test-token', {});

    expect(result).toBe('No files found.');
  });

  it('handles errors from Graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 401, message: 'Graph token expired. Use ms_auth_status to reconnect.' },
    });

    const result = await executeFiles('expired-token', {});

    expect(result).toBe('Error: Graph token expired. Use ms_auth_status to reconnect.');
  });

  it('clamps count to maximum of 50', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { count: 100 });

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=50'), 'test-token', {
      timezone: false,
    });
  });

  it('clamps count to minimum of 1', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { count: 0 });

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=1'), 'test-token', {
      timezone: false,
    });
  });

  it('defaults count to 20', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', {});

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=20'), 'test-token', {
      timezone: false,
    });
  });

  it('search takes precedence over path', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeFiles('test-token', { search: 'test', path: '/Documents' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/root/search'),
      'test-token',
      { timezone: false },
    );
  });
});

describe('formatFileSize', () => {
  it('formats bytes', () => {
    expect(formatFileSize(500)).toBe('500 B');
  });

  it('formats kilobytes', () => {
    expect(formatFileSize(1024)).toBe('1.0 KB');
  });

  it('formats megabytes', () => {
    expect(formatFileSize(1048576)).toBe('1.0 MB');
  });

  it('formats gigabytes', () => {
    expect(formatFileSize(1073741824)).toBe('1.0 GB');
  });

  it('handles zero bytes', () => {
    expect(formatFileSize(0)).toBe('0 B');
  });

  it('handles undefined', () => {
    expect(formatFileSize(undefined)).toBe('N/A');
  });

  it('formats fractional sizes', () => {
    expect(formatFileSize(1536)).toBe('1.5 KB');
  });
});
