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

const { executePeople, peopleToolDefinition } = await import('../../lib/tools/people.js');

describe('peopleToolDefinition', () => {
  it('is declared read-only', () => {
    expect(peopleToolDefinition.name).toBe('ms_people');
    expect(peopleToolDefinition.annotations.readOnlyHint).toBe(true);
  });
});

describe('executePeople', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('explains itself when given no arguments', async () => {
    const result = await executePeople('test-token', {});

    expect(result).toContain('Provide search');
    expect(mockGraphFetch).not.toHaveBeenCalled();
  });

  it('searches the directory with the eventual consistency header', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'user-1',
            displayName: 'Stuart Mason',
            mail: 'stuart.mason@example.com',
            jobTitle: 'Senior Developer',
            department: 'Platforms',
            officeLocation: 'Remote',
            mobilePhone: '+44 7000 000000',
          },
        ],
      },
    });

    const result = await executePeople('test-token', { search: 'Mason' });

    expect(result).toContain('Stuart Mason');
    expect(result).toContain('Email: stuart.mason@example.com');
    expect(result).toContain('Job Title: Senior Developer');
    expect(result).toContain('Phone: +44 7000 000000');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$search='),
      'test-token',
      expect.objectContaining({ headers: { ConsistencyLevel: 'eventual' } }),
    );
  });

  it('falls back to userPrincipalName and business phone', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'user-2',
            displayName: 'Jamie Hook',
            userPrincipalName: 'jamie.hook@example.com',
            businessPhones: ['+44 7000 111111'],
          },
        ],
      },
    });

    const result = await executePeople('test-token', { search: 'Hook' });

    expect(result).toContain('Email: jamie.hook@example.com');
    expect(result).toContain('Phone: +44 7000 111111');
  });

  it('names users with no display name', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'u' }] } });

    expect(await executePeople('test-token', { search: 'x' })).toContain('Unknown');
  });

  it('reports when a search finds nobody', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executePeople('test-token', { search: 'Nobody' })).toBe(
      'No directory matches for "Nobody".',
    );
  });

  it('surfaces a search error', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    expect(await executePeople('test-token', { search: 'x' })).toBe(
      'Error: Insufficient permissions.',
    );
  });

  it('reads one person with manager and direct reports', async () => {
    mockGraphFetch
      .mockResolvedValueOnce({
        ok: true,
        data: {
          id: 'user-1',
          displayName: 'Stuart Mason',
          mail: 'stuart.mason@example.com',
        },
      })
      .mockResolvedValueOnce({
        ok: true,
        data: { id: 'mgr-1', displayName: 'Jason Stanbery', mail: 'jason@example.com' },
      })
      .mockResolvedValueOnce({
        ok: true,
        data: {
          value: [{ id: 'rep-1', displayName: 'Erik Araujo', mail: 'erik@example.com' }],
        },
      });

    const result = await executePeople('test-token', { user: 'stuart.mason@example.com' });

    expect(result).toContain('Stuart Mason');
    expect(result).toContain('## Manager');
    expect(result).toContain('Jason Stanbery');
    expect(result).toContain('## Direct Reports (1)');
    expect(result).toContain('Erik Araujo (erik@example.com)');
  });

  it('omits manager and reports sections when there are none', async () => {
    mockGraphFetch
      .mockResolvedValueOnce({
        ok: true,
        data: { id: 'user-1', displayName: 'Solo Person', mail: 'solo@example.com' },
      })
      .mockResolvedValueOnce({ ok: false, error: { status: 404, message: 'Not found.' } })
      .mockResolvedValueOnce({ ok: true, data: { value: [] } });

    const result = await executePeople('test-token', { user: 'solo@example.com' });

    expect(result).toContain('Solo Person');
    expect(result).not.toContain('## Manager');
    expect(result).not.toContain('## Direct Reports');
  });

  it('treats an unreadable direct-reports call as no reports', async () => {
    mockGraphFetch
      .mockResolvedValueOnce({
        ok: true,
        data: { id: 'user-1', displayName: 'Person', mail: 'p@example.com' },
      })
      .mockResolvedValueOnce({ ok: false, error: { status: 403, message: 'Denied.' } })
      .mockResolvedValueOnce({ ok: false, error: { status: 403, message: 'Denied.' } });

    const result = await executePeople('test-token', { user: 'p@example.com' });

    expect(result).toContain('Person');
    expect(result).not.toContain('## Direct Reports');
  });

  it('surfaces an error reading one person', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    expect(await executePeople('test-token', { user: 'ghost@example.com' })).toBe(
      'Error: Resource not found.',
    );
  });

  it('lists group memberships', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'group-1',
            displayName: 'Platforms Guild',
            description: 'Engineering guild',
            mail: 'guild@example.com',
            groupTypes: ['Unified'],
          },
          // Directory roles come back in the same collection with no displayName.
          { id: 'role-1' },
        ],
      },
    });

    const result = await executePeople('test-token', { groups: true });

    expect(result).toContain('Platforms Guild');
    expect(result).toContain('Types: Unified');
    expect(result).not.toContain('role-1');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/memberOf'),
      'test-token',
      expect.any(Object),
    );
  });

  it('reports when there are no group memberships', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executePeople('test-token', { groups: true })).toBe('No group memberships found.');
  });

  it('surfaces an error listing groups', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    expect(await executePeople('test-token', { groups: true })).toBe(
      'Error: Insufficient permissions.',
    );
  });

  it('clamps count to the documented range', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    await executePeople('test-token', { groups: true, count: 999 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );

    await executePeople('test-token', { groups: true, count: 0 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=20'),
      'test-token',
      expect.any(Object),
    );
  });
});
