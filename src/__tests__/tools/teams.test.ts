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

const { executeTeams, teamsToolDefinition } = await import('../../lib/tools/teams.js');

describe('teamsToolDefinition', () => {
  it('is declared read-only', () => {
    expect(teamsToolDefinition.name).toBe('ms_teams');
    expect(teamsToolDefinition.annotations.readOnlyHint).toBe(true);
    expect(teamsToolDefinition.annotations.destructiveHint).toBe(false);
  });
});

describe('executeTeams', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('lists joined teams by default', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'team-1',
            displayName: 'Engineering',
            description: 'Product team',
            webUrl: 'https://teams.microsoft.com/l/team/team-1',
          },
        ],
      },
    });

    const result = await executeTeams('test-token', {});

    expect(result).toContain('Engineering');
    expect(result).toContain('Team ID: team-1');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/joinedTeams'),
      'test-token',
      expect.any(Object),
    );
  });

  it('marks archived teams', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [{ id: 'team-2', displayName: 'Old Project', isArchived: true }] },
    });

    const result = await executeTeams('test-token', {});

    expect(result).toContain('Old Project (archived)');
  });

  it('falls back to a placeholder when a team has no name', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'team-3' }] } });

    const result = await executeTeams('test-token', {});

    expect(result).toContain('Unnamed Team');
  });

  it('reports when there are no joined teams', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    const result = await executeTeams('test-token', {});

    expect(result).toBe('No joined teams found.');
  });

  it('surfaces an error listing teams', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeTeams('test-token', {});

    expect(result).toBe('Error: Insufficient permissions.');
  });

  it('lists channels for a team', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'channel-1',
            displayName: 'General',
            description: 'Team-wide',
            membershipType: 'standard',
            webUrl: 'https://teams.microsoft.com/l/channel/channel-1',
          },
        ],
      },
    });

    const result = await executeTeams('test-token', { team_id: 'team-1' });

    expect(result).toContain('General');
    expect(result).toContain('Channel ID: channel-1');
    expect(result).toContain('Type: standard');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/teams/team-1/channels'),
      'test-token',
      expect.any(Object),
    );
  });

  it('reports when a team has no channels', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    const result = await executeTeams('test-token', { team_id: 'team-1' });

    expect(result).toBe('No channels found in this team.');
  });

  it('names unnamed channels and surfaces channel errors', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'c' }] } });
    expect(await executeTeams('test-token', { team_id: 'team-1' })).toContain('Unnamed Channel');

    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });
    expect(await executeTeams('test-token', { team_id: 'team-1' })).toBe(
      'Error: Resource not found.',
    );
  });

  it('reads channel messages and strips HTML', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'msg-1',
            createdDateTime: '2026-09-10T10:00:00Z',
            from: { user: { displayName: 'Stuart Mason' } },
            body: { content: '<p>Deploy is <b>green</b></p>' },
            attachments: [{ name: 'notes.txt' }],
          },
        ],
      },
    });

    const result = await executeTeams('test-token', {
      team_id: 'team-1',
      channel_id: 'channel-1',
    });

    expect(result).toContain('Deploy is green');
    expect(result).not.toContain('<b>');
    expect(result).toContain('From: Stuart Mason');
    expect(result).toContain('Attachments: notes.txt');
  });

  it('uses the subject as the heading when present', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [{ id: 'm', subject: 'Release notes', body: { content: 'Shipped' } }],
      },
    });

    const result = await executeTeams('test-token', { team_id: 't', channel_id: 'c' });

    expect(result).toContain('## Release notes');
    expect(result).toContain('Unknown date');
  });

  it('drops system events that carry no message body', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'sys-1', body: { content: '' } },
          { id: 'msg-2', body: { content: 'Real message' } },
        ],
      },
    });

    const result = await executeTeams('test-token', { team_id: 't', channel_id: 'c' });

    expect(result).toContain('Real message');
    expect(result).not.toContain('sys-1');
  });

  it('reports when a channel has only system events', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [{ id: 'sys-1', body: { content: '' } }] },
    });

    const result = await executeTeams('test-token', { team_id: 't', channel_id: 'c' });

    expect(result).toBe('No messages found in this channel.');
  });

  it('surfaces an error reading channel messages', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeTeams('test-token', { team_id: 't', channel_id: 'c' });

    expect(result).toBe('Error: Insufficient permissions.');
  });

  it('reads replies on a message thread', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'reply-1',
            from: { user: { displayName: 'Jamie Hook' } },
            body: { content: 'Agreed' },
          },
        ],
      },
    });

    const result = await executeTeams('test-token', {
      team_id: 't',
      channel_id: 'c',
      message_id: 'm',
    });

    expect(result).toContain('Agreed');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/messages/m/replies'),
      'test-token',
      expect.any(Object),
    );
  });

  it('reports when a message has no replies', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    const result = await executeTeams('test-token', {
      team_id: 't',
      channel_id: 'c',
      message_id: 'm',
    });

    expect(result).toBe('No replies on this message.');
  });

  it('surfaces an error reading replies', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    const result = await executeTeams('test-token', {
      team_id: 't',
      channel_id: 'c',
      message_id: 'm',
    });

    expect(result).toBe('Error: Resource not found.');
  });

  // /me/joinedTeams and /teams/{id}/channels reject $top, so those two are
  // trimmed client-side; the message collections take $top on the query.
  it('limits teams client-side rather than with $top', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'a', displayName: 'Team A' },
          { id: 'b', displayName: 'Team B' },
        ],
      },
    });

    const result = await executeTeams('test-token', { count: 1 });

    expect(result).toContain('Team A');
    expect(result).not.toContain('Team B');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      '/me/joinedTeams',
      'test-token',
      expect.any(Object),
    );
  });

  it('limits channels client-side rather than with $top', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'c1', displayName: 'General' },
          { id: 'c2', displayName: 'Random' },
        ],
      },
    });

    const result = await executeTeams('test-token', { team_id: 't', count: 1 });

    expect(result).toContain('General');
    expect(result).not.toContain('Random');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      '/teams/t/channels',
      'test-token',
      expect.any(Object),
    );
  });

  it('clamps the message count to the documented range', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    await executeTeams('test-token', { team_id: 't', channel_id: 'c', count: 999 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );

    await executeTeams('test-token', { team_id: 't', channel_id: 'c', count: 0 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=20'),
      'test-token',
      expect.any(Object),
    );
  });
});
