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
  graphPost: jest.fn(),
}));

const { executeTasks } = await import('../../lib/tools/tasks.js');
const { executeTeams } = await import('../../lib/tools/teams.js');
const { executePeople } = await import('../../lib/tools/people.js');

/** Builds a JWT-shaped Graph token granting exactly these scopes. */
function tokenGranting(...scopes: string[]): string {
  const body = Buffer.from(JSON.stringify({ scp: scopes.join(' ') })).toString('base64url');
  return `header.${body}.sig`;
}

const EVERYTHING = tokenGranting(
  'Tasks.Read',
  'Channel.ReadBasic.All',
  'User.Read.All',
  'Group.Read.All',
);

afterEach(() => {
  mockGraphFetch.mockReset();
});

describe('ms_tasks without Tasks.Read', () => {
  const token = tokenGranting('Mail.Read');

  it.each([
    ['To Do lists', {}],
    ['a To Do list', { list_id: 'list-1' }],
    ['Planner', { planner: true }],
  ])('explains the gap instead of calling Graph for %s', async (_label, args) => {
    const result = await executeTasks(token, args);
    expect(result).toMatch(/^Not available:/);
    expect(result).toContain('Tasks.Read');
    expect(mockGraphFetch).not.toHaveBeenCalled();
  });

  it('still calls Graph when the scope is granted', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } } as never);
    await executeTasks(EVERYTHING, {});
    expect(mockGraphFetch).toHaveBeenCalled();
  });
});

describe('ms_teams without Channel.ReadBasic.All', () => {
  const token = tokenGranting('Team.ReadBasic.All');

  it('explains why channels cannot be listed', async () => {
    const result = await executeTeams(token, { team_id: 'team-1' });
    expect(result).toMatch(/^Not available:/);
    expect(result).toContain('Channel.ReadBasic.All');
    expect(mockGraphFetch).not.toHaveBeenCalled();
  });

  // Reading a channel whose id came from elsewhere needs only ChannelMessage.Read.All,
  // so the guard must not reach past channel enumeration.
  it('still reads messages from a known channel', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } } as never);
    const result = await executeTeams(token, { team_id: 'team-1', channel_id: 'chan-1' });
    expect(result).toBe('No messages found in this channel.');
    expect(mockGraphFetch).toHaveBeenCalled();
  });

  it('still lists joined teams', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } } as never);
    await executeTeams(token, {});
    expect(mockGraphFetch).toHaveBeenCalledWith('/me/joinedTeams', token, expect.anything());
  });
});

describe('ms_people without User.Read.All', () => {
  const token = tokenGranting('User.Read');

  it.each([
    ['a directory search', { search: 'mason' }],
    ['one person', { user: 'someone@example.com' }],
  ])('explains the gap instead of calling Graph for %s', async (_label, args) => {
    const result = await executePeople(token, args);
    expect(result).toMatch(/^Not available:/);
    expect(result).toContain('User.Read.All');
    expect(mockGraphFetch).not.toHaveBeenCalled();
  });
});

describe('ms_people groups with redacted memberships', () => {
  // Graph answers /me/memberOf under plain User.Read but nulls every property
  // except id, so the memberships are real and only their names are missing.
  it('reports the hidden names rather than claiming there are none', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'g1', displayName: null },
          { id: 'g2', displayName: null },
        ],
      },
    } as never);

    const result = await executePeople(tokenGranting('User.Read'), { groups: true });
    expect(result).toMatch(/^Not available:/);
    expect(result).toContain('Group.Read.All');
    expect(result).toContain('2 groups');
  });

  it('says there are none when Graph really returns none', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } } as never);
    const result = await executePeople(EVERYTHING, { groups: true });
    expect(result).toBe('No group memberships found.');
  });

  it('lists the groups when their names are readable', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [{ id: 'g1', displayName: 'Platforms' }] },
    } as never);
    const result = await executePeople(EVERYTHING, { groups: true });
    expect(result).toContain('Platforms');
  });
});
