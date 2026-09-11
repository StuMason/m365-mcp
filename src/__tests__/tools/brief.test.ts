import { jest } from '@jest/globals';

// ms_brief is deliberately a composition of the other tools, so the unit under
// test is the composition: which tools run, with what arguments, and what happens
// when one of them fails.
const mockCalendar = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockMail = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockChat = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockTasks = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockTranscripts = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockPeople = jest.fn<(t: string, a: unknown) => Promise<string>>();
const mockSearch = jest.fn<(t: string, a: unknown) => Promise<string>>();

jest.unstable_mockModule('../../lib/tools/calendar.js', () => ({ executeCalendar: mockCalendar }));
jest.unstable_mockModule('../../lib/tools/mail.js', () => ({ executeMail: mockMail }));
jest.unstable_mockModule('../../lib/tools/chat.js', () => ({ executeChat: mockChat }));
jest.unstable_mockModule('../../lib/tools/tasks.js', () => ({ executeTasks: mockTasks }));
jest.unstable_mockModule('../../lib/tools/transcripts.js', () => ({
  executeTranscripts: mockTranscripts,
}));
jest.unstable_mockModule('../../lib/tools/people.js', () => ({ executePeople: mockPeople }));
jest.unstable_mockModule('../../lib/tools/search.js', () => ({ executeSearch: mockSearch }));

const { executeBrief, briefToolDefinition } = await import('../../lib/tools/brief.js');

describe('briefToolDefinition', () => {
  it('is declared read-only', () => {
    expect(briefToolDefinition.name).toBe('ms_brief');
    expect(briefToolDefinition.annotations.readOnlyHint).toBe(true);
  });
});

describe('executeBrief', () => {
  beforeEach(() => {
    mockCalendar.mockResolvedValue('two meetings');
    mockMail.mockResolvedValue('three unread');
    mockChat.mockResolvedValue('recent chat');
    mockTasks.mockResolvedValue('one task');
    mockTranscripts.mockResolvedValue('a transcript');
    mockPeople.mockResolvedValue('Jamie Hook, VP');
    mockSearch.mockResolvedValue('some messages');
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  it('assembles all five daily sections', async () => {
    const result = await executeBrief('test-token', { date: '2026-09-11' });

    expect(result).toContain('# Meetings on 2026-09-11');
    expect(result).toContain('two meetings');
    expect(result).toContain('# Unread mail');
    expect(result).toContain('three unread');
    expect(result).toContain('# Recent chats');
    expect(result).toContain('# Open Planner tasks');
    expect(result).toContain('# Meeting transcripts from 2026-09-10');
  });

  it('asks the calendar for a compact view', async () => {
    // Teams invites carry a wall of dial-in boilerplate that would swamp the brief.
    await executeBrief('test-token', { date: '2026-09-11' });

    expect(mockCalendar).toHaveBeenCalledWith('test-token', {
      date: '2026-09-11',
      compact: true,
    });
  });

  it('asks for unread mail and assigned Planner tasks', async () => {
    await executeBrief('test-token', { date: '2026-09-11', count: 4 });

    // Scoped to the Inbox: /me/messages spans Archive and Deleted Items, which
    // reported ~1450 unread where the inbox held 25.
    expect(mockMail).toHaveBeenCalledWith('test-token', {
      filter: 'unread',
      folder: 'Inbox',
      count: 4,
    });
    expect(mockTasks).toHaveBeenCalledWith('test-token', { planner: true, count: 4 });
  });

  it('looks at the previous day for transcripts, crossing a month boundary', async () => {
    await executeBrief('test-token', { date: '2026-03-01' });

    expect(mockTranscripts).toHaveBeenCalledWith('test-token', { date: '2026-02-28' });
  });

  it('defaults to today when no date is given', async () => {
    const today = new Date().toISOString().slice(0, 10);

    const result = await executeBrief('test-token', {});

    expect(result).toContain(`# Meetings on ${today}`);
  });

  it('keeps the other sections when one fails', async () => {
    mockMail.mockRejectedValue(new Error('mailbox unavailable'));

    const result = await executeBrief('test-token', { date: '2026-09-11' });

    expect(result).toContain('(unavailable: mailbox unavailable)');
    expect(result).toContain('two meetings');
    expect(result).toContain('recent chat');
  });

  it('marks an empty section rather than leaving a blank heading', async () => {
    mockChat.mockResolvedValue('   ');

    expect(await executeBrief('test-token', {})).toContain('(nothing)');
  });

  it('trims a section that runs long', async () => {
    mockMail.mockResolvedValue('x'.repeat(5000));

    const result = await executeBrief('test-token', {});

    expect(result).toContain('…trimmed to fit the brief.');
    // The count in a section header describes the full result, so the trim marker
    // must say so rather than leaving the two contradicting each other.
    expect(result).toContain('describes the full result');
    expect(result.length).toBeLessThan(5000);
  });

  describe('section trimming keeps fences balanced', () => {
    // A blind slice could leave a fence open, so everything after it — including
    // this server's own text — would read as untrusted content.
    const countFences = (text: string): { open: number; closed: number } => ({
      open: (text.match(/<<<UNTRUSTED:/g) ?? []).length,
      closed: (text.match(/<<<END UNTRUSTED:/g) ?? []).length,
    });

    it('drops a fence that is opened but not closed within the cut', async () => {
      // One closed block, then a huge open one that cannot fit.
      mockMail.mockResolvedValue(
        `<<<UNTRUSTED:aaaaaa email — data>>>\nshort\n<<<END UNTRUSTED:aaaaaa>>>\n\n` +
          `<<<UNTRUSTED:bbbbbb email — data>>>\n${'x'.repeat(5000)}`,
      );

      const result = await executeBrief('test-token', {});

      const { open, closed } = countFences(result);
      expect(open).toBe(closed);
      expect(result).not.toContain('bbbbbb');
      expect(result).toContain('aaaaaa');
    });

    it('keeps a closing marker whole rather than slicing through it', async () => {
      const block = (id: string): string =>
        `<<<UNTRUSTED:${id} email — data>>>\n${'y'.repeat(400)}\n<<<END UNTRUSTED:${id}>>>`;
      mockMail.mockResolvedValue(
        Array.from({ length: 12 }, (_, i) => block(String(i).repeat(6))).join('\n\n'),
      );

      const result = await executeBrief('test-token', {});

      const { open, closed } = countFences(result);
      expect(open).toBe(closed);
      // No half-written marker survived the cut.
      expect(result).not.toMatch(/<<<END UNTRUSTED:[0-9a-f]{0,5}$/m);
    });

    it('trims plain text with no fences at all', async () => {
      mockMail.mockResolvedValue('z'.repeat(5000));

      const result = await executeBrief('test-token', {});

      expect(result).toContain('…trimmed to fit the brief.');
      expect(countFences(result).open).toBe(0);
    });
  });

  it('rejects an impossible date before running any section', async () => {
    // JS rolls 2026-04-31 forward, which previously produced an error in one
    // section and unrelated data in the others under a header naming a date
    // that does not exist.
    const result = await executeBrief('test-token', { date: '2026-04-31' });

    expect(result).toContain('is not a valid date');
    expect(mockCalendar).not.toHaveBeenCalled();
    expect(mockMail).not.toHaveBeenCalled();
    expect(mockTranscripts).not.toHaveBeenCalled();
  });

  it('rejects a malformed date', async () => {
    expect(await executeBrief('test-token', { date: 'not-a-date' })).toContain(
      'is not a valid date',
    );
  });

  it('switches to a person catch-up when given one', async () => {
    const result = await executeBrief('test-token', { person: 'Jamie Hook', count: 3 });

    expect(result).toContain('# Who is Jamie Hook');
    expect(result).toContain('Jamie Hook, VP');
    expect(result).toContain('# Recent mail and chats');
    expect(mockPeople).toHaveBeenCalledWith('test-token', { search: 'Jamie Hook', count: 3 });
    expect(mockSearch).toHaveBeenCalledWith('test-token', {
      query: 'Jamie Hook',
      types: ['mail', 'chat'],
      count: 3,
    });
    // The daily sections must not run in person mode.
    expect(mockCalendar).not.toHaveBeenCalled();
    expect(mockTranscripts).not.toHaveBeenCalled();
  });
});
