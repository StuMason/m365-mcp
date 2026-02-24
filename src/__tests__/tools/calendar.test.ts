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
const { executeCalendar } = await import('../../lib/tools/calendar.js');

describe('executeCalendar', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('formats events correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Standup',
            start: { dateTime: '2025-01-15T09:00:00', timeZone: 'Europe/London' },
            end: { dateTime: '2025-01-15T09:30:00', timeZone: 'Europe/London' },
            location: { displayName: 'Room A' },
            organizer: { emailAddress: { name: 'Alice' } },
            attendees: [{ emailAddress: { name: 'Bob' } }, { emailAddress: { name: 'Charlie' } }],
            isAllDay: false,
            body: { contentType: 'text', content: 'Daily standup meeting' },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('## Standup');
    expect(result).toContain('Time: 2025-01-15T09:00:00 - 2025-01-15T09:30:00');
    expect(result).toContain('Location: Room A');
    expect(result).toContain('Organizer: Alice');
    expect(result).toContain('Attendees: Bob, Charlie');
    expect(result).toContain('Daily standup meeting');
  });

  it('handles all-day events', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Holiday',
            isAllDay: true,
            start: { dateTime: '2025-01-15T00:00:00' },
            end: { dateTime: '2025-01-16T00:00:00' },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('## Holiday');
    expect(result).toContain('Time: All day');
  });

  it('returns message when no events found', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toBe('No calendar events found for the specified date range.');
  });

  it('handles error from graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toBe(
      'Error: Insufficient permissions. Check granted scopes with ms_auth_status.',
    );
  });

  it('uses date param to build calendarView path', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeCalendar('test-token', { date: '2025-06-15' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('startDateTime=2025-06-15T00:00:00.000Z'),
      'test-token',
      { timezone: true },
    );
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('endDateTime=2025-06-16T00:00:00.000Z'),
      'test-token',
      { timezone: true },
    );
  });

  it('uses start and end params directly when provided', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeCalendar('test-token', {
      start: '2025-06-15T09:00:00Z',
      end: '2025-06-15T17:00:00Z',
    });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('startDateTime=2025-06-15T09:00:00Z'),
      'test-token',
      { timezone: true },
    );
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('endDateTime=2025-06-15T17:00:00Z'),
      'test-token',
      { timezone: true },
    );
  });

  it('defaults to today when no date params given', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const todayStr = `${year}-${month}-${day}`;

    await executeCalendar('test-token', {});

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining(`startDateTime=${todayStr}`),
      'test-token',
      { timezone: true },
    );
  });

  it('handles events with minimal fields', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: undefined,
            isAllDay: false,
            start: { dateTime: undefined },
            end: { dateTime: undefined },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('## Untitled');
    expect(result).toContain('Time: N/A - N/A');
  });

  it('formats multiple events separated by blank lines', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Meeting 1',
            isAllDay: false,
            start: { dateTime: '2025-01-15T09:00:00' },
            end: { dateTime: '2025-01-15T10:00:00' },
          },
          {
            subject: 'Meeting 2',
            isAllDay: false,
            start: { dateTime: '2025-01-15T11:00:00' },
            end: { dateTime: '2025-01-15T12:00:00' },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('## Meeting 1');
    expect(result).toContain('## Meeting 2');
    // Two events separated by double newline
    const parts = result.split('\n\n');
    expect(parts.length).toBe(2);
  });

  it('strips HTML from event body', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'HTML Body Meeting',
            isAllDay: false,
            start: { dateTime: '2025-01-15T09:00:00' },
            end: { dateTime: '2025-01-15T10:00:00' },
            body: {
              contentType: 'html',
              content:
                '<html><body><p>Join the meeting here</p><p>Agenda items below</p></body></html>',
            },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('Join the meeting here');
    expect(result).toContain('Agenda items below');
    expect(result).not.toContain('<p>');
    expect(result).not.toContain('<html>');
  });

  it('truncates long event bodies to 500 chars', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Long Body',
            isAllDay: false,
            start: { dateTime: '2025-01-15T09:00:00' },
            end: { dateTime: '2025-01-15T10:00:00' },
            body: { contentType: 'text', content: 'A'.repeat(800) },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).not.toContain('A'.repeat(800));
    expect(result).toContain('A'.repeat(500) + '...');
  });

  it('strips Teams meeting boilerplate from event body', async () => {
    const teamsBody =
      'Sprint planning session\n\n' +
      '________________________________________________________________________________\n' +
      'Microsoft Teams meeting\n' +
      'Join on your computer, mobile app or room device\n' +
      'Click here to join the meeting\n' +
      'Meeting ID: 123 456 789\n' +
      'Passcode: abc123\n' +
      'Dial in: +1 555-0100';

    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Teams Meeting',
            isAllDay: false,
            start: { dateTime: '2025-01-15T09:00:00' },
            end: { dateTime: '2025-01-15T10:00:00' },
            body: { contentType: 'text', content: teamsBody },
          },
        ],
      },
    });

    const result = await executeCalendar('test-token', { date: '2025-01-15' });

    expect(result).toContain('Sprint planning session');
    expect(result).not.toContain('Microsoft Teams meeting');
    expect(result).not.toContain('Passcode');
    expect(result).not.toContain('Dial in');
  });

  it('requests body field in $select', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeCalendar('test-token', { date: '2025-01-15' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('body');
    expect(calledPath).not.toContain('bodyPreview');
  });

  describe('calendars list mode', () => {
    it('lists calendars with default marker', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          value: [
            { name: 'Calendar', color: 'auto', isDefaultCalendar: true, canEdit: true },
            { name: 'Work', color: 'lightBlue', isDefaultCalendar: false, canEdit: false },
          ],
        },
      });

      const result = await executeCalendar('test-token', { calendars: true });

      expect(result).toContain('## Calendar (default)');
      expect(result).toContain('Color: auto');
      expect(result).toContain('Can edit: Yes');
      expect(result).toContain('## Work');
      expect(result).toContain('Can edit: No');
      expect(mockGraphFetch).toHaveBeenCalledWith(
        expect.stringContaining('/me/calendars'),
        'test-token',
      );
    });

    it('returns message when no calendars found', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: { value: [] },
      });

      const result = await executeCalendar('test-token', { calendars: true });

      expect(result).toBe('No calendars found.');
    });

    it('handles error when listing calendars', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: false,
        error: { status: 403, message: 'Insufficient permissions.' },
      });

      const result = await executeCalendar('test-token', { calendars: true });

      expect(result).toBe('Error: Insufficient permissions.');
    });
  });

  describe('event detail mode', () => {
    it('shows full event detail with attendees and Teams URL', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          subject: 'Sprint Planning',
          start: { dateTime: '2025-01-15T09:00:00', timeZone: 'Europe/London' },
          end: { dateTime: '2025-01-15T10:00:00', timeZone: 'Europe/London' },
          organizer: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
          attendees: [
            {
              emailAddress: { name: 'Bob', address: 'bob@example.com' },
              status: { response: 'accepted' },
            },
            {
              emailAddress: { name: 'Carol', address: 'carol@example.com' },
              status: { response: 'tentative' },
            },
          ],
          body: {
            contentType: 'html',
            content:
              '<html><body><p>Sprint planning session</p><p>Please prepare your updates</p></body></html>',
          },
          location: { displayName: 'Conference Room B' },
          onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/abc123' },
          hasAttachments: true,
          showAs: 'busy',
          importance: 'high',
          categories: ['Work', 'Sprint'],
        },
      });

      const result = await executeCalendar('test-token', { event_id: 'event-123' });

      expect(result).toContain('## Sprint Planning');
      expect(result).toContain('Time: 2025-01-15T09:00:00 - 2025-01-15T10:00:00');
      expect(result).toContain('Location: Conference Room B');
      expect(result).toContain('Organizer: Alice (alice@example.com)');
      expect(result).toContain('Bob (accepted)');
      expect(result).toContain('Carol (tentative)');
      expect(result).toContain('Sprint planning session');
      expect(result).toContain('Please prepare your updates');
      expect(result).not.toContain('<p>');
      expect(result).toContain('https://teams.microsoft.com/l/meetup-join/abc123');
      expect(result).toContain('Categories: Work, Sprint');
      expect(result).toContain('Importance: high');
      expect(result).toContain('Show as: busy');
      expect(mockGraphFetch).toHaveBeenCalledWith(
        expect.stringContaining('/me/events/event-123'),
        'test-token',
      );
    });

    it('handles event with no online meeting', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          subject: 'In-Person Meeting',
          start: { dateTime: '2025-01-15T14:00:00', timeZone: 'Europe/London' },
          end: { dateTime: '2025-01-15T15:00:00', timeZone: 'Europe/London' },
          organizer: { emailAddress: { name: 'Dave', address: 'dave@example.com' } },
          attendees: [],
          body: { contentType: 'text', content: 'Office meeting' },
          location: { displayName: 'Room C' },
          onlineMeeting: null,
        },
      });

      const result = await executeCalendar('test-token', { event_id: 'event-456' });

      expect(result).toContain('## In-Person Meeting');
      expect(result).toContain('Location: Room C');
      expect(result).not.toContain('Teams');
      expect(result).not.toContain('joinUrl');
    });

    it('handles event_id error', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: false,
        error: {
          status: 404,
          message: 'Event not found.',
        },
      });

      const result = await executeCalendar('test-token', { event_id: 'bad-id' });

      expect(result).toBe('Error: Event not found.');
    });
  });
});
