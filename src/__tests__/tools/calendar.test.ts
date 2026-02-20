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
});
