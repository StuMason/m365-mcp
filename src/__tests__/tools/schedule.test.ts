import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphPost =
  jest.fn<
    <TBody, TResult>(
      path: string,
      token: string,
      body: TBody,
      options?: { beta?: boolean; timezone?: boolean; headers?: Record<string, string> },
    ) => Promise<GraphResult<TResult>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphPost: mockGraphPost,
}));

const { executeSchedule } = await import('../../lib/tools/schedule.js');

describe('executeSchedule', () => {
  afterEach(() => {
    mockGraphPost.mockReset();
  });

  it('checks availability for a single person', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'alice@example.com',
            availabilityView: '000022220000',
            scheduleItems: [
              {
                subject: 'Sprint Planning',
                start: { dateTime: '2026-02-23T10:00:00' },
                end: { dateTime: '2026-02-23T12:00:00' },
                status: 'busy',
              },
            ],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['alice@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('alice@example.com');
    expect(result).toContain('free');
    expect(result).toContain('busy');
    expect(mockGraphPost).toHaveBeenCalledWith(
      '/me/calendar/getSchedule',
      'test-token',
      expect.objectContaining({
        schedules: ['alice@example.com'],
      }),
      expect.any(Object),
    );
  });

  it('checks availability for multiple people', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'alice@example.com',
            availabilityView: '0022',
            scheduleItems: [],
          },
          {
            scheduleId: 'bob@example.com',
            availabilityView: '0000',
            scheduleItems: [],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['alice@example.com', 'bob@example.com'],
      date: '2026-02-23',
      start: '09:00',
      end: '11:00',
      interval: 30,
    });

    expect(result).toContain('alice@example.com');
    expect(result).toContain('bob@example.com');
  });

  it('formats OOF and tentative statuses', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'user@example.com',
            availabilityView: '01340',
            scheduleItems: [],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('free');
    expect(result).toContain('tentative');
    expect(result).toContain('out of office');
    expect(result).toContain('working elsewhere');
  });

  it('handles error from Graph API', async () => {
    mockGraphPost.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
    });

    expect(result).toContain('Error');
  });

  it('defaults to 08:00-18:00 and 30-minute intervals', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(mockGraphPost).toHaveBeenCalledWith(
      '/me/calendar/getSchedule',
      'test-token',
      expect.objectContaining({
        startTime: expect.objectContaining({ dateTime: '2026-02-23T08:00:00' }),
        endTime: expect.objectContaining({ dateTime: '2026-02-23T18:00:00' }),
        availabilityViewInterval: 30,
      }),
      expect.any(Object),
    );
  });

  it('returns error when emails is empty', async () => {
    const result = await executeSchedule('test-token', { emails: [] });
    expect(result).toBe('Error: At least one email address is required.');
    expect(mockGraphPost).not.toHaveBeenCalled();
  });

  it('returns message when API returns empty value array', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(result).toBe('No schedule data returned.');
  });

  it('handles error with exact error message', async () => {
    mockGraphPost.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
    });

    expect(result).toBe('Error: Insufficient permissions.');
  });

  it('formats schedule items with missing fields gracefully', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'user@example.com',
            availabilityView: '02',
            scheduleItems: [
              {
                subject: undefined,
                start: undefined,
                end: undefined,
                status: undefined,
              },
            ],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('Untitled');
    expect(result).toContain('? to ?');
    expect(result).toContain('[unknown]');
  });

  it('handles per-user error for non-existent email', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'fake@example.com',
            error: {
              responseCode: 'ErrorMailRecipientNotFound',
              message: 'The specified recipient was not found.',
            },
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['fake@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('fake@example.com');
    expect(result).toContain('Unable to retrieve schedule');
    expect(result).toContain('The specified recipient was not found.');
    expect(result).not.toContain('Availability:');
  });

  it('handles mixed valid and error entries', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'alice@example.com',
            availabilityView: '0022',
            scheduleItems: [],
          },
          {
            scheduleId: 'fake@example.com',
            error: {
              responseCode: 'ErrorMailRecipientNotFound',
            },
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['alice@example.com', 'fake@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('alice@example.com');
    expect(result).toContain('free');
    expect(result).toContain('fake@example.com');
    expect(result).toContain('Unable to retrieve schedule');
  });

  it('handles undefined scheduleItems without crashing', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'user@example.com',
            availabilityView: '00',
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('user@example.com');
    expect(result).toContain('free');
  });
});
