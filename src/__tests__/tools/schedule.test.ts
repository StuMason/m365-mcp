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
});
