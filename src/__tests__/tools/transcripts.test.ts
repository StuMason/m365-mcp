import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

// ── Mocks ──────────────────────────────────────────────────────────────────

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

// Mock global fetch for VTT content downloads
const mockFetch = jest.fn<typeof globalThis.fetch>();
globalThis.fetch = mockFetch as typeof globalThis.fetch;

// Dynamic import AFTER mocks are registered
const { executeTranscripts, extractMeetingId, parseTranscriptId } =
  await import('../../lib/tools/transcripts.js');

// ── Helpers ────────────────────────────────────────────────────────────────

function makeJoinUrl(threadId: string, oid: string, tid = 'tenant-123'): string {
  const context = JSON.stringify({ Tid: tid, Oid: oid });
  return `https://teams.microsoft.com/l/meetup-join/${encodeURIComponent(threadId)}/0?context=${encodeURIComponent(context)}`;
}

function mockResponse(body: string, status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    text: () => Promise.resolve(body),
    headers: new Headers(),
    redirected: false,
    statusText: status === 200 ? 'OK' : 'Error',
    type: 'basic',
    url: '',
    clone: () => mockResponse(body, status),
    body: null,
    bodyUsed: false,
    arrayBuffer: () => Promise.resolve(new ArrayBuffer(0)),
    blob: () => Promise.resolve(new Blob()),
    formData: () => Promise.resolve(new FormData()),
    json: () => Promise.resolve(JSON.parse(body)),
    bytes: () => Promise.resolve(new Uint8Array()),
  } as Response;
}

// ── extractMeetingId ───────────────────────────────────────────────────────

describe('extractMeetingId', () => {
  it('extracts valid meeting ID from standard Teams join URL', () => {
    const threadId = '19:meeting_abc123@thread.v2';
    const oid = 'user-oid-456';
    const url = makeJoinUrl(threadId, oid);

    const result = extractMeetingId(url);

    const expected = Buffer.from(`1*${oid}*0**${threadId}`).toString('base64');
    expect(result).toBe(expected);
  });

  it('returns null for URL without meetup-join path', () => {
    const result = extractMeetingId('https://teams.microsoft.com/l/channel/foo/bar');
    expect(result).toBeNull();
  });

  it('returns null for URL missing context param', () => {
    const result = extractMeetingId('https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0');
    expect(result).toBeNull();
  });

  it('returns null for URL with context missing Oid', () => {
    const context = JSON.stringify({ Tid: 'tenant-123' });
    const url = `https://teams.microsoft.com/l/meetup-join/19:abc@thread.v2/0?context=${encodeURIComponent(context)}`;
    const result = extractMeetingId(url);
    expect(result).toBeNull();
  });

  it('returns null for completely invalid URLs', () => {
    expect(extractMeetingId('not-a-url')).toBeNull();
    expect(extractMeetingId('')).toBeNull();
  });

  it('returns null when meetup-join has no thread segment after it', () => {
    const context = JSON.stringify({ Tid: 'tenant-123', Oid: 'oid-456' });
    const url = `https://teams.microsoft.com/l/meetup-join?context=${encodeURIComponent(context)}`;
    const result = extractMeetingId(url);
    expect(result).toBeNull();
  });

  it('handles URL-encoded thread IDs', () => {
    const threadId = '19:meeting_special+chars@thread.v2';
    const oid = 'oid-789';
    const url = makeJoinUrl(threadId, oid);

    const result = extractMeetingId(url);

    const expected = Buffer.from(`1*${oid}*0**${threadId}`).toString('base64');
    expect(result).toBe(expected);
  });
});

// ── parseTranscriptId ──────────────────────────────────────────────────────

describe('parseTranscriptId', () => {
  it('parses valid compound ID', () => {
    const result = parseTranscriptId('meetingABC/transcript123');
    expect(result).toEqual({
      meetingId: 'meetingABC',
      transcriptId: 'transcript123',
    });
  });

  it('handles transcript IDs containing slashes after the first', () => {
    const result = parseTranscriptId('meetingABC/transcript/with/slashes');
    expect(result).toEqual({
      meetingId: 'meetingABC',
      transcriptId: 'transcript/with/slashes',
    });
  });

  it('returns null for input without slash', () => {
    expect(parseTranscriptId('noslashhere')).toBeNull();
  });

  it('returns null for empty string', () => {
    expect(parseTranscriptId('')).toBeNull();
  });

  it('returns null when slash is at start', () => {
    expect(parseTranscriptId('/transcriptOnly')).toBeNull();
  });

  it('returns null when slash is at end', () => {
    expect(parseTranscriptId('meetingOnly/')).toBeNull();
  });

  it('handles base64 meeting IDs with = padding', () => {
    const result = parseTranscriptId('MSo1Njc4OTAx==/transcript-id');
    expect(result).toEqual({
      meetingId: 'MSo1Njc4OTAx==',
      transcriptId: 'transcript-id',
    });
  });
});

// ── executeTranscripts ─────────────────────────────────────────────────────

describe('executeTranscripts', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
    mockFetch.mockReset();
  });

  // ── Drill-down mode ──────────────────────────────────────────────────

  describe('drill-down mode', () => {
    it('returns full transcript for valid compound ID', async () => {
      const vttContent = 'WEBVTT\n\n00:00:00.000 --> 00:00:05.000\nHello world';

      // VTT content fetch (raw fetch, v1.0 succeeds)
      mockFetch.mockResolvedValueOnce(mockResponse(vttContent));

      // Meeting subject fetch
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { subject: 'Team Standup' },
      });

      const result = await executeTranscripts('test-token', {
        transcript_id: 'meetingId123/transcriptId456',
      });

      expect(result).toContain('# Transcript: Team Standup');
      expect(result).toContain(vttContent);
    });

    it('falls back to beta when v1.0 returns 403', async () => {
      const vttContent = 'WEBVTT\n\nFallback content';

      // v1.0 returns 403
      mockFetch.mockResolvedValueOnce(mockResponse('Forbidden', 403));
      // beta returns successfully
      mockFetch.mockResolvedValueOnce(mockResponse(vttContent));

      // Meeting subject
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { subject: 'Beta Meeting' },
      });

      const result = await executeTranscripts('test-token', {
        transcript_id: 'meetingId/transcriptId',
      });

      expect(result).toContain('# Transcript: Beta Meeting');
      expect(result).toContain(vttContent);
      expect(mockFetch).toHaveBeenCalledTimes(2);
      expect(mockFetch).toHaveBeenNthCalledWith(
        1,
        expect.stringContaining('graph.microsoft.com/v1.0'),
        expect.any(Object),
      );
      expect(mockFetch).toHaveBeenNthCalledWith(
        2,
        expect.stringContaining('graph.microsoft.com/beta'),
        expect.any(Object),
      );
    });

    it('falls back to beta when v1.0 returns 400', async () => {
      const vttContent = 'WEBVTT\n\nBeta fallback';

      mockFetch.mockResolvedValueOnce(mockResponse('Bad Request', 400));
      mockFetch.mockResolvedValueOnce(mockResponse(vttContent));

      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { subject: 'Meeting 400' },
      });

      const result = await executeTranscripts('test-token', {
        transcript_id: 'meeting/transcript',
      });

      expect(result).toContain(vttContent);
    });

    it('returns error when both v1.0 and beta fail', async () => {
      mockFetch.mockResolvedValueOnce(mockResponse('Forbidden', 403));
      mockFetch.mockResolvedValueOnce(mockResponse('Forbidden', 403));

      const result = await executeTranscripts('test-token', {
        transcript_id: 'meeting/transcript',
      });

      expect(result).toContain('Error: Could not fetch transcript content');
    });

    it('returns error for invalid compound ID format', async () => {
      const result = await executeTranscripts('test-token', {
        transcript_id: 'no-slash-here',
      });

      expect(result).toContain('Error: Invalid transcript_id format');
    });

    it('uses (Unknown meeting) when subject fetch fails', async () => {
      const vttContent = 'WEBVTT\n\nContent here';

      mockFetch.mockResolvedValueOnce(mockResponse(vttContent));

      mockGraphFetch.mockResolvedValueOnce({
        ok: false,
        error: { status: 404, message: 'Not found' },
      });

      const result = await executeTranscripts('test-token', {
        transcript_id: 'meeting/transcript',
      });

      expect(result).toContain('# Transcript: (Unknown meeting)');
      expect(result).toContain(vttContent);
    });
  });

  // ── List mode ────────────────────────────────────────────────────────

  describe('list mode', () => {
    it('returns previews with transcript IDs', async () => {
      const threadId = '19:meeting_abc@thread.v2';
      const oid = 'organizer-oid';
      const joinUrl = makeJoinUrl(threadId, oid);
      const meetingId = Buffer.from(`1*${oid}*0**${threadId}`).toString('base64');

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'Sprint Planning',
              start: { dateTime: '2025-06-15T10:00:00' },
              end: { dateTime: '2025-06-15T11:00:00' },
              attendees: [{ emailAddress: { name: 'Alice' } }, { emailAddress: { name: 'Bob' } }],
              organizer: { emailAddress: { name: 'Alice' } },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // Transcripts list (v1.0)
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [{ id: 'transcript-001' }] },
      });

      // VTT content
      const vttContent = 'WEBVTT\n\n00:00:00.000 --> 00:00:05.000\nHello from sprint';
      mockFetch.mockResolvedValueOnce(mockResponse(vttContent));

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toContain('Found 1 meetings, 1 with transcripts.');
      expect(result).toContain('## Sprint Planning');
      expect(result).toContain('Date: 2025-06-15T10:00:00');
      expect(result).toContain('Attendees: Alice, Bob');
      expect(result).toContain(`Transcript ID: ${meetingId}/transcript-001`);
      expect(result).toContain(vttContent);
    });

    it('truncates VTT preview to ~3000 chars', async () => {
      const threadId = '19:meeting_long@thread.v2';
      const oid = 'oid-long';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'Long Meeting',
              start: { dateTime: '2025-06-15T10:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // Transcripts list
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [{ id: 'tx-long' }] },
      });

      // Long VTT content
      const longVtt = 'WEBVTT\n\n' + 'A'.repeat(4000);
      mockFetch.mockResolvedValueOnce(mockResponse(longVtt));

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toContain('... [truncated — use transcript_id for full content]');
      // The preview should not contain the full 4000-char string
      expect(result).not.toContain('A'.repeat(4000));
    });

    it('handles empty calendar (no meetings)', async () => {
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [] },
      });

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toBe('No Teams meetings found in the given date range.');
    });

    it('handles meetings with no transcripts', async () => {
      const threadId = '19:meeting_notx@thread.v2';
      const oid = 'oid-notx';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'No Recording',
              start: { dateTime: '2025-06-15T14:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // Transcripts list returns empty
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [] },
      });

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toContain('none have transcripts recorded');
      expect(result).toContain('No Recording');
    });

    it('handles Graph API error in calendar fetch', async () => {
      mockGraphFetch.mockResolvedValueOnce({
        ok: false,
        error: {
          status: 403,
          message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
        },
      });

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toBe(
        'Error: Insufficient permissions. Check granted scopes with ms_auth_status.',
      );
    });

    it('uses start/end params directly when provided', async () => {
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [] },
      });

      await executeTranscripts('test-token', {
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
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [] },
      });

      const now = new Date();
      const year = now.getFullYear();
      const month = String(now.getMonth() + 1).padStart(2, '0');
      const day = String(now.getDate()).padStart(2, '0');
      const todayStr = `${year}-${month}-${day}`;

      await executeTranscripts('test-token', {});

      expect(mockGraphFetch).toHaveBeenCalledWith(
        expect.stringContaining(`startDateTime=${todayStr}`),
        'test-token',
        { timezone: true },
      );
    });

    it('returns null VTT gracefully when v1.0 returns non-retryable status (e.g. 500)', async () => {
      const threadId = '19:meeting_500@thread.v2';
      const oid = 'oid-500';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'Server Error Meeting',
              start: { dateTime: '2025-06-15T10:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // Transcripts list succeeds
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [{ id: 'tx-500' }] },
      });

      // v1.0 VTT fetch returns 500 (non-retryable)
      mockFetch.mockResolvedValueOnce(mockResponse('Internal Server Error', 500));

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      // Meeting has transcripts but VTT download failed — still shows meeting without preview
      expect(result).toContain('1 with transcripts');
      expect(result).toContain('Server Error Meeting');
      expect(mockFetch).toHaveBeenCalledTimes(1); // No beta retry for 500
    });

    it('skips meeting when both v1.0 and beta transcript listing fail with non-retryable error', async () => {
      const threadId = '19:meeting_fail@thread.v2';
      const oid = 'oid-fail';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'Failed Transcript List',
              start: { dateTime: '2025-06-15T10:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // v1.0 transcript list fails with 500 (non-retryable)
      mockGraphFetch.mockResolvedValueOnce({
        ok: false,
        error: { status: 500, message: 'Server error' },
      });

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      // Meeting found but transcript listing failed — shows as "none have transcripts"
      expect(result).toContain('none have transcripts');
    });

    it('returns error for invalid date format', async () => {
      const result = await executeTranscripts('test-token', { date: 'not-a-date' });

      expect(result).toBe('Error: Invalid date format. Expected YYYY-MM-DD.');
    });

    it('falls back to beta for transcript listing on 403', async () => {
      const threadId = '19:meeting_beta@thread.v2';
      const oid = 'oid-beta';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'Beta Transcripts',
              start: { dateTime: '2025-06-15T10:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // v1.0 transcripts list fails with 403
      mockGraphFetch.mockResolvedValueOnce({
        ok: false,
        error: { status: 403, message: 'Forbidden' },
      });

      // beta transcripts list succeeds
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [{ id: 'tx-beta' }] },
      });

      // VTT content
      mockFetch.mockResolvedValueOnce(mockResponse('WEBVTT\n\nBeta content'));

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      expect(result).toContain('1 with transcripts');
      expect(result).toContain('Beta Transcripts');
    });

    it('skips events without onlineMeeting joinUrl', async () => {
      const threadId = '19:meeting_ok@thread.v2';
      const oid = 'oid-ok';
      const joinUrl = makeJoinUrl(threadId, oid);

      // Calendar view with mix of online and non-online meetings
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: {
          value: [
            {
              subject: 'In-Person Meeting',
              start: { dateTime: '2025-06-15T09:00:00' },
              onlineMeeting: null,
            },
            {
              subject: 'Online Meeting',
              start: { dateTime: '2025-06-15T10:00:00' },
              onlineMeeting: { joinUrl },
            },
          ],
        },
      });

      // Transcripts list for the online meeting
      mockGraphFetch.mockResolvedValueOnce({
        ok: true,
        data: { value: [{ id: 'tx-1' }] },
      });

      // VTT content
      mockFetch.mockResolvedValueOnce(mockResponse('WEBVTT\n\nOnline content'));

      const result = await executeTranscripts('test-token', { date: '2025-06-15' });

      // Only 1 meeting (with joinUrl) counted, 1 with transcripts
      expect(result).toContain('Found 1 meetings, 1 with transcripts.');
      expect(result).toContain('Online Meeting');
      expect(result).not.toContain('In-Person Meeting');
    });
  });
});
