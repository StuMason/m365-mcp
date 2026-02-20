import { graphFetch } from '../graph.js';

const DEFAULT_CHUNK_SIZE = 10_000;
const MAX_CHUNK_SIZE = 50_000;

export const transcriptsToolDefinition = {
  name: 'ms_transcripts',
  description:
    'Fetch meeting transcripts from Microsoft Teams. ' +
    'Without transcript_id: lists meetings with ~3000 char previews. ' +
    'With transcript_id: returns transcript content in chunks (default 10,000 chars). ' +
    'Use offset to paginate through long transcripts.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      date: { type: 'string', description: 'Date (YYYY-MM-DD)' },
      start: { type: 'string', description: 'Start of date range (ISO 8601)' },
      end: { type: 'string', description: 'End of date range (ISO 8601)' },
      transcript_id: {
        type: 'string',
        description: 'Transcript ID for content drill-down (from a previous list call)',
      },
      offset: {
        type: 'integer',
        description:
          'Character offset for pagination (default 0). Use the value from the previous response to continue reading.',
      },
      length: {
        type: 'integer',
        description: 'Max characters to return (default 10000, max 50000)',
      },
    },
  },
};

interface CalendarEvent {
  subject?: string;
  start?: { dateTime?: string };
  end?: { dateTime?: string };
  attendees?: Array<{ emailAddress?: { name?: string } }>;
  organizer?: { emailAddress?: { name?: string } };
  onlineMeeting?: { joinUrl?: string } | null;
}

interface CalendarViewResponse {
  value: CalendarEvent[];
}

interface TranscriptEntry {
  id: string;
  createdDateTime?: string;
}

interface TranscriptsResponse {
  value: TranscriptEntry[];
}

interface OnlineMeetingResponse {
  subject?: string;
}

/**
 * Extracts a Graph online meeting ID from a Teams join URL.
 *
 * Join URL format:
 *   https://teams.microsoft.com/l/meetup-join/{threadId}/0?context={"Tid":"...","Oid":"..."}
 *
 * Meeting ID = base64("1*{organizerOid}*0**{threadId}")
 */
export function extractMeetingId(joinUrl: string): string | null {
  try {
    const url = new URL(joinUrl);
    const parts = url.pathname.split('/');
    const joinIdx = parts.indexOf('meetup-join');
    if (joinIdx === -1 || joinIdx + 1 >= parts.length) return null;

    const threadId = decodeURIComponent(parts[joinIdx + 1]);
    const contextParam = url.searchParams.get('context');
    if (!contextParam) return null;

    const context = JSON.parse(contextParam) as { Oid?: string };
    const organizerOid = context.Oid;
    if (!organizerOid) return null;

    return Buffer.from(`1*${organizerOid}*0**${threadId}`).toString('base64');
  } catch {
    return null;
  }
}

/**
 * Parses compound transcript ID format: {meetingId}/{transcriptId}
 */
export function parseTranscriptId(
  transcriptId: string,
): { meetingId: string; transcriptId: string } | null {
  const slashIdx = transcriptId.indexOf('/');
  if (slashIdx === -1 || slashIdx === 0 || slashIdx === transcriptId.length - 1) return null;
  return {
    meetingId: transcriptId.slice(0, slashIdx),
    transcriptId: transcriptId.slice(slashIdx + 1),
  };
}

/**
 * Computes the start and end ISO strings for a given YYYY-MM-DD date.
 */
function dateRangeForDay(dateStr: string): { start: string; end: string } | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return null;
  const date = new Date(`${dateStr}T00:00:00.000Z`);
  if (isNaN(date.getTime())) return null;
  const next = new Date(date);
  next.setUTCDate(next.getUTCDate() + 1);
  return { start: date.toISOString(), end: next.toISOString() };
}

/**
 * Returns today's date range (local midnight to next midnight) in ISO format.
 */
function todayRange(): { start: string; end: string } {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return dateRangeForDay(`${year}-${month}-${day}`)!;
}

/**
 * Matches transcripts to a specific calendar event occurrence.
 * For recurring meetings (multiple transcripts sharing the same meeting ID),
 * finds the transcript whose createdDateTime is closest to the event's start time.
 *
 * When there's only one transcript, returns it without filtering.
 * When no createdDateTime data is available, falls back to returning all transcripts.
 */
export function matchTranscriptsToEvent(
  transcripts: TranscriptEntry[],
  event: CalendarEvent,
): TranscriptEntry[] {
  if (transcripts.length <= 1) {
    return transcripts;
  }

  const eventStartStr = event.start?.dateTime;
  if (!eventStartStr) {
    return transcripts;
  }

  const eventStartMs = new Date(eventStartStr).getTime();
  if (isNaN(eventStartMs)) {
    return transcripts;
  }

  // Find transcript with createdDateTime closest to event start
  let closest: TranscriptEntry | null = null;
  let closestDiff = Infinity;

  for (const t of transcripts) {
    if (!t.createdDateTime) continue;
    const createdMs = new Date(t.createdDateTime).getTime();
    if (isNaN(createdMs)) continue;
    const diff = Math.abs(createdMs - eventStartMs);
    if (diff < closestDiff) {
      closest = t;
      closestDiff = diff;
    }
  }

  // Only match if within 24 hours — handles timezone discrepancies between
  // event times (user's preferred timezone) and transcript UTC timestamps,
  // while still distinguishing daily recurring meeting occurrences.
  const MAX_DIFF_MS = 24 * 60 * 60 * 1000;
  if (closest && closestDiff <= MAX_DIFF_MS) {
    return [closest];
  }

  if (!closest) {
    // No transcripts have valid createdDateTime — fall back to returning all
    return transcripts;
  }

  // Closest transcript is too far from this event — no match for this occurrence
  return [];
}

/**
 * Fetches VTT transcript content using raw fetch (not graphFetch, since it returns text).
 * Tries v1.0 first, falls back to beta on 403/400.
 */
async function fetchVttContent(
  token: string,
  meetingId: string,
  transcriptId: string,
): Promise<string | null> {
  const encodedMeetingId = encodeURIComponent(meetingId);
  const bases = ['https://graph.microsoft.com/v1.0', 'https://graph.microsoft.com/beta'];

  for (const base of bases) {
    const url = `${base}/me/onlineMeetings/${encodedMeetingId}/transcripts/${encodeURIComponent(transcriptId)}/content?$format=text/vtt`;
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` },
    });

    if (response.ok) {
      return response.text();
    }

    // Only retry with beta if 403 or 400; otherwise give up
    if (response.status !== 403 && response.status !== 400) {
      return null;
    }
  }

  return null;
}

/**
 * Fetches transcript list for a meeting. Tries v1.0, falls back to beta on 403/400.
 */
async function fetchTranscriptsList(
  token: string,
  meetingId: string,
): Promise<TranscriptsResponse | null> {
  const encodedMeetingId = encodeURIComponent(meetingId);
  const path = `/me/onlineMeetings/${encodedMeetingId}/transcripts`;

  const v1Result = await graphFetch<TranscriptsResponse>(path, token);
  if (v1Result.ok) {
    return v1Result.data;
  }

  if (v1Result.error.status === 403 || v1Result.error.status === 400) {
    const betaResult = await graphFetch<TranscriptsResponse>(path, token, { beta: true });
    if (betaResult.ok) {
      return betaResult.data;
    }
  }

  return null;
}

/**
 * Handles drill-down mode: fetch transcript by compound ID with pagination.
 */
async function executeDrillDown(
  token: string,
  compoundId: string,
  offset: number,
  length: number,
): Promise<string> {
  const parsed = parseTranscriptId(compoundId);
  if (!parsed) {
    return 'Error: Invalid transcript_id format. Expected "{meetingId}/{transcriptId}".';
  }

  const { meetingId, transcriptId } = parsed;

  const vtt = await fetchVttContent(token, meetingId, transcriptId);
  if (!vtt) {
    return 'Error: Could not fetch transcript content. The transcript may have been deleted or you may lack permissions.';
  }

  // Fetch meeting subject (non-critical)
  let subject = '(Unknown meeting)';
  const meetingResult = await graphFetch<OnlineMeetingResponse>(
    `/me/onlineMeetings/${encodeURIComponent(meetingId)}?$select=subject`,
    token,
  );
  if (meetingResult.ok && meetingResult.data.subject) {
    subject = meetingResult.data.subject;
  }

  const totalLength = vtt.length;

  // Short transcript — return it all, no pagination needed
  if (totalLength <= length) {
    return `# Transcript: ${subject}\nLength: ${totalLength} chars (complete)\n\n${vtt}`;
  }

  // Paginated: slice the requested chunk
  const chunk = vtt.slice(offset, offset + length);
  const end = offset + chunk.length;
  const remaining = totalLength - end;

  const lines: string[] = [];
  lines.push(`# Transcript: ${subject}`);
  lines.push(`Length: ${totalLength} chars | Showing: ${offset}–${end} | Remaining: ${remaining}`);
  lines.push('');
  lines.push(chunk);

  if (remaining > 0) {
    lines.push('');
    lines.push(
      `--- To continue reading, call again with transcript_id="${compoundId}" offset=${end} ---`,
    );
  }

  return lines.join('\n');
}

/**
 * Handles list mode: find meetings with transcripts in date range.
 */
async function executeList(
  token: string,
  args: { date?: string; start?: string; end?: string },
): Promise<string> {
  let start: string;
  let end: string;

  if (args.date) {
    const range = dateRangeForDay(args.date);
    if (!range) return 'Error: Invalid date format. Expected YYYY-MM-DD.';
    start = range.start;
    end = range.end;
  } else if (args.start && args.end) {
    start = args.start;
    end = args.end;
  } else {
    const range = todayRange();
    start = range.start;
    end = range.end;
  }

  const select = 'subject,start,end,attendees,organizer,onlineMeeting';
  const path =
    `/me/calendarView?startDateTime=${start}&endDateTime=${end}` +
    `&$orderby=start/dateTime&$top=50&$select=${select}`;

  const result = await graphFetch<CalendarViewResponse>(path, token, { timezone: true });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const events = result.data.value;
  if (!events || events.length === 0) {
    return 'No Teams meetings found in the given date range.';
  }

  // Filter to events with Teams join URLs
  const meetingEvents = events.filter((e) => e.onlineMeeting?.joinUrl);

  if (meetingEvents.length === 0) {
    return 'No Teams meetings found in the given date range.';
  }

  // Extract unique meeting IDs and map events to them
  const meetingIdMap = new Map<string, string>(); // joinUrl -> meetingId
  const uniqueMeetingIds = new Set<string>();

  for (const event of meetingEvents) {
    const joinUrl = event.onlineMeeting?.joinUrl;
    if (!joinUrl || meetingIdMap.has(joinUrl)) continue;
    const meetingId = extractMeetingId(joinUrl);
    if (!meetingId) continue;
    meetingIdMap.set(joinUrl, meetingId);
    uniqueMeetingIds.add(meetingId);
  }

  // Fetch all transcript lists in parallel (one per unique meeting ID)
  const transcriptCache = new Map<string, TranscriptEntry[]>();
  const fetchPromises = [...uniqueMeetingIds].map(async (meetingId) => {
    const data = await fetchTranscriptsList(token, meetingId);
    transcriptCache.set(meetingId, data?.value ?? []);
  });
  await Promise.all(fetchPromises);

  const sections: string[] = [];
  let transcriptCount = 0;

  for (const event of meetingEvents) {
    const joinUrl = event.onlineMeeting?.joinUrl;
    if (!joinUrl) continue;

    const meetingId = meetingIdMap.get(joinUrl);
    if (!meetingId) continue;

    const allTranscripts = transcriptCache.get(meetingId);
    if (!allTranscripts || allTranscripts.length === 0) continue;

    // Match transcripts to this specific event occurrence (handles recurring meetings)
    const eventTranscripts = matchTranscriptsToEvent(allTranscripts, event);
    if (eventTranscripts.length === 0) continue;

    transcriptCount++;

    const firstTranscript = eventTranscripts[0];
    const compoundId = `${meetingId}/${firstTranscript.id}`;

    const lines: string[] = [];
    lines.push(`## ${event.subject || 'Untitled'}`);
    lines.push(`Date: ${event.start?.dateTime || 'N/A'}`);

    const attendeeNames = event.attendees
      ?.map((a) => a.emailAddress?.name)
      .filter(Boolean)
      .join(', ');
    if (attendeeNames) {
      lines.push(`Attendees: ${attendeeNames}`);
    }

    lines.push(`Transcript ID: ${compoundId}`);

    sections.push(lines.join('\n'));
  }

  if (transcriptCount === 0) {
    // Meetings found but none have transcripts
    const meetingList = meetingEvents
      .map((e) => `- ${e.subject || 'Untitled'} (${e.start?.dateTime || 'N/A'})`)
      .join('\n');
    return `Found ${meetingEvents.length} Teams meetings, but none have transcripts recorded.\n\n${meetingList}`;
  }

  const header = `Found ${meetingEvents.length} meetings, ${transcriptCount} with transcripts.`;
  return header + '\n\n' + sections.join('\n\n---\n\n');
}

/**
 * Fetches meeting transcripts from Microsoft Teams.
 * Supports two modes: list (date range) and drill-down (specific transcript with pagination).
 */
export async function executeTranscripts(
  token: string,
  args: {
    date?: string;
    start?: string;
    end?: string;
    transcript_id?: string;
    offset?: number;
    length?: number;
  },
): Promise<string> {
  if (args.transcript_id) {
    const offset = Math.max(args.offset ?? 0, 0);
    const length = Math.min(Math.max(args.length ?? DEFAULT_CHUNK_SIZE, 1), MAX_CHUNK_SIZE);
    return executeDrillDown(token, args.transcript_id, offset, length);
  }
  return executeList(token, args);
}
