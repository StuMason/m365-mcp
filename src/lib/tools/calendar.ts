import { graphFetch } from '../graph.js';

export const calendarToolDefinition = {
  name: 'ms_calendar',
  description:
    "Fetch the user's Microsoft 365 calendar events. Defaults to today if no date params given.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      date: { type: 'string', description: 'Fetch events for a specific date (YYYY-MM-DD)' },
      start: { type: 'string', description: 'Start of date range (ISO 8601)' },
      end: { type: 'string', description: 'End of date range (ISO 8601)' },
    },
  },
};

interface CalendarEvent {
  subject?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
  location?: { displayName?: string };
  organizer?: { emailAddress?: { name?: string } };
  attendees?: Array<{ emailAddress?: { name?: string } }>;
  isAllDay?: boolean;
  bodyPreview?: string;
  onlineMeeting?: { joinUrl?: string } | null;
  webLink?: string;
}

interface CalendarViewResponse {
  value: CalendarEvent[];
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
 * Formats a calendar event into readable multi-line text.
 */
function formatEvent(event: CalendarEvent): string {
  const lines: string[] = [];

  lines.push(`## ${event.subject || 'Untitled'}`);

  if (event.isAllDay) {
    lines.push('Time: All day');
  } else {
    const startTime = event.start?.dateTime || 'N/A';
    const endTime = event.end?.dateTime || 'N/A';
    lines.push(`Time: ${startTime} - ${endTime}`);
  }

  if (event.location?.displayName) {
    lines.push(`Location: ${event.location.displayName}`);
  }

  if (event.organizer?.emailAddress?.name) {
    lines.push(`Organizer: ${event.organizer.emailAddress.name}`);
  }

  if (event.attendees && event.attendees.length > 0) {
    const names = event.attendees
      .map((a) => a.emailAddress?.name)
      .filter(Boolean)
      .join(', ');
    if (names) {
      lines.push(`Attendees: ${names}`);
    }
  }

  if (event.bodyPreview) {
    lines.push(event.bodyPreview);
  }

  return lines.join('\n');
}

/**
 * Fetches calendar events for the specified date range and returns
 * a human-readable summary.
 */
export async function executeCalendar(
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

  const select = [
    'subject',
    'start',
    'end',
    'location',
    'attendees',
    'isAllDay',
    'bodyPreview',
    'organizer',
    'onlineMeeting',
    'webLink',
  ].join(',');

  const path =
    `/me/calendarView?startDateTime=${start}&endDateTime=${end}` +
    `&$orderby=start/dateTime&$top=50&$select=${select}`;

  const result = await graphFetch<CalendarViewResponse>(path, token, { timezone: true });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const events = result.data.value;
  if (!events || events.length === 0) {
    return 'No calendar events found for the specified date range.';
  }

  return events.map(formatEvent).join('\n\n');
}
