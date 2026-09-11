import { z } from 'zod';
import { graphPost } from '../graph.js';
import { formatTime, timezone } from '../format.js';

export const scheduleToolDefinition = {
  name: 'ms_schedule',
  title: 'Free/Busy Schedule',
  description:
    "Check people's availability / free-busy status for a given time window. Accepts one or more email addresses and returns their schedule with time slots showing free, busy, tentative, out of office, or working elsewhere.",
  inputSchema: z.object({
    emails: z.array(z.string()).describe('Email addresses to check availability for (required)'),
    date: z.string().optional().describe('Date to check (YYYY-MM-DD). Defaults to today.'),
    start: z
      .string()
      .regex(/^([01]\d|2[0-3]):[0-5]\d$/, 'Expected a 24-hour time like 08:00')
      .optional()
      .describe('Start time (HH:MM, 24h). Defaults to 08:00.'),
    end: z
      .string()
      .regex(/^([01]\d|2[0-3]):[0-5]\d$/, 'Expected a 24-hour time like 18:00')
      .optional()
      .describe('End time (HH:MM, 24h). Defaults to 18:00.'),
    interval: z
      .int()
      .min(5)
      .max(1440)
      .optional()
      .describe('Slot duration in minutes. Graph accepts 5-1440. Defaults to 30.'),
  }),
  annotations: {
    title: 'Free/Busy Schedule',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

const STATUS_MAP: Record<string, string> = {
  '0': 'free',
  '1': 'tentative',
  '2': 'busy',
  '3': 'out of office',
  '4': 'working elsewhere',
};

interface ScheduleItem {
  subject?: string;
  start?: { dateTime?: string };
  end?: { dateTime?: string };
  status?: string;
}

interface ScheduleEntry {
  scheduleId: string;
  availabilityView?: string;
  scheduleItems?: ScheduleItem[];
  error?: { responseCode?: string; message?: string };
}

interface ScheduleResponse {
  value: ScheduleEntry[];
}

interface ScheduleArgs {
  emails: string[];
  date?: string;
  start?: string;
  end?: string;
  interval?: number;
}

/**
 * Returns today's date as a YYYY-MM-DD string.
 */
function todayDate(): string {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Decodes an availabilityView string into an array of time-slot labels.
 * Each character represents one interval-length slot.
 */
function decodeAvailabilityView(
  view: string,
  startTime: string,
  intervalMinutes: number,
): string[] {
  const lines: string[] = [];
  const [startHour, startMin] = startTime.split(':').map(Number);

  for (let i = 0; i < view.length; i++) {
    const totalMinutes = startHour * 60 + startMin + i * intervalMinutes;
    const h = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
    const m = String(totalMinutes % 60).padStart(2, '0');
    const status = STATUS_MAP[view[i]] ?? 'unknown';
    lines.push(`  ${h}:${m} - ${status}`);
  }

  return lines;
}

/**
 * Formats schedule items (meetings) into readable lines.
 */
function formatScheduleItems(items?: ScheduleItem[]): string[] {
  if (!items || items.length === 0) return [];

  const lines: string[] = ['', 'Scheduled items:'];
  for (const item of items) {
    const subject = item.subject || 'Untitled';
    const start = formatTime(item.start?.dateTime);
    const end = formatTime(item.end?.dateTime);
    const status = item.status || 'unknown';
    lines.push(`  - ${subject} (${start} to ${end}) [${status}]`);
  }
  return lines;
}

/**
 * Check people's free/busy availability via the Graph API getSchedule endpoint.
 */
export async function executeSchedule(token: string, args: ScheduleArgs): Promise<string> {
  if (!args.emails || args.emails.length === 0) {
    return 'Error: At least one email address is required.';
  }

  const date = args.date || todayDate();
  const start = args.start || '08:00';
  const end = args.end || '18:00';
  const interval = args.interval ?? 30;

  // The times are wall-clock in the user's zone, so they must be sent with that
  // zone. Labelling them UTC shifted every free/busy window by the offset —
  // asking for 08:00 actually queried 09:00 in London, 10:00 in Brussels.
  const tz = timezone();
  const body = {
    schedules: args.emails,
    startTime: { dateTime: `${date}T${start}:00`, timeZone: tz },
    endTime: { dateTime: `${date}T${end}:00`, timeZone: tz },
    availabilityViewInterval: interval,
  };

  // The Prefer header matters here, not just on the request times. Without it
  // getSchedule returns scheduleItems in UTC as offset-less strings — shapes that
  // are indistinguishable from local wall-clock, so a 14:00 meeting came back as
  // "13:00" and would be labelled with the local zone. With it, Graph returns the
  // items already in the requested zone, matching ms_calendar.
  const result = await graphPost<typeof body, ScheduleResponse>(
    '/me/calendar/getSchedule',
    token,
    body,
  );

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const entries = result.data.value;
  if (!entries || entries.length === 0) {
    return 'No schedule data returned.';
  }

  const sections: string[] = [];

  for (const entry of entries) {
    const lines: string[] = [];
    lines.push(`## ${entry.scheduleId}`);

    if (entry.error) {
      lines.push(
        `Error: Unable to retrieve schedule — ${entry.error.message || entry.error.responseCode || 'unknown error'}`,
      );
      sections.push(lines.join('\n'));
      continue;
    }

    lines.push(`Date: ${date} | ${start} - ${end} ${tz} (${interval}-min slots)`);
    lines.push('');

    if (entry.availabilityView) {
      lines.push('Availability:');
      lines.push(...decodeAvailabilityView(entry.availabilityView, start, interval));
    }

    lines.push(...formatScheduleItems(entry.scheduleItems));

    sections.push(lines.join('\n'));
  }

  return sections.join('\n\n');
}
