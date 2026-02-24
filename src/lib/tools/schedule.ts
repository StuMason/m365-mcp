import { graphPost } from '../graph.js';

export const scheduleToolDefinition = {
  name: 'ms_schedule',
  description:
    "Check people's availability / free-busy status for a given time window. Accepts one or more email addresses and returns their schedule with time slots showing free, busy, tentative, out of office, or working elsewhere.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      emails: {
        type: 'array',
        items: { type: 'string' },
        description: 'Email addresses to check availability for (required)',
      },
      date: {
        type: 'string',
        description: 'Date to check (YYYY-MM-DD). Defaults to today.',
      },
      start: {
        type: 'string',
        description: 'Start time (HH:MM, 24h). Defaults to 08:00.',
      },
      end: {
        type: 'string',
        description: 'End time (HH:MM, 24h). Defaults to 18:00.',
      },
      interval: {
        type: 'number',
        description: 'Slot duration in minutes. Defaults to 30.',
      },
    },
    required: ['emails'],
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
    const start = item.start?.dateTime || '?';
    const end = item.end?.dateTime || '?';
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

  const body = {
    schedules: args.emails,
    startTime: { dateTime: `${date}T${start}:00`, timeZone: 'UTC' },
    endTime: { dateTime: `${date}T${end}:00`, timeZone: 'UTC' },
    availabilityViewInterval: interval,
  };

  const result = await graphPost<typeof body, ScheduleResponse>(
    '/me/calendar/getSchedule',
    token,
    body,
    { timezone: false },
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

    lines.push(`Date: ${date} | ${start} - ${end} (${interval}-min slots)`);
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
