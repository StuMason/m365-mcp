import { z } from 'zod';
import { executeCalendar } from './calendar.js';
import { executeMail } from './mail.js';
import { executeChat } from './chat.js';
import { executeTasks } from './tasks.js';
import { executeTranscripts } from './transcripts.js';
import { executePeople } from './people.js';
import { executeSearch } from './search.js';
import { echo } from '../format.js';

export const briefToolDefinition = {
  name: 'ms_brief',
  title: 'Daily Brief',
  description:
    'One call that assembles a catch-up. With no arguments: today’s meetings, unread mail, recent chats, open tasks and yesterday’s meeting transcripts. With person: who they are and your recent mail and chats involving them. Use this instead of calling four or five tools separately.',
  inputSchema: z.object({
    person: z
      .string()
      .optional()
      .describe('Catch up on one person instead — their name or email address'),
    date: z.string().optional().describe('Date for the brief (YYYY-MM-DD). Defaults to today.'),
    count: z.int().min(1).max(15).optional().describe('Max items per section (1-15, default 5)'),
  }),
  annotations: {
    title: 'Daily Brief',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

/** Longest a single section may run before it is trimmed. */
const SECTION_CAP = 2500;

/**
 * Trims a section to the cap without cutting a fence in half.
 *
 * Sections contain <<<UNTRUSTED:nonce …>>> … <<<END UNTRUSTED:nonce>>> blocks. A
 * blind slice left a fence open, so everything after it — including this
 * server's own text — read as untrusted content, or worse, the reverse.
 * Trimming back to the last completed block keeps every fence balanced.
 */
function trimSection(body: string): string {
  if (body.length <= SECTION_CAP) {
    return body;
  }

  const cut = body.slice(0, SECTION_CAP);
  const lastEnd = cut.lastIndexOf('<<<END UNTRUSTED');
  const lastStart = cut.lastIndexOf('<<<UNTRUSTED:');

  // An unclosed fence was opened inside the cut: drop back to just before it.
  let safe = cut;
  if (lastStart > lastEnd) {
    safe = cut.slice(0, lastStart);
  } else if (lastEnd > -1) {
    // Keep the closing marker whole rather than slicing through it.
    const endOfMarker = cut.indexOf('>>>', lastEnd);
    safe = endOfMarker === -1 ? cut.slice(0, lastEnd) : cut.slice(0, endOfMarker + 3);
  }

  return (
    `${safe.trimEnd()}\n\n` +
    '…trimmed to fit the brief. Any count above describes the full result, not what ' +
    'is shown here — call the underlying tool for the rest.'
  );
}

/**
 * Runs one section, converting a failure into a short note rather than losing the
 * whole brief. One dead section should not cost you the other five.
 */
async function section(title: string, run: () => Promise<string>): Promise<string> {
  let body: string;
  try {
    body = await run();
  } catch (error) {
    body = `(unavailable: ${error instanceof Error ? error.message : String(error)})`;
  }

  // A scope the registration lacks is a permanent, whole-section condition. The
  // tool explains it at length because it is the whole answer there; here it is one
  // line among six, so only the first line earns its place.
  if (body.startsWith('Not available:')) {
    return `# ${title}\n\n${body.split('\n')[0]}`;
  }

  return `# ${title}\n\n${trimSection(body).trim() || '(nothing)'}`;
}

/**
 * Returns the ISO date (YYYY-MM-DD) `daysAgo` days before the given date.
 * Assumes the date has already been validated by isRealDate.
 */
function shiftDate(date: string, daysAgo: number): string {
  const d = new Date(`${date}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() - daysAgo);
  return d.toISOString().slice(0, 10);
}

/**
 * True only for a date that exists.
 *
 * JavaScript rolls impossible dates forward, so 2026-04-31 silently becomes
 * 2026-05-01. Round-tripping is the only way to catch it.
 */
function isRealDate(date: string): boolean {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
  const d = new Date(`${date}T00:00:00.000Z`);
  return !isNaN(d.getTime()) && d.toISOString().slice(0, 10) === date;
}

/**
 * Assembles a daily brief, or a catch-up on one person.
 *
 * Deliberately thin: every section is an existing tool. Nothing here talks to
 * Graph directly, so the formatting and error handling stay in one place per area.
 */
export async function executeBrief(
  token: string,
  args: { person?: string; date?: string; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count || 5, 1), 15);
  const date = args.date || new Date().toISOString().slice(0, 10);

  // Fail the whole call rather than letting sections disagree about the date.
  if (!isRealDate(date)) {
    return `Error: "${echo(date)}" is not a valid date. Expected YYYY-MM-DD, and the day must exist in that month.`;
  }

  // Catch up on one person.
  if (args.person) {
    const sections = await Promise.all([
      section(`Who is ${args.person}`, () =>
        executePeople(token, { search: args.person!, count: 3 }),
      ),
      section('Recent mail and chats', () =>
        executeSearch(token, { query: args.person!, types: ['mail', 'chat'], count }),
      ),
    ]);
    return sections.join('\n\n');
  }

  // The daily brief.
  const yesterday = shiftDate(date, 1);

  const sections = await Promise.all([
    section(`Meetings on ${date}`, () => executeCalendar(token, { date, compact: true })),
    section('Unread mail', () =>
      // Inbox explicitly, matching ms_mail's default for a filter.
      executeMail(token, { filter: 'unread', folder: 'Inbox', count }),
    ),
    section('Recent chats', () => executeChat(token, { count })),
    section('Open Planner tasks', () => executeTasks(token, { planner: true, count })),
    section(`Meeting transcripts from ${yesterday}`, () =>
      executeTranscripts(token, { date: yesterday }),
    ),
  ]);

  return sections.join('\n\n');
}
