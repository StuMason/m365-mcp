import { z } from 'zod';
import { executeCalendar } from './calendar.js';
import { executeMail } from './mail.js';
import { executeChat } from './chat.js';
import { executeTasks } from './tasks.js';
import { executeTranscripts } from './transcripts.js';
import { executePeople } from './people.js';
import { executeSearch } from './search.js';

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

  const trimmed =
    body.length > SECTION_CAP
      ? `${body.slice(0, SECTION_CAP)}\n\n…trimmed. Use the underlying tool for the full list.`
      : body;

  return `# ${title}\n\n${trimmed.trim() || '(nothing)'}`;
}

/**
 * Returns the ISO date (YYYY-MM-DD) `daysAgo` days before the given date.
 */
function shiftDate(date: string, daysAgo: number): string {
  const d = new Date(`${date}T12:00:00`);
  d.setDate(d.getDate() - daysAgo);
  return d.toISOString().slice(0, 10);
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
      // Inbox, not /me/messages: the latter spans Archive and Deleted Items too,
      // reporting ~1450 unread where the inbox holds 25.
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
