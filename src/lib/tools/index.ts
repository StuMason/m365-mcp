import { authStatusToolDefinition } from './auth-status.js';
import { profileToolDefinition } from './profile.js';
import { calendarToolDefinition } from './calendar.js';
import { mailToolDefinition } from './mail.js';
import { chatToolDefinition } from './chat.js';
import { filesToolDefinition } from './files.js';
import { transcriptsToolDefinition } from './transcripts.js';
import { scheduleToolDefinition } from './schedule.js';
import { sharepointToolDefinition } from './sharepoint.js';
import { teamsToolDefinition } from './teams.js';
import { tasksToolDefinition } from './tasks.js';
import { peopleToolDefinition } from './people.js';
import { serverInfoToolDefinition } from './server-info.js';

/**
 * Every tool this server exposes, in the order they are advertised.
 *
 * Single source of truth for the roster. `index.ts` registers exactly these and
 * `ms_server_info` counts them, so the advertised list cannot drift from the
 * registered one — that drift is what shipped in 0.7.0, where three separate
 * hand-maintained lists disagreed about how many tools there were.
 *
 * Handlers are deliberately *not* here: `registerTool` infers each handler's
 * argument type from its own zod schema, and a single array of mixed handler
 * signatures would erase that inference.
 */
export const TOOL_DEFINITIONS = [
  authStatusToolDefinition,
  profileToolDefinition,
  calendarToolDefinition,
  mailToolDefinition,
  chatToolDefinition,
  filesToolDefinition,
  transcriptsToolDefinition,
  scheduleToolDefinition,
  sharepointToolDefinition,
  teamsToolDefinition,
  tasksToolDefinition,
  peopleToolDefinition,
  serverInfoToolDefinition,
];

/** Tool names in advertised order. Used by ms_server_info and the drift guard. */
export function toolNames(): string[] {
  return TOOL_DEFINITIONS.map((t) => t.name);
}
