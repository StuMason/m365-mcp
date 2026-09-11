import { z } from 'zod';
import { formatTime, untrusted, echo } from '../format.js';
import { graphPost } from '../graph.js';

export const searchToolDefinition = {
  name: 'ms_search',
  title: 'Search Everything',
  description:
    'Search across Microsoft 365 in one call — mail, Teams chats, calendar events, OneDrive files and SharePoint. Use this when you do not already know where something lives; use the per-area tools (ms_mail, ms_chat, ms_files) when you do. Supports KQL, so "from:jane subject:budget" works.',
  inputSchema: z.object({
    query: z
      .string()
      .min(1)
      .describe('What to search for. KQL is supported, e.g. \'from:jane "Q3 budget"\''),
    types: z
      .array(z.enum(['mail', 'chat', 'calendar', 'files', 'sharepoint']))
      .optional()
      .describe('Limit to these areas. Defaults to all of them.'),
    count: z.int().min(1).max(25).optional().describe('Max results per area (1-25, default 5)'),
  }),
  annotations: {
    title: 'Search Everything',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

/**
 * Graph rejects arbitrary entity-type combinations ("Invalid entity type
 * combination") and accepts only one entityRequest per call, so a search across
 * everything is three parallel calls along these lines. Verified against the
 * live API — do not merge these groups.
 *
 * `person` is deliberately absent: it needs People.Read, which is not consented,
 * and ms_people covers directory lookup properly anyway.
 */
const GROUPS: Array<{ label: string; entityTypes: string[] }> = [
  { label: 'Mail & Chat', entityTypes: ['message', 'chatMessage'] },
  { label: 'Calendar', entityTypes: ['event'] },
  { label: 'Files & SharePoint', entityTypes: ['driveItem', 'listItem', 'site'] },
];

const TYPE_MAP: Record<string, string[]> = {
  mail: ['message'],
  chat: ['chatMessage'],
  calendar: ['event'],
  files: ['driveItem'],
  sharepoint: ['site', 'listItem'],
};

interface SearchHit {
  hitId?: string;
  rank?: number;
  summary?: string;
  resource?: {
    '@odata.type'?: string;
    subject?: string;
    name?: string;
    displayName?: string;
    bodyPreview?: string;
    webUrl?: string;
    webLink?: string;
    id?: string;
    size?: number;
    createdDateTime?: string;
    receivedDateTime?: string;
    lastModifiedDateTime?: string;
    from?: { emailAddress?: { name?: string }; user?: { displayName?: string } };
    start?: { dateTime?: string };
    description?: string;
  };
}

interface SearchResponse {
  value?: Array<{
    hitsContainers?: Array<{ total?: number; hits?: SearchHit[] }>;
  }>;
}

/**
 * Strips the <c0></c0> hit-highlight markers Graph wraps around matched terms.
 */
export function stripHighlights(text: string): string {
  return (
    text
      .replace(/<\/?c\d+>/g, '')
      // <ddd/> is Graph's elision marker for the omitted middle of a snippet.
      .replace(/<ddd\s*\/?>/gi, '…')
      .replace(/\s+/g, ' ')
      .trim()
  );
}

/**
 * Turns "#microsoft.graph.driveItem" into "driveItem".
 */
function shortType(odataType?: string): string {
  return (odataType ?? '').replace(/^#?microsoft\.graph\./, '') || 'result';
}

/**
 * Formats a single search hit. Search returns a different (thinner) projection
 * than the per-area endpoints, so this deliberately shows only what is present.
 */
function formatHit(hit: SearchHit): string {
  const r = hit.resource ?? {};
  const lines: string[] = [];

  const type = shortType(r['@odata.type']);
  const who = r.from?.emailAddress?.name || r.from?.user?.displayName;

  // Chat messages have no subject, so "(untitled)" would be every result.
  const title = r.subject || r.name || r.displayName || (who ? `Message from ${who}` : `(${type})`);
  lines.push(`### ${title}`);
  lines.push(`Type: ${type}`);

  if (who) {
    lines.push(`From: ${who}`);
  }

  const when = r.receivedDateTime || r.start?.dateTime || r.lastModifiedDateTime;
  if (when) {
    lines.push(`Date: ${formatTime(when)}`);
  }

  const summary = stripHighlights(hit.summary || r.bodyPreview || r.description || '');
  if (summary) {
    const provenance = who
      ? `${type} from ${who}`
      : r.name || r.displayName
        ? `${type}: ${r.name || r.displayName}`
        : type;
    lines.push(untrusted(provenance, summary));
  }

  const url = r.webUrl || r.webLink;
  if (url) {
    lines.push(`URL: ${url}`);
  }

  return lines.join('\n');
}

/**
 * Searches across Microsoft 365 via the unified /search/query endpoint.
 * Each compatible entity-type group is a separate call, run in parallel.
 */
export async function executeSearch(
  token: string,
  args: { query: string; types?: string[]; count?: number },
): Promise<string> {
  const size = Math.min(Math.max(args.count || 5, 1), 25);

  // Work out which entity types the caller asked for, then keep only the groups
  // that still have something in them.
  const wanted = args.types?.length ? new Set(args.types.flatMap((t) => TYPE_MAP[t] ?? [])) : null;

  const groups = GROUPS.map((g) => ({
    label: g.label,
    entityTypes: wanted ? g.entityTypes.filter((e) => wanted.has(e)) : g.entityTypes,
  })).filter((g) => g.entityTypes.length > 0);

  if (groups.length === 0) {
    return 'No searchable areas selected.';
  }

  const results = await Promise.all(
    groups.map((g) =>
      graphPost<unknown, SearchResponse>(
        '/search/query',
        token,
        {
          requests: [
            { entityTypes: g.entityTypes, query: { queryString: args.query }, from: 0, size },
          ],
        },
        { timezone: false },
      ),
    ),
  );

  const sections: string[] = [];
  const errors: string[] = [];

  for (const [i, result] of results.entries()) {
    const group = groups[i]!;
    if (!result.ok) {
      errors.push(`${group.label}: ${result.error.message}`);
      continue;
    }

    const container = result.data.value?.[0]?.hitsContainers?.[0];
    const hits = container?.hits ?? [];
    if (hits.length === 0) {
      continue;
    }

    const total = container?.total ?? hits.length;
    const more = total > hits.length ? ` (showing ${hits.length} of ${total})` : '';
    sections.push(`## ${group.label}${more}\n\n${hits.map(formatHit).join('\n\n')}`);
  }

  if (sections.length === 0) {
    const nothing = `No results for "${echo(args.query)}".`;
    return errors.length > 0 ? `${nothing}\n\nErrors:\n${errors.join('\n')}` : nothing;
  }

  const body = sections.join('\n\n');
  return errors.length > 0 ? `${body}\n\nSome areas failed:\n${errors.join('\n')}` : body;
}
