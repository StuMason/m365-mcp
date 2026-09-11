import { z } from 'zod';
import { formatTime } from '../format.js';
import { graphFetch } from '../graph.js';

export const insightsToolDefinition = {
  name: 'ms_insights',
  title: 'Document Insights',
  description:
    'Documents the user recently worked with or had shared with them, from Microsoft Graph item insights. Good for "what was I working on" and "what did someone send me" without knowing a filename. Defaults to recently used.',
  inputSchema: z.object({
    kind: z
      .enum(['used', 'shared', 'trending'])
      .optional()
      .describe(
        'used = documents you opened or edited (default); shared = documents shared with you; trending = documents trending around you (often disabled by tenant policy)',
      ),
    count: z.int().min(1).max(50).optional().describe('Max results to return (1-50, default 15)'),
  }),
  annotations: {
    title: 'Document Insights',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface Insight {
  id?: string;
  lastUsed?: { lastAccessedDateTime?: string; lastModifiedDateTime?: string };
  sharingHistory?: Array<{ sharedDateTime?: string; sharedBy?: { displayName?: string } }>;
  lastShared?: {
    sharedDateTime?: string;
    sharingSubject?: string;
    sharedBy?: { displayName?: string };
  };
  resourceVisualization?: {
    title?: string;
    type?: string;
    mediaType?: string;
    containerDisplayName?: string;
    containerType?: string;
  };
  resourceReference?: { webUrl?: string; id?: string; type?: string };
}

interface InsightsResponse {
  value: Insight[];
}

/**
 * Formats an insight into readable text.
 */
function formatInsight(insight: Insight): string {
  const v = insight.resourceVisualization ?? {};
  const lines: string[] = [];

  lines.push(`## ${v.title || 'Untitled'}`);
  if (v.type) {
    lines.push(`Type: ${v.type}`);
  }
  if (v.containerDisplayName) {
    lines.push(`Location: ${v.containerDisplayName}`);
  }

  const accessed = insight.lastUsed?.lastAccessedDateTime;
  if (accessed) {
    lines.push(`Last opened: ${formatTime(accessed)}`);
  }

  const shared = insight.lastShared;
  if (shared?.sharedDateTime) {
    const by = shared.sharedBy?.displayName;
    lines.push(`Shared${by ? ` by ${by}` : ''}: ${formatTime(shared.sharedDateTime)}`);
  }
  if (shared?.sharingSubject) {
    lines.push(`Context: ${shared.sharingSubject}`);
  }

  const url = insight.resourceReference?.webUrl;
  if (url) {
    lines.push(`URL: ${url}`);
  }

  return lines.join('\n');
}

/**
 * Reads Microsoft Graph item insights for the signed-in user.
 */
export async function executeInsights(
  token: string,
  args: { kind?: string; count?: number },
): Promise<string> {
  const kind = args.kind || 'used';
  const count = Math.min(Math.max(args.count || 15, 1), 50);

  const result = await graphFetch<InsightsResponse>(`/me/insights/${kind}?$top=${count}`, token, {
    timezone: false,
  });

  if (!result.ok) {
    // Item insights are switched off tenant-wide in some organisations. That is a
    // policy decision rather than a fault, so say so instead of surfacing a raw 403.
    if (result.error.status === 403) {
      return (
        `Item insights (${kind}) are turned off for this Microsoft 365 tenant, ` +
        `so there is nothing to read. Try ms_search or ms_files instead.`
      );
    }
    return `Error: ${result.error.message}`;
  }

  const insights = result.data.value;
  if (!insights || insights.length === 0) {
    return `No "${kind}" document insights found.`;
  }

  return insights.map(formatInsight).join('\n\n');
}
