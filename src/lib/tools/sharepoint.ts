import { z } from 'zod';
import { formatTime } from '../format.js';
import { graphFetch } from '../graph.js';

export const sharepointToolDefinition = {
  name: 'ms_sharepoint',
  title: 'SharePoint',
  description:
    'Search SharePoint sites, list site lists, or browse list items. Without parameters, searches all accessible sites. Provide site_id to see its lists, or site_id + list_id to browse items.',
  inputSchema: z
    .object({
      search: z
        .string()
        .optional()
        .describe("Search query for finding sites (default '*' for all sites)"),
      site_id: z
        .string()
        .optional()
        .describe('Site ID to list its lists, or combined with list_id to browse items'),
      list_id: z
        .string()
        .optional()
        .describe('List ID (requires site_id) to browse list items with expanded fields'),
      count: z.int().min(1).max(50).optional().describe('Max results to return (1-50, default 10)'),
    })
    .refine((a) => !a.list_id || !!a.site_id, {
      // Without this, list_id alone silently fell through to "search all sites" and
      // returned a site list that looks like a successful drill-down.
      message: 'list_id requires site_id — pass the site the list belongs to.',
      path: ['list_id'],
    }),
  annotations: {
    title: 'SharePoint',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface SharePointSite {
  displayName?: string;
  webUrl?: string;
  description?: string;
  id?: string;
}

interface SitesResponse {
  value: SharePointSite[];
}

interface SiteList {
  id?: string;
  displayName?: string;
  name?: string;
  description?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  list?: { template?: string; hidden?: boolean };
}

interface SiteListsResponse {
  value: SiteList[];
}

interface ListItem {
  id?: string;
  fields?: Record<string, unknown>;
}

interface ListItemsResponse {
  value: ListItem[];
}

/**
 * Formats a SharePoint site into readable text.
 */
function formatSite(site: SharePointSite): string {
  const lines: string[] = [];
  lines.push(`## ${site.displayName || 'Unnamed Site'}`);
  if (site.description) {
    lines.push(site.description);
  }
  if (site.id) {
    lines.push(`ID: ${site.id}`);
  }
  if (site.webUrl) {
    lines.push(`URL: ${site.webUrl}`);
  }
  return lines.join('\n');
}

/**
 * Formats a site list into readable text.
 */
function formatList(list: SiteList): string {
  const lines: string[] = [];
  lines.push(`## ${list.displayName || 'Unnamed List'}`);
  if (list.id) {
    lines.push(`List ID: ${list.id}`);
  }
  if (list.name) {
    lines.push(`Name: ${list.name}`);
  }
  if (list.description) {
    lines.push(`Description: ${list.description}`);
  }
  if (list.list?.template) {
    lines.push(`Template: ${list.list.template}`);
  }
  if (list.lastModifiedDateTime) {
    lines.push(`Modified: ${formatTime(list.lastModifiedDateTime)}`);
  }
  if (list.webUrl) {
    lines.push(`URL: ${list.webUrl}`);
  }
  return lines.join('\n');
}

const HIDDEN_FIELDS = new Set([
  'AuthorLookupId',
  'EditorLookupId',
  'ContentType',
  'ItemChildCount',
  'FolderChildCount',
  'Attachments',
  'Edit',
  'LinkTitleNoMenu',
  'LinkTitle',
  'DocIcon',
  'FileSystemObjectType',
  'ServerRedirectedEmbedUri',
  'ServerRedirectedEmbedUrl',
  'ComplianceAssetId',
  'OData__ColorTag',
  'OData__UIVersionString',
  'GUID',
  'id',
]);

/**
 * Formats a list item into readable text by iterating over its fields.
 * Filters out internal/boilerplate fields.
 */
function formatItem(item: ListItem): string {
  const lines: string[] = [];
  if (item.id) {
    lines.push(`## Item ${item.id}`);
  }
  if (item.fields) {
    for (const [key, value] of Object.entries(item.fields)) {
      if (key.startsWith('@odata') || key.startsWith('_')) continue;
      if (HIDDEN_FIELDS.has(key)) continue;
      const display =
        value !== null && typeof value === 'object' ? JSON.stringify(value) : String(value ?? '');
      lines.push(`${key}: ${display}`);
    }
  }
  return lines.join('\n');
}

/**
 * Searches SharePoint sites, lists site lists, or browses list items
 * depending on which parameters are provided.
 */
export async function executeSharepoint(
  token: string,
  args: { search?: string; site_id?: string; list_id?: string; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count || 10, 1), 50);

  // Mode 1: List items from a specific list
  if (args.site_id && args.list_id) {
    const path = `/sites/${encodeURIComponent(args.site_id)}/lists/${encodeURIComponent(args.list_id)}/items?$expand=fields&$top=${count}`;
    const result = await graphFetch<ListItemsResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const items = result.data.value;
    if (!items || items.length === 0) {
      return 'No items found in this list.';
    }

    return items.map(formatItem).join('\n\n');
  }

  // Mode 2: List lists for a specific site
  if (args.site_id) {
    const path = `/sites/${encodeURIComponent(args.site_id)}/lists?$top=${count}&$select=id,displayName,name,description,webUrl,lastModifiedDateTime,list&$filter=list/hidden eq false`;
    const result = await graphFetch<SiteListsResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const lists = result.data.value;
    if (!lists || lists.length === 0) {
      return 'No lists found for this site.';
    }

    return lists.map(formatList).join('\n\n');
  }

  // Mode 3: Search sites (default)
  const query = args.search || '*';
  const path = `/sites?search=${encodeURIComponent(query)}&$top=${count}`;
  const result = await graphFetch<SitesResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const sites = result.data.value;
  if (!sites || sites.length === 0) {
    return 'No sites found.';
  }

  return sites.map(formatSite).join('\n\n');
}
