import { graphFetch } from '../graph.js';

export const sharepointToolDefinition = {
  name: 'ms_sharepoint',
  description:
    'Search SharePoint sites, list site lists, or browse list items. Without parameters, searches all accessible sites. Provide site_id to see its lists, or site_id + list_id to browse items.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      search: {
        type: 'string',
        description: "Search query for finding sites (default '*' for all sites)",
      },
      site_id: {
        type: 'string',
        description: 'Site ID to list its lists, or combined with list_id to browse items',
      },
      list_id: {
        type: 'string',
        description: 'List ID (requires site_id) to browse list items with expanded fields',
      },
      count: {
        type: 'integer',
        description: 'Max results to return (1-50, default 10)',
      },
    },
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
  displayName?: string;
  name?: string;
  description?: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  list?: { template?: string };
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
    lines.push(`Modified: ${new Date(list.lastModifiedDateTime).toLocaleString()}`);
  }
  if (list.webUrl) {
    lines.push(`URL: ${list.webUrl}`);
  }
  return lines.join('\n');
}

/**
 * Formats a list item into readable text by iterating over its fields.
 */
function formatItem(item: ListItem): string {
  const lines: string[] = [];
  if (item.id) {
    lines.push(`## Item ${item.id}`);
  }
  if (item.fields) {
    for (const [key, value] of Object.entries(item.fields)) {
      if (key.startsWith('@odata') || key.startsWith('_')) continue;
      lines.push(`${key}: ${value}`);
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
    const path = `/sites/${args.site_id}/lists/${args.list_id}/items?$expand=fields&$top=${count}`;
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
    const path = `/sites/${args.site_id}/lists?$top=${count}&$select=displayName,name,description,webUrl,lastModifiedDateTime,list`;
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
  const path = `/sites?search=${query}&$top=${count}`;
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
