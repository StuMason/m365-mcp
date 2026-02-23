import { graphFetch } from '../graph.js';

export const filesToolDefinition = {
  name: 'ms_files',
  description: "Browse or search the user's OneDrive files.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      path: { type: 'string', description: "Folder path (e.g. '/Documents')" },
      search: { type: 'string', description: 'Search across OneDrive' },
      count: { type: 'integer', description: 'Max items (1-50, default 20)' },
      item_id: {
        type: 'string',
        description: 'File/folder ID for detailed metadata and download URL',
      },
      shared: { type: 'boolean', description: 'List files shared with me' },
    },
  },
};

interface DriveItem {
  name?: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  file?: Record<string, unknown>;
  folder?: { childCount?: number };
}

interface DriveResponse {
  value: DriveItem[];
}

interface DriveItemDetail {
  name?: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  '@microsoft.graph.downloadUrl'?: string;
  file?: Record<string, unknown>;
  folder?: { childCount?: number };
}

interface SharedDriveItem {
  name?: string;
  size?: number;
  lastModifiedDateTime?: string;
  webUrl?: string;
  file?: Record<string, unknown>;
  folder?: { childCount?: number };
  remoteItem?: {
    shared?: {
      sharedBy?: { user?: { displayName?: string } };
    };
  };
}

interface SharedDriveResponse {
  value: SharedDriveItem[];
}

/**
 * Formats a byte count into a human-readable size string.
 */
export function formatFileSize(bytes: number | undefined): string {
  if (bytes === undefined || bytes === null) return 'N/A';
  if (bytes === 0) return '0 B';

  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.floor(Math.log(bytes) / Math.log(1024));
  const index = Math.min(i, units.length - 1);
  const value = bytes / Math.pow(1024, index);

  return `${value.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
}

/**
 * Formats a single drive item into a readable line.
 */
function formatItem(item: DriveItem): string {
  const icon = item.folder ? '\u{1F4C1}' : '\u{1F4C4}';
  const name = item.name || 'Unnamed';
  const size = formatFileSize(item.size);
  const modified = item.lastModifiedDateTime
    ? new Date(item.lastModifiedDateTime).toLocaleString()
    : 'N/A';
  const url = item.webUrl || '';

  const lines = [`${icon} ${name}`, `  Size: ${size}  Modified: ${modified}`];
  if (url) {
    lines.push(`  URL: ${url}`);
  }
  return lines.join('\n');
}

/**
 * Fetches detail for a single drive item by ID, including download URL if available.
 */
async function executeFileDetail(token: string, itemId: string): Promise<string> {
  const result = await graphFetch<DriveItemDetail>(`/me/drive/items/${itemId}`, token, {
    timezone: false,
  });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const item = result.data;
  const name = item.name || 'Unnamed';
  const size = formatFileSize(item.size);
  const modified = item.lastModifiedDateTime
    ? new Date(item.lastModifiedDateTime).toLocaleString()
    : 'N/A';
  const isFolder = !!item.folder;
  const type = isFolder ? 'folder' : 'file';

  const lines = [`# ${name}`, `Type: ${type}`, `Size: ${size}`, `Modified: ${modified}`];

  if (isFolder && item.folder?.childCount !== undefined) {
    lines.push(`Children: ${item.folder.childCount}`);
  }

  if (item.webUrl) {
    lines.push(`Web URL: ${item.webUrl}`);
  }

  if (item['@microsoft.graph.downloadUrl']) {
    lines.push(`Download URL: ${item['@microsoft.graph.downloadUrl']}`);
  }

  return lines.join('\n');
}

/**
 * Fetches files shared with the current user.
 */
async function executeSharedFiles(token: string, count: number): Promise<string> {
  const result = await graphFetch<SharedDriveResponse>(
    `/me/drive/sharedWithMe?$top=${count}`,
    token,
    { timezone: false },
  );

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const items = result.data.value;
  if (!items || items.length === 0) {
    return 'No shared files found.';
  }

  return items.map(formatSharedItem).join('\n\n');
}

/**
 * Formats a shared drive item into a readable block, including who shared it.
 */
function formatSharedItem(item: SharedDriveItem): string {
  const icon = item.folder ? '\u{1F4C1}' : '\u{1F4C4}';
  const name = item.name || 'Unnamed';
  const size = formatFileSize(item.size);
  const modified = item.lastModifiedDateTime
    ? new Date(item.lastModifiedDateTime).toLocaleString()
    : 'N/A';
  const url = item.webUrl || '';
  const sharedBy = item.remoteItem?.shared?.sharedBy?.user?.displayName || 'Unknown';

  const lines = [
    `${icon} ${name}`,
    `  Size: ${size}  Modified: ${modified}`,
    `  Shared by: ${sharedBy}`,
  ];
  if (url) {
    lines.push(`  URL: ${url}`);
  }
  return lines.join('\n');
}

/**
 * Fetches OneDrive files from a given path, by search query, or from root,
 * and returns a human-readable listing.
 */
export async function executeFiles(
  token: string,
  args: { path?: string; search?: string; count?: number; item_id?: string; shared?: boolean },
): Promise<string> {
  // File detail mode
  if (args.item_id) {
    return executeFileDetail(token, args.item_id);
  }

  // Shared files mode
  if (args.shared) {
    const count = Math.min(Math.max(args.count ?? 20, 1), 50);
    return executeSharedFiles(token, count);
  }

  const count = Math.min(Math.max(args.count ?? 20, 1), 50);
  const select = 'name,size,lastModifiedDateTime,webUrl,file,folder';

  let path: string;

  if (args.search) {
    path = `/me/drive/root/search(q='${encodeURIComponent(args.search)}')?$top=${count}&$select=${select}`;
  } else if (args.path && args.path !== '/') {
    path = `/me/drive/root:${encodeURIComponent(args.path)}:/children?$top=${count}&$select=${select}`;
  } else {
    path = `/me/drive/root/children?$top=${count}&$select=${select}`;
  }

  const result = await graphFetch<DriveResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const items = result.data.value;
  if (!items || items.length === 0) {
    return 'No files found.';
  }

  return items.map(formatItem).join('\n\n');
}
