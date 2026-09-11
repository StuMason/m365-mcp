import { z } from 'zod';
import { formatTime, untrusted, echo } from '../format.js';
import { graphFetch } from '../graph.js';

export const mailToolDefinition = {
  name: 'ms_mail',
  title: 'Mail',
  description:
    "Read the user's recent emails from Microsoft 365. " +
    'Without message_id: lists emails with preview text. ' +
    'With message_id: returns the full email body. ' +
    'Use folders: true to list mail folders, folder to read from a specific folder, ' +
    'attachments with message_id to list attachments, or filter for quick filters.',
  inputSchema: z
    .object({
      search: z.string().optional().describe('Search keyword to filter emails (KQL)'),
      count: z
        .int()
        .min(1)
        .max(25)
        .optional()
        .describe('Number of emails to return (1-25, default 10)'),
      message_id: z.string().optional().describe('Email message ID for full body drill-down'),
      folder: z
        .string()
        .optional()
        .describe(
          'Folder name or ID (e.g. "Inbox", "Sent Items"). With filter, defaults to Inbox — pass "all" to search every folder including Deleted Items.',
        ),
      folders: z.boolean().optional().describe('List all mail folders with unread counts'),
      attachments: z
        .boolean()
        .optional()
        .describe('When used with message_id, list attachments instead of body'),
      filter: z
        .string()
        .optional()
        .describe('Filter shortcut: "unread", "flagged", "attachments", "important"'),
    })
    .refine((a) => !a.attachments || !!a.message_id, {
      message: 'attachments requires message_id — pass the message to list attachments for.',
      path: ['attachments'],
    }),
  annotations: {
    title: 'Mail',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface EmailAddress {
  name?: string;
  address?: string;
}

interface MailMessage {
  id?: string;
  subject?: string;
  from?: { emailAddress?: EmailAddress };
  receivedDateTime?: string;
  bodyPreview?: string;
  isRead?: boolean;
  importance?: string;
  hasAttachments?: boolean;
}

interface MailResponse {
  value: MailMessage[];
  '@odata.count'?: number;
}

/**
 * Prefixes a listing with "showing N of M" when the collection is larger than the
 * page. Without it, "2 unread" reads as the whole inbox rather than the first two.
 */
function withTotal(messages: MailMessage[], total: number | undefined, scope: string): string {
  const body = messages.map(formatMessage).join('\n\n');
  // /me/messages spans every folder, not just the inbox — 1450 unread across the
  // mailbox versus 25 in the inbox. Naming the scope keeps the number honest.
  const head =
    total !== undefined && total > messages.length
      ? `Showing ${messages.length} of ${total} (${scope})`
      : `${messages.length} message${messages.length === 1 ? '' : 's'} (${scope})`;
  return `${head}\n\n${body}`;
}

interface MailMessageFull {
  subject?: string;
  from?: { emailAddress?: EmailAddress };
  receivedDateTime?: string;
  body?: { contentType?: string; content?: string };
  isRead?: boolean;
  importance?: string;
  toRecipients?: Array<{ emailAddress?: EmailAddress }>;
  ccRecipients?: Array<{ emailAddress?: EmailAddress }>;
}

interface MailFolder {
  id?: string;
  displayName?: string;
  unreadItemCount?: number;
  totalItemCount?: number;
}

interface MailFoldersResponse {
  value: MailFolder[];
}

interface Attachment {
  name?: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
}

interface AttachmentsResponse {
  value: Attachment[];
}

const FILTER_MAP: Record<string, string> = {
  unread: 'isRead eq false',
  flagged: "flag/flagStatus eq 'flagged'",
  attachments: 'hasAttachments eq true',
  important: "importance eq 'high'",
};

/**
 * Strips HTML tags and decodes common entities for email body content.
 */
function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<style[^>]*>.*?<\/style>/gis, '')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Formats a date string into a human-readable format.
 */
function formatDate(dateStr: string | undefined): string {
  return dateStr ? formatTime(dateStr) : 'N/A';
}

/**
 * Formats a recipient address as "Name <email>".
 */
function formatAddress(addr?: EmailAddress): string {
  if (!addr) return 'unknown';
  const name = addr.name || addr.address || 'unknown';
  const email = addr.address || 'unknown';
  return name !== email ? `${name} <${email}>` : email;
}

/**
 * Formats a file size in bytes to a human-readable string.
 */
function formatAttachmentSize(bytes: number | undefined): string {
  if (bytes === undefined || bytes === null) return 'unknown size';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/**
 * Formats a single email message into a readable multi-line block.
 */
function formatMessage(msg: MailMessage): string {
  const lines: string[] = [];

  lines.push(`## ${msg.subject || 'No Subject'}`);

  const fromName = msg.from?.emailAddress?.name || 'Unknown';
  const fromAddr = msg.from?.emailAddress?.address || 'unknown';
  lines.push(`From: ${fromName} <${fromAddr}>`);
  lines.push(`Date: ${formatDate(msg.receivedDateTime)}`);
  lines.push(`Importance: ${msg.importance || 'normal'} | Read: ${msg.isRead ? 'Yes' : 'No'}`);

  if (msg.bodyPreview) {
    // Graph truncates bodyPreview itself, mid-word and unmarked — "Fixed PDF
    // download detectio" reads as the end of the message rather than a cut.
    // Say it is a preview, and point at the drill-down that returns the whole body.
    const preview = msg.bodyPreview.trim();
    const clipped = preview.length >= 250;
    lines.push(
      untrusted(
        `email from ${fromName}`,
        clipped ? `${preview}… [preview truncated — use message_id for the full body]` : preview,
      ),
    );
  }

  if (msg.id) {
    lines.push(`Message ID: ${msg.id}`);
  }

  return lines.join('\n');
}

/**
 * Formats a full email message for drill-down view.
 */
function formatFullMessage(msg: MailMessageFull): string {
  const lines: string[] = [];

  lines.push(`# ${msg.subject || 'No Subject'}`);

  lines.push(`From: ${formatAddress(msg.from?.emailAddress)}`);

  const to = msg.toRecipients?.map((r) => formatAddress(r.emailAddress)).join(', ');
  if (to) lines.push(`To: ${to}`);

  const cc = msg.ccRecipients?.map((r) => formatAddress(r.emailAddress)).join(', ');
  if (cc) lines.push(`Cc: ${cc}`);

  lines.push(`Date: ${formatDate(msg.receivedDateTime)}`);
  lines.push(`Importance: ${msg.importance || 'normal'} | Read: ${msg.isRead ? 'Yes' : 'No'}`);
  lines.push('');

  if (msg.body?.content) {
    const content =
      msg.body.contentType === 'html' ? stripHtml(msg.body.content) : msg.body.content;
    lines.push(
      untrusted(`email from ${msg.from?.emailAddress?.name || 'unknown sender'}`, content),
    );
  } else {
    lines.push('(no body)');
  }

  return lines.join('\n');
}

/**
 * Fetches recent emails from the user's mailbox with optional search filtering
 * and returns a human-readable summary.
 *
 * Supports multiple modes:
 * - folders: true → list all mail folders with unread counts
 * - message_id + attachments → list attachments for that message
 * - message_id → drill-down to full email body
 * - folder → list messages from a specific folder
 * - filter → filtered list from main inbox
 * - search → search with ConsistencyLevel header
 * - Default → list recent messages
 */
export async function executeMail(
  token: string,
  args: {
    search?: string;
    count?: number;
    message_id?: string;
    folder?: string;
    folders?: boolean;
    attachments?: boolean;
    filter?: string;
  },
): Promise<string> {
  // 1. List folders
  if (args.folders) {
    return executeFolders(token);
  }

  // 2. List attachments for a message
  if (args.message_id && args.attachments) {
    return executeAttachments(token, args.message_id);
  }

  // 3. Drill-down to full body
  if (args.message_id) {
    return executeDrillDown(token, args.message_id);
  }

  // 4. List messages from a specific folder
  if (args.folder && !args.filter) {
    return executeFolderMessages(token, args.folder, args.count);
  }

  // 5. Filtered list
  if (args.filter) {
    return executeFiltered(token, args.filter, args.count, args.folder);
  }

  const count = Math.min(Math.max(args.count ?? 10, 1), 25);

  const select = 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance';

  let path: string;
  if (args.search) {
    // 6. Search with ConsistencyLevel header
    // $orderBy is not supported with $search — Graph returns results by relevance
    path = `/me/messages?$top=${count}&$select=${select}&$search="${encodeURIComponent(args.search)}"`;

    const result = await graphFetch<MailResponse>(path, token, {
      timezone: false,
      headers: { ConsistencyLevel: 'eventual' },
    });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const messages = result.data.value;
    if (!messages || messages.length === 0) {
      return 'No emails found.';
    }

    return withTotal(messages, result.data['@odata.count'], 'search results, all folders');
  }

  // 7. Default — list recent messages
  path = `/me/messages?$top=${count}&$count=true&$orderby=receivedDateTime desc&$select=${select}`;

  const result = await graphFetch<MailResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const messages = result.data.value;
  if (!messages || messages.length === 0) {
    return 'No emails found.';
  }

  return withTotal(messages, result.data['@odata.count'], 'all folders');
}

/**
 * Lists all mail folders with unread and total item counts.
 */
async function executeFolders(token: string): Promise<string> {
  const path = '/me/mailFolders?$top=50&$select=id,displayName,unreadItemCount,totalItemCount';
  const result = await graphFetch<MailFoldersResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const folders = result.data.value;
  if (!folders || folders.length === 0) {
    return 'No folders found.';
  }

  return folders
    .map((f) => {
      const name = f.displayName || 'Unknown Folder';
      const unread = f.unreadItemCount ?? 0;
      const total = f.totalItemCount ?? 0;
      return `## ${name}\n${unread} unread / ${total} total\nFolder ID: ${f.id || 'N/A'}`;
    })
    .join('\n\n');
}

/**
 * Resolves a folder name to its ID via Graph API lookup.
 * If the value already looks like a Graph ID (contains '='), returns it as-is.
 */
async function resolveFolderId(
  token: string,
  folderNameOrId: string,
): Promise<{ ok: true; id: string } | { ok: false; error: string }> {
  // Graph folder IDs are long base64 strings containing '=' padding
  if (folderNameOrId.includes('=')) {
    return { ok: true, id: folderNameOrId };
  }

  const path = `/me/mailFolders?$filter=displayName eq '${encodeURIComponent(folderNameOrId)}'&$select=id,displayName&$top=1`;
  const result = await graphFetch<MailFoldersResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return { ok: false, error: result.error.message };
  }

  const folders = result.data.value;
  if (!folders || folders.length === 0) {
    return { ok: false, error: `Folder "${echo(folderNameOrId)}" not found.` };
  }

  return { ok: true, id: folders[0].id || folderNameOrId };
}

/**
 * Lists messages from a specific mail folder. Accepts folder name or ID.
 */
async function executeFolderMessages(
  token: string,
  folder: string,
  countArg?: number,
): Promise<string> {
  const resolved = await resolveFolderId(token, folder);
  if (!resolved.ok) {
    return `Error: ${resolved.error}`;
  }

  const count = Math.min(Math.max(countArg ?? 10, 1), 25);
  const select = 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance,hasAttachments';
  const path = `/me/mailFolders/${encodeURIComponent(resolved.id)}/messages?$top=${count}&$count=true&$orderby=receivedDateTime desc&$select=${select}`;

  const result = await graphFetch<MailResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const messages = result.data.value;
  if (!messages || messages.length === 0) {
    return 'No emails found.';
  }

  return withTotal(messages, result.data['@odata.count'], folder);
}

/**
 * Lists attachments for a specific email message.
 */
async function executeAttachments(token: string, messageId: string): Promise<string> {
  const path = `/me/messages/${encodeURIComponent(messageId)}/attachments?$select=name,contentType,size,isInline`;

  const result = await graphFetch<AttachmentsResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const attachments = result.data.value;
  if (!attachments || attachments.length === 0) {
    return 'No attachments found.';
  }

  return attachments
    .map((a) => {
      const name = a.name || 'Unknown';
      const type = a.contentType || 'unknown type';
      const size = formatAttachmentSize(a.size);
      const inline = a.isInline ? ' (inline)' : '';
      return `- ${name} — ${type}, ${size}${inline}`;
    })
    .join('\n');
}

/**
 * Lists messages matching a filter shortcut.
 */
async function executeFiltered(
  token: string,
  filter: string,
  countArg?: number,
  folder?: string,
): Promise<string> {
  const count = Math.min(Math.max(countArg ?? 10, 1), 25);
  const select = 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance';
  const filterExpr = FILTER_MAP[filter];

  if (!filterExpr) {
    return `Error: Unknown filter "${filter}". Valid filters: ${Object.keys(FILTER_MAP).join(', ')}`;
  }

  // Default a filter to the Inbox. /me/messages spans Deleted Items and Junk, so
  // "unread" across everything counted 1450 where the inbox held 25 — a number
  // nobody wants, and one that made ms_mail and ms_brief disagree. Pass
  // folder="all" to opt back into every folder.
  const target = folder ?? 'Inbox';
  let base = '/me/messages';
  let scope = 'all folders';
  if (target !== 'all') {
    const resolved = await resolveFolderId(token, target);
    if (!resolved.ok) {
      return `Error: ${resolved.error}`;
    }
    base = `/me/mailFolders/${encodeURIComponent(resolved.id)}/messages`;
    scope = target;
  }

  const path = `${base}?$top=${count}&$count=true&$select=${select}&$filter=${filterExpr}`;

  const result = await graphFetch<MailResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const messages = result.data.value;
  if (!messages || messages.length === 0) {
    return 'No emails found.';
  }

  return withTotal(messages, result.data['@odata.count'], `${filter} in ${scope}`);
}

/**
 * Fetches a single email by ID and returns its full body content.
 */
async function executeDrillDown(token: string, messageId: string): Promise<string> {
  const select = 'subject,from,receivedDateTime,body,isRead,importance,toRecipients,ccRecipients';
  const path = `/me/messages/${encodeURIComponent(messageId)}?$select=${select}`;

  const result = await graphFetch<MailMessageFull>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  return formatFullMessage(result.data);
}
