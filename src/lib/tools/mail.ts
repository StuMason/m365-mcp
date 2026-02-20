import { graphFetch } from '../graph.js';

export const mailToolDefinition = {
  name: 'ms_mail',
  description:
    "Read the user's recent emails from Microsoft 365. " +
    'Without message_id: lists emails with preview text. ' +
    'With message_id: returns the full email body.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      search: { type: 'string', description: 'Search keyword to filter emails' },
      count: { type: 'integer', description: 'Number of emails to return (1-25, default 10)' },
      message_id: {
        type: 'string',
        description: 'Email message ID for full body drill-down (from a previous list call)',
      },
    },
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
}

interface MailResponse {
  value: MailMessage[];
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
  if (!dateStr) return 'N/A';
  const d = new Date(dateStr);
  return d.toLocaleString();
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
    lines.push(msg.bodyPreview);
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
    lines.push(content);
  } else {
    lines.push('(no body)');
  }

  return lines.join('\n');
}

/**
 * Fetches recent emails from the user's mailbox with optional search filtering
 * and returns a human-readable summary.
 *
 * Supports two modes:
 * - List mode: returns email summaries with previews and message IDs
 * - Drill-down mode: when message_id is provided, returns the full email body
 */
export async function executeMail(
  token: string,
  args: { search?: string; count?: number; message_id?: string },
): Promise<string> {
  if (args.message_id) {
    return executeDrillDown(token, args.message_id);
  }

  const count = Math.min(Math.max(args.count ?? 10, 1), 25);

  const select = 'id,subject,from,receivedDateTime,bodyPreview,isRead,importance';

  let path: string;
  if (args.search) {
    // $orderBy is not supported with $search — Graph returns results by relevance
    path = `/me/messages?$top=${count}&$select=${select}&$search="${encodeURIComponent(args.search)}"`;
  } else {
    path = `/me/messages?$top=${count}&$orderby=receivedDateTime desc&$select=${select}`;
  }

  const result = await graphFetch<MailResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const messages = result.data.value;
  if (!messages || messages.length === 0) {
    return 'No emails found.';
  }

  return messages.map(formatMessage).join('\n\n');
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
