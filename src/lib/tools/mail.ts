import { graphFetch } from '../graph.js';

export const mailToolDefinition = {
  name: 'ms_mail',
  description: "Read the user's recent emails from Microsoft 365. Optionally search by keyword.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      search: { type: 'string', description: 'Search keyword to filter emails' },
      count: { type: 'integer', description: 'Number of emails to return (1-25, default 10)' },
    },
  },
};

interface EmailAddress {
  name?: string;
  address?: string;
}

interface MailMessage {
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

/**
 * Formats a date string into a human-readable format.
 */
function formatDate(dateStr: string | undefined): string {
  if (!dateStr) return 'N/A';
  const d = new Date(dateStr);
  return d.toLocaleString();
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

  return lines.join('\n');
}

/**
 * Fetches recent emails from the user's mailbox with optional search filtering
 * and returns a human-readable summary.
 */
export async function executeMail(
  token: string,
  args: { search?: string; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count ?? 10, 1), 25);

  const select = 'subject,from,receivedDateTime,bodyPreview,isRead,importance';

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
