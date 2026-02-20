import { graphFetch } from '../graph.js';

export const chatToolDefinition = {
  name: 'ms_chat',
  description:
    "Read the user's recent Microsoft Teams chats. Without chat_id lists recent chats; with chat_id returns messages from that chat.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      chat_id: { type: 'string', description: 'Specific chat thread ID to read messages from' },
      count: { type: 'integer', description: 'Number of chats/messages (1-25, default 10)' },
    },
  },
};

interface ChatMessageBody {
  content?: string;
  contentType?: string;
}

interface ChatMessage {
  from?: { user?: { displayName?: string } };
  createdDateTime?: string;
  body?: ChatMessageBody;
}

interface ChatMessagesResponse {
  value: ChatMessage[];
}

interface LastMessagePreview {
  body?: { content?: string };
  createdDateTime?: string;
}

interface Chat {
  id?: string;
  topic?: string;
  chatType?: string;
  lastMessagePreview?: LastMessagePreview;
}

interface ChatsResponse {
  value: Chat[];
}

/**
 * Converts Teams HTML message content to plain text.
 * Handles <br>, <p>, <emoji alt="...">, <at>, <attachment>, and other tags.
 */
export function stripHtml(html: string): string {
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<emoji[^>]*alt="([^"]*)"[^>]*\/?>/gi, '$1')
    .replace(/<attachment[^>]*>.*?<\/attachment>/gis, '')
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
 * Formats a chat message into a readable line.
 */
function formatChatMessage(msg: ChatMessage): string {
  const sender = msg.from?.user?.displayName || 'Unknown';
  const time = msg.createdDateTime ? new Date(msg.createdDateTime).toLocaleString() : 'N/A';
  let content = msg.body?.content || '';
  if (msg.body?.contentType === 'html') {
    content = stripHtml(content);
  }
  return `**${sender}** (${time}):\n${content || '(empty message)'}`;
}

/**
 * Formats a chat listing entry with topic, type, preview, and ID for drill-down.
 */
function formatChatListing(chat: Chat): string {
  const lines: string[] = [];
  const topic = chat.topic || `${chat.chatType || 'chat'} chat`;
  lines.push(`## ${topic}`);
  lines.push(`Type: ${chat.chatType || 'unknown'}`);

  if (chat.lastMessagePreview) {
    const rawPreview = chat.lastMessagePreview.body?.content || '(no preview)';
    const preview = rawPreview === '(no preview)' ? rawPreview : stripHtml(rawPreview);
    const time = chat.lastMessagePreview.createdDateTime
      ? new Date(chat.lastMessagePreview.createdDateTime).toLocaleString()
      : '';
    lines.push(`Last message${time ? ` (${time})` : ''}: ${preview}`);
  }

  lines.push(`Chat ID: ${chat.id || 'N/A'}`);
  return lines.join('\n');
}

/**
 * Fetches Teams chats or messages from a specific chat thread and returns
 * a human-readable summary.
 */
export async function executeChat(
  token: string,
  args: { chat_id?: string; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count ?? 10, 1), 25);

  if (args.chat_id) {
    const chatId = encodeURIComponent(args.chat_id);
    const path = `/me/chats/${chatId}/messages?$top=${count}&$orderby=createdDateTime desc`;

    const result = await graphFetch<ChatMessagesResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const messages = result.data.value;
    if (!messages || messages.length === 0) {
      return 'No messages found in this chat.';
    }

    return messages.map(formatChatMessage).join('\n\n');
  }

  const path =
    `/me/chats?$top=${count}&$orderby=lastMessagePreview/createdDateTime desc` +
    `&$expand=lastMessagePreview&$select=id,topic,chatType,lastMessagePreview`;

  const result = await graphFetch<ChatsResponse>(path, token, { timezone: false });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const chats = result.data.value;
  if (!chats || chats.length === 0) {
    return 'No Teams chats found.';
  }

  return chats.map(formatChatListing).join('\n\n');
}
