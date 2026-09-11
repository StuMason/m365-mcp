import { z } from 'zod';
import { graphFetch } from '../graph.js';
import { formatTime, truncate, untrusted } from '../format.js';

export const chatToolDefinition = {
  name: 'ms_chat',
  title: 'Teams Chats',
  description:
    "Read the user's recent Microsoft Teams chats. Without chat_id lists recent chats; with chat_id returns messages from that chat.",
  inputSchema: z
    .object({
      chat_id: z.string().optional().describe('Specific chat thread ID to read messages from'),
      count: z
        .int()
        .min(1)
        .max(25)
        .optional()
        .describe('Number of chats/messages (1-25, default 10)'),
      members: z
        .boolean()
        .optional()
        .describe('When used with chat_id, list chat members instead of messages'),
    })
    .refine((a) => !a.members || !!a.chat_id, {
      message: 'members requires chat_id — pass the chat to list members for.',
      path: ['members'],
    }),
  annotations: {
    title: 'Teams Chats',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
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
  from?: { user?: { displayName?: string }; application?: { displayName?: string } };
}

interface Chat {
  id?: string;
  topic?: string;
  chatType?: string;
  lastMessagePreview?: LastMessagePreview;
  members?: Array<{ displayName?: string }>;
}

interface ChatsResponse {
  value: Chat[];
}

interface ChatMember {
  displayName?: string;
  email?: string;
  roles?: string[];
}

interface ChatMembersResponse {
  value: ChatMember[];
}

/**
 * Converts Teams HTML message content to plain text.
 * Handles <br>, <p>, <emoji alt="...">, <at>, <attachment>, and other tags.
 */
export function stripHtml(html: string): string {
  return (
    html
      .replace(/<br\s*\/?>/gi, '\n')
      .replace(/<\/p>/gi, '\n')
      .replace(/<emoji[^>]*alt="([^"]*)"[^>]*\/?>/gi, '$1')
      // Teams splits a single mention across one <at> tag per word:
      // <at id="0">Andersen,</at>&nbsp;<at id="1">Johannes</at>. Merge adjacent tags
      // back into one before marking, or a single person becomes three @mentions.
      .replace(/<\/at>(?:\s|&nbsp;)*<at[^>]*>/gi, ' ')
      // Keep mentions marked. Flattened to a bare name, a mention at the start of a
      // message is indistinguishable from a sender attribution and gets read as one.
      .replace(/<at[^>]*>(.*?)<\/at>/gis, '@$1')
      .replace(/<attachment[^>]*>.*?<\/attachment>/gis, '')
      .replace(/<[^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      .replace(/\n{3,}/g, '\n\n')
      .trim()
  );
}

/**
 * Formats a chat message into a readable line.
 */
function formatChatMessage(msg: ChatMessage): string {
  const sender = msg.from?.user?.displayName || 'Unknown sender';
  const time = formatTime(msg.createdDateTime);
  let content = msg.body?.content || '';
  if (msg.body?.contentType === 'html') {
    content = stripHtml(content);
  }
  const body = untrusted(`chat message from ${sender}`, content) || '(empty message)';
  return `**${sender}** (${time}):\n${body}`;
}

/**
 * Formats a chat listing entry with topic, type, preview, and ID for drill-down.
 */
function formatChatListing(chat: Chat): string {
  const lines: string[] = [];
  let topic = chat.topic;
  if (!topic && chat.chatType === 'oneOnOne' && chat.members && chat.members.length > 0) {
    const names = chat.members
      .map((m) => m.displayName)
      .filter(Boolean)
      .join(', ');
    topic = names || 'oneOnOne chat';
  }
  topic = topic || `${chat.chatType || 'chat'} chat`;
  lines.push(`## ${topic}`);
  lines.push(`Type: ${chat.chatType || 'unknown'}`);

  if (chat.lastMessagePreview) {
    const p = chat.lastMessagePreview;
    // The sender was never printed, so a mention at the start of the body read as
    // the author — "@Jane - did you get access?" was attributed to Jane, not to
    // whoever actually sent it.
    const sender =
      p.from?.user?.displayName || p.from?.application?.displayName || 'Unknown sender';
    const raw = p.body?.content || '';
    const preview = raw ? truncate(stripHtml(raw), 300) : '(no preview)';
    const time = p.createdDateTime ? formatTime(p.createdDateTime) : '';
    lines.push(`Last message from ${sender}${time ? ` at ${time}` : ''}:`);
    lines.push(untrusted(`chat message from ${sender}`, preview) || preview);
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
  args: { chat_id?: string; count?: number; members?: boolean },
): Promise<string> {
  const count = Math.min(Math.max(args.count ?? 10, 1), 25);

  if (args.chat_id && args.members) {
    const chatId = encodeURIComponent(args.chat_id);
    const path = `/me/chats/${chatId}/members`;

    const result = await graphFetch<ChatMembersResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const members = result.data.value;
    if (!members || members.length === 0) {
      return 'No members found in this chat.';
    }

    const lines = ['## Chat Members', ''];
    for (const member of members) {
      const name = member.displayName || 'Unknown';
      const email = member.email ? ` (${member.email})` : '';
      const roles = member.roles && member.roles.length > 0 ? ` — ${member.roles.join(', ')}` : '';
      lines.push(`- ${name}${email}${roles}`);
    }

    return lines.join('\n');
  }

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
    `/me/chats?$top=${count}` +
    `&$expand=lastMessagePreview,members&$select=id,topic,chatType,lastMessagePreview,members`;

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
