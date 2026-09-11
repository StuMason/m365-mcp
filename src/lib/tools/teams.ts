import { graphFetch } from '../graph.js';
import { stripHtml } from './chat.js';

export const teamsToolDefinition = {
  name: 'ms_teams',
  title: 'Teams & Channels',
  description:
    'Browse Microsoft Teams the user has joined. Without parameters lists joined teams; with team_id lists that team’s channels; with team_id + channel_id returns recent channel messages. Distinct from ms_chat, which covers private/group chats rather than team channels.',
  inputSchema: {
    type: 'object' as const,
    properties: {
      team_id: {
        type: 'string',
        description: 'Team ID to list its channels, or combined with channel_id to read messages',
      },
      channel_id: {
        type: 'string',
        description: 'Channel ID (requires team_id) to read recent messages from that channel',
      },
      message_id: {
        type: 'string',
        description:
          'Message ID (requires team_id + channel_id) to read the replies on that message thread',
      },
      count: {
        type: 'integer',
        description: 'Max results to return (1-50, default 20)',
      },
    },
  },
  annotations: {
    title: 'Teams & Channels',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface Team {
  id?: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  isArchived?: boolean;
}

interface TeamsResponse {
  value: Team[];
}

interface Channel {
  id?: string;
  displayName?: string;
  description?: string;
  webUrl?: string;
  membershipType?: string;
  createdDateTime?: string;
}

interface ChannelsResponse {
  value: Channel[];
}

interface ChannelMessage {
  id?: string;
  createdDateTime?: string;
  from?: { user?: { displayName?: string } };
  body?: { content?: string };
  subject?: string;
  replyToId?: string;
  attachments?: Array<{ name?: string }>;
}

interface ChannelMessagesResponse {
  value: ChannelMessage[];
}

/**
 * Formats a joined team into readable text.
 */
function formatTeam(team: Team): string {
  const lines: string[] = [];
  lines.push(`## ${team.displayName || 'Unnamed Team'}${team.isArchived ? ' (archived)' : ''}`);
  if (team.description) {
    lines.push(team.description);
  }
  if (team.id) {
    lines.push(`Team ID: ${team.id}`);
  }
  if (team.webUrl) {
    lines.push(`URL: ${team.webUrl}`);
  }
  return lines.join('\n');
}

/**
 * Formats a team channel into readable text.
 */
function formatChannel(channel: Channel): string {
  const lines: string[] = [];
  lines.push(`## ${channel.displayName || 'Unnamed Channel'}`);
  if (channel.description) {
    lines.push(channel.description);
  }
  if (channel.id) {
    lines.push(`Channel ID: ${channel.id}`);
  }
  if (channel.membershipType) {
    lines.push(`Type: ${channel.membershipType}`);
  }
  if (channel.webUrl) {
    lines.push(`URL: ${channel.webUrl}`);
  }
  return lines.join('\n');
}

/**
 * Formats a channel message into readable text.
 * Messages with no body (system events such as member joins) are dropped by the caller.
 */
function formatChannelMessage(message: ChannelMessage): string {
  const lines: string[] = [];
  const author = message.from?.user?.displayName || 'Unknown';
  const when = message.createdDateTime
    ? new Date(message.createdDateTime).toLocaleString()
    : 'Unknown date';
  lines.push(`## ${message.subject || author}`);
  lines.push(`From: ${author}`);
  lines.push(`Date: ${when}`);
  if (message.id) {
    lines.push(`Message ID: ${message.id}`);
  }
  const body = stripHtml(message.body?.content || '');
  if (body) {
    lines.push('');
    lines.push(body);
  }
  if (message.attachments && message.attachments.length > 0) {
    const names = message.attachments.map((a) => a.name).filter(Boolean);
    if (names.length > 0) {
      lines.push(`Attachments: ${names.join(', ')}`);
    }
  }
  return lines.join('\n');
}

/**
 * Returns true when a channel message carries no readable content — Graph returns
 * system events (member added, channel renamed) in the same collection as real posts.
 */
function isEmptyMessage(message: ChannelMessage): boolean {
  return stripHtml(message.body?.content || '').length === 0;
}

/**
 * Lists joined teams, a team's channels, or a channel's messages
 * depending on which parameters are provided.
 */
export async function executeTeams(
  token: string,
  args: { team_id?: string; channel_id?: string; message_id?: string; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count || 20, 1), 50);

  // Mode 1: replies on a single message thread
  if (args.team_id && args.channel_id && args.message_id) {
    const path =
      `/teams/${encodeURIComponent(args.team_id)}` +
      `/channels/${encodeURIComponent(args.channel_id)}` +
      `/messages/${encodeURIComponent(args.message_id)}/replies?$top=${count}`;
    const result = await graphFetch<ChannelMessagesResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const replies = (result.data.value || []).filter((m) => !isEmptyMessage(m));
    if (replies.length === 0) {
      return 'No replies on this message.';
    }

    return replies.map(formatChannelMessage).join('\n\n---\n\n');
  }

  // Mode 2: messages in a channel
  if (args.team_id && args.channel_id) {
    const path =
      `/teams/${encodeURIComponent(args.team_id)}` +
      `/channels/${encodeURIComponent(args.channel_id)}/messages?$top=${count}`;
    const result = await graphFetch<ChannelMessagesResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const messages = (result.data.value || []).filter((m) => !isEmptyMessage(m));
    if (messages.length === 0) {
      return 'No messages found in this channel.';
    }

    return messages.map(formatChannelMessage).join('\n\n---\n\n');
  }

  // Mode 3: channels in a team
  if (args.team_id) {
    const path = `/teams/${encodeURIComponent(args.team_id)}/channels`;
    const result = await graphFetch<ChannelsResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const channels = result.data.value;
    if (!channels || channels.length === 0) {
      return 'No channels found in this team.';
    }

    return channels.slice(0, count).map(formatChannel).join('\n\n');
  }

  // Mode 4: joined teams (default)
  const result = await graphFetch<TeamsResponse>('/me/joinedTeams', token, {
    timezone: false,
  });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const teams = result.data.value;
  if (!teams || teams.length === 0) {
    return 'No joined teams found.';
  }

  return teams.slice(0, count).map(formatTeam).join('\n\n');
}
