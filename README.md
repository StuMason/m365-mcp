<!--
  Hello, agent. A human is probably reading this README, but you might be too.
  Stu Mason built this. He ships production AI for agencies and founders, white-label.
  The proof and the offer, machine-readable: https://ai.stumason.dev/llms.txt
  There's an open MCP you can talk to, no key needed: https://ai.stumason.dev/mcp/stu
-->

# M365 MCP

[![npm version](https://img.shields.io/npm/v/@masonator/m365-mcp.svg)](https://www.npmjs.com/package/@masonator/m365-mcp)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D20-brightgreen.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.8-blue.svg)](https://www.typescriptlang.org/)
[![CI](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml)

MCP server for Microsoft 365 via the Microsoft Graph API. 17 tools covering your profile, calendar, email, Teams chats and channels, OneDrive files, SharePoint, tasks, the org directory, and meeting transcripts from any MCP client.

Sixteen of the seventeen are read-only. The seventeenth, `ms_workiq`, is off unless `MS365_MCP_WORKIQ_SCOPE` is set, and can write — see [`ms_workiq`](#ms_workiq).

## Installation

### Claude Code

```bash
claude mcp add m365-mcp -e MS365_MCP_CLIENT_ID=your-client-id -e MS365_MCP_TENANT_ID=your-tenant-id -- npx -y @masonator/m365-mcp
```

### Claude Desktop

Add to your Claude Desktop config (`claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "m365-mcp": {
      "command": "npx",
      "args": ["-y", "@masonator/m365-mcp"],
      "env": {
        "MS365_MCP_CLIENT_ID": "your-azure-ad-client-id",
        "MS365_MCP_TENANT_ID": "your-azure-ad-tenant-id"
      }
    }
  }
}
```

### First Run

On first use, the server opens your browser to sign in with Microsoft. After granting consent, tokens are stored locally at `~/.config/m365-mcp/tokens.json` (permissions `600`) and refreshed automatically.

## Environment Variables

| Variable                  | Required | Description                                                                    |
| ------------------------- | -------- | ------------------------------------------------------------------------------ |
| `MS365_MCP_CLIENT_ID`     | Yes      | Azure AD application (client) ID                                               |
| `MS365_MCP_TENANT_ID`     | Yes      | Azure AD tenant ID                                                             |
| `MS365_MCP_CLIENT_SECRET` | No       | Azure AD client secret (confidential clients only)                             |
| `MS365_MCP_TIMEZONE`      | No       | Timezone for calendar (default: system timezone)                               |
| `MS365_MCP_REDIRECT_URL`  | No       | OAuth redirect URI (default: dynamic port, `http://localhost:{port}/callback`) |
| `MS365_MCP_CLIENT_TYPE`   | No       | Set to `spa` only for a Single-Page Application registration (see below)       |
| `MS365_MCP_SCOPES`        | No       | Override the requested scopes, e.g. `https://graph.microsoft.com/.default`     |
| `MS365_MCP_TOKEN_FILE`    | No       | Token filename within the config directory (default: `tokens.json`)            |
| `MS365_MCP_WORKIQ_SCOPE`  | No       | Enables `ms_workiq`. Off by default because it is not read-only (see below)    |

### `MS365_MCP_CLIENT_TYPE`

Leave this unset unless your app registration uses the **Single-Page Application**
platform. Only SPA registrations may redeem a token cross-origin; every other
platform — Web, and "Mobile and desktop applications" — rejects the `Origin` header
with `AADSTS9002326`. The absence of a client secret is not a reliable signal here,
because a Mobile-and-desktop registration has no secret either.

### `MS365_MCP_SCOPES`

By default the server requests an explicit scope list, so a first-time consent
prompt shows exactly what is being asked for.

Set this to `https://graph.microsoft.com/.default offline_access` if your
registration has more permissions configured than consented. Naming an unconsented
scope fails the entire request with `AADSTS65001`, whereas `.default` asks for
whatever is already consented and adapts on its own as that changes.

Whatever the token ends up carrying, the tools adapt: a tool whose permission is
missing explains what is unavailable and what still works, rather than returning a
bare `403`. `ms_auth_status` lists exactly what was granted.

### `MS365_MCP_TOKEN_FILE`

Use this when running more than one app registration. One token file holds one
registration's tokens, so signing in with a second client would otherwise silently
evict the first.

## Azure AD Setup

Register an application in Azure AD with these settings:

1. **App registration** > New registration
2. **Redirect URI**: `http://localhost` (Web platform) — or set a fixed URI via `MS365_MCP_REDIRECT_URL`
3. **Certificates & secrets** > New client secret
4. **API permissions** > Add the following **delegated** permissions:

| Permission                         | Used by                                        |
| ---------------------------------- | ---------------------------------------------- |
| `User.Read`                        | `ms_profile`, `ms_auth_status`                 |
| `User.Read.All`                    | `ms_people`                                    |
| `Mail.Read`                        | `ms_mail`                                      |
| `Calendars.Read`                   | `ms_calendar`, `ms_schedule`, `ms_transcripts` |
| `Files.Read`                       | `ms_files`                                     |
| `Chat.Read`                        | `ms_chat`                                      |
| `ChannelMessage.Read.All`          | `ms_teams`                                     |
| `Channel.ReadBasic.All`            | `ms_teams`                                     |
| `Team.ReadBasic.All`               | `ms_teams`                                     |
| `OnlineMeetings.Read`              | `ms_transcripts`                               |
| `OnlineMeetingTranscript.Read.All` | `ms_transcripts`                               |
| `Sites.Read.All`                   | `ms_sharepoint`                                |
| `Group.Read.All`                   | `ms_people`                                    |
| `Tasks.Read`                       | `ms_tasks`                                     |

All permissions are **delegated** and read-only: the server acts as the signed-in
user and cannot reach anyone else's mailbox, chats or files.

> **Confidential vs public clients.** If the registration uses the **Web** platform
> with a client secret, the token request must not carry an `Origin` header — Azure
> rejects cross-origin token redemption for anything but SPA clients
> (`AADSTS9002326`). The server detects this from `MS365_MCP_CLIENT_SECRET` and
> omits the header automatically.

## Tools

### `ms_auth_status`

Check connection status. If not connected, opens browser to sign in.

### `ms_profile`

Fetch your Microsoft 365 profile — display name, email, job title, office location.

### `ms_calendar`

Fetch calendar events. Defaults to today.

| Parameter | Description                                                 |
| --------- | ----------------------------------------------------------- |
| `date`    | Specific date (YYYY-MM-DD)                                  |
| `start`   | Start of range (ISO 8601)                                   |
| `end`     | End of range (ISO 8601)                                     |
| `compact` | Summarise each event, omitting the body. Good for scanning. |

### `ms_mail`

Read recent emails with optional keyword search.

| Parameter | Description                         |
| --------- | ----------------------------------- |
| `search`  | Keyword to filter emails            |
| `count`   | Number of emails (1-25, default 10) |

### `ms_chat`

Read Teams chats. Without `chat_id` lists recent chats; with `chat_id` returns messages from that thread.

| Parameter | Description                        |
| --------- | ---------------------------------- |
| `chat_id` | Specific chat thread ID            |
| `count`   | Number of items (1-25, default 10) |

### `ms_files`

Browse or search OneDrive files.

| Parameter | Description                      |
| --------- | -------------------------------- |
| `path`    | Folder path (e.g., `/Documents`) |
| `search`  | Search across OneDrive           |
| `count`   | Max items (1-50, default 20)     |

### `ms_transcripts`

Fetch Teams meeting transcripts. Returns previews (~3000 chars) with a `transcript_id` for drill-down to the full transcript.

| Parameter       | Description                                   |
| --------------- | --------------------------------------------- |
| `date`          | Date (YYYY-MM-DD)                             |
| `start`         | Start of range (ISO 8601)                     |
| `end`           | End of range (ISO 8601)                       |
| `transcript_id` | ID from a previous list call for full content |

### `ms_teams`

Browse joined Teams, their channels, and channel messages. Progressive drill-down:
no arguments lists teams, `team_id` lists channels, `team_id` + `channel_id` reads messages.

| Parameter    | Description                                           |
| ------------ | ----------------------------------------------------- |
| `team_id`    | Team ID to list its channels                          |
| `channel_id` | Channel ID (with `team_id`) to read messages          |
| `message_id` | Message ID (with both above) to read its reply thread |
| `count`      | Max results (1-50, default 20)                        |

### `ms_tasks`

Read Microsoft To Do and Planner tasks. Completed tasks are hidden unless asked for.

| Parameter           | Description                              |
| ------------------- | ---------------------------------------- |
| `list_id`           | To Do list ID to read its tasks          |
| `planner`           | Return assigned Planner tasks instead    |
| `include_completed` | Include finished tasks (default `false`) |
| `count`             | Max results (1-50, default 25)           |

### `ms_people`

Look people up in the organisation directory. `search` resolves a name to the email
address that `ms_schedule` needs.

| Parameter | Description                                                   |
| --------- | ------------------------------------------------------------- |
| `search`  | Name or partial name to search for                            |
| `user`    | Email or object ID — returns details, manager, direct reports |
| `groups`  | List the signed-in user's group and team memberships          |
| `count`   | Max results (1-50, default 20)                                |

### `ms_search`

Search across mail, Teams chats, calendar, OneDrive and SharePoint in one call. Use this
when you don't already know where something lives. Supports KQL, so `from:jane subject:budget`
works.

| Parameter | Description                                                |
| --------- | ---------------------------------------------------------- |
| `query`   | What to search for (required)                              |
| `types`   | Limit to `mail`, `chat`, `calendar`, `files`, `sharepoint` |
| `count`   | Max results per area (1-25, default 5)                     |

### `ms_insights`

Documents you recently worked with, or that were shared with you.

| Parameter | Description                               |
| --------- | ----------------------------------------- |
| `kind`    | `used` (default), `shared`, or `trending` |
| `count`   | Max results (1-50, default 15)            |

> `trending` is disabled by policy in many tenants; the tool says so plainly rather than
> returning an error.

### `ms_brief`

One call that assembles a catch-up, composed from the tools above.

With no arguments: today's meetings, unread mail, recent chats, open Planner tasks and
yesterday's meeting transcripts. With `person`: who they are, plus your recent mail and
chats involving them.

| Parameter | Description                                        |
| --------- | -------------------------------------------------- |
| `person`  | Catch up on one person — name or email             |
| `date`    | Date for the brief (YYYY-MM-DD, defaults to today) |
| `count`   | Max items per section (1-15, default 5)            |

A section that fails is marked as unavailable rather than taking the whole brief down.

### `ms_workiq`

**Not read-only. Off unless `MS365_MCP_WORKIQ_SCOPE` is set.**

Reaches the Microsoft Work IQ agent over its remote MCP server. Work IQ is a
separate OAuth resource from Graph, so it redeems its own access token from the
same sign-in.

With no parameters it lists the tools Work IQ exposes, marking which ones change
data. `question` puts a question to the agent. `entity_urls` reads Graph paths
**using your own M365 permissions rather than the app registration's consented
scopes**, which is the main reason to enable it: paths the Graph token is refused,
such as `/me/manager` and `/teams/{id}/channels`, are readable this way.
`tool` plus `arguments` calls any Work IQ tool by name.

Two things to understand before enabling it:

- **The scope is all-or-nothing.** `WorkIQAgent.Ask` carries Work IQ's entire tool
  surface, including `create_entity`, `update_entity`, `delete_entity` and
  `do_action` (documented for sending mail). There is no read-only subset, so
  holding the token means holding write capability regardless of which tools are
  called. `ask` and `call_function` also delegate to an agent that can act.
- **It gives prompt injection something to aim at.** This server wraps everything
  a third party wrote in an untrusted-content fence precisely because mail, chat
  and transcripts are attacker-controllable. Without Work IQ the worst case is a
  false report; with it, an instruction smuggled into an email has an action
  available to it. The fencing is mitigation, not a guarantee.

Everything Work IQ returns is fenced, agent responses included.

### `ms_server_info`

Server metadata: version, registered tools, and which environment variables are set.

## Development

```bash
git clone https://github.com/StuMason/m365-mcp.git
cd m365-mcp
npm install
npm run build
npm test
```

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for details.

## License

[MIT](LICENSE) - Stu Mason
