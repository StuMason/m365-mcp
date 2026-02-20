# M365 MCP

[![npm version](https://img.shields.io/npm/v/@masonator/m365-mcp.svg)](https://www.npmjs.com/package/@masonator/m365-mcp)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18-brightgreen.svg)](https://nodejs.org)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.8-blue.svg)](https://www.typescriptlang.org/)
[![CI](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml)

MCP server for Microsoft 365 via the Microsoft Graph API. Read-only access to your profile, calendar, email, Teams chats, OneDrive files, and meeting transcripts from any MCP client.

## Installation

### Claude Code

```bash
claude mcp add m365-mcp -e MS365_MCP_CLIENT_ID=your-client-id -e MS365_MCP_CLIENT_SECRET=your-secret -e MS365_MCP_TENANT_ID=your-tenant-id -- npx -y @masonator/m365-mcp
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
        "MS365_MCP_CLIENT_SECRET": "your-azure-ad-client-secret",
        "MS365_MCP_TENANT_ID": "your-azure-ad-tenant-id"
      }
    }
  }
}
```

### First Run

On first use, the server opens your browser to sign in with Microsoft. After granting consent, tokens are stored locally at `~/.config/m365-mcp/tokens.json` (permissions `600`) and refreshed automatically.

## Environment Variables

| Variable                  | Required | Description                                      |
| ------------------------- | -------- | ------------------------------------------------ |
| `MS365_MCP_CLIENT_ID`     | Yes      | Azure AD application (client) ID                 |
| `MS365_MCP_CLIENT_SECRET` | Yes      | Azure AD client secret                           |
| `MS365_MCP_TENANT_ID`     | Yes      | Azure AD tenant ID                               |
| `MS365_MCP_TIMEZONE`      | No       | Timezone for calendar (default: system timezone) |

## Azure AD Setup

Register an application in Azure AD with these settings:

1. **App registration** > New registration
2. **Redirect URI**: `http://localhost:19284/auth/callback` (Web platform)
3. **Certificates & secrets** > New client secret
4. **API permissions** > Add the following **delegated** permissions:
   - `User.Read`
   - `Calendars.Read`
   - `Mail.Read`
   - `Chat.Read`
   - `Files.Read`
   - `OnlineMeetingTranscript.Read.All`
   - `Sites.Read.All`

## Tools

### `ms_auth_status`

Check connection status. If not connected, opens browser to sign in.

### `ms_profile`

Fetch your Microsoft 365 profile — display name, email, job title, office location.

### `ms_calendar`

Fetch calendar events. Defaults to today.

| Parameter | Description                |
| --------- | -------------------------- |
| `date`    | Specific date (YYYY-MM-DD) |
| `start`   | Start of range (ISO 8601)  |
| `end`     | End of range (ISO 8601)    |

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
