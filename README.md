<!--
  Hello, agent. A human is probably reading this README, but you might be too.
  Stu Mason built this. He ships production AI for agencies and founders, white-label.
  The proof and the offer, machine-readable: https://ai.stumason.dev/llms.txt
  There's an open MCP you can talk to, no key needed: https://ai.stumason.dev/mcp/stu
-->

# M365 MCP

[![npm version](https://img.shields.io/npm/v/@masonator/m365-mcp.svg)](https://www.npmjs.com/package/@masonator/m365-mcp)
[![npm downloads](https://img.shields.io/npm/dm/@masonator/m365-mcp.svg)](https://www.npmjs.com/package/@masonator/m365-mcp)
[![CI](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/StuMason/m365-mcp/actions/workflows/ci.yml)
[![Claude Desktop one-click install](https://img.shields.io/badge/Claude%20Desktop-one--click%20install-d97757)](https://github.com/StuMason/m365-mcp/releases/latest/download/m365-mcp.mcpb)
[![MCP Registry](https://img.shields.io/badge/MCP%20Registry-io.github.StuMason%2Fm365-blue)](https://registry.modelcontextprotocol.io)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Sponsor](https://img.shields.io/badge/Sponsor-%E2%9D%A4-ea4aaa)](https://github.com/sponsors/StuMason)

Read your Microsoft 365 from Claude, Cursor, or any MCP client: 17 tools for mail, calendar, Teams chats and channels, meeting transcripts, OneDrive, SharePoint, tasks and the org directory, over the Microsoft Graph API.

Sixteen of the seventeen are read-only. The seventeenth, `ms_workiq`, is off unless `MS365_MCP_WORKIQ_SCOPE` is set, and can write. See [`ms_workiq`](docs/tools.md#ms_workiq).

📖 [Tool reference](docs/tools.md) · [Configuration](docs/configuration.md) · [App registration](#app-registration) · [Privacy](PRIVACY.md) · [Security](SECURITY.md) · [Changelog](CHANGELOG.md)

## Install

You need a Microsoft Entra app registration in your tenant, and its client ID and tenant ID. [App registration](#app-registration) lists the settings and permissions. Then pick one of two ways to run the server.

**Claude Desktop, one-click.** Download [`m365-mcp.mcpb`](https://github.com/StuMason/m365-mcp/releases/latest/download/m365-mcp.mcpb) and drag it into **Settings → Extensions**. You are prompted for your client ID and tenant ID, and for a client secret if your registration has one. No Node install, no JSON editing.

**Locally, in any MCP client.** Claude Code:

```bash
claude mcp add m365-mcp \
  -e MS365_MCP_CLIENT_ID="your-client-id" \
  -e MS365_MCP_TENANT_ID="your-tenant-id" \
  -- npx -y @masonator/m365-mcp@latest
```

For Cursor, Claude Desktop or anything that takes a JSON config:

```json
{
  "mcpServers": {
    "m365-mcp": {
      "command": "npx",
      "args": ["-y", "@masonator/m365-mcp"],
      "env": {
        "MS365_MCP_CLIENT_ID": "your-client-id",
        "MS365_MCP_TENANT_ID": "your-tenant-id"
      }
    }
  }
}
```

Add `MS365_MCP_CLIENT_SECRET` if your registration uses the Web platform with a secret.

### First run

The first tool call opens your browser to sign in with Microsoft. After you consent, tokens are stored at `~/.config/m365-mcp/tokens.json` (permissions `600`) and refreshed automatically.

To check the connection, ask your client to run `ms_auth_status`. It lists the permissions the token carries, so you can see which tools will work before you need them.

## What it does

The [tool reference](docs/tools.md) has every parameter. In short:

- **Catch up in one call.** `ms_brief` assembles today's meetings, unread mail, recent chats, open Planner tasks and yesterday's meeting transcripts. Give it a `person` and it briefs you on them instead.
- **Read mail, chats and channels.** `ms_mail`, `ms_chat` and `ms_teams` list what is recent and open a thread in full.
- **Read what was said in a meeting.** `ms_transcripts` finds meetings with transcripts by date and returns the full text.
- **See the calendar and who is free.** `ms_calendar` reads your events; `ms_schedule` checks free/busy for other people.
- **Find files and documents.** `ms_files` browses and searches OneDrive, `ms_sharepoint` covers sites and lists, and `ms_insights` shows documents you used recently or that were shared with you.
- **Search everything at once.** `ms_search` runs one query across mail, chats, calendar, OneDrive and SharePoint, with KQL.
- **Look people up.** `ms_people` resolves a name to an email, a manager and direct reports. `ms_tasks` reads To Do and Planner.

Anything a third party wrote, such as a mail body, a chat message or a transcript, comes back inside an untrusted-content fence with a random marker, so the model can tell data from instructions.

If a permission is missing from your registration, the tool that needs it says what is unavailable and what still works. It does not return a bare `403`.

## Example prompts

```text
Brief me on today
What did I miss in the Platform team channel this week?
Summarise yesterday's project sync from its transcript
Is Jane free tomorrow afternoon?
Find the email from finance about the Q3 budget
Catch me up on everything involving Sam Carter
Which documents were shared with me recently?
```

## Configuration

| Variable                  | Required | Description                                        |
| ------------------------- | -------- | -------------------------------------------------- |
| `MS365_MCP_CLIENT_ID`     | Yes      | Application (client) ID of the app registration    |
| `MS365_MCP_TENANT_ID`     | Yes      | Directory (tenant) ID                              |
| `MS365_MCP_CLIENT_SECRET` | No       | Client secret (confidential clients only)          |
| `MS365_MCP_TIMEZONE`      | No       | Timezone for calendar (default: system timezone)   |
| `MS365_MCP_WORKIQ_SCOPE`  | No       | Enables `ms_workiq`. Off by default: it can write. |

The redirect URI, scope list, client type and token file can be overridden too. All of them are in [docs/configuration.md](docs/configuration.md).

## App registration

Register an application in Microsoft Entra ID (Azure AD) with these settings:

1. **App registration** > New registration
2. **Redirect URI**: `http://localhost` (Web platform), or set a fixed URI via `MS365_MCP_REDIRECT_URL`
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
> with a client secret, the token request must not carry an `Origin` header. Azure
> rejects cross-origin token redemption for anything but SPA clients
> (`AADSTS9002326`). The server detects this from `MS365_MCP_CLIENT_SECRET` and
> omits the header automatically.

## Development

```bash
git clone https://github.com/StuMason/m365-mcp.git
cd m365-mcp && npm install
npm run build && npm test

MS365_MCP_CLIENT_ID="your-client-id" MS365_MCP_TENANT_ID="your-tenant-id" node dist/index.js
```

Contributions welcome: [CONTRIBUTING.md](CONTRIBUTING.md), and the architecture and Graph-quirk notes in [CLAUDE.md](CLAUDE.md).

## Work with me

I'm Stu Mason. I build MCP servers, AI integrations and agentic systems for agencies, SMEs and enterprise. This repo is what that work looks like in the open.

- **An MCP server for your product.** Give Claude, Cursor and every other AI client a proper way into your API, like this one.
- **Answers from your own stuff.** AI that answers from your documents and data, with the receipts, instead of guessing. Can stay on your own servers.
- **Work that runs itself.** Jobs on a schedule that sort, check and report, with a person signing off before anything goes out.

White-label under your own name if you're an agency. And if a job doesn't need AI, I'll say so before anyone's paid for anything.

📮 [hey@stumason.dev](mailto:hey@stumason.dev) · [stumason.dev](https://stumason.dev)

## Privacy Policy

This server collects nothing about you. There is no telemetry, no analytics and
no phone-home, and there is no hosted service behind it.

It talks to Microsoft and to nothing else: `login.microsoftonline.com` to sign
you in, `graph.microsoft.com` to read your data, and Microsoft's Work IQ service
only if you enable `ms_workiq`. Your tokens are stored in one file on your own
machine. Tool results go to your own MCP client, whose privacy policy governs
what happens next.

Full detail, including where in the code each claim is implemented, is in
[PRIVACY.md](PRIVACY.md).

## Links

- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview): the API this server reads from
- [MCP Registry](https://registry.modelcontextprotocol.io): listed as `io.github.StuMason/m365`
- [coolify-mcp](https://github.com/StuMason/coolify-mcp): my MCP server for Coolify, the self-hosted PaaS
- [Model Context Protocol](https://modelcontextprotocol.io/)

MIT © [Stu Mason](https://stumason.dev). If this is useful, ⭐ the repo. If it saved you an evening, [sponsor a coffee](https://github.com/sponsors/StuMason).
