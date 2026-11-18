# Tool reference

Every tool the server registers, with its parameters. `ms_server_info` returns the
same roster from a running server.

Anything a third party wrote, such as a mail body, a chat message or a transcript,
comes back inside an untrusted-content fence so the model can tell data from
instructions.

## `ms_auth_status`

Check connection status. If not connected, opens browser to sign in.

## `ms_profile`

Fetch your Microsoft 365 profile — display name, email, job title, office location.

## `ms_calendar`

Fetch calendar events. Defaults to today.

| Parameter | Description                                                 |
| --------- | ----------------------------------------------------------- |
| `date`    | Specific date (YYYY-MM-DD)                                  |
| `start`   | Start of range (ISO 8601)                                   |
| `end`     | End of range (ISO 8601)                                     |
| `compact` | Summarise each event, omitting the body. Good for scanning. |

## `ms_mail`

Read recent emails with optional keyword search.

| Parameter | Description                         |
| --------- | ----------------------------------- |
| `search`  | Keyword to filter emails            |
| `count`   | Number of emails (1-25, default 10) |

## `ms_chat`

Read Teams chats. Without `chat_id` lists recent chats; with `chat_id` returns messages from that thread.

| Parameter | Description                        |
| --------- | ---------------------------------- |
| `chat_id` | Specific chat thread ID            |
| `count`   | Number of items (1-25, default 10) |

## `ms_files`

Browse or search OneDrive files.

| Parameter | Description                      |
| --------- | -------------------------------- |
| `path`    | Folder path (e.g., `/Documents`) |
| `search`  | Search across OneDrive           |
| `count`   | Max items (1-50, default 20)     |

## `ms_transcripts`

Fetch Teams meeting transcripts. Returns previews (~3000 chars) with a `transcript_id` for drill-down to the full transcript.

| Parameter       | Description                                   |
| --------------- | --------------------------------------------- |
| `date`          | Date (YYYY-MM-DD)                             |
| `start`         | Start of range (ISO 8601)                     |
| `end`           | End of range (ISO 8601)                       |
| `transcript_id` | ID from a previous list call for full content |

## `ms_schedule`

Check free/busy for one or more people over a time window. Slots come back as free,
tentative, busy, out of office or working elsewhere.

| Parameter  | Description                                   |
| ---------- | --------------------------------------------- |
| `emails`   | Email addresses to check (required)           |
| `date`     | Date to check (YYYY-MM-DD, defaults to today) |
| `start`    | Start time (HH:MM, 24h, default 08:00)        |
| `end`      | End time (HH:MM, 24h, default 18:00)          |
| `interval` | Slot length in minutes (5-1440, default 30)   |

## `ms_sharepoint`

Search SharePoint sites, list a site's lists, or browse list items. With no parameters
it searches every site you can reach.

| Parameter | Description                                      |
| --------- | ------------------------------------------------ |
| `search`  | Query for finding sites (default `*`, all sites) |
| `site_id` | Site ID to see its lists                         |
| `list_id` | List ID (with `site_id`) to browse its items     |
| `count`   | Max results (1-50, default 10)                   |

## `ms_teams`

Browse joined Teams, their channels, and channel messages. Progressive drill-down:
no arguments lists teams, `team_id` lists channels, `team_id` + `channel_id` reads messages.

| Parameter    | Description                                           |
| ------------ | ----------------------------------------------------- |
| `team_id`    | Team ID to list its channels                          |
| `channel_id` | Channel ID (with `team_id`) to read messages          |
| `message_id` | Message ID (with both above) to read its reply thread |
| `count`      | Max results (1-50, default 20)                        |

## `ms_tasks`

Read Microsoft To Do and Planner tasks. Completed tasks are hidden unless asked for.

| Parameter           | Description                              |
| ------------------- | ---------------------------------------- |
| `list_id`           | To Do list ID to read its tasks          |
| `planner`           | Return assigned Planner tasks instead    |
| `include_completed` | Include finished tasks (default `false`) |
| `count`             | Max results (1-50, default 25)           |

## `ms_people`

Look people up in the organisation directory. `search` resolves a name to the email
address that `ms_schedule` needs.

| Parameter | Description                                                   |
| --------- | ------------------------------------------------------------- |
| `search`  | Name or partial name to search for                            |
| `user`    | Email or object ID — returns details, manager, direct reports |
| `groups`  | List the signed-in user's group and team memberships          |
| `count`   | Max results (1-50, default 20)                                |

## `ms_search`

Search across mail, Teams chats, calendar, OneDrive and SharePoint in one call. Use this
when you don't already know where something lives. Supports KQL, so `from:jane subject:budget`
works.

| Parameter | Description                                                |
| --------- | ---------------------------------------------------------- |
| `query`   | What to search for (required)                              |
| `types`   | Limit to `mail`, `chat`, `calendar`, `files`, `sharepoint` |
| `count`   | Max results per area (1-25, default 5)                     |

## `ms_insights`

Documents you recently worked with, or that were shared with you.

| Parameter | Description                               |
| --------- | ----------------------------------------- |
| `kind`    | `used` (default), `shared`, or `trending` |
| `count`   | Max results (1-50, default 15)            |

> `trending` is disabled by policy in many tenants; the tool says so plainly rather than
> returning an error.

## `ms_brief`

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

## `ms_workiq`

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

## `ms_server_info`

Server metadata: version, registered tools, and which environment variables are set.
