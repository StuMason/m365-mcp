# CLAUDE.md

## Project Overview

Standalone MCP server for Microsoft 365 via the Microsoft Graph API. Provides read-only tools for profile, calendar, mail, Teams chat, OneDrive files, and meeting transcripts. Any MCP client can use it.

**Package:** `@masonator/m365-mcp`

## Commands

```bash
npm install          # Install dependencies
npm run build        # Compile TypeScript (tsc)
npm test             # Run unit tests (jest, 80% coverage threshold)
npm run lint         # Run ESLint
npm run format       # Format with Prettier
```

## Architecture

```text
src/
├── index.ts              # MCP server entry point (stdio transport, tool dispatch)
├── types/
│   └── tokens.ts         # TokenData, AuthConfig interfaces
├── lib/
│   ├── auth.ts           # OAuth2 confidential client, token storage, refresh
│   ├── version.ts        # single source of truth for the version (package.json)
│   ├── graph.ts          # graphFetch() wrapper with error mapping
│   └── tools/
│       ├── index.ts        # TOOL_DEFINITIONS — the single source of truth for the roster
│       ├── auth-status.ts  # ms_auth_status — connection check + re-auth
│       ├── profile.ts      # ms_profile — /me
│       ├── calendar.ts     # ms_calendar — /me/calendarView
│       ├── mail.ts         # ms_mail — /me/messages
│       ├── chat.ts         # ms_chat — /me/chats
│       ├── files.ts        # ms_files — /me/drive
│       ├── sharepoint.ts   # ms_sharepoint — /sites
│       ├── schedule.ts     # ms_schedule — /me/calendar/getSchedule
│       ├── teams.ts        # ms_teams — /me/joinedTeams → channels → messages
│       ├── tasks.ts        # ms_tasks — /me/todo, /me/planner/tasks
│       ├── people.ts       # ms_people — /users, /me/memberOf
│       ├── search.ts       # ms_search — /search/query across M365
│       ├── insights.ts     # ms_insights — /me/insights/{used,shared,trending}
│       ├── brief.ts        # ms_brief — composes the tools above, no Graph calls
│       ├── server-info.ts  # ms_server_info — version + registered tools
│       └── transcripts.ts  # ms_transcripts — calendar → meeting ID → VTT
└── __tests__/            # Jest tests (393 tests, ~95% coverage)
```

### Auth Flow

OAuth2 confidential client (client_secret). On first run, opens browser for Microsoft consent. Tokens stored at `~/.config/m365-mcp/tokens.json` (chmod 600). Auto-refresh with 2-minute expiry buffer.

### Key Patterns

- **graphFetch()** wraps all Graph API calls with typed results (`GraphResult<T>`) and maps HTTP errors to user-friendly messages
- **Each tool** exports a `toolDefinition` (with a zod `inputSchema`) and an `execute` function.
  `index.ts` registers them with `server.registerTool`, which types each handler's args from
  its own schema — that is why handlers are not held in one array.
- **The roster** lives in `src/lib/tools/index.ts` (`TOOL_DEFINITIONS`). `ms_server_info`
  counts it and `roster.test.ts` asserts `index.ts` registers exactly those tools and that
  the "N tools" claims in README.md and CLAUDE.md agree. 0.7.0 shipped three lists that
  disagreed; this is the guard against a repeat.
- **`server-info.ts` takes the roster as an argument** rather than importing it. It is
  itself in the roster, so importing it back is a module cycle (a real TDZ crash at startup).
- **Transcript drill-down**: compound `{meetingId}/{transcriptId}` IDs for HATEOAS-style lazy loading of full VTT content
- **Timezone**: uses system timezone by default, configurable via `MS365_MCP_TIMEZONE` env var
- **Confidential clients**: when `MS365_MCP_CLIENT_SECRET` is set, the token and
  refresh requests must NOT send an `Origin` header. Azure allows cross-origin
  token redemption only for SPA-platform registrations and otherwise fails with
  `AADSTS9002326`. Do not "restore" that header.
- **Scopes**: `SCOPES` in `auth.ts` must stay a subset of what is actually consented
  on the app registration. Requesting an unconsented scope produces a consent prompt
  the user cannot complete.
- **Graph query quirks**: `/me/joinedTeams` and `/teams/{id}/channels` reject `$top`
  and are trimmed client-side; `/users?$search` needs the `ConsistencyLevel: eventual`
  header.
- **`/search/query` constraints** (verified live, do not "simplify"): only ONE
  `entityRequest` per call, and entity types cannot be combined freely —
  `message`+`chatMessage` is legal, `message`+`event` is not. `ms_search` therefore
  issues one call per compatible group, in parallel. `person` needs `People.Read`,
  which is not consented.
- **`ms_brief` composes, it does not call Graph.** Every section is an existing
  `execute*` function. Keep it that way: formatting and error handling belong with
  the area they came from. A failing section degrades to a note rather than taking
  the brief down.
- **Item insights are often disabled tenant-wide** (`trending` returns 403
  `ItemInsightsDisabled`). That is policy, not a fault, so `ms_insights` explains it
  instead of surfacing a raw error.

## Adding a New Tool

1. Create `src/lib/tools/my-tool.ts` exporting `myToolDefinition` (name, title, description,
   zod `inputSchema`, read-only `annotations`) and `executeMyTool(token, args)`
2. Add the definition to `TOOL_DEFINITIONS` in `src/lib/tools/index.ts`
3. Register it in `src/index.ts` with `server.registerTool(def.name, def, withToken(execute))`
4. Add tests in `src/__tests__/tools/my-tool.test.ts` (mock `graphFetch`)
5. Update the tool count and docs in README.md — `roster.test.ts` fails until they agree

### Schema conventions

Use zod 4. Give every numeric bound explicitly (`z.int().min(1).max(50)`): a bare `z.int()`
publishes `minimum: -9007199254740991` into the tool schema. Out-of-range arguments are now
rejected by the SDK before the handler runs, so keep the clamping in `execute*` anyway —
tests call those functions directly.

## Testing

- 80% coverage threshold on `src/lib/**/*.ts`
- Mock `graphFetch` for tool tests, mock `fetch` for graph/auth tests
- Utility functions (extractMeetingId, parseTranscriptId, formatFileSize) tested directly

## TypeScript Standards

- Explicit return types on all functions (eslint warn)
- No implicit `any`
- Strict mode, NodeNext module resolution
- `.js` extensions in relative imports

## Git Workflow

- Conventional commits: `feat:`, `fix:`, `chore:`
- Pre-commit hooks: eslint + prettier via lint-staged
- Commits are SSH-signed (`commit.gpgsign=true` globally). Never disable signing.

## Publishing

Auto-publish to npm on version bump via CI.
