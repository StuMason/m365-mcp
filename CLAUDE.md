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
│   ├── graph.ts          # graphFetch() wrapper with error mapping
│   └── tools/
│       ├── auth-status.ts  # ms_auth_status — connection check + re-auth
│       ├── profile.ts      # ms_profile — /me
│       ├── calendar.ts     # ms_calendar — /me/calendarView
│       ├── mail.ts         # ms_mail — /me/messages
│       ├── chat.ts         # ms_chat — /me/chats
│       ├── files.ts        # ms_files — /me/drive
│       └── transcripts.ts  # ms_transcripts — calendar → meeting ID → VTT
└── __tests__/            # Jest tests (144 tests, ~94% coverage)
```

### Auth Flow

OAuth2 confidential client (client_secret). On first run, opens browser for Microsoft consent. Tokens stored at `~/.config/m365-mcp/tokens.json` (chmod 600). Auto-refresh with 2-minute expiry buffer.

### Key Patterns

- **graphFetch()** wraps all Graph API calls with typed results (`GraphResult<T>`) and maps HTTP errors to user-friendly messages
- **Each tool** exports a `toolDefinition` and `execute` function. `index.ts` wires them into the MCP protocol.
- **Transcript drill-down**: compound `{meetingId}/{transcriptId}` IDs for HATEOAS-style lazy loading of full VTT content
- **Timezone**: uses system timezone by default, configurable via `MS365_MCP_TIMEZONE` env var

## Adding a New Tool

1. Create `src/lib/tools/my-tool.ts` with exported `myToolDefinition` and `executeMyTool(token, args)`
2. Register in `src/index.ts`: add to `ListToolsRequestSchema` array and `CallToolRequestSchema` switch
3. Add tests in `src/__tests__/tools/my-tool.test.ts` (mock `graphFetch`)
4. Update README.md

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
- GPG signing via 1Password (use `--no-gpg-sign` if agent unavailable)

## Publishing

Auto-publish to npm on version bump via CI.
