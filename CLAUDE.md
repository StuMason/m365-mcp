# CLAUDE.md

## Project Overview

MCP server for Microsoft 365 via the Microsoft Graph API. Provides token-optimised tools for interacting with M365 services (Mail, Calendar, OneDrive, Teams, etc.) through the Model Context Protocol.

**Package:** `@masonator/m365-mcp`

## Commands

```bash
npm install          # Install dependencies
npm run build        # Compile TypeScript
npm test             # Run unit tests
npm run test:coverage # Run tests with coverage
npm run lint         # Run ESLint
npm run format       # Format with Prettier
npm run format:check # Check formatting
```

## Architecture

```text
src/
├── index.ts              # Entry point - stdio transport
├── lib/                  # Core implementation
├── types/                # TypeScript type definitions
└── __tests__/            # Jest tests
```

### Key Patterns

- **Microsoft Graph API**: All M365 operations go through the Graph API (`https://graph.microsoft.com/v1.0/`)
- **Context-optimised responses**: List endpoints return summaries, use get\_\* for full details
- **Token efficiency**: Minimise response sizes for LLM consumption

## Adding New Endpoints

1. Add types in `src/types/`
2. Add client method in `src/lib/`
3. Add MCP tool definition
4. Add tests in `src/__tests__/`
5. Update CHANGELOG.md, README.md

## Testing Requirements

- 80% coverage threshold enforced via codecov
- All new endpoints need mocked tests
- Integration tests in `src/__tests__/integration/`

## TypeScript Standards

- Explicit return types on all functions
- No implicit `any`
- Strict mode enabled

## Git Workflow

- Commit frequently, stage all files, push after commit
- Work on feature branches, PR to main
- Conventional commits: `feat:`, `fix:`, `chore:`, etc.

## Publishing

Auto-publish to npm on version bump via CI. Update version in `package.json`, add changelog entry, merge to main.
