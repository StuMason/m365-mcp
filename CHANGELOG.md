# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-09-11

First stable release. 13 read-only tools over the Microsoft Graph API, verified
against a live tenant.

### Breaking

- **Tool arguments are now validated against the published schema.** Previously an
  out-of-range `count` was silently clamped to the maximum; it now returns a clear
  validation error (`count: Too big: expected number to be <=50`) before the handler
  runs. Every numeric bound is published in the tool schema, so callers can get it
  right first time. The defensive clamping inside each tool is retained for direct
  callers. Any caller relying on the old tolerance needs to send in-range values.
- `ms_auth_status` is no longer annotated `readOnlyHint` / `idempotentHint`. It writes
  `tokens.json` and can open a browser, so by the spec's definition it modifies its
  environment. Clients that auto-approve read-only tools will now prompt for it.

### Changed

- **Migrated to MCP SDK v2** (`@modelcontextprotocol/core` + `@modelcontextprotocol/server`
  2.0.0, replacing `@modelcontextprotocol/sdk` 1.x, which is the end of that line). Tools
  are now registered with `server.registerTool` and zod 4 schemas, which removes the
  ~130-line dispatch switch and every `args as {...}` cast from `index.ts`. The negotiated
  wire protocol is unchanged at `2025-11-25` — both SDK lines top out there; v2 implements
  the newer 2026-07-28 spec revision's semantics.
- `ms_server_info` takes the tool roster as an argument instead of importing it, which
  also removes a module cycle.

### Added

- `src/lib/tools/index.ts` — `TOOL_DEFINITIONS`, a single source of truth for the tool
  roster, replacing three separately hand-maintained lists.
- `roster.test.ts` — fails the build if `index.ts` registers a different set of tools than
  the roster declares, if any tool is not annotated read-only, or if the "N tools" claims
  in README.md and CLAUDE.md disagree with reality. 0.7.0 shipped with all three lists out
  of step; this makes that a test failure rather than a release note.

### Fixed

- README Node badge said `>=18`; the package has required `>=20` since 0.7.0, and the v2
  SDK requires it too.

## [0.8.0] - 2026-09-11

### Added

- `ms_teams` — joined teams, channels, channel messages and reply threads
- `ms_tasks` — Microsoft To Do lists/tasks and assigned Planner tasks
- `ms_people` — directory search, one-person lookup with manager and direct
  reports, and the signed-in user's group memberships
- Seven newly-requested delegated scopes: `User.Read.All`,
  `ChannelMessage.Read.All`, `Channel.ReadBasic.All`, `Team.ReadBasic.All`,
  `OnlineMeetings.Read`, `Group.Read.All`, `Tasks.Read`
- MCP tool `title` and read-only `annotations` on every tool, plus server
  `instructions` in the initialize handshake
- A clear error when the OAuth callback port is already in use, rather than a
  silent hang — this matters when `MS365_MCP_REDIRECT_URL` pins a fixed port

### Fixed

- Token exchange and refresh no longer send an `Origin` header for confidential
  (secret-bearing) clients. Azure permits cross-origin token redemption only for
  SPA-platform registrations and rejects it with `AADSTS9002326`, which broke
  sign-in against any Web-platform app registration.
- The refresh grant now requests the full scope set explicitly, so a token cached
  by an earlier release picks up newly added scopes instead of silently
  refreshing with the old, narrower grant.

### Changed

- `@modelcontextprotocol/sdk` 1.26 → 1.30; all other dependencies updated to
  their latest in-range versions. `npm audit` reports zero vulnerabilities.
- Server version is now read from `package.json` rather than hardcoded, so the
  MCP handshake and `ms_server_info` cannot drift from the released version.

## [0.1.0] - 2026-02-20

### Added

- Initial project scaffolding
- CI/CD pipeline with GitHub Actions
- Dependabot with auto-merge for patch/minor updates
- Auto-publish to npm on version bump
