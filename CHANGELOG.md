# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
