# v0.7 Full Scope Coverage

**Date:** 2026-02-24
**Commit:** a3be695

## What changed

Expanded m365-mcp from 8 to 10 tools and ~13 to ~26 Graph API endpoints. Version bumped from 0.6.0 to 0.7.0.

**New tools:**

- `ms_schedule` — POST `/me/calendar/getSchedule` for free/busy availability checking
- `ms_sharepoint` — Search sites, list site lists, browse list items with expanded fields

**Expanded existing tools:**

- `ms_profile` — 6→16 fields, `include` param for manager/reports/groups/photo sub-fetches
- `ms_calendar` — Added calendars list mode and event detail drill-down
- `ms_mail` — Added folders list, folder messages, attachments, keyword filter
- `ms_chat` — Added members listing mode
- `ms_files` — Added file detail with download URL and shared-with-me listing

**Infrastructure:**

- `graphPost()` added alongside `graphFetch()` with shared internals (`buildHeaders`, `buildUrl`, `handleResponse`, `handleNetworkError`)
- Custom headers support in `GraphFetchOptions`
- Fixed `$orderby` on `/me/chats` (unsupported by Graph API)

**Quality hardening (from PR review):**

- `encodeURIComponent` on all user-supplied URL path segments
- Try-catch around `response.json()` on success path and `response.text()` on error path
- `fetchPhoto` distinguishes 404 from other errors
- Unknown `include` values produce warnings instead of silent ignoring
- Empty emails validation in schedule tool
- Error prefix pattern (`Error fetching X:`) for sub-fetch failures
- `BaseDriveItem` extracted to reduce type duplication across drive item interfaces
- Non-primitive values in SharePoint `formatItem` use `JSON.stringify` instead of implicit `toString()`
- Removed dead `recurrence` field from `EventDetail`

**Tests:** 144→270 tests, coverage 94.42% statements / 84.32% branches / 95.23% lines.

## Why

v0.6 covered the basics but left significant Graph API surface area unused. Users needed schedule checking, SharePoint access, mail folder browsing, file detail retrieval, and richer profile data. The goal was to reach functional parity with what a power user would expect from a Microsoft 365 MCP integration.

## Decisions made

- **Read-only scope maintained**: All new endpoints are GET or non-mutating POST (getSchedule). No send/create/delete operations to keep the trust model simple.
- **Mode-based routing over separate tools**: Each tool uses parameter presence to select behavior (e.g. `event_id` → detail, `calendars: true` → list calendars) rather than creating separate MCP tools for each endpoint. This keeps the tool count manageable for LLM tool selection.
- **Shared graph client internals**: Extracted `buildHeaders`, `buildUrl`, `handleResponse`, `handleNetworkError` as private helpers shared between `graphFetch` and `graphPost` rather than duplicating logic.
- **`include` array for profile sub-fetches**: Rather than separate tools or always-on fetches, the profile tool accepts `include: ["manager", "reports", "groups", "photo"]` for opt-in expansion. This avoids extra API calls when only basic info is needed.
- **`BaseDriveItem` extraction**: Three drive item interfaces shared 6 fields. Used interface inheritance with a type alias (`type DriveItem = BaseDriveItem`) to satisfy the linter's no-empty-object-type rule.
- **JSON.stringify for non-primitive SharePoint fields**: SharePoint list item fields can contain nested objects. `JSON.stringify` is a pragmatic choice over `[object Object]`, and the `@odata`/`_` prefix filter keeps internal metadata out of output.

## Rejected alternatives

- **Separate MCP tools per endpoint**: Would have meant ~26 tools instead of 10. LLMs handle a smaller tool set more reliably, and the mode-based approach mirrors how users think ("check my calendar" vs "list my calendars" as variants of one concept).
- **Shared `stripHtml` utility**: Three files (calendar, mail, chat) each have their own `stripHtml`. Considered extracting to a shared utility but each has slightly different behavior tuned to its content type. The duplication is low-cost and avoids coupling.
- **Type aliases everywhere instead of interfaces**: Considered making `DriveItemDetail` and `SharedDriveItem` type intersections (`BaseDriveItem & { ... }`) but interfaces with `extends` are more readable and provide better error messages.

## Context

- Design doc: `docs/plans/2026-02-23-v07-full-scope-coverage-design.md`
- Implementation plan: `docs/plans/2026-02-23-v07-implementation-plan.md`
- Built using subagent-driven development across two sessions (12 tasks + 4-agent PR review)
- PR review caught 3 critical issues (URL encoding, unprotected JSON parse, photo error swallowing) and 4 important issues — all resolved before this CDR
