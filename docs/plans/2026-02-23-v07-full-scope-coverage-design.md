# v0.7 — Full Scope Coverage Design

**Date:** 2026-02-23
**Status:** Approved

## Goal

Expand from 8 tools / ~13 endpoints to 10 tools / ~26 endpoints, using all capabilities proven working against the Azure AD tenant within the 8 granted scopes.

## Graph Client Changes

### Extended GraphFetchOptions

Add optional `headers` field to `GraphFetchOptions`, merged after default headers (Authorization, timezone). Allows tools to pass custom headers (e.g., `ConsistencyLevel: eventual` for mail search).

```typescript
interface GraphFetchOptions {
  beta?: boolean;
  timezone?: boolean;
  headers?: Record<string, string>;
}
```

### New graphPost Function

`graphPost<TBody, TResult>()` alongside `graphFetch`. Same error handling, same options type. Sets `Content-Type: application/json` and uses POST method. Both functions share a private `buildHeaders()` helper.

```typescript
export async function graphPost<TBody, TResult>(
  path: string,
  token: string,
  body: TBody,
  options?: GraphFetchOptions,
): Promise<GraphResult<TResult>>;
```

## Bug Fix: ms_chat $orderby

Remove `$orderby=lastMessagePreview/createdDateTime desc` from `/me/chats` query — endpoint returns 400. Keep `$orderby=createdDateTime desc` on `/me/chats/{id}/messages` (different endpoint, works).

## Expanded Tools

### ms_profile — Full profile + org context

Expand `$select` to 16 fields. Add optional `include` array parameter:

- `"photo"` — GET `/me/photo/$value`, return size confirmation only
- `"manager"` — GET `/me/manager`, handle 403 gracefully (contractor accounts)
- `"reports"` — GET `/me/directReports`, may be empty
- `"groups"` — GET `/me/memberOf`, use `displayName ?? mail ?? id ?? "(unnamed)"` for null names

### ms_mail — Folders, attachments, filters

Four new modes via new parameters:

- `folders: true` — GET `/me/mailFolders` with unread counts
- `folder: string` — GET `/me/mailFolders/{folder}/messages` (well-known names supported)
- `message_id + attachments: true` — GET `/me/messages/{id}/attachments` (no contentBytes)
- `filter: string` — Map shortcuts to OData: "unread", "flagged", "attachments", "important"

Mail search uses `ConsistencyLevel: eventual` header via extended GraphFetchOptions.

### ms_calendar — Event detail, calendars list

Two new modes:

- `calendars: true` — GET `/me/calendars`
- `event_id: string` — GET `/me/events/{id}` with attendees, response status, Teams join URL, stripped HTML body

### ms_chat — Members + bug fix

- Fix: Remove `$orderby` from chat list
- New mode: `chat_id + members: true` — GET `/me/chats/{id}/members` (no `$select`)

### ms_files — Detail/download URL, shared

Two new modes:

- `item_id: string` — GET `/me/drive/items/{id}` with `@microsoft.graph.downloadUrl`
- `shared: true` — GET `/me/drive/sharedWithMe`

## New Tools

### ms_schedule — Availability checker

POST `/me/calendar/getSchedule`. Accepts emails, date, start/end times, interval. Parses `availabilityView` string (0=free, 1=tentative, 2=busy, 3=OOF, 4=working elsewhere). Formats readable schedule per person. Uses UTC timezone.

### ms_sharepoint — Sites, lists, documents

Three modes:

- Search sites (default `search=*`)
- Site lists (`site_id`)
- List items (`site_id + list_id`) with expanded fields

Uses already-granted `Sites.Read.All` scope.

## Registration

- `index.ts`: Register 2 new tools in ListTools and CallTools handlers
- `server-info.ts`: Update tool list (10 tools total)
- Version bump: `package.json` and `index.ts` to 0.7.0

## Tests

Follow existing pattern: `jest.unstable_mockModule` + dynamic imports. Add `mockGraphPost` for schedule. Key cases:

- Chat list without `$orderby`
- Chat members with no `$select`
- Profile manager 403 graceful handling
- Profile groups with null displayName
- Mail folders, filters
- Schedule POST with multiple emails
- SharePoint default search
- Files detail with `@microsoft.graph.downloadUrl`

80% coverage threshold maintained.

## API Quirks Reference

1. `/me/chats` does NOT support `$orderby` (400)
2. `/me/chats/{id}/messages` does NOT support `$select` (400)
3. `/me/chats/{id}/members` does NOT support `$select` (400)
4. `/me/memberOf` groups may have null displayName
5. `/me/manager` returns 403 for contractor accounts
6. Mail `$search` requires `ConsistencyLevel: eventual` header
7. `/me/drive/items/{id}` returns `@microsoft.graph.downloadUrl` (~1hr TTL, no auth needed)
8. POST `/me/calendar/getSchedule` is the only POST call
