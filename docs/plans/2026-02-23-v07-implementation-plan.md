# v0.7 Full Scope Coverage — Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Expand from 8 tools / ~13 Graph endpoints to 10 tools / ~26 endpoints, using all proven capabilities within the 8 granted Azure AD scopes.

**Architecture:** Existing tool-per-file pattern. Each tool exports a definition + execute function. `graphFetch` wrapper handles all GET calls; a new `graphPost` handles POST. `index.ts` wires everything via MCP protocol. Tests mock `graphFetch`/`graphPost` at module level.

**Tech Stack:** TypeScript (strict, NodeNext), @modelcontextprotocol/sdk, Jest (ESM via unstable_mockModule)

**Reference:** See `docs/plans/2026-02-23-v07-full-scope-coverage-design.md` for full design.

---

## Task 1: Fix ms_chat $orderby bug

Files:

- Modify: `src/lib/tools/chat.ts:131-133`
- Modify: `src/__tests__/tools/chat.test.ts`

### Step 1: Write test verifying no $orderby in chat list URL

In `src/__tests__/tools/chat.test.ts`, add a test inside the existing `describe('executeChat')`:

```typescript
it('does not include $orderby in chat list URL', async () => {
  mockGraphFetch.mockResolvedValue({
    ok: true,
    data: { value: [] },
  });

  await executeChat('test-token', {});

  const calledPath = mockGraphFetch.mock.calls[0][0] as string;
  expect(calledPath).not.toContain('$orderby');
});
```

### Step 2: Run test to verify it fails

Run: `npm test -- --testPathPattern=chat.test`

Expected: FAIL — the current code includes `$orderby=lastMessagePreview/createdDateTime desc`

### Step 3: Fix the code

In `src/lib/tools/chat.ts`, replace lines 131-133:

```typescript
// BEFORE (broken — /me/chats does NOT support $orderby):
const path =
  `/me/chats?$top=${count}&$orderby=lastMessagePreview/createdDateTime desc` +
  `&$expand=lastMessagePreview&$select=id,topic,chatType,lastMessagePreview`;

// AFTER:
const path =
  `/me/chats?$top=${count}` +
  `&$expand=lastMessagePreview&$select=id,topic,chatType,lastMessagePreview`;
```

### Step 4: Run tests

Run: `npm test -- --testPathPattern=chat.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/chat.ts src/__tests__/tools/chat.test.ts
git commit --no-gpg-sign -m "fix: remove unsupported \$orderby from /me/chats endpoint"
```

---

## Task 2: Add graphPost and extend GraphFetchOptions with headers

Files:

- Modify: `src/lib/graph.ts`
- Modify: `src/__tests__/graph.test.ts`

### Step 1: Write tests for custom headers on graphFetch

Add to `src/__tests__/graph.test.ts`:

```typescript
it('merges custom headers into request', async () => {
  const mock = mockFetch({
    ok: true,
    json: () => Promise.resolve({ value: [] }),
  } as Partial<Response>);

  await graphFetch('/me/messages', 'test-token', {
    timezone: false,
    headers: { ConsistencyLevel: 'eventual' },
  });

  const callHeaders = mock.mock.calls[0][1]!.headers as Record<string, string>;
  expect(callHeaders).toHaveProperty('ConsistencyLevel', 'eventual');
  expect(callHeaders).toHaveProperty('Authorization', 'Bearer test-token');
});
```

### Step 2: Write tests for graphPost

Add a new `describe('graphPost')` block in `src/__tests__/graph.test.ts`:

```typescript
describe('graphPost', () => {
  it('sends POST request with JSON body', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({ value: [{ scheduleId: 'user@example.com' }] }),
    } as Partial<Response>);

    const result = await graphPost<{ schedules: string[] }, { value: unknown[] }>(
      '/me/calendar/getSchedule',
      'test-token',
      { schedules: ['user@example.com'] },
    );

    expect(result).toEqual({
      ok: true,
      data: { value: [{ scheduleId: 'user@example.com' }] },
    });

    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/v1.0/me/calendar/getSchedule',
      expect.objectContaining({
        method: 'POST',
        body: JSON.stringify({ schedules: ['user@example.com'] }),
        headers: expect.objectContaining({
          Authorization: 'Bearer test-token',
          'Content-Type': 'application/json',
        }),
      }),
    );
  });

  it('returns error on failed POST', async () => {
    mockFetch({
      ok: false,
      status: 403,
    } as Partial<Response>);

    const result = await graphPost('/me/calendar/getSchedule', 'test-token', {});

    expect(result).toEqual({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });
  });

  it('handles network error on POST', async () => {
    globalThis.fetch = jest.fn<typeof fetch>().mockRejectedValue(new Error('Network failure'));

    const result = await graphPost('/me/calendar/getSchedule', 'test-token', {});

    expect(result).toEqual({
      ok: false,
      error: { status: 0, message: 'Network error: Network failure' },
    });
  });

  it('uses beta URL when beta option is true', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({}),
    } as Partial<Response>);

    await graphPost('/me/calendar/getSchedule', 'test-token', {}, { beta: true });

    expect(mock).toHaveBeenCalledWith(
      'https://graph.microsoft.com/beta/me/calendar/getSchedule',
      expect.any(Object),
    );
  });

  it('merges custom headers into POST request', async () => {
    const mock = mockFetch({
      ok: true,
      json: () => Promise.resolve({}),
    } as Partial<Response>);

    await graphPost(
      '/test',
      'test-token',
      {},
      {
        timezone: false,
        headers: { 'X-Custom': 'value' },
      },
    );

    const callHeaders = mock.mock.calls[0][1]!.headers as Record<string, string>;
    expect(callHeaders).toHaveProperty('X-Custom', 'value');
  });
});
```

### Step 3: Run tests to verify they fail

Run: `npm test -- --testPathPattern=graph.test`

Expected: FAIL — graphPost doesn't exist, headers option not supported

### Step 4: Implement

Rewrite `src/lib/graph.ts`. Extract a `buildHeaders()` helper, add `headers` to `GraphFetchOptions`, add `graphPost`:

```typescript
export interface GraphError {
  status: number;
  message: string;
}

export type GraphResult<T> = { ok: true; data: T } | { ok: false; error: GraphError };

export interface GraphFetchOptions {
  beta?: boolean;
  timezone?: boolean;
  headers?: Record<string, string>;
}

function buildUrl(path: string, options?: GraphFetchOptions): string {
  const base = options?.beta
    ? 'https://graph.microsoft.com/beta'
    : 'https://graph.microsoft.com/v1.0';
  return `${base}${path}`;
}

function buildHeaders(token: string, options?: GraphFetchOptions): Record<string, string> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  if (options?.timezone !== false) {
    const tz =
      process.env.MS365_MCP_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    headers['Prefer'] = `outlook.timezone="${tz}"`;
  }

  if (options?.headers) {
    Object.assign(headers, options.headers);
  }

  return headers;
}

function mapError(status: number, response: Response): string | Promise<string> {
  switch (status) {
    case 401:
      return 'Graph token expired. Use ms_auth_status to reconnect.';
    case 403:
      return 'Insufficient permissions. Check granted scopes with ms_auth_status.';
    case 404:
      return 'Resource not found. The item may not exist or you may lack access.';
    default:
      return response.text().then((text) => `Graph API error (${status}): ${text}`);
  }
}

async function handleResponse<T>(response: Response): Promise<GraphResult<T>> {
  if (response.ok) {
    const data = (await response.json()) as T;
    return { ok: true, data };
  }

  const message = await mapError(response.status, response);
  return { ok: false, error: { status: response.status, message } };
}

function handleNetworkError<T>(err: unknown): GraphResult<T> {
  return {
    ok: false,
    error: {
      status: 0,
      message: `Network error: ${err instanceof Error ? err.message : String(err)}`,
    },
  };
}

/**
 * Thin fetch wrapper for Microsoft Graph API GET calls.
 * Translates HTTP errors into typed GraphError results.
 */
export async function graphFetch<T>(
  path: string,
  token: string,
  options?: GraphFetchOptions,
): Promise<GraphResult<T>> {
  let response: Response;
  try {
    response = await fetch(buildUrl(path, options), {
      headers: buildHeaders(token, options),
    });
  } catch (err) {
    return handleNetworkError(err);
  }

  return handleResponse<T>(response);
}

/**
 * Thin fetch wrapper for Microsoft Graph API POST calls.
 * Same error handling as graphFetch but sends a JSON body.
 */
export async function graphPost<TBody, TResult>(
  path: string,
  token: string,
  body: TBody,
  options?: GraphFetchOptions,
): Promise<GraphResult<TResult>> {
  let response: Response;
  try {
    response = await fetch(buildUrl(path, options), {
      method: 'POST',
      headers: buildHeaders(token, options),
      body: JSON.stringify(body),
    });
  } catch (err) {
    return handleNetworkError(err);
  }

  return handleResponse<TResult>(response);
}
```

### Step 5: Run tests

Run: `npm test -- --testPathPattern=graph.test`

Expected: All PASS

### Step 6: Run full suite to check for regressions

Run: `npm test`

Expected: All PASS — existing tool tests mock graphFetch so they're unaffected

### Step 7: Commit

```bash
git add src/lib/graph.ts src/__tests__/graph.test.ts
git commit --no-gpg-sign -m "feat: add graphPost and custom headers support to graph client"
```

---

## Task 3: Expand ms_profile

Files:

- Modify: `src/lib/tools/profile.ts`
- Modify: `src/__tests__/tools/profile.test.ts`

### Step 1: Write tests for expanded profile

Replace the contents of `src/__tests__/tools/profile.test.ts` with tests covering:

1. Expanded field set (16 fields)
2. `include: ["manager"]` success case
3. `include: ["manager"]` with 403 (graceful handling)
4. `include: ["reports"]` with results
5. `include: ["reports"]` empty (normal)
6. `include: ["groups"]` with null displayName
7. `include: ["photo"]` returning size confirmation
8. `include: ["manager", "reports", "groups"]` all at once
9. Error from base profile call

The mock pattern stays the same. Key test for manager 403:

```typescript
it('handles manager 403 gracefully', async () => {
  // First call: profile success
  mockGraphFetch.mockResolvedValueOnce({
    ok: true,
    data: {
      displayName: 'Stuart Mason',
      mail: 'stuart@example.com',
      jobTitle: 'Engineer',
    },
  });
  // Second call: manager 403
  mockGraphFetch.mockResolvedValueOnce({
    ok: false,
    error: { status: 403, message: 'Insufficient permissions.' },
  });

  const result = await executeProfile('test-token', { include: ['manager'] });

  expect(result).toContain('Name: Stuart Mason');
  expect(result).toContain('Manager info not available (tenant policy)');
  expect(result).not.toContain('Error');
});
```

Key test for groups with null displayName:

```typescript
it('handles groups with null displayName', async () => {
  mockGraphFetch.mockResolvedValueOnce({
    ok: true,
    data: { displayName: 'Stuart', mail: 'stuart@example.com' },
  });
  mockGraphFetch.mockResolvedValueOnce({
    ok: true,
    data: {
      value: [
        { displayName: 'Engineering', id: 'g1' },
        { displayName: null, mail: 'secret-group@example.com', id: 'g2' },
        { displayName: null, mail: null, id: 'g3' },
        { displayName: null, mail: null, id: null },
      ],
    },
  });

  const result = await executeProfile('test-token', { include: ['groups'] });

  expect(result).toContain('Engineering');
  expect(result).toContain('secret-group@example.com');
  expect(result).toContain('g3');
  expect(result).toContain('(unnamed)');
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=profile.test`

Expected: FAIL — executeProfile doesn't accept args yet

### Step 3: Implement expanded profile

Rewrite `src/lib/tools/profile.ts`:

- Update `profileToolDefinition` to include the `include` parameter
- Expand `$select` to 16 fields
- Update `ProfileResponse` interface to include all fields
- Add interfaces for manager, reports, groups, photo responses
- Add conditional `include` fetches after the main profile fetch
- Handle manager 403 specifically: if `result.error.status === 403`, return note instead of error
- Handle photo: check response status, return `Photo available ({size} bytes)` or similar
- For groups: `(item.displayName ?? item.mail ?? item.id ?? '(unnamed)')`
- Update `executeProfile` signature to accept `(token: string, args?: { include?: string[] })`

### Step 4: Run tests

Run: `npm test -- --testPathPattern=profile.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/profile.ts src/__tests__/tools/profile.test.ts
git commit --no-gpg-sign -m "feat(profile): expand to 16 fields, add manager/reports/groups/photo include"
```

---

## Task 4: Expand ms_mail

Files:

- Modify: `src/lib/tools/mail.ts`
- Modify: `src/__tests__/tools/mail.test.ts`

### Step 1: Write tests for new mail modes

Add to `src/__tests__/tools/mail.test.ts`:

Folders mode:

```typescript
describe('folder listing mode', () => {
  it('lists mail folders with counts', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { displayName: 'Inbox', unreadItemCount: 3, totalItemCount: 150 },
          { displayName: 'Sent Items', unreadItemCount: 0, totalItemCount: 50 },
        ],
      },
    });

    const result = await executeMail('test-token', { folders: true });

    expect(result).toContain('Inbox');
    expect(result).toContain('3 unread');
    expect(result).toContain('150 total');
    expect(result).toContain('Sent Items');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/mailFolders'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

Folder messages mode:

```typescript
describe('folder messages mode', () => {
  it('lists messages from a specific folder', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'msg1',
            subject: 'Hello',
            from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
            receivedDateTime: '2026-01-15T10:00:00Z',
            bodyPreview: 'Hi there',
            isRead: true,
            importance: 'normal',
            hasAttachments: false,
          },
        ],
      },
    });

    const result = await executeMail('test-token', { folder: 'Inbox' });

    expect(result).toContain('Hello');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/mailFolders/Inbox/messages'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

Attachments mode:

```typescript
describe('attachments mode', () => {
  it('lists attachments for a message', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { name: 'report.pdf', contentType: 'application/pdf', size: 102400, isInline: false },
          { name: 'logo.png', contentType: 'image/png', size: 5120, isInline: true },
        ],
      },
    });

    const result = await executeMail('test-token', {
      message_id: 'msg1',
      attachments: true,
    });

    expect(result).toContain('report.pdf');
    expect(result).toContain('application/pdf');
    expect(result).toContain('logo.png');
    expect(result).toContain('inline');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/messages/msg1/attachments'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

Filter mode:

```typescript
describe('filter mode', () => {
  it('applies unread filter', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { filter: 'unread' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('$filter=isRead eq false');
  });

  it('applies flagged filter', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { filter: 'flagged' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain("$filter=flag/flagStatus eq 'flagged'");
  });

  it('applies attachments filter', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { filter: 'attachments' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('$filter=hasAttachments eq true');
  });

  it('applies important filter', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { filter: 'important' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain("$filter=importance eq 'high'");
  });
});
```

Search with ConsistencyLevel header:

```typescript
it('includes ConsistencyLevel header for search', async () => {
  mockGraphFetch.mockResolvedValue({
    ok: true,
    data: { value: [] },
  });

  await executeMail('test-token', { search: 'test' });

  expect(mockGraphFetch).toHaveBeenCalledWith(
    expect.any(String),
    'test-token',
    expect.objectContaining({
      headers: { ConsistencyLevel: 'eventual' },
    }),
  );
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=mail.test`

Expected: FAIL — new modes don't exist

### Step 3: Implement

In `src/lib/tools/mail.ts`:

- Update `mailToolDefinition.inputSchema.properties` with: `folder`, `folders`, `attachments`, `filter`
- Add interfaces: `MailFolderResponse`, `AttachmentResponse`
- Add `FILTER_MAP` constant mapping filter names to OData expressions
- Update `executeMail` signature to include new params
- Add mode routing at top of `executeMail`:
  1. `folders: true` → `executeFolderList(token)`
  2. `message_id && attachments` → `executeAttachments(token, message_id)`
  3. `message_id` → `executeDrillDown(token, message_id)` (existing)
  4. `folder` → `executeFolderMessages(token, folder, count)`
  5. `filter` → add `$filter` to the existing list path
  6. Default: existing list behavior
- For search: pass `{ timezone: false, headers: { ConsistencyLevel: 'eventual' } }` to graphFetch

### Step 4: Run tests

Run: `npm test -- --testPathPattern=mail.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/mail.ts src/__tests__/tools/mail.test.ts
git commit --no-gpg-sign -m "feat(mail): add folders, attachments, filters, and search consistency header"
```

---

## Task 5: Expand ms_calendar

Files:

- Modify: `src/lib/tools/calendar.ts`
- Modify: `src/__tests__/tools/calendar.test.ts`

### Step 1: Write tests for new calendar modes

Add to `src/__tests__/tools/calendar.test.ts`:

Calendars list:

```typescript
describe('calendars list mode', () => {
  it('lists all calendars', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { name: 'Calendar', color: 'auto', isDefaultCalendar: true, canEdit: true },
          { name: 'Birthdays', color: 'lightBlue', isDefaultCalendar: false, canEdit: false },
        ],
      },
    });

    const result = await executeCalendar('test-token', { calendars: true });

    expect(result).toContain('Calendar');
    expect(result).toContain('(default)');
    expect(result).toContain('Birthdays');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/calendars'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

Event detail:

```typescript
describe('event detail mode', () => {
  it('fetches full event detail with attendees and response status', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        subject: 'Sprint Planning',
        start: { dateTime: '2026-01-15T09:00:00', timeZone: 'Europe/London' },
        end: { dateTime: '2026-01-15T10:00:00', timeZone: 'Europe/London' },
        organizer: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
        attendees: [
          {
            emailAddress: { name: 'Bob', address: 'bob@example.com' },
            status: { response: 'accepted' },
          },
          {
            emailAddress: { name: 'Carol', address: 'carol@example.com' },
            status: { response: 'tentative' },
          },
        ],
        body: { contentType: 'html', content: '<p>Agenda here</p>' },
        location: { displayName: 'Room A' },
        onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/...' },
        showAs: 'busy',
        importance: 'normal',
        categories: ['Work'],
      },
    });

    const result = await executeCalendar('test-token', { event_id: 'evt-123' });

    expect(result).toContain('Sprint Planning');
    expect(result).toContain('Bob (accepted)');
    expect(result).toContain('Carol (tentative)');
    expect(result).toContain('Agenda here');
    expect(result).not.toContain('<p>');
    expect(result).toContain('teams.microsoft.com');
    expect(result).toContain('Room A');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/events/evt-123'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=calendar.test`

### Step 3: Implement

In `src/lib/tools/calendar.ts`:

- Add `event_id` and `calendars` to `calendarToolDefinition.inputSchema.properties`
- Add `CalendarsResponse` and `EventDetailResponse` interfaces (attendees with `status.response`)
- Add `executeCalendarsList(token)` function
- Add `executeEventDetail(token, eventId)` function — format attendees as `"Name (response)"`, strip HTML from body (reuse existing `stripHtml`), include Teams join URL from `onlineMeeting.joinUrl`
- Update `executeCalendar` to route: `calendars: true` → list, `event_id` → detail, else existing view

### Step 4: Run tests

Run: `npm test -- --testPathPattern=calendar.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/calendar.ts src/__tests__/tools/calendar.test.ts
git commit --no-gpg-sign -m "feat(calendar): add event detail drill-down and calendars list"
```

---

## Task 6: Expand ms_chat with members

Files:

- Modify: `src/lib/tools/chat.ts`
- Modify: `src/__tests__/tools/chat.test.ts`

### Step 1: Write tests for chat members mode

Add to `src/__tests__/tools/chat.test.ts`:

```typescript
describe('chat members mode', () => {
  it('lists chat members', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { displayName: 'Alice Smith', email: 'alice@example.com', roles: ['owner'] },
          { displayName: 'Bob Jones', email: 'bob@example.com', roles: ['guest'] },
        ],
      },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123', members: true });

    expect(result).toContain('Alice Smith');
    expect(result).toContain('alice@example.com');
    expect(result).toContain('owner');
    expect(result).toContain('Bob Jones');
  });

  it('does not include $select in members URL', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeChat('test-token', { chat_id: 'chat-123', members: true });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).not.toContain('$select');
    expect(calledPath).toContain('/me/chats/chat-123/members');
  });

  it('handles empty members list', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123', members: true });

    expect(result).toContain('No members found');
  });

  it('handles error fetching members', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123', members: true });

    expect(result).toContain('Error');
  });
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=chat.test`

### Step 3: Implement

In `src/lib/tools/chat.ts`:

- Add `members` boolean to `chatToolDefinition.inputSchema.properties`
- Add `ChatMember` interface and `ChatMembersResponse`
- Add routing in `executeChat`: if `args.chat_id && args.members` → `executeChatMembers(token, chatId)`
- `executeChatMembers`: GET `/me/chats/${chatId}/members` — NO `$select` (returns 400)
- Format each member: `displayName`, `email`, `roles`
- Update `executeChat` args type

### Step 4: Run tests

Run: `npm test -- --testPathPattern=chat.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/chat.ts src/__tests__/tools/chat.test.ts
git commit --no-gpg-sign -m "feat(chat): add members listing mode"
```

---

## Task 7: Expand ms_files with detail and shared

Files:

- Modify: `src/lib/tools/files.ts`
- Modify: `src/__tests__/tools/files.test.ts`

### Step 1: Write tests for file detail and shared modes

Add to `src/__tests__/tools/files.test.ts`:

File detail with download URL:

```typescript
describe('file detail mode', () => {
  it('returns file metadata with download URL', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        name: 'report.docx',
        size: 25600,
        lastModifiedDateTime: '2026-01-15T10:00:00Z',
        webUrl: 'https://onedrive.example.com/report.docx',
        '@microsoft.graph.downloadUrl': 'https://download.example.com/report.docx?token=abc',
        file: { mimeType: 'application/vnd.openxmlformats' },
      },
    });

    const result = await executeFiles('test-token', { item_id: 'item-123' });

    expect(result).toContain('report.docx');
    expect(result).toContain('25.0 KB');
    expect(result).toContain('https://download.example.com/report.docx?token=abc');
    expect(result).toContain('Download URL');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/items/item-123'),
      'test-token',
      expect.any(Object),
    );
  });

  it('handles file without download URL', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        name: 'folder',
        size: 0,
        folder: { childCount: 5 },
      },
    });

    const result = await executeFiles('test-token', { item_id: 'folder-123' });

    expect(result).toContain('folder');
    expect(result).not.toContain('Download URL');
  });
});
```

Shared files:

```typescript
describe('shared with me mode', () => {
  it('lists files shared with user', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            name: 'shared-doc.xlsx',
            size: 51200,
            lastModifiedDateTime: '2026-01-14T09:00:00Z',
            webUrl: 'https://example.com/shared-doc.xlsx',
            remoteItem: {
              shared: {
                sharedBy: { user: { displayName: 'Alice' } },
              },
            },
          },
        ],
      },
    });

    const result = await executeFiles('test-token', { shared: true });

    expect(result).toContain('shared-doc.xlsx');
    expect(result).toContain('Alice');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/drive/sharedWithMe'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=files.test`

### Step 3: Implement

In `src/lib/tools/files.ts`:

- Add `item_id` and `shared` to `filesToolDefinition.inputSchema.properties`
- Add `DriveItemDetail` interface (includes `@microsoft.graph.downloadUrl`)
- Add `SharedDriveItem` interface (includes `remoteItem.shared.sharedBy`)
- Add `executeFileDetail(token, itemId)` — GET `/me/drive/items/{itemId}`, format name/size/modified/webUrl/downloadUrl
- Add `executeSharedFiles(token, count)` — GET `/me/drive/sharedWithMe?$top={count}`, include sharer name
- Route in `executeFiles`: `item_id` → detail, `shared` → shared list, else existing

### Step 4: Run tests

Run: `npm test -- --testPathPattern=files.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/files.ts src/__tests__/tools/files.test.ts
git commit --no-gpg-sign -m "feat(files): add file detail with download URL and shared files"
```

---

## Task 8: New tool — ms_schedule

Files:

- Create: `src/lib/tools/schedule.ts`
- Create: `src/__tests__/tools/schedule.test.ts`

### Step 1: Write tests

Create `src/__tests__/tools/schedule.test.ts`:

```typescript
import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphPost =
  jest.fn<
    <TBody, TResult>(
      path: string,
      token: string,
      body: TBody,
      options?: { beta?: boolean; timezone?: boolean; headers?: Record<string, string> },
    ) => Promise<GraphResult<TResult>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphPost: mockGraphPost,
}));

const { executeSchedule } = await import('../../lib/tools/schedule.js');

describe('executeSchedule', () => {
  afterEach(() => {
    mockGraphPost.mockReset();
  });

  it('checks availability for a single person', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'alice@example.com',
            availabilityView: '000022220000',
            scheduleItems: [
              {
                subject: 'Sprint Planning',
                start: { dateTime: '2026-02-23T10:00:00' },
                end: { dateTime: '2026-02-23T12:00:00' },
                status: 'busy',
              },
            ],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['alice@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('alice@example.com');
    expect(result).toContain('free');
    expect(result).toContain('busy');
    expect(mockGraphPost).toHaveBeenCalledWith(
      '/me/calendar/getSchedule',
      'test-token',
      expect.objectContaining({
        schedules: ['alice@example.com'],
      }),
      expect.any(Object),
    );
  });

  it('checks availability for multiple people', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'alice@example.com',
            availabilityView: '0022',
            scheduleItems: [],
          },
          {
            scheduleId: 'bob@example.com',
            availabilityView: '0000',
            scheduleItems: [],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['alice@example.com', 'bob@example.com'],
      date: '2026-02-23',
      start: '09:00',
      end: '11:00',
      interval: 30,
    });

    expect(result).toContain('alice@example.com');
    expect(result).toContain('bob@example.com');
  });

  it('formats OOF and tentative statuses', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            scheduleId: 'user@example.com',
            availabilityView: '01340',
            scheduleItems: [],
          },
        ],
      },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(result).toContain('free');
    expect(result).toContain('tentative');
    expect(result).toContain('out of office');
    expect(result).toContain('working elsewhere');
  });

  it('handles error from Graph API', async () => {
    mockGraphPost.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSchedule('test-token', {
      emails: ['user@example.com'],
    });

    expect(result).toContain('Error');
  });

  it('defaults to 08:00-18:00 and 30-minute intervals', async () => {
    mockGraphPost.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSchedule('test-token', {
      emails: ['user@example.com'],
      date: '2026-02-23',
    });

    expect(mockGraphPost).toHaveBeenCalledWith(
      '/me/calendar/getSchedule',
      'test-token',
      expect.objectContaining({
        startTime: expect.objectContaining({ dateTime: '2026-02-23T08:00:00' }),
        endTime: expect.objectContaining({ dateTime: '2026-02-23T18:00:00' }),
        availabilityViewInterval: 30,
      }),
      expect.any(Object),
    );
  });
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=schedule.test`

Expected: FAIL — module doesn't exist

### Step 3: Implement

Create `src/lib/tools/schedule.ts`:

- Export `scheduleToolDefinition` with name `ms_schedule`, description about checking availability
- Input schema: `emails` (required array of strings), `date` (optional string YYYY-MM-DD), `start` (optional HH:MM, default `08:00`), `end` (optional HH:MM, default `18:00`), `interval` (optional number, default 30)
- Export `executeSchedule(token, args)`:
  - Build date from `args.date` or today's date
  - Build start/end dateTime strings: `${date}T${start}:00` and `${date}T${end}:00`
  - POST body: `{ schedules: args.emails, startTime: { dateTime, timeZone: 'UTC' }, endTime: { dateTime, timeZone: 'UTC' }, availabilityViewInterval: args.interval ?? 30 }`
  - Call `graphPost('/me/calendar/getSchedule', token, body, { timezone: false })`
  - Parse response: for each person, decode `availabilityView` string (each char = one slot)
  - Status map: `{ '0': 'free', '1': 'tentative', '2': 'busy', '3': 'out of office', '4': 'working elsewhere' }`
  - Format: time slots with their status, plus schedule items with subjects

### Step 4: Run tests

Run: `npm test -- --testPathPattern=schedule.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/schedule.ts src/__tests__/tools/schedule.test.ts
git commit --no-gpg-sign -m "feat: add ms_schedule tool for availability checking"
```

---

## Task 9: New tool — ms_sharepoint

Files:

- Create: `src/lib/tools/sharepoint.ts`
- Create: `src/__tests__/tools/sharepoint.test.ts`

### Step 1: Write tests

Create `src/__tests__/tools/sharepoint.test.ts`:

```typescript
import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean; headers?: Record<string, string> },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
}));

const { executeSharepoint } = await import('../../lib/tools/sharepoint.js');

describe('executeSharepoint', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('searches sites with default wildcard', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            displayName: 'Engineering Hub',
            webUrl: 'https://example.sharepoint.com/sites/engineering',
            description: 'Engineering team site',
            id: 'site-1',
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('Engineering Hub');
    expect(result).toContain('https://example.sharepoint.com/sites/engineering');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites?search=*'),
      'test-token',
      expect.any(Object),
    );
  });

  it('searches sites with custom query', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSharepoint('test-token', { search: 'marketing' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites?search=marketing'),
      'test-token',
      expect.any(Object),
    );
  });

  it('lists site lists', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            displayName: 'Documents',
            name: 'documents',
            description: 'Shared documents',
            webUrl: 'https://example.sharepoint.com/sites/eng/Documents',
            lastModifiedDateTime: '2026-01-15T10:00:00Z',
            list: { template: 'documentLibrary' },
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', { site_id: 'site-1' });

    expect(result).toContain('Documents');
    expect(result).toContain('documentLibrary');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites/site-1/lists'),
      'test-token',
      expect.any(Object),
    );
  });

  it('lists items from a specific list', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'item-1',
            fields: { Title: 'Q1 Report', Status: 'Published' },
          },
        ],
      },
    });

    const result = await executeSharepoint('test-token', {
      site_id: 'site-1',
      list_id: 'list-1',
    });

    expect(result).toContain('Q1 Report');
    expect(result).toContain('Published');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/sites/site-1/lists/list-1/items'),
      'test-token',
      expect.any(Object),
    );
  });

  it('handles empty results', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('No sites found');
  });

  it('handles error from Graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    const result = await executeSharepoint('test-token', {});

    expect(result).toContain('Error');
  });

  it('clamps count', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeSharepoint('test-token', { count: 100 });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );
  });
});
```

### Step 2: Run tests to verify they fail

Run: `npm test -- --testPathPattern=sharepoint.test`

### Step 3: Implement

Create `src/lib/tools/sharepoint.ts`:

- Export `sharepointToolDefinition` with name `ms_sharepoint`
- Input schema: `search`, `site_id`, `list_id`, `count` (1-50, default 10)
- Export `executeSharepoint(token, args)`:
  - `site_id && list_id` → GET `/sites/{siteId}/lists/{listId}/items?$expand=fields&$top={count}`
  - `site_id` → GET `/sites/{siteId}/lists?$top={count}&$select=displayName,name,description,webUrl,lastModifiedDateTime,list`
  - Default → GET `/sites?search={query}&$top={count}` (default query `*`)
- Format each mode's output as readable text

### Step 4: Run tests

Run: `npm test -- --testPathPattern=sharepoint.test`

Expected: All PASS

### Step 5: Commit

```bash
git add src/lib/tools/sharepoint.ts src/__tests__/tools/sharepoint.test.ts
git commit --no-gpg-sign -m "feat: add ms_sharepoint tool for sites, lists, and documents"
```

---

## Task 10: Register all tools in index.ts

Files:

- Modify: `src/index.ts`

### Step 1: Add imports for new tools

Add after the existing imports (line 14):

```typescript
import { scheduleToolDefinition, executeSchedule } from './lib/tools/schedule.js';
import { sharepointToolDefinition, executeSharepoint } from './lib/tools/sharepoint.js';
```

### Step 2: Add to ListToolsRequestSchema

Add `scheduleToolDefinition` and `sharepointToolDefinition` to the tools array (after `serverInfoToolDefinition` at line 44).

### Step 3: Add to CallToolRequestSchema switch

Add cases in the switch block:

```typescript
case 'ms_schedule':
  result = await executeSchedule(
    token,
    args as {
      emails: string[];
      date?: string;
      start?: string;
      end?: string;
      interval?: number;
    },
  );
  break;
case 'ms_sharepoint':
  result = await executeSharepoint(
    token,
    args as {
      search?: string;
      site_id?: string;
      list_id?: string;
      count?: number;
    },
  );
  break;
```

### Step 4: Update type casts for expanded existing tools

Update the `ms_profile` case to pass args:

```typescript
case 'ms_profile':
  result = await executeProfile(
    token,
    args as {
      include?: string[];
    },
  );
  break;
```

Update `ms_chat` args type:

```typescript
case 'ms_chat':
  result = await executeChat(
    token,
    args as {
      chat_id?: string;
      count?: number;
      members?: boolean;
    },
  );
  break;
```

Update `ms_mail` args type:

```typescript
case 'ms_mail':
  result = await executeMail(
    token,
    args as {
      search?: string;
      count?: number;
      message_id?: string;
      folder?: string;
      folders?: boolean;
      attachments?: boolean;
      filter?: string;
    },
  );
  break;
```

Update `ms_calendar` args type:

```typescript
case 'ms_calendar':
  result = await executeCalendar(
    token,
    args as {
      date?: string;
      start?: string;
      end?: string;
      event_id?: string;
      calendars?: boolean;
    },
  );
  break;
```

Update `ms_files` args type:

```typescript
case 'ms_files':
  result = await executeFiles(
    token,
    args as {
      path?: string;
      search?: string;
      count?: number;
      item_id?: string;
      shared?: boolean;
    },
  );
  break;
```

### Step 5: Run full test suite

Run: `npm test`

Expected: All PASS

### Step 6: Commit

```bash
git add src/index.ts
git commit --no-gpg-sign -m "feat: register ms_schedule and ms_sharepoint, update expanded tool args"
```

---

## Task 11: Update server-info and version bump

Files:

- Modify: `src/lib/tools/server-info.ts`
- Modify: `src/index.ts:33`
- Modify: `package.json:3`

### Step 1: Update server-info TOOL_NAMES

In `src/lib/tools/server-info.ts`, replace the `TOOL_NAMES` array (lines 15-24):

```typescript
const TOOL_NAMES = [
  'ms_auth_status',
  'ms_profile',
  'ms_calendar',
  'ms_mail',
  'ms_chat',
  'ms_files',
  'ms_transcripts',
  'ms_schedule',
  'ms_sharepoint',
  'ms_server_info',
];
```

### Step 2: Version bump

In `package.json` line 3: change `"version": "0.6.0"` to `"version": "0.7.0"`

In `src/index.ts` line 33: change `version: '0.6.0'` to `version: '0.7.0'`

### Step 3: Run server-info tests

Run: `npm test -- --testPathPattern=server-info.test`

Expected: PASS (or update expected tool count if hardcoded in tests)

### Step 4: Run full suite

Run: `npm test`

Expected: All PASS

### Step 5: Commit

```bash
git add package.json src/index.ts src/lib/tools/server-info.ts
git commit --no-gpg-sign -m "chore: bump version to 0.7.0, update server-info tool list"
```

---

## Task 12: Final verification

### Step 1: Run full test suite with coverage

Run: `npm run test:coverage`

Expected: All tests pass, 80%+ coverage on `src/lib/**/*.ts`

### Step 2: Run lint

Run: `npm run lint`

Expected: No errors

### Step 3: Run build

Run: `npm run build`

Expected: Clean compile, no TypeScript errors

### Step 4: Verify format

Run: `npm run format:check`

Expected: All files formatted

### Step 5: If any issues, fix and commit

If lint/format issues:

```bash
npm run lint:fix && npm run format
git add -A
git commit --no-gpg-sign -m "chore: fix lint and formatting"
```

---

## API Quirks Quick Reference

Keep this handy while implementing — these are the things that will 400 if you get them wrong:

| Endpoint                  | Quirk                                        |
| ------------------------- | -------------------------------------------- |
| `/me/chats`               | NO `$orderby`                                |
| `/me/chats/{id}/messages` | NO `$select`                                 |
| `/me/chats/{id}/members`  | NO `$select`                                 |
| `/me/memberOf`            | `displayName` can be null                    |
| `/me/manager`             | 403 for contractors (not a bug)              |
| Mail `$search`            | Requires `ConsistencyLevel: eventual` header |
| `/me/drive/items/{id}`    | Returns `@microsoft.graph.downloadUrl`       |
| `getSchedule`             | Only POST endpoint                           |
