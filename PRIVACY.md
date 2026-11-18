# Privacy Policy

**Last updated:** 18 November 2026

`@masonator/m365-mcp` is a Model Context Protocol server that you run yourself, as
a local process started by your MCP client. There is no hosted service behind it
and no account to create.

This policy describes what the software does with data. It is written against the
code, and the "Where the code does this" notes point at the file that implements
each claim.

## What is collected about you

Nothing.

The server contains no telemetry, analytics, crash reporting, usage counting or
phone-home of any kind. The author receives no data from your installation,
including the fact that you installed it.

## Your Microsoft 365 data

The server signs in as you, with delegated permissions, and reads what you ask it
to read: mail, calendar, chats, files, transcripts and the rest of the
[tool reference](docs/tools.md). It holds none of it. Each tool call fetches from
Microsoft Graph, formats the result and returns it to your MCP client. Nothing
you read is cached or written to disk.

## Outbound network connections

The server talks to Microsoft and to nothing else.

1. **`login.microsoftonline.com`.** Sign-in and token refresh for the tenant you
   set in `MS365_MCP_TENANT_ID`. Your client ID travels with these requests, and
   your client secret if you set one.
2. **`graph.microsoft.com`.** Every data tool calls Microsoft Graph, on the `v1.0`
   and `beta` endpoints, with your access token.
3. **`workiq.svc.cloud.microsoft`, only if you enable it.** `ms_workiq` is off
   unless `MS365_MCP_WORKIQ_SCOPE` is set. When it is on, the questions and
   arguments you pass to that tool go to Microsoft's Work IQ service.

During sign-in the server also listens on `localhost` for the redirect from
Microsoft, and closes that listener when sign-in ends.

There are no other outbound connections. No third-party APIs, no update checks
and no analytics endpoints.

Where the code does this: `src/lib/auth.ts`, `src/lib/graph.ts`,
`src/lib/workiq.ts`, `src/lib/tools/transcripts.ts`.

## Stored data

One file. After sign-in the server writes your access token, refresh token, token
expiry and granted scopes to `~/.config/m365-mcp/tokens.json`, or under
`XDG_CONFIG_HOME` if you set it, with `0600` permissions. `MS365_MCP_TOKEN_FILE`
changes the filename. Deleting the file signs you out.

Your client ID, tenant ID and client secret are read from environment variables
and are not written to disk by the server. `ms_server_info` reports whether each
is set, never its value.

Where the code does this: `src/lib/auth.ts`, `src/lib/tools/server-info.ts`.

## Data that reaches your MCP client

Tool results are returned to whichever MCP client you connected, such as Claude
Desktop or Claude Code. Those results contain whatever you asked for, which can
include the content of mail, chats, meetings and documents, and the names and
addresses of the people in them.

What that client then does with the data is governed by that client's own privacy
policy, not this one. If you would rather a category of data never reached the
model, do not grant the permission that reads it: a tool whose permission is
missing says so and makes no Graph call.

That holds only while `ms_workiq` is off. Work IQ reads with your own Microsoft 365
permissions, not the app registration's, so it can return data the Graph tools are
refused. Leave `MS365_MCP_WORKIQ_SCOPE` unset if you are relying on withheld
permissions.

## Logs

The server writes diagnostics to standard error on the machine running it: the
sign-in URL when a browser cannot be opened, and the error body Microsoft
returned when a Graph call or a token refresh fails. Those error bodies can carry
identifiers from your tenant. The server does not add tokens or the client secret
to them.

Standard error is not transmitted anywhere. Where it ends up is decided by your
MCP client.

Where the code does this: `src/lib/graph.ts`, `src/lib/auth.ts`.

## Retention and deletion

The author holds no data about you, so there is nothing to retain and nothing to
delete. The token file lives on your machine and is removed when you delete it.
Revoke the app's access from your Microsoft account, or delete the app
registration, to invalidate the tokens themselves.

## Third parties

Microsoft, as the service you are reading from. Nobody else. Your data is not
shared with, sold to, or processed by anyone on the author's behalf.

## Children

This is developer tooling and is not directed at children.

## Changes

Material changes to this policy are recorded in [CHANGELOG.md](CHANGELOG.md)
alongside the release that makes them.

## Contact

Open an issue at <https://github.com/StuMason/m365-mcp/issues>, or email
<hey@stumason.dev>.
