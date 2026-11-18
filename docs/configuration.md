# Configuration

The server is configured through environment variables. Only the first two are
required.

| Variable                  | Required | Description                                                                    |
| ------------------------- | -------- | ------------------------------------------------------------------------------ |
| `MS365_MCP_CLIENT_ID`     | Yes      | Azure AD application (client) ID                                               |
| `MS365_MCP_TENANT_ID`     | Yes      | Azure AD tenant ID                                                             |
| `MS365_MCP_CLIENT_SECRET` | No       | Azure AD client secret (confidential clients only)                             |
| `MS365_MCP_TIMEZONE`      | No       | Timezone for calendar (default: system timezone)                               |
| `MS365_MCP_REDIRECT_URL`  | No       | OAuth redirect URI (default: dynamic port, `http://localhost:{port}/callback`) |
| `MS365_MCP_CLIENT_TYPE`   | No       | Set to `spa` only for a Single-Page Application registration (see below)       |
| `MS365_MCP_SCOPES`        | No       | Override the requested scopes, e.g. `https://graph.microsoft.com/.default`     |
| `MS365_MCP_TOKEN_FILE`    | No       | Token filename within the config directory (default: `tokens.json`)            |
| `MS365_MCP_WORKIQ_SCOPE`  | No       | Enables `ms_workiq`. Off by default because it is not read-only (see below)    |

## `MS365_MCP_CLIENT_TYPE`

Leave this unset unless your app registration uses the **Single-Page Application**
platform. Only SPA registrations may redeem a token cross-origin; every other
platform — Web, and "Mobile and desktop applications" — rejects the `Origin` header
with `AADSTS9002326`. The absence of a client secret is not a reliable signal here,
because a Mobile-and-desktop registration has no secret either.

## `MS365_MCP_SCOPES`

By default the server requests an explicit scope list, so a first-time consent
prompt shows exactly what is being asked for.

Set this to `https://graph.microsoft.com/.default offline_access` if your
registration has more permissions configured than consented. Naming an unconsented
scope fails the entire request with `AADSTS65001`, whereas `.default` asks for
whatever is already consented and adapts on its own as that changes.

Whatever the token ends up carrying, the tools adapt: a tool whose permission is
missing explains what is unavailable and what still works, rather than returning a
bare `403`. `ms_auth_status` lists exactly what was granted.

## `MS365_MCP_TOKEN_FILE`

Use this when running more than one app registration. One token file holds one
registration's tokens, so signing in with a second client would otherwise silently
evict the first.
