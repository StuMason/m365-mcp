# OAuth Flow: SPA Platform, PKCE, and Configurable Redirect URI

**Date:** 2026-02-20
**Commit:** 20bd424

## What changed

- OAuth redirect URI is now configurable via `MS365_MCP_REDIRECT_URL` env var (default: dynamic port with `/callback`)
- Added PKCE (S256 code challenge) to the authorization code flow
- Made `MS365_MCP_CLIENT_SECRET` optional to support both public and confidential Azure AD clients
- Added `Origin` header to token exchange and refresh requests for SPA-registered redirect URIs

## Why

During live testing with Claude Desktop, the auth flow hit three sequential Azure AD errors:

1. **AADSTS9002325** — Azure AD now requires PKCE for authorization code redemption. The original flow used a confidential client with `client_secret` only, no PKCE.
2. **AADSTS700025** — The test Azure AD app was registered as a public client (SPA platform), which rejects `client_secret` in token requests.
3. **AADSTS9002327** — SPA-platform redirect URIs require an `Origin` header in the token exchange POST, proving the request is cross-origin.

These three fixes make the server work with any Azure AD app registration (public SPA, public mobile/desktop, or confidential web).

## Decisions made

- **Configurable redirect URI via single env var**: Rather than hardcoding a port or requiring multiple config vars (`REDIRECT_PORT`, `REDIRECT_PATH`), a single `MS365_MCP_REDIRECT_URL` overrides everything. The server parses port and path from the URL. Without it, the original dynamic-port behaviour is preserved.
- **PKCE always on**: PKCE is sent on every auth flow regardless of client type. It's required for SPA clients and harmless for confidential clients. No config flag needed.
- **client_secret optional**: With PKCE in place, public clients don't need a secret. The server includes `client_secret` only when `MS365_MCP_CLIENT_SECRET` is set. This keeps the server agnostic — users choose their Azure AD app type.
- **Origin header derived from redirect URI**: Rather than a separate config, the `Origin` header is derived from `MS365_MCP_REDIRECT_URL` when set. This satisfies Azure AD's SPA cross-origin requirement without exposing internal concerns to the user.

## Rejected alternatives

- **Hardcoded fixed port (19284)**: Initially implemented a fixed port, but this was too opinionated. The configurable approach via env var is more flexible while keeping the dynamic-port default for zero-config setups.
- **Separate `MS365_MCP_REDIRECT_PORT` env var**: Considered alongside `MS365_MCP_REDIRECT_URL`, but two vars for the same concern is redundant. A single URL is cleaner.
- **Requiring users to choose "Web" platform in Azure AD**: Would avoid the SPA/Origin complexity, but dictating Azure AD configuration makes the server less agnostic.

## Context

- Azure AD SPA platform was chosen in the test app registration because the redirect URI was added under "Single-page application" rather than "Web" in the Azure portal. Both work now.
- The `findAvailablePort()` function is retained for the default (no `MS365_MCP_REDIRECT_URL`) path. It binds to port 0 and reads the assigned port.
- Token refresh also includes the `Origin` header when `MS365_MCP_REDIRECT_URL` is set, since Azure AD applies the same SPA rules to refresh requests.
