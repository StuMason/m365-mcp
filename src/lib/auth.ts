import { mkdirSync, readFileSync, writeFileSync, unlinkSync } from 'node:fs';
import { join } from 'node:path';
import { homedir } from 'node:os';
import { createServer } from 'node:net';
import { createServer as createHttpServer, type Server } from 'node:http';
import { URL } from 'node:url';
import { randomBytes } from 'node:crypto';
import { execFile } from 'node:child_process';
import type { TokenData, AuthConfig } from '../types/tokens.js';

const TOKEN_FILENAME = 'tokens.json';
const EXPIRY_BUFFER_MS = 120_000; // 2 minutes
const AUTH_TIMEOUT_MS = 300_000; // 5 minutes

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#x27;');
}

export const SCOPES = [
  'openid',
  'profile',
  'email',
  'offline_access',
  'User.Read',
  'Calendars.Read',
  'Mail.Read',
  'Chat.Read',
  'Files.Read',
  'OnlineMeetingTranscript.Read.All',
  'Sites.Read.All',
];

/**
 * Returns the config directory for m365-mcp.
 * Respects XDG_CONFIG_HOME, falls back to ~/.config/m365-mcp.
 * Creates the directory (recursively) if it doesn't exist.
 */
export function getConfigDir(): string {
  const base = process.env['XDG_CONFIG_HOME'] || join(homedir(), '.config');
  const configDir = join(base, 'm365-mcp');
  mkdirSync(configDir, { recursive: true });
  return configDir;
}

/**
 * Loads tokens from tokens.json in the given config directory.
 * Returns null if the file doesn't exist or contains invalid JSON.
 */
export function loadTokens(configDir?: string): TokenData | null {
  const dir = configDir ?? getConfigDir();
  const filePath = join(dir, TOKEN_FILENAME);
  try {
    const raw = readFileSync(filePath, 'utf-8');
    return JSON.parse(raw) as TokenData;
  } catch {
    return null;
  }
}

/**
 * Saves tokens to tokens.json in the given config directory.
 * Sets file permissions to 0o600 (user read/write only).
 */
export function saveTokens(tokens: TokenData, configDir?: string): void {
  const dir = configDir ?? getConfigDir();
  const filePath = join(dir, TOKEN_FILENAME);
  writeFileSync(filePath, JSON.stringify(tokens, null, 2), { mode: 0o600, encoding: 'utf-8' });
}

/**
 * Deletes tokens.json from the given config directory.
 * Does nothing if the file doesn't exist.
 */
export function deleteTokens(configDir?: string): void {
  const dir = configDir ?? getConfigDir();
  const filePath = join(dir, TOKEN_FILENAME);
  try {
    unlinkSync(filePath);
  } catch {
    // File doesn't exist — nothing to do
  }
}

/**
 * Returns true if the token expires within 2 minutes (120 000 ms safety buffer).
 */
export function isTokenExpired(tokens: TokenData): boolean {
  const expiresAt = new Date(tokens.expires_at).getTime();
  return expiresAt < Date.now() + EXPIRY_BUFFER_MS;
}

/**
 * Loads auth configuration from environment variables.
 * Throws with a clear message if any required variable is missing.
 */
export function loadAuthConfig(): AuthConfig {
  const clientId = process.env['MS365_MCP_CLIENT_ID'];
  const clientSecret = process.env['MS365_MCP_CLIENT_SECRET'];
  const tenantId = process.env['MS365_MCP_TENANT_ID'];

  const missing: string[] = [];
  if (!clientId) missing.push('MS365_MCP_CLIENT_ID');
  if (!clientSecret) missing.push('MS365_MCP_CLIENT_SECRET');
  if (!tenantId) missing.push('MS365_MCP_TENANT_ID');

  if (missing.length > 0) {
    throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
  }

  return {
    clientId: clientId!,
    clientSecret: clientSecret!,
    tenantId: tenantId!,
  };
}

/**
 * Finds an available port by binding to port 0 and reading the assigned port.
 * @internal Exported for testing only.
 */
export function findAvailablePort(): Promise<number> {
  return new Promise((resolve, reject) => {
    const srv = createServer();
    srv.listen(0, () => {
      const addr = srv.address();
      if (addr && typeof addr === 'object') {
        const port = addr.port;
        srv.close(() => resolve(port));
      } else {
        /* istanbul ignore next -- defensive branch never reached with port 0 */
        srv.close(() => reject(new Error('Could not determine port')));
      }
    });
    srv.on('error', reject);
  });
}

/**
 * Opens a URL in the default browser.
 * Falls back to printing the URL to stderr if the browser cannot be opened.
 * @internal Exported for testing only.
 */
export function openBrowser(url: string): void {
  /* istanbul ignore next -- platform-specific browser launch */
  const platform = process.platform;
  try {
    /* istanbul ignore next */
    if (platform === 'darwin') {
      execFile('open', [url]);
    } else if (platform === 'win32') {
      execFile('cmd', ['/c', 'start', '', url]);
    } else {
      execFile('xdg-open', [url]);
    }
  } catch {
    process.stderr.write(`Could not open browser. Please visit:\n${url}\n`);
  }
}

/**
 * Exchanges an authorization code for tokens via Azure AD token endpoint.
 * @internal Exported for testing only.
 */
export async function exchangeCodeForTokens(
  config: AuthConfig,
  code: string,
  redirectUri: string,
): Promise<TokenData> {
  const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'authorization_code',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    code,
    redirect_uri: redirectUri,
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Token exchange failed (${response.status}): ${errorText}`);
  }

  const data = (await response.json()) as {
    access_token: string;
    refresh_token: string;
    expires_in: number;
    scope: string;
  };

  return {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    expires_at: new Date(Date.now() + data.expires_in * 1000).toISOString(),
    scopes: data.scope,
  };
}

/**
 * Starts an HTTP server on the given port and waits for the OAuth callback.
 * Returns the authorization code from the callback query string.
 * @internal Exported for testing only.
 */
export function waitForAuthCallback(
  port: number,
  expectedState: string,
  timeoutMs: number = AUTH_TIMEOUT_MS,
): { promise: Promise<string>; server: Server } {
  let httpServer!: Server;

  const promise = new Promise<string>((resolve, reject) => {
    const closeServer = (): void => {
      httpServer.close();
      httpServer.closeAllConnections();
    };

    const timeout = setTimeout(() => {
      closeServer();
      reject(new Error('Authentication timed out after 5 minutes. Please try again.'));
    }, timeoutMs);

    httpServer = createHttpServer((req, res) => {
      if (!req.url?.startsWith('/callback')) {
        res.writeHead(404);
        res.end('Not found');
        return;
      }

      const url = new URL(req.url, `http://localhost:${port}`);
      const callbackState = url.searchParams.get('state');
      const callbackCode = url.searchParams.get('code');
      const error = url.searchParams.get('error');

      if (error) {
        const errorDesc = url.searchParams.get('error_description') || error;
        const safeDesc = escapeHtml(errorDesc);
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(
          `<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:40px"><h1>Sign-in failed</h1><p>${safeDesc}</p></body></html>`,
        );
        clearTimeout(timeout);
        closeServer();
        reject(new Error(`Authentication failed: ${errorDesc}`));
        return;
      }

      if (callbackState !== expectedState) {
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(
          '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:40px"><h1>Error</h1><p>State mismatch — possible CSRF attack.</p></body></html>',
        );
        clearTimeout(timeout);
        closeServer();
        reject(new Error('State mismatch in OAuth callback'));
        return;
      }

      if (!callbackCode) {
        res.writeHead(400, { 'Content-Type': 'text/html' });
        res.end(
          '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:40px"><h1>Error</h1><p>No authorization code received.</p></body></html>',
        );
        clearTimeout(timeout);
        closeServer();
        reject(new Error('No authorization code in callback'));
        return;
      }

      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(
        '<!DOCTYPE html><html><body style="font-family:sans-serif;text-align:center;padding:40px"><h1>Signed in!</h1><p>You can close this tab.</p></body></html>',
      );
      clearTimeout(timeout);
      closeServer();
      resolve(callbackCode);
    });

    httpServer.listen(port);
  });

  return { promise, server: httpServer };
}

/**
 * Browser-popup + localhost-callback OAuth2 flow.
 * Opens the browser for Microsoft 365 sign-in, listens for the callback,
 * exchanges the authorization code for tokens, saves them, and returns the TokenData.
 */
export async function startAuthFlow(config: AuthConfig): Promise<TokenData> {
  const port = await findAvailablePort();
  const redirectUri = `http://localhost:${port}/callback`;
  const state = randomBytes(16).toString('hex');

  const authUrl = new URL(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize`,
  );
  authUrl.searchParams.set('client_id', config.clientId);
  authUrl.searchParams.set('response_type', 'code');
  authUrl.searchParams.set('redirect_uri', redirectUri);
  authUrl.searchParams.set('scope', SCOPES.join(' '));
  authUrl.searchParams.set('state', state);

  process.stderr.write('Opening browser for Microsoft 365 sign-in...\n');
  openBrowser(authUrl.toString());

  const { promise } = waitForAuthCallback(port, state);
  const code = await promise;

  const tokenData = await exchangeCodeForTokens(config, code, redirectUri);
  saveTokens(tokenData);
  return tokenData;
}

/**
 * Refreshes an access token using a refresh token.
 * On success, saves and returns the new TokenData.
 * On failure, deletes stored tokens and returns null.
 */
export async function refreshAccessToken(
  config: AuthConfig,
  refreshToken: string,
): Promise<TokenData | null> {
  const tokenUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: config.clientId,
    client_secret: config.clientSecret,
    refresh_token: refreshToken,
  });

  try {
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body.toString(),
    });

    if (!response.ok) {
      const errorText = await response.text();
      process.stderr.write(`Token refresh failed (${response.status}): ${errorText}\n`);
      deleteTokens();
      return null;
    }

    const data = (await response.json()) as {
      access_token: string;
      refresh_token: string;
      expires_in: number;
      scope: string;
    };

    const tokenData: TokenData = {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      expires_at: new Date(Date.now() + data.expires_in * 1000).toISOString(),
      scopes: data.scope,
    };

    saveTokens(tokenData);
    return tokenData;
  } catch (err) {
    process.stderr.write(`Token refresh error: ${err instanceof Error ? err.message : err}\n`);
    deleteTokens();
    return null;
  }
}

/**
 * Main entry point — returns a valid access token.
 * Loads cached tokens, refreshes if expired, or starts a new auth flow if needed.
 */
export async function getAccessToken(config: AuthConfig): Promise<string> {
  const tokens = loadTokens();

  if (!tokens) {
    const newTokens = await startAuthFlow(config);
    return newTokens.access_token;
  }

  if (!isTokenExpired(tokens)) {
    return tokens.access_token;
  }

  const refreshed = await refreshAccessToken(config, tokens.refresh_token);
  if (refreshed) {
    return refreshed.access_token;
  }

  const newTokens = await startAuthFlow(config);
  return newTokens.access_token;
}
