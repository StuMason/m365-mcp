import { randomUUID } from 'node:crypto';
import { loadTokens, saveTokens } from './auth.js';
import type { AuthConfig } from '../types/tokens.js';

/**
 * Client for the Work IQ remote MCP server.
 *
 * Work IQ is a separate OAuth resource from Graph, so it needs its own access
 * token: Azure will not issue one token covering both. Both tokens come from the
 * same sign-in, by redeeming the stored refresh token twice.
 *
 * Its `fetch` tool proxies Graph using the signed-in user's own M365 permissions
 * rather than the app registration's consented scopes, so it reads paths the Graph
 * token is refused — which is why it is worth having at all.
 */

const MCP_ENDPOINT = 'https://workiq.svc.cloud.microsoft/mcp';
const PROTOCOL_VERSION = '2025-06-18';
const EXPIRY_BUFFER_MS = 120_000;

/** Tools that only read. Everything else can change or send something. */
export const READ_ONLY_TOOLS = ['fetch', 'fetch_blob', 'get_schema', 'search_paths', 'list_agents'];

/**
 * Returns the configured Work IQ scope, or null when Work IQ is not enabled.
 *
 * Opt-in is deliberate. The scope carries the server's whole tool surface,
 * writes included, so enabling it has to be a decision the operator makes rather
 * than a capability that arrives with an upgrade.
 */
export function workIqScope(): string | null {
  return process.env['MS365_MCP_WORKIQ_SCOPE']?.trim() || null;
}

let cached: { token: string; expiresAt: number } | null = null;

/** Discards the cached token. Exported for tests. */
export function resetWorkIqToken(): void {
  cached = null;
}

/**
 * Redeems the stored refresh token against the Work IQ resource.
 *
 * The refresh grant rotates the refresh token, and the rotated one is what the
 * Graph side must use from then on — so it is written back to the token file.
 * Dropping it strands the Graph session on a token Azure has already retired.
 */
export async function getWorkIqToken(config: AuthConfig): Promise<string> {
  const scope = workIqScope();
  if (!scope) {
    throw new Error('Work IQ is not enabled. Set MS365_MCP_WORKIQ_SCOPE to use it.');
  }

  if (cached && cached.expiresAt > Date.now() + EXPIRY_BUFFER_MS) {
    return cached.token;
  }

  const stored = loadTokens();
  if (!stored?.refresh_token) {
    throw new Error('Not signed in. Run ms_auth_status first.');
  }

  const params: Record<string, string> = {
    client_id: config.clientId,
    grant_type: 'refresh_token',
    refresh_token: stored.refresh_token,
    scope,
  };
  if (config.clientSecret) {
    params['client_secret'] = config.clientSecret;
  }

  const response = await fetch(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(params).toString(),
    },
  );

  const text = await response.text();
  if (!response.ok) {
    // AADSTS65001 here means the registration has no Work IQ consent, which is a
    // configuration answer rather than a fault.
    if (text.includes('AADSTS65001')) {
      throw new Error(
        'This app registration has not been granted the Work IQ scope, so Work IQ is unavailable. ' +
          'The Graph tools are unaffected.',
      );
    }
    throw new Error(`Could not get a Work IQ token (${response.status}).`);
  }

  const data = JSON.parse(text) as {
    access_token: string;
    refresh_token?: string;
    expires_in: number;
  };

  if (data.refresh_token && data.refresh_token !== stored.refresh_token) {
    saveTokens({ ...stored, refresh_token: data.refresh_token });
  }

  cached = {
    token: data.access_token,
    expiresAt: Date.now() + data.expires_in * 1000,
  };
  return cached.token;
}

interface RpcResult {
  [key: string]: unknown;
}

let sessionId: string | null = null;
let initialised = false;

/** Drops the MCP session. Exported for tests. */
export function resetWorkIqSession(): void {
  sessionId = null;
  initialised = false;
}

/**
 * One JSON-RPC call against the Work IQ MCP server.
 *
 * The transport is streamable HTTP, so a reply arrives either as plain JSON or as
 * an SSE stream that has to be reassembled from its `data:` lines.
 */
async function rpc(
  token: string,
  method: string,
  params?: Record<string, unknown>,
  notify = false,
): Promise<RpcResult> {
  const body: Record<string, unknown> = { jsonrpc: '2.0', method };
  if (params) body['params'] = params;
  if (!notify) body['id'] = randomUUID();

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
    Accept: 'application/json, text/event-stream',
  };
  if (sessionId) headers['Mcp-Session-Id'] = sessionId;

  const response = await fetch(MCP_ENDPOINT, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });

  const returnedSession = response.headers.get('mcp-session-id');
  if (returnedSession) sessionId = returnedSession;

  if (notify) return {};

  const text = await response.text();
  if (!response.ok) {
    throw new Error(`Work IQ returned HTTP ${response.status}.`);
  }

  const payload = /^(event:|data:)/.test(text)
    ? text
        .split('\n')
        .filter((line) => line.startsWith('data:'))
        .map((line) => line.slice(5).trim())
        .join('')
    : text;

  const parsed = JSON.parse(payload) as {
    result?: RpcResult;
    error?: { message?: string; code?: number };
  };
  if (parsed.error) {
    throw new Error(parsed.error.message || `Work IQ error ${parsed.error.code}`);
  }
  return parsed.result ?? {};
}

/** Performs the MCP handshake once per process. */
async function ensureInitialised(token: string): Promise<void> {
  if (initialised) return;
  await rpc(token, 'initialize', {
    protocolVersion: PROTOCOL_VERSION,
    capabilities: {},
    clientInfo: { name: 'm365-mcp', version: '1' },
  });
  await rpc(token, 'notifications/initialized', undefined, true);
  initialised = true;
}

export interface WorkIqTool {
  name: string;
  description?: string;
  inputSchema?: unknown;
}

/** Lists the tools the Work IQ server exposes. */
export async function listWorkIqTools(config: AuthConfig): Promise<WorkIqTool[]> {
  const token = await getWorkIqToken(config);
  await ensureInitialised(token);
  const result = await rpc(token, 'tools/list', {});
  return (result['tools'] as WorkIqTool[]) ?? [];
}

export interface WorkIqCall {
  content?: Array<{ type?: string; text?: string }>;
  structuredContent?: Record<string, unknown>;
}

/** Calls one Work IQ tool by name. */
export async function callWorkIqTool(
  config: AuthConfig,
  name: string,
  args: Record<string, unknown>,
): Promise<WorkIqCall> {
  const token = await getWorkIqToken(config);
  await ensureInitialised(token);
  return (await rpc(token, 'tools/call', { name, arguments: args })) as WorkIqCall;
}
