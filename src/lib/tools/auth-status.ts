import type { AuthConfig, TokenData } from '../../types/tokens.js';
import { loadTokens, isTokenExpired, startAuthFlow, SCOPES } from '../auth.js';
import { refreshAccessToken } from '../auth.js';
import { graphFetch } from '../graph.js';

export const authStatusToolDefinition = {
  name: 'ms_auth_status',
  description: 'Check Microsoft 365 connection status. If not connected, opens browser to sign in.',
  inputSchema: {
    type: 'object' as const,
    properties: {},
  },
};

interface ProfileResponse {
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
}

/**
 * Fetches a short profile summary for the status display.
 * Returns null if the fetch fails.
 */
async function fetchProfileSummary(
  token: string,
): Promise<{ displayName: string; email: string } | null> {
  const result = await graphFetch<ProfileResponse>('/me', token);
  if (!result.ok) {
    return null;
  }
  return {
    displayName: result.data.displayName || 'Unknown',
    email: result.data.mail || result.data.userPrincipalName || 'Unknown',
  };
}

/**
 * Formats a "just signed in" status message.
 */
function formatJustConnected(profile: { displayName: string; email: string } | null): string {
  const lines = ['Status: Connected \u2713 (just signed in)'];
  if (profile) {
    lines.push(`User: ${profile.displayName} (${profile.email})`);
  }
  return lines.join('\n');
}

/**
 * Formats a full connected status message with token details.
 */
function formatConnectedStatus(
  tokens: TokenData,
  profile: { displayName: string; email: string } | null,
): string {
  const lines = ['Status: Connected \u2713'];
  if (profile) {
    lines.push(`User: ${profile.displayName} (${profile.email})`);
  }
  lines.push(`Token expires: ${tokens.expires_at}`);
  lines.push(`Scopes: ${tokens.scopes || SCOPES.join(' ')}`);
  return lines.join('\n');
}

/**
 * Formats an error status message.
 */
function formatError(message: string): string {
  return [
    'Status: Not connected',
    `Error: ${message}`,
    'Action: Run ms_auth_status again to sign in.',
  ].join('\n');
}

/**
 * Check Microsoft 365 connection status.
 * Handles the full lifecycle: no tokens, expired tokens, valid tokens.
 * Triggers auth flow if not connected.
 */
export async function executeAuthStatus(config: AuthConfig): Promise<string> {
  try {
    const tokens = loadTokens();

    // No tokens — start auth flow
    if (!tokens) {
      process.stderr.write('Not connected. Starting auth flow...\n');
      const newTokens = await startAuthFlow(config);
      const profile = await fetchProfileSummary(newTokens.access_token);
      return formatJustConnected(profile);
    }

    // Tokens exist but expired — try refresh
    if (isTokenExpired(tokens)) {
      const refreshed = await refreshAccessToken(config, tokens.refresh_token);
      if (!refreshed) {
        return formatError('Token expired and refresh failed. Please sign in again.');
      }
      const profile = await fetchProfileSummary(refreshed.access_token);
      return formatConnectedStatus(refreshed, profile);
    }

    // Tokens valid — show status
    const profile = await fetchProfileSummary(tokens.access_token);
    return formatConnectedStatus(tokens, profile);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return formatError(message);
  }
}
