import { mkdirSync, readFileSync, writeFileSync, unlinkSync, chmodSync } from 'node:fs';
import { join } from 'node:path';
import { homedir } from 'node:os';
import type { TokenData, AuthConfig } from '../types/tokens.js';

const TOKEN_FILENAME = 'tokens.json';
const EXPIRY_BUFFER_MS = 120_000; // 2 minutes

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
  writeFileSync(filePath, JSON.stringify(tokens, null, 2), 'utf-8');
  chmodSync(filePath, 0o600);
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
