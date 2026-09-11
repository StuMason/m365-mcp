import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * Reads the package version from package.json.
 * Single source of truth: the MCP handshake and ms_server_info both report this,
 * so a release bump cannot drift from what the server advertises on the wire.
 */
export function getVersion(): string {
  try {
    const dir = dirname(fileURLToPath(import.meta.url));
    const pkgPath = join(dir, '..', '..', 'package.json');
    const pkg = JSON.parse(readFileSync(pkgPath, 'utf-8')) as { version?: string };
    return pkg.version ?? 'unknown';
  } catch {
    return 'unknown';
  }
}

export interface BuildInfo {
  builtAt: string;
  commit?: string;
  branch?: string;
  dirty?: boolean;
}

/**
 * Reads the build fingerprint written by scripts/build-info.mjs.
 * Returns null when running from source, or from a package built without git.
 */
export function getBuildInfo(): BuildInfo | null {
  try {
    const dir = dirname(fileURLToPath(import.meta.url));
    return JSON.parse(readFileSync(join(dir, '..', 'build-info.json'), 'utf-8')) as BuildInfo;
  } catch {
    return null;
  }
}
