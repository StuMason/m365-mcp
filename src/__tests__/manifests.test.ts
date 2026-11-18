import { readFileSync } from 'node:fs';

/**
 * Guards against the distribution manifests drifting from package.json. The
 * registry rejects a server.json whose version or name disagrees with the npm
 * package, and it does so at publish time, after npm has already shipped.
 * scripts/sync-manifests.mjs (npm `version` lifecycle) keeps them aligned; this
 * fails the build if they diverge anyway.
 */
const read = (file: string): Record<string, unknown> =>
  JSON.parse(readFileSync(file, 'utf8')) as Record<string, unknown>;

describe('distribution manifests', () => {
  const pkg = read('package.json') as unknown as { version: string; mcpName: string };

  it('server.json (MCP registry) matches package.json', () => {
    const serverJson = read('server.json') as unknown as {
      name: string;
      description: string;
      version: string;
      packages: Array<{ identifier: string; version: string }>;
    };
    expect(serverJson.version).toBe(pkg.version);
    expect(serverJson.name).toBe(pkg.mcpName);
    // The registry schema caps the description at 100 characters.
    expect(serverJson.description.length).toBeLessThanOrEqual(100);
    for (const p of serverJson.packages) {
      expect(p.version).toBe(pkg.version);
    }
  });

  it('manifest.json (MCPB bundle) matches package.json', () => {
    const manifest = read('manifest.json') as unknown as { version: string; server: unknown };
    expect(manifest.version).toBe(pkg.version);
    expect(manifest.server).toBeDefined();
  });
});
