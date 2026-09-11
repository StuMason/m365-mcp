import { readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { TOOL_DEFINITIONS, toolNames } from '../../lib/tools/index.js';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..', '..', '..');

/**
 * 0.7.0 shipped three hand-maintained lists of tools that disagreed with each
 * other — index.ts registered ten, server-info claimed ten, the README documented
 * seven. These tests make that class of drift a test failure rather than a
 * release note.
 */
describe('tool roster', () => {
  it('every definition is well formed', () => {
    for (const def of TOOL_DEFINITIONS) {
      expect(def.name).toMatch(/^ms_[a-z_]+$/);
      expect(def.title).toBeTruthy();
      expect(def.description.length).toBeGreaterThan(20);
      expect(def.inputSchema).toBeDefined();
    }
  });

  it('has no duplicate names', () => {
    const names = toolNames();
    expect(new Set(names).size).toBe(names.length);
  });

  it('declares every tool read-only and non-destructive', () => {
    // The whole server is read-only. A tool that mutates would need a deliberate
    // change here, not a silently different annotation.
    for (const def of TOOL_DEFINITIONS) {
      expect(def.annotations.readOnlyHint).toBe(true);
      expect(def.annotations.destructiveHint).toBe(false);
      expect(def.annotations.title).toBe(def.title);
    }
  });

  it('registers in index.ts exactly the tools in the roster', () => {
    const source = readFileSync(join(repoRoot, 'src', 'index.ts'), 'utf-8');
    const registered = [...source.matchAll(/server\.registerTool\(\s*(\w+)\.name/g)].map(
      (m) => m[1],
    );

    const rosterVars = TOOL_DEFINITIONS.map((d) => {
      // ms_auth_status -> authStatusToolDefinition, ms_server_info -> serverInfoToolDefinition
      const camel = d.name
        .replace(/^ms_/, '')
        .replace(/_([a-z])/g, (_, c: string) => c.toUpperCase());
      return `${camel}ToolDefinition`;
    });

    expect(registered.sort()).toEqual(rosterVars.sort());
  });

  it('matches the tool count claimed in the README and CLAUDE.md', () => {
    const expected = TOOL_DEFINITIONS.length;
    for (const file of ['README.md', 'CLAUDE.md']) {
      const text = readFileSync(join(repoRoot, file), 'utf-8');
      for (const [claim, n] of text.matchAll(/(\d+) tools/g)) {
        expect({ file, claim, n: Number(n) }).toEqual({ file, claim, n: expected });
      }
    }
  });
});
