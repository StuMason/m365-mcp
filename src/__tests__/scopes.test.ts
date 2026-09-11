import { describe, it, expect } from '@jest/globals';
import { grantedScopes, lacksScope, missingScope } from '../lib/scopes.js';

/** Builds a JWT-shaped token with the given payload. Signature is never checked. */
function tokenWith(payload: Record<string, unknown>): string {
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  return `header.${body}.signature`;
}

describe('grantedScopes', () => {
  it('reads the space-delimited scp claim', () => {
    const scopes = grantedScopes(tokenWith({ scp: 'Mail.Read User.Read' }));
    expect(scopes).toEqual(new Set(['Mail.Read', 'User.Read']));
  });

  it('reads an array-valued scp claim', () => {
    expect(grantedScopes(tokenWith({ scp: ['Mail.Read'] }))).toEqual(new Set(['Mail.Read']));
  });

  it('ignores the empty strings a doubled separator produces', () => {
    expect(grantedScopes(tokenWith({ scp: 'Mail.Read  User.Read' }))).toEqual(
      new Set(['Mail.Read', 'User.Read']),
    );
  });

  it('returns null for a token with no scp claim', () => {
    expect(grantedScopes(tokenWith({ aud: 'graph' }))).toBeNull();
  });

  it('returns null for a token that is not a JWT', () => {
    expect(grantedScopes('opaque-token')).toBeNull();
  });

  it('returns null when the payload is not JSON', () => {
    expect(grantedScopes('header.bm90LWpzb24.sig')).toBeNull();
  });
});

describe('lacksScope', () => {
  it('is true when the token positively omits the scope', () => {
    expect(lacksScope(tokenWith({ scp: 'Mail.Read' }), 'Tasks.Read')).toBe(true);
  });

  it('is false when the scope is present', () => {
    expect(lacksScope(tokenWith({ scp: 'Mail.Read Tasks.Read' }), 'Tasks.Read')).toBe(false);
  });

  // An unreadable token must not lock tools out: not knowing is not the same as
  // knowing the scope is absent, and guessing wrong here disables working calls.
  it('is false when the scopes cannot be read', () => {
    expect(lacksScope('opaque-token', 'Tasks.Read')).toBe(false);
  });
});

describe('missingScope', () => {
  it('names the scope, the consequence and the alternative', () => {
    const text = missingScope('Tasks.Read', 'Nothing to return.', 'Try ms_search.');
    expect(text).toContain('Tasks.Read');
    expect(text).toContain('Nothing to return.');
    expect(text).toContain('Try ms_search.');
    expect(text).toContain('ms_auth_status');
  });

  it('omits the alternative paragraph when there is none', () => {
    const text = missingScope('Tasks.Read', 'Nothing to return.');
    expect(text.split('\n\n')).toHaveLength(3);
  });

  // brief.ts keys its one-line collapse off this prefix.
  it('starts with the prefix the brief collapses on', () => {
    expect(missingScope('Tasks.Read', 'x')).toMatch(/^Not available:/);
  });
});
