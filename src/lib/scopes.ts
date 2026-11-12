/**
 * Scope inspection for the current access token.
 *
 * Graph delegated tokens carry the granted scopes in the `scp` claim, so a tool can
 * tell "you were never granted this" apart from "this failed" and say which it was.
 * Registrations differ in what they have consented — the same build runs against a
 * client with Tasks.Read and one without — so the difference has to be discovered at
 * runtime rather than assumed at compile time.
 */

/**
 * Returns the scopes in a Graph access token, or null when they cannot be read.
 *
 * Null means "do not know", not "none": the JWT shape of Graph tokens is an
 * implementation detail, so every caller must treat null as permission to proceed.
 */
export function grantedScopes(token: string): Set<string> | null {
  const payload = token.split('.')[1];
  if (!payload) return null;
  try {
    const claims = JSON.parse(Buffer.from(payload, 'base64url').toString('utf-8')) as {
      scp?: string | string[];
    };
    const scp = claims.scp;
    if (!scp) return null;
    const list = Array.isArray(scp) ? scp : scp.split(' ');
    return new Set(list.filter(Boolean));
  } catch {
    return null;
  }
}

/**
 * Returns true only when the token positively does not carry the scope.
 * An unreadable token answers false, so an unexpected token format degrades to
 * attempting the call rather than refusing it.
 */
export function lacksScope(token: string, scope: string): boolean {
  const granted = grantedScopes(token);
  return granted !== null && !granted.has(scope);
}

/**
 * Builds the message shown when a scope is missing: what is not granted, what that
 * costs, and what still works instead. Tools return this in place of the Graph call
 * so the caller gets an answer rather than a 403 it cannot act on.
 */
export function missingScope(scope: string, consequence: string, alternative?: string): string {
  const lines = [
    `Not available: this app registration has not been granted ${scope}.`,
    consequence,
  ];
  if (alternative) {
    lines.push(alternative);
  }
  lines.push('Run ms_auth_status to see the full list of granted scopes.');
  return lines.join('\n\n');
}
