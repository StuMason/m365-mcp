import { timezone } from './format.js';
export interface GraphError {
  status: number;
  message: string;
}

export type GraphResult<T> = { ok: true; data: T } | { ok: false; error: GraphError };

export interface GraphFetchOptions {
  beta?: boolean;
  timezone?: boolean;
  headers?: Record<string, string>;
}

/**
 * Build the full Graph API URL from a path and options.
 */
function buildUrl(path: string, options?: GraphFetchOptions): string {
  const base = options?.beta
    ? 'https://graph.microsoft.com/beta'
    : 'https://graph.microsoft.com/v1.0';
  return `${base}${path}`;
}

/**
 * Build request headers including auth, timezone, and any custom headers.
 * Custom headers are merged after defaults, so they can override them.
 */
function buildHeaders(token: string, options?: GraphFetchOptions): Record<string, string> {
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  if (options?.timezone !== false) {
    // Same resolution as format.ts: formatTime treats offset-less Graph times as
    // already being in this zone, so the two must not drift apart.
    headers['Prefer'] = `outlook.timezone="${timezone()}"`;
  }

  if (options?.headers) {
    Object.assign(headers, options.headers);
  }

  return headers;
}

/**
 * Known Graph error codes mapped to what the caller can actually act on.
 * Anything matched here is reported without the raw body.
 */
const ERROR_HINTS: Array<[RegExp, string]> = [
  [/AutoDiscover|Availability(Config|Service)|InfoWorker/i, 'That mailbox could not be found.'],
  [/ErrorInvalidIdMalformed|invalid.*id|ErrorInvalidId\b/i, 'That ID is not valid for this tool.'],
  [/ItemNotFound|ErrorItemNotFound/i, 'That item no longer exists, or you cannot see it.'],
  [/ErrorAccessDenied|Forbidden/i, 'You do not have access to that item.'],
  [
    /ThrottledRequest|TooManyRequests/i,
    'Microsoft Graph is throttling requests — try again shortly.',
  ],
  [/MailboxNotEnabled|ErrorNonExistentMailbox/i, 'That user has no mailbox.'],
  [/recipient was not found|RecipientNotFound/i, 'That mailbox could not be found.'],
];

/**
 * Maps an error string to the caller-facing hint, or null if unrecognised.
 *
 * Separate from sanitiseGraphError because not every Graph error arrives as an
 * HTTP failure: getSchedule returns 200 with per-mailbox errors inside the body,
 * which bypassed the HTTP path entirely and leaked an Autodiscover exception
 * complete with EWS endpoint, backend server name and diagnostic LID.
 */
export function sanitiseErrorText(text: string): string | null {
  for (const [pattern, hint] of ERROR_HINTS) {
    if (pattern.test(text)) {
      return hint;
    }
  }
  return null;
}

/**
 * Turns a raw Graph error body into something safe and useful.
 *
 * Graph error bodies carry internal detail with no value to the caller: EWS
 * endpoints, .NET exception class names, backend server names and diagnostic
 * LIDs. That is needless disclosure through an assistant's context, so the raw
 * body goes to stderr and the caller gets the intent.
 */
export function sanitiseGraphError(status: number, body: string): string {
  const hint = sanitiseErrorText(body);
  if (hint) {
    return `${hint} (Graph error ${status})`;
  }

  // Unrecognised: surface the code only, never the surrounding prose.
  let code: string | undefined;
  try {
    code = (JSON.parse(body) as { error?: { code?: string } }).error?.code;
  } catch {
    code = undefined;
  }

  process.stderr.write(`Graph API error (${status}): ${body}\n`);
  return code
    ? `Microsoft Graph returned an error (${status}, ${code}).`
    : `Microsoft Graph returned an error (${status}).`;
}

/**
 * Handle a successful or error response from the Graph API.
 */
async function handleResponse<T>(response: Response): Promise<GraphResult<T>> {
  if (response.ok) {
    try {
      const data = (await response.json()) as T;
      return { ok: true, data };
    } catch {
      return {
        ok: false,
        error: {
          status: response.status,
          message: `Graph API returned status ${response.status} but the response body was not valid JSON.`,
        },
      };
    }
  }

  const status = response.status;
  let message: string;

  switch (status) {
    case 401:
      message = 'Graph token expired. Use ms_auth_status to reconnect.';
      break;
    case 403:
      message = 'Insufficient permissions. Check granted scopes with ms_auth_status.';
      break;
    case 404:
      message = 'Resource not found. The item may not exist or you may lack access.';
      break;
    default: {
      let text: string;
      try {
        text = await response.text();
      } catch {
        text = '(unable to read error response body)';
      }
      message = sanitiseGraphError(status, text);
      break;
    }
  }

  return { ok: false, error: { status, message } };
}

/**
 * Wrap a network-level error into a GraphResult.
 */
function handleNetworkError<T>(err: unknown): GraphResult<T> {
  return {
    ok: false,
    error: {
      status: 0,
      message: `Network error: ${err instanceof Error ? err.message : String(err)}`,
    },
  };
}

/**
 * Thin fetch wrapper for Microsoft Graph API GET calls.
 * Translates HTTP errors into typed GraphError results.
 */
export async function graphFetch<T>(
  path: string,
  token: string,
  options?: GraphFetchOptions,
): Promise<GraphResult<T>> {
  let response: Response;
  try {
    response = await fetch(buildUrl(path, options), {
      headers: buildHeaders(token, options),
    });
  } catch (err) {
    return handleNetworkError(err);
  }
  return handleResponse<T>(response);
}

/**
 * Thin fetch wrapper for Microsoft Graph API POST calls.
 * Sends a JSON body and translates HTTP errors into typed GraphError results.
 */
export async function graphPost<TBody, TResult>(
  path: string,
  token: string,
  body: TBody,
  options?: GraphFetchOptions,
): Promise<GraphResult<TResult>> {
  let response: Response;
  try {
    response = await fetch(buildUrl(path, options), {
      method: 'POST',
      headers: buildHeaders(token, options),
      body: JSON.stringify(body),
    });
  } catch (err) {
    return handleNetworkError(err);
  }
  return handleResponse<TResult>(response);
}
