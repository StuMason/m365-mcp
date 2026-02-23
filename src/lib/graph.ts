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
    const tz =
      process.env.MS365_MCP_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    headers['Prefer'] = `outlook.timezone="${tz}"`;
  }

  if (options?.headers) {
    Object.assign(headers, options.headers);
  }

  return headers;
}

/**
 * Handle a successful or error response from the Graph API.
 */
async function handleResponse<T>(response: Response): Promise<GraphResult<T>> {
  if (response.ok) {
    const data = (await response.json()) as T;
    return { ok: true, data };
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
      const text = await response.text();
      message = `Graph API error (${status}): ${text}`;
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
