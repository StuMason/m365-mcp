export interface GraphError {
  status: number;
  message: string;
}

export type GraphResult<T> = { ok: true; data: T } | { ok: false; error: GraphError };

export interface GraphFetchOptions {
  beta?: boolean;
  timezone?: boolean;
}

/**
 * Thin fetch wrapper for Microsoft Graph API calls.
 * Translates HTTP errors into typed GraphError results.
 */
export async function graphFetch<T>(
  path: string,
  token: string,
  options?: GraphFetchOptions,
): Promise<GraphResult<T>> {
  const base = options?.beta
    ? 'https://graph.microsoft.com/beta'
    : 'https://graph.microsoft.com/v1.0';

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  if (options?.timezone !== false) {
    const tz =
      process.env.MS365_MCP_TIMEZONE || Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
    headers['Prefer'] = `outlook.timezone="${tz}"`;
  }

  let response: Response;
  try {
    response = await fetch(`${base}${path}`, { headers });
  } catch (err) {
    return {
      ok: false,
      error: {
        status: 0,
        message: `Network error: ${err instanceof Error ? err.message : String(err)}`,
      },
    };
  }

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
      message = 'Resource not found. Your account may not have an Exchange Online license.';
      break;
    default: {
      const text = await response.text();
      message = `Graph API error (${status}): ${text}`;
      break;
    }
  }

  return { ok: false, error: { status, message } };
}
