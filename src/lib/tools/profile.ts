import { graphFetch } from '../graph.js';

export const profileToolDefinition = {
  name: 'ms_profile',
  description:
    "Fetch the user's Microsoft 365 profile (display name, email, job title, office location)",
  inputSchema: {
    type: 'object' as const,
    properties: {},
  },
};

interface ProfileResponse {
  displayName?: string;
  mail?: string;
  jobTitle?: string;
  officeLocation?: string;
  userPrincipalName?: string;
}

/**
 * Fetches the authenticated user's Microsoft 365 profile and returns
 * a human-readable summary.
 */
export async function executeProfile(token: string): Promise<string> {
  const result = await graphFetch<ProfileResponse>('/me', token);

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const p = result.data;
  const lines = [
    `Name: ${p.displayName || 'N/A'}`,
    `Email: ${p.mail || p.userPrincipalName || 'N/A'}`,
    `Job Title: ${p.jobTitle || 'N/A'}`,
    `Office: ${p.officeLocation || 'N/A'}`,
  ];
  return lines.join('\n');
}
