import { z } from 'zod';
import { graphFetch } from '../graph.js';

export const peopleToolDefinition = {
  name: 'ms_people',
  title: 'People Directory',
  description:
    'Look people up in the organisation directory. Use search to resolve a name to an email address (the input ms_schedule needs), user to read one person’s details plus their manager and direct reports, or groups to list the groups and teams the signed-in user belongs to.',
  inputSchema: z.object({
    search: z.string().optional().describe('Name or partial name to search the directory for'),
    user: z
      .string()
      .optional()
      .describe(
        'User principal name (email) or object ID to fetch details, manager and direct reports for',
      ),
    groups: z.boolean().optional().describe("List the signed-in user's group and team memberships"),
    count: z.int().min(1).max(50).optional().describe('Max results to return (1-50, default 20)'),
  }),
  annotations: {
    title: 'People Directory',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface DirectoryUser {
  id?: string;
  displayName?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
}

interface UsersResponse {
  value: DirectoryUser[];
}

interface DirectoryGroup {
  id?: string;
  displayName?: string;
  description?: string;
  mail?: string;
  groupTypes?: string[];
  '@odata.type'?: string;
}

interface GroupsResponse {
  value: DirectoryGroup[];
}

const USER_SELECT =
  '$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones';

/**
 * Formats a directory user into readable text.
 */
function formatUser(user: DirectoryUser, heading = true): string {
  const lines: string[] = [];
  const name = user.displayName || 'Unknown';
  lines.push(heading ? `## ${name}` : name);
  const email = user.mail || user.userPrincipalName;
  if (email) {
    lines.push(`Email: ${email}`);
  }
  if (user.jobTitle) {
    lines.push(`Job Title: ${user.jobTitle}`);
  }
  if (user.department) {
    lines.push(`Department: ${user.department}`);
  }
  if (user.officeLocation) {
    lines.push(`Office: ${user.officeLocation}`);
  }
  const phone = user.mobilePhone || user.businessPhones?.[0];
  if (phone) {
    lines.push(`Phone: ${phone}`);
  }
  if (user.id) {
    lines.push(`User ID: ${user.id}`);
  }
  return lines.join('\n');
}

/**
 * Formats a group or team membership into readable text.
 */
function formatGroup(group: DirectoryGroup): string {
  const lines: string[] = [];
  lines.push(`## ${group.displayName || 'Unnamed Group'}`);
  if (group.description) {
    lines.push(group.description);
  }
  if (group.mail) {
    lines.push(`Email: ${group.mail}`);
  }
  if (group.groupTypes && group.groupTypes.length > 0) {
    lines.push(`Types: ${group.groupTypes.join(', ')}`);
  }
  if (group.id) {
    lines.push(`Group ID: ${group.id}`);
  }
  return lines.join('\n');
}

/**
 * Fetches a user's manager. Returns null when there is none, or when the
 * caller cannot see it — neither is an error worth failing the whole lookup for.
 */
async function fetchManager(token: string, user: string): Promise<DirectoryUser | null> {
  const result = await graphFetch<DirectoryUser>(
    `/users/${encodeURIComponent(user)}/manager?${USER_SELECT}`,
    token,
    { timezone: false },
  );
  return result.ok ? result.data : null;
}

/**
 * Fetches a user's direct reports. Returns an empty list when there are none
 * or the caller cannot see them.
 */
async function fetchDirectReports(token: string, user: string): Promise<DirectoryUser[]> {
  const result = await graphFetch<UsersResponse>(
    `/users/${encodeURIComponent(user)}/directReports?${USER_SELECT}`,
    token,
    { timezone: false },
  );
  return result.ok ? result.data.value || [] : [];
}

/**
 * Searches the directory, reads one person's details, or lists the signed-in
 * user's group memberships depending on which parameters are provided.
 */
export async function executePeople(
  token: string,
  args: { search?: string; user?: string; groups?: boolean; count?: number },
): Promise<string> {
  const count = Math.min(Math.max(args.count || 20, 1), 50);

  // Mode 1: group and team memberships
  if (args.groups) {
    const result = await graphFetch<GroupsResponse>(`/me/memberOf?$top=${count}`, token, {
      timezone: false,
    });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const groups = (result.data.value || []).filter((g) => g.displayName);
    if (groups.length === 0) {
      return 'No group memberships found.';
    }

    return groups.map(formatGroup).join('\n\n');
  }

  // Mode 2: one person, with manager and direct reports
  if (args.user) {
    const result = await graphFetch<DirectoryUser>(
      `/users/${encodeURIComponent(args.user)}?${USER_SELECT}`,
      token,
      { timezone: false },
    );

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const sections = [formatUser(result.data)];

    const [manager, reports] = await Promise.all([
      fetchManager(token, args.user),
      fetchDirectReports(token, args.user),
    ]);

    if (manager) {
      sections.push(`## Manager\n${formatUser(manager, false)}`);
    }
    if (reports.length > 0) {
      const names = reports.map((r) => `- ${r.displayName} (${r.mail || r.userPrincipalName})`);
      sections.push(`## Direct Reports (${reports.length})\n${names.join('\n')}`);
    }

    return sections.join('\n\n');
  }

  // Mode 3: directory search
  if (!args.search) {
    return 'Provide search to find people, user to read one person, or groups to list your memberships.';
  }

  // $search on /users requires the eventual consistency level and a $count hint.
  const query = `"displayName:${args.search}" OR "mail:${args.search}"`;
  const path = `/users?$search=${encodeURIComponent(query)}&$top=${count}&$count=true&${USER_SELECT}`;
  const result = await graphFetch<UsersResponse>(path, token, {
    timezone: false,
    headers: { ConsistencyLevel: 'eventual' },
  });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const users = result.data.value;
  if (!users || users.length === 0) {
    return `No directory matches for "${args.search}".`;
  }

  return users.map((u) => formatUser(u)).join('\n\n');
}
