import { graphFetch } from '../graph.js';

export const profileToolDefinition = {
  name: 'ms_profile',
  description:
    "Fetch the user's Microsoft 365 profile. Optionally include manager, reports, groups, or photo.",
  inputSchema: {
    type: 'object' as const,
    properties: {
      include: {
        type: 'array',
        items: { type: 'string' },
        description: 'Additional data to include: "manager", "reports", "groups", "photo"',
      },
    },
  },
};

const PROFILE_SELECT_FIELDS = [
  'displayName',
  'mail',
  'jobTitle',
  'department',
  'companyName',
  'officeLocation',
  'city',
  'country',
  'employeeId',
  'employeeType',
  'userPrincipalName',
  'mobilePhone',
  'businessPhones',
  'preferredLanguage',
  'givenName',
  'surname',
].join(',');

interface ProfileResponse {
  displayName?: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  companyName?: string;
  officeLocation?: string;
  city?: string;
  country?: string;
  employeeId?: string;
  employeeType?: string;
  userPrincipalName?: string;
  mobilePhone?: string;
  businessPhones?: string[];
  preferredLanguage?: string;
  givenName?: string;
  surname?: string;
}

interface ManagerResponse {
  displayName?: string;
  mail?: string;
  jobTitle?: string;
}

interface DirectReport {
  displayName?: string;
  mail?: string;
}

interface GroupMember {
  displayName?: string | null;
  mail?: string | null;
  id?: string | null;
}

interface PhotoMetadata {
  '@odata.mediaContentType'?: string;
  height?: number;
  width?: number;
}

/**
 * Format the base profile fields into a multi-line string.
 */
function formatProfile(p: ProfileResponse): string {
  const lines = [
    `Name: ${p.displayName || 'N/A'}`,
    `Email: ${p.mail || p.userPrincipalName || 'N/A'}`,
    `Job Title: ${p.jobTitle || 'N/A'}`,
    `Department: ${p.department || 'N/A'}`,
    `Company: ${p.companyName || 'N/A'}`,
    `Office: ${p.officeLocation || 'N/A'}`,
    `City: ${p.city || 'N/A'}`,
    `Country: ${p.country || 'N/A'}`,
    `Employee ID: ${p.employeeId || 'N/A'}`,
    `Employee Type: ${p.employeeType || 'N/A'}`,
    `Mobile: ${p.mobilePhone || 'N/A'}`,
    `Business Phones: ${p.businessPhones?.length ? p.businessPhones.join(', ') : 'N/A'}`,
    `Language: ${p.preferredLanguage || 'N/A'}`,
    `Given Name: ${p.givenName || 'N/A'}`,
    `Surname: ${p.surname || 'N/A'}`,
  ];
  return lines.join('\n');
}

/**
 * Fetch manager info.
 */
async function fetchManager(token: string): Promise<string> {
  const result = await graphFetch<ManagerResponse>(
    '/me/manager?$select=displayName,mail,jobTitle',
    token,
  );

  if (!result.ok) {
    if (result.error.status === 403) {
      return 'Manager info not available (tenant policy)';
    }
    return `Error fetching manager: ${result.error.message}`;
  }

  const m = result.data;
  const parts = [m.displayName || 'N/A'];
  if (m.mail) parts.push(`<${m.mail}>`);
  if (m.jobTitle) parts.push(`- ${m.jobTitle}`);
  return parts.join(' ');
}

/**
 * Fetch direct reports.
 */
async function fetchReports(token: string): Promise<string> {
  const result = await graphFetch<{ value: DirectReport[] }>(
    '/me/directReports?$select=displayName,mail&$top=10',
    token,
  );

  if (!result.ok) {
    return `Error fetching reports: ${result.error.message}`;
  }

  const reports = result.data.value;
  if (!reports.length) {
    return 'No direct reports';
  }

  return reports.map((r) => `${r.displayName || 'N/A'} <${r.mail || 'N/A'}>`).join('\n');
}

/**
 * Fetch group memberships.
 */
async function fetchGroups(token: string): Promise<string> {
  const result = await graphFetch<{ value: GroupMember[] }>(
    '/me/memberOf?$top=50&$select=displayName,mail,id',
    token,
  );

  if (!result.ok) {
    return `Error fetching groups: ${result.error.message}`;
  }

  const groups = result.data.value;
  if (!groups.length) {
    return 'No group memberships';
  }

  return groups.map((g) => g.displayName ?? g.mail ?? g.id ?? '(unnamed)').join('\n');
}

/**
 * Fetch photo metadata.
 */
async function fetchPhoto(token: string): Promise<string> {
  const result = await graphFetch<PhotoMetadata>('/me/photo', token);

  if (!result.ok) {
    if (result.error.status === 404) {
      return 'No photo available';
    }
    return `Photo error: ${result.error.message}`;
  }

  const photo = result.data;
  if (photo.height && photo.width) {
    return `Photo available (${photo.width}x${photo.height})`;
  }
  return 'Photo available';
}

/**
 * Fetches the authenticated user's Microsoft 365 profile and returns
 * a human-readable summary. Optionally includes manager, reports, groups, or photo.
 */
export async function executeProfile(
  token: string,
  args?: { include?: string[] },
): Promise<string> {
  const result = await graphFetch<ProfileResponse>(`/me?$select=${PROFILE_SELECT_FIELDS}`, token);

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const sections: string[] = [formatProfile(result.data)];

  const include = args?.include ?? [];

  for (const item of include) {
    switch (item) {
      case 'manager': {
        const managerInfo = await fetchManager(token);
        sections.push(`\nManager: ${managerInfo}`);
        break;
      }
      case 'reports': {
        const reportsInfo = await fetchReports(token);
        sections.push(`\nDirect Reports:\n${reportsInfo}`);
        break;
      }
      case 'groups': {
        const groupsInfo = await fetchGroups(token);
        sections.push(`\nGroups:\n${groupsInfo}`);
        break;
      }
      case 'photo': {
        const photoInfo = await fetchPhoto(token);
        sections.push(`\nPhoto: ${photoInfo}`);
        break;
      }
      default:
        sections.push(
          `\nWarning: Unknown include option "${item}". Valid options: manager, reports, groups, photo`,
        );
        break;
    }
  }

  return sections.join('\n');
}
