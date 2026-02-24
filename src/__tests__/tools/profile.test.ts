import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
}));

// Dynamic import AFTER the mock is registered
const { executeProfile } = await import('../../lib/tools/profile.js');

describe('executeProfile', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('formats a full profile correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
        jobTitle: 'Engineer',
        department: 'Engineering',
        companyName: 'Acme Corp',
        officeLocation: 'London',
        city: 'London',
        country: 'United Kingdom',
        employeeId: 'E12345',
        employeeType: 'Employee',
        userPrincipalName: 'stuart@example.onmicrosoft.com',
        mobilePhone: '+44 7700 900000',
        businessPhones: ['+44 20 7946 0958'],
        preferredLanguage: 'en-GB',
        givenName: 'Stuart',
        surname: 'Mason',
      },
    });

    const result = await executeProfile('test-token');

    expect(result).toContain('Name: Stuart Mason');
    expect(result).toContain('Email: stuart@example.com');
    expect(result).toContain('Job Title: Engineer');
    expect(result).toContain('Department: Engineering');
    expect(result).toContain('Company: Acme Corp');
    expect(result).toContain('Office: London');
    expect(result).toContain('City: London');
    expect(result).toContain('Country: United Kingdom');
    expect(result).toContain('Employee ID: E12345');
    expect(result).toContain('Employee Type: Employee');
    expect(result).toContain('Mobile: +44 7700 900000');
    expect(result).toContain('Business Phones: +44 20 7946 0958');
    expect(result).toContain('Language: en-GB');
    expect(result).toContain('Given Name: Stuart');
    expect(result).toContain('Surname: Mason');
  });

  it('uses userPrincipalName when mail is missing', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        userPrincipalName: 'stuart@example.onmicrosoft.com',
      },
    });

    const result = await executeProfile('test-token');

    expect(result).toContain('Email: stuart@example.onmicrosoft.com');
  });

  it('handles all missing fields gracefully', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {},
    });

    const result = await executeProfile('test-token');

    expect(result).toContain('Name: N/A');
    expect(result).toContain('Email: N/A');
    expect(result).toContain('Job Title: N/A');
    expect(result).toContain('Office: N/A');
  });

  it('returns error message on graph failure', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 401, message: 'Graph token expired. Use ms_auth_status to reconnect.' },
    });

    const result = await executeProfile('expired-token');

    expect(result).toBe('Error: Graph token expired. Use ms_auth_status to reconnect.');
  });

  it('sends expanded $select fields in the query', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });

    await executeProfile('test-token');

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('displayName');
    expect(calledPath).toContain('department');
    expect(calledPath).toContain('companyName');
    expect(calledPath).toContain('city');
    expect(calledPath).toContain('country');
    expect(calledPath).toContain('employeeId');
    expect(calledPath).toContain('employeeType');
    expect(calledPath).toContain('mobilePhone');
    expect(calledPath).toContain('businessPhones');
    expect(calledPath).toContain('preferredLanguage');
    expect(calledPath).toContain('givenName');
    expect(calledPath).toContain('surname');
  });

  it('include: ["manager"] success — shows manager info', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
        jobTitle: 'Engineer',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Jane Boss',
        mail: 'jane@example.com',
        jobTitle: 'Director',
      },
    });

    const result = await executeProfile('test-token', { include: ['manager'] });

    expect(result).toContain('Jane Boss');
    expect(result).toContain('jane@example.com');
    expect(result).toContain('Director');
    expect(mockGraphFetch).toHaveBeenCalledTimes(2);
    expect(mockGraphFetch.mock.calls[1][0]).toContain('/me/manager');
  });

  it('include: ["manager"] with 403 — shows graceful message', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });

    const result = await executeProfile('test-token', { include: ['manager'] });

    expect(result).toContain('Manager info not available (tenant policy)');
    expect(result).not.toMatch(/^Error/);
    expect(result).toContain('Name: Stuart Mason');
  });

  it('include: ["reports"] with results — lists reports', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Jane Boss',
        mail: 'jane@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [
          { displayName: 'Alice', mail: 'alice@example.com' },
          { displayName: 'Bob', mail: 'bob@example.com' },
        ],
      },
    });

    const result = await executeProfile('test-token', { include: ['reports'] });

    expect(result).toContain('Alice <alice@example.com>');
    expect(result).toContain('Bob <bob@example.com>');
  });

  it('include: ["reports"] empty — shows no direct reports', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [],
      },
    });

    const result = await executeProfile('test-token', { include: ['reports'] });

    expect(result).toContain('No direct reports');
  });

  it('include: ["groups"] empty — shows no group memberships', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [],
      },
    });

    const result = await executeProfile('test-token', { include: ['groups'] });

    expect(result).toContain('No group memberships');
  });

  it('include: ["groups"] with null displayName — uses fallback chain', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [
          { displayName: 'Engineering', mail: null, id: 'g1' },
          { displayName: null, mail: 'secret-group@example.com', id: 'g2' },
          { displayName: null, mail: null, id: 'g3' },
          { displayName: null, mail: null, id: null },
        ],
      },
    });

    const result = await executeProfile('test-token', { include: ['groups'] });

    expect(result).toContain('Engineering');
    expect(result).toContain('secret-group@example.com');
    expect(result).toContain('g3');
    expect(result).toContain('(unnamed)');
  });

  it('include: ["photo"] — returns photo metadata', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        '@odata.mediaContentType': 'image/jpeg',
        height: 648,
        width: 648,
      },
    });

    const result = await executeProfile('test-token', { include: ['photo'] });

    expect(result).toContain('Photo available');
    expect(result).toContain('648');
  });

  it('include: ["photo"] without dimensions — shows photo available without size', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        '@odata.mediaContentType': 'image/jpeg',
      },
    });

    const result = await executeProfile('test-token', { include: ['photo'] });

    expect(result).toContain('Photo: Photo available');
    expect(result).not.toMatch(/\d+x\d+/);
  });

  it('include: ["photo"] when no photo — shows not available', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    const result = await executeProfile('test-token', { include: ['photo'] });

    expect(result).toContain('No photo available');
  });

  it('include: ["manager", "reports", "groups"] all at once — all sections appear', async () => {
    // 1st call: profile
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Stuart Mason',
        mail: 'stuart@example.com',
      },
    });
    // 2nd call: manager
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        displayName: 'Jane Boss',
        mail: 'jane@example.com',
        jobTitle: 'Director',
      },
    });
    // 3rd call: reports
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [{ displayName: 'Alice', mail: 'alice@example.com' }],
      },
    });
    // 4th call: groups
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: {
        value: [{ displayName: 'Engineering', mail: null, id: 'g1' }],
      },
    });

    const result = await executeProfile('test-token', {
      include: ['manager', 'reports', 'groups'],
    });

    expect(result).toContain('Jane Boss');
    expect(result).toContain('Alice <alice@example.com>');
    expect(result).toContain('Engineering');
  });

  it('returns error when base profile call fails', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 500, message: 'Graph API error (500): Internal Server Error' },
    });

    const result = await executeProfile('test-token', { include: ['manager'] });

    expect(result).toMatch(/^Error:/);
    expect(result).toContain('Graph API error (500)');
  });

  it('include: ["manager"] with non-403 error shows prefixed message', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 500, message: 'Internal Server Error' },
    });

    const result = await executeProfile('test-token', { include: ['manager'] });
    expect(result).toContain('Error fetching manager: Internal Server Error');
    expect(result).toContain('Name: Stuart');
  });

  it('include: ["reports"] with error shows prefixed message', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 500, message: 'Internal Server Error' },
    });

    const result = await executeProfile('test-token', { include: ['reports'] });
    expect(result).toContain('Error fetching reports: Internal Server Error');
  });

  it('include: ["groups"] with error shows prefixed message', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 500, message: 'Internal Server Error' },
    });

    const result = await executeProfile('test-token', { include: ['groups'] });
    expect(result).toContain('Error fetching groups: Internal Server Error');
  });

  it('include: ["photo"] with non-404 error shows photo error message', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 401, message: 'Graph token expired.' },
    });

    const result = await executeProfile('test-token', { include: ['photo'] });
    expect(result).toContain('Photo error: Graph token expired.');
  });

  it('include: ["photo"] with 404 shows no photo available', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });
    mockGraphFetch.mockResolvedValueOnce({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    const result = await executeProfile('test-token', { include: ['photo'] });
    expect(result).toContain('Photo: No photo available');
  });

  it('warns about unknown include options', async () => {
    mockGraphFetch.mockResolvedValueOnce({
      ok: true,
      data: { displayName: 'Stuart', mail: 'stuart@example.com' },
    });

    const result = await executeProfile('test-token', { include: ['typo'] });
    expect(result).toContain('Warning: Unknown include option "typo"');
    expect(result).toContain('Valid options: manager, reports, groups, photo');
  });
});
