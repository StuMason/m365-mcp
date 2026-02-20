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
        officeLocation: 'London',
        userPrincipalName: 'stuart@example.onmicrosoft.com',
      },
    });

    const result = await executeProfile('test-token');

    expect(result).toBe(
      'Name: Stuart Mason\nEmail: stuart@example.com\nJob Title: Engineer\nOffice: London',
    );
    expect(mockGraphFetch).toHaveBeenCalledWith('/me', 'test-token');
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

    expect(result).toBe('Name: N/A\nEmail: N/A\nJob Title: N/A\nOffice: N/A');
  });

  it('returns error message on graph failure', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 401, message: 'Graph token expired. Use ms_auth_status to reconnect.' },
    });

    const result = await executeProfile('expired-token');

    expect(result).toBe('Error: Graph token expired. Use ms_auth_status to reconnect.');
  });
});
