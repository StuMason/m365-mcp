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
const { executeMail } = await import('../../lib/tools/mail.js');

describe('executeMail', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('formats messages correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Meeting Notes',
            from: { emailAddress: { name: 'Alice Smith', address: 'alice@example.com' } },
            receivedDateTime: '2025-06-15T10:30:00Z',
            bodyPreview: 'Here are the meeting notes from today.',
            isRead: true,
            importance: 'high',
          },
        ],
      },
    });

    const result = await executeMail('test-token', {});

    expect(result).toContain('## Meeting Notes');
    expect(result).toContain('From: Alice Smith <alice@example.com>');
    expect(result).toContain('Importance: high | Read: Yes');
    expect(result).toContain('Here are the meeting notes from today.');
  });

  it('handles search param without $orderBy', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { search: 'quarterly report' });

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('$search="quarterly%20report"');
    expect(calledPath).not.toContain('$orderby');
  });

  it('includes $orderBy when no search is provided', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', {});

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('$orderby=receivedDateTime desc');
    expect(calledPath).not.toContain('$search');
  });

  it('handles empty results', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeMail('test-token', {});

    expect(result).toBe('No emails found.');
  });

  it('handles errors from Graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 401, message: 'Graph token expired. Use ms_auth_status to reconnect.' },
    });

    const result = await executeMail('expired-token', {});

    expect(result).toBe('Error: Graph token expired. Use ms_auth_status to reconnect.');
  });

  it('clamps count to minimum of 1', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { count: 0 });

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=1'), 'test-token', {
      timezone: false,
    });
  });

  it('clamps count to maximum of 25', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', { count: 100 });

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=25'), 'test-token', {
      timezone: false,
    });
  });

  it('defaults count to 10', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', {});

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=10'), 'test-token', {
      timezone: false,
    });
  });

  it('formats unread messages correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'New Message',
            from: { emailAddress: { name: 'Bob', address: 'bob@example.com' } },
            receivedDateTime: '2025-06-15T12:00:00Z',
            bodyPreview: 'Hello there.',
            isRead: false,
            importance: 'normal',
          },
        ],
      },
    });

    const result = await executeMail('test-token', {});

    expect(result).toContain('Read: No');
  });

  it('handles messages with missing fields', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: undefined,
            from: undefined,
            receivedDateTime: undefined,
            bodyPreview: undefined,
            isRead: false,
            importance: undefined,
          },
        ],
      },
    });

    const result = await executeMail('test-token', {});

    expect(result).toContain('## No Subject');
    expect(result).toContain('From: Unknown <unknown>');
    expect(result).toContain('Date: N/A');
    expect(result).toContain('Importance: normal');
  });

  it('formats multiple messages separated by blank lines', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            subject: 'Email 1',
            from: { emailAddress: { name: 'A', address: 'a@example.com' } },
            receivedDateTime: '2025-06-15T10:00:00Z',
            bodyPreview: 'Body 1',
            isRead: true,
            importance: 'normal',
          },
          {
            subject: 'Email 2',
            from: { emailAddress: { name: 'B', address: 'b@example.com' } },
            receivedDateTime: '2025-06-15T11:00:00Z',
            bodyPreview: 'Body 2',
            isRead: false,
            importance: 'high',
          },
        ],
      },
    });

    const result = await executeMail('test-token', {});

    expect(result).toContain('## Email 1');
    expect(result).toContain('## Email 2');
    const parts = result.split('\n\n');
    expect(parts.length).toBe(2);
  });

  it('passes timezone false to graphFetch', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', {});

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.any(String), 'test-token', {
      timezone: false,
    });
  });

  it('includes message ID in list results', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'AAMkAGI2',
            subject: 'Test Email',
            from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
            receivedDateTime: '2025-06-15T10:00:00Z',
            bodyPreview: 'Preview text',
            isRead: true,
            importance: 'normal',
          },
        ],
      },
    });

    const result = await executeMail('test-token', {});

    expect(result).toContain('Message ID: AAMkAGI2');
  });

  it('includes id in $select for list mode', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeMail('test-token', {});

    const calledPath = mockGraphFetch.mock.calls[0][0] as string;
    expect(calledPath).toContain('$select=id,');
  });

  describe('drill-down mode', () => {
    it('fetches full email body by message_id', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          subject: 'Quarterly Report',
          from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
          receivedDateTime: '2025-06-15T10:00:00Z',
          body: { contentType: 'text', content: 'Full body content here.' },
          isRead: true,
          importance: 'high',
          toRecipients: [{ emailAddress: { name: 'Bob', address: 'bob@example.com' } }],
          ccRecipients: [{ emailAddress: { name: 'Carol', address: 'carol@example.com' } }],
        },
      });

      const result = await executeMail('test-token', { message_id: 'AAMkAGI2' });

      expect(result).toContain('# Quarterly Report');
      expect(result).toContain('From: Alice <alice@example.com>');
      expect(result).toContain('To: Bob <bob@example.com>');
      expect(result).toContain('Cc: Carol <carol@example.com>');
      expect(result).toContain('Full body content here.');
      expect(mockGraphFetch).toHaveBeenCalledWith(
        expect.stringContaining('/me/messages/AAMkAGI2'),
        'test-token',
        { timezone: false },
      );
    });

    it('strips HTML from email body', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          subject: 'HTML Email',
          from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
          body: {
            contentType: 'html',
            content: '<html><body><p>Hello world</p><p>Second paragraph</p></body></html>',
          },
          isRead: true,
          importance: 'normal',
        },
      });

      const result = await executeMail('test-token', { message_id: 'msg-html' });

      expect(result).toContain('Hello world');
      expect(result).toContain('Second paragraph');
      expect(result).not.toContain('<p>');
      expect(result).not.toContain('<html>');
    });

    it('handles missing body gracefully', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: true,
        data: {
          subject: 'Empty Email',
          from: { emailAddress: { name: 'Alice', address: 'alice@example.com' } },
          isRead: false,
          importance: 'normal',
        },
      });

      const result = await executeMail('test-token', { message_id: 'msg-empty' });

      expect(result).toContain('# Empty Email');
      expect(result).toContain('(no body)');
    });

    it('returns error when message not found', async () => {
      mockGraphFetch.mockResolvedValue({
        ok: false,
        error: {
          status: 404,
          message: 'Resource not found. The item may not exist or you may lack access.',
        },
      });

      const result = await executeMail('test-token', { message_id: 'nonexistent' });

      expect(result).toContain('Error:');
      expect(result).toContain('not found');
    });
  });
});
