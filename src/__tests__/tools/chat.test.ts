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
const { executeChat, stripHtml } = await import('../../lib/tools/chat.js');

describe('executeChat', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('lists chats correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'chat-123',
            topic: 'Project Alpha',
            chatType: 'group',
            lastMessagePreview: {
              body: { content: 'See you tomorrow' },
              createdDateTime: '2025-06-15T10:00:00Z',
            },
          },
        ],
      },
    });

    const result = await executeChat('test-token', {});

    expect(result).toContain('## Project Alpha');
    expect(result).toContain('Type: group');
    expect(result).toContain('See you tomorrow');
    expect(result).toContain('Chat ID: chat-123');
  });

  it('drills down into chat messages', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            from: { user: { displayName: 'Alice' } },
            createdDateTime: '2025-06-15T10:30:00Z',
            body: { content: 'Hello everyone!', contentType: 'text' },
          },
          {
            from: { user: { displayName: 'Bob' } },
            createdDateTime: '2025-06-15T10:31:00Z',
            body: { content: '<p>Hi Alice!</p>', contentType: 'html' },
          },
        ],
      },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123' });

    expect(result).toContain('**Alice**');
    expect(result).toContain('Hello everyone!');
    expect(result).toContain('**Bob**');
    expect(result).toContain('Hi Alice!');
    // HTML should be stripped
    expect(result).not.toContain('<p>');
  });

  it('handles empty chat list', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeChat('test-token', {});

    expect(result).toBe('No Teams chats found.');
  });

  it('handles empty messages in a chat', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123' });

    expect(result).toBe('No messages found in this chat.');
  });

  it('handles errors from Graph API', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: {
        status: 403,
        message: 'Insufficient permissions. Check granted scopes with ms_auth_status.',
      },
    });

    const result = await executeChat('test-token', {});

    expect(result).toBe(
      'Error: Insufficient permissions. Check granted scopes with ms_auth_status.',
    );
  });

  it('handles errors when drilling into a chat', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: {
        status: 404,
        message: 'Resource not found. Your account may not have an Exchange Online license.',
      },
    });

    const result = await executeChat('test-token', { chat_id: 'nonexistent' });

    expect(result).toBe(
      'Error: Resource not found. Your account may not have an Exchange Online license.',
    );
  });

  it('clamps count correctly', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeChat('test-token', { count: 50 });

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=25'), 'test-token', {
      timezone: false,
    });
  });

  it('defaults count to 10', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeChat('test-token', {});

    expect(mockGraphFetch).toHaveBeenCalledWith(expect.stringContaining('$top=10'), 'test-token', {
      timezone: false,
    });
  });

  it('encodes chat_id in the URL', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [] },
    });

    await executeChat('test-token', { chat_id: '19:abc@thread.v2' });

    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/chats/19%3Aabc%40thread.v2/messages'),
      'test-token',
      { timezone: false },
    );
  });

  it('handles chat listing with no topic', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'chat-456',
            topic: null,
            chatType: 'oneOnOne',
            lastMessagePreview: null,
          },
        ],
      },
    });

    const result = await executeChat('test-token', {});

    expect(result).toContain('## oneOnOne chat');
    expect(result).toContain('Type: oneOnOne');
    expect(result).toContain('Chat ID: chat-456');
  });

  it('handles messages with empty body', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            from: { user: { displayName: 'Alice' } },
            createdDateTime: '2025-06-15T10:30:00Z',
            body: { content: '', contentType: 'text' },
          },
        ],
      },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123' });

    expect(result).toContain('(empty message)');
  });

  it('strips HTML from chat listing preview', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'chat-html',
            topic: 'HTML Preview Chat',
            chatType: 'group',
            lastMessagePreview: {
              body: { content: '<p>Hey <at id="0">Stuart</at>, check this out</p>' },
              createdDateTime: '2025-06-15T10:00:00Z',
            },
          },
        ],
      },
    });

    const result = await executeChat('test-token', {});

    expect(result).toContain('Hey Stuart, check this out');
    expect(result).not.toContain('<p>');
    expect(result).not.toContain('<at');
  });

  it('handles emoji tags in messages', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            from: { user: { displayName: 'Alice' } },
            createdDateTime: '2025-06-15T10:30:00Z',
            body: {
              content: '<p>Great work <emoji id="thumbsup" alt="👍"></emoji></p>',
              contentType: 'html',
            },
          },
        ],
      },
    });

    const result = await executeChat('test-token', { chat_id: 'chat-123' });

    expect(result).toContain('Great work 👍');
    expect(result).not.toContain('<emoji');
    expect(result).not.toContain('<p>');
  });
});

describe('stripHtml', () => {
  it('converts br tags to newlines', () => {
    expect(stripHtml('line1<br>line2<br/>line3')).toBe('line1\nline2\nline3');
  });

  it('converts closing p tags to newlines', () => {
    expect(stripHtml('<p>First</p><p>Second</p>')).toBe('First\nSecond');
  });

  it('extracts emoji alt text', () => {
    expect(stripHtml('<emoji id="thumbsup" alt="👍"></emoji>')).toBe('👍');
    expect(stripHtml('<emoji id="smile" alt="😊"/>')).toBe('😊');
  });

  it('preserves text inside at-mention tags', () => {
    expect(stripHtml('<at id="0">Stuart Mason</at>')).toBe('Stuart Mason');
  });

  it('removes attachment tags and content', () => {
    expect(stripHtml('Hello<attachment id="abc">file.pdf</attachment> world')).toBe('Hello world');
  });

  it('decodes HTML entities', () => {
    expect(stripHtml('a &amp; b &lt; c &gt; d &quot;e&quot; f&#39;s')).toBe(
      'a & b < c > d "e" f\'s',
    );
  });

  it('collapses excessive newlines', () => {
    expect(stripHtml('<p></p><p></p><p>Content</p>')).toBe('Content');
  });

  it('trims whitespace', () => {
    expect(stripHtml('  <p>Hello</p>  ')).toBe('Hello');
  });
});
