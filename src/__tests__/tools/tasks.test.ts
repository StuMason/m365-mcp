import { jest } from '@jest/globals';
import type { GraphResult } from '../../lib/graph.js';

const mockGraphFetch =
  jest.fn<
    <T>(
      path: string,
      token: string,
      options?: { beta?: boolean; timezone?: boolean; headers?: Record<string, string> },
    ) => Promise<GraphResult<T>>
  >();

jest.unstable_mockModule('../../lib/graph.js', () => ({
  graphFetch: mockGraphFetch,
}));

const { executeTasks, tasksToolDefinition } = await import('../../lib/tools/tasks.js');

describe('tasksToolDefinition', () => {
  it('is declared read-only', () => {
    expect(tasksToolDefinition.name).toBe('ms_tasks');
    expect(tasksToolDefinition.annotations.readOnlyHint).toBe(true);
  });
});

describe('executeTasks', () => {
  afterEach(() => {
    mockGraphFetch.mockReset();
  });

  it('lists To Do lists by default', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'list-1', displayName: 'Tasks', wellknownListName: 'defaultList' },
          { id: 'list-2', displayName: 'Shopping', isShared: true, wellknownListName: 'none' },
        ],
      },
    });

    const result = await executeTasks('test-token', {});

    expect(result).toContain('Tasks');
    expect(result).toContain('List ID: list-1');
    expect(result).toContain('Well-known list: defaultList');
    expect(result).toContain('Shopping (shared)');
    // wellknownListName of 'none' is noise, not information.
    expect(result).not.toContain('Well-known list: none');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('/me/todo/lists'),
      'test-token',
      expect.any(Object),
    );
  });

  it('names unnamed lists', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'l' }] } });

    expect(await executeTasks('test-token', {})).toContain('Unnamed List');
  });

  it('reports when there are no lists', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executeTasks('test-token', {})).toBe('No To Do lists found.');
  });

  it('surfaces an error listing lists', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    expect(await executeTasks('test-token', {})).toBe('Error: Insufficient permissions.');
  });

  it('reads tasks in a list and excludes completed ones by default', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'task-1',
            title: 'Ship the release',
            status: 'notStarted',
            importance: 'high',
            dueDateTime: { dateTime: '2026-09-15T09:00:00' },
            body: { content: 'Cut v0.8' },
          },
        ],
      },
    });

    const result = await executeTasks('test-token', { list_id: 'list-1' });

    expect(result).toContain('[ ] Ship the release');
    expect(result).toContain('Importance: high');
    expect(result).toContain('Cut v0.8');
    const path = mockGraphFetch.mock.calls[0]![0];
    expect(path).toContain('/me/todo/lists/list-1/tasks');
    expect(path).toContain('status');
  });

  it('includes completed tasks when asked', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'task-2',
            title: 'Old task',
            status: 'completed',
            importance: 'normal',
            completedDateTime: { dateTime: '2026-09-01T09:00:00' },
          },
        ],
      },
    });

    const result = await executeTasks('test-token', {
      list_id: 'list-1',
      include_completed: true,
    });

    expect(result).toContain('[x] Old task');
    expect(result).toContain('Completed:');
    // normal importance is the default and not worth a line.
    expect(result).not.toContain('Importance: normal');
    expect(mockGraphFetch.mock.calls[0]![0]).not.toContain('$filter');
  });

  it('names untitled tasks', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 't' }] } });

    expect(await executeTasks('test-token', { list_id: 'l' })).toContain('Untitled task');
  });

  it('distinguishes an empty list from a fully completed one', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executeTasks('test-token', { list_id: 'l' })).toContain('No open tasks');
    expect(await executeTasks('test-token', { list_id: 'l', include_completed: true })).toBe(
      'No tasks found in this list.',
    );
  });

  it('surfaces an error reading tasks', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 404, message: 'Resource not found.' },
    });

    expect(await executeTasks('test-token', { list_id: 'l' })).toBe('Error: Resource not found.');
  });

  it('reads Planner tasks and hides finished ones by default', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          {
            id: 'p-1',
            title: 'Design review',
            planId: 'plan-1',
            percentComplete: 50,
            dueDateTime: '2026-09-20T09:00:00Z',
          },
          { id: 'p-2', title: 'Done already', percentComplete: 100 },
        ],
      },
    });

    const result = await executeTasks('test-token', { planner: true });

    expect(result).toContain('[ ] Design review');
    expect(result).toContain('Progress: 50%');
    expect(result).toContain('Plan ID: plan-1');
    expect(result).not.toContain('Done already');
    expect(mockGraphFetch).toHaveBeenCalledWith(
      '/me/planner/tasks',
      'test-token',
      expect.any(Object),
    );
  });

  it('includes finished Planner tasks when asked', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: { value: [{ id: 'p-2', title: 'Done already', percentComplete: 100 }] },
    });

    const result = await executeTasks('test-token', { planner: true, include_completed: true });

    expect(result).toContain('[x] Done already');
  });

  it('treats a Planner task with no progress field as open', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [{ id: 'p-3' }] } });

    const result = await executeTasks('test-token', { planner: true });

    expect(result).toContain('[ ] Untitled task');
    expect(result).toContain('Progress: 0%');
  });

  it('distinguishes no Planner tasks from none open', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    expect(await executeTasks('test-token', { planner: true })).toContain('No open Planner tasks');
    expect(await executeTasks('test-token', { planner: true, include_completed: true })).toBe(
      'No Planner tasks assigned to you.',
    );
  });

  it('surfaces an error reading Planner tasks', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: false,
      error: { status: 403, message: 'Insufficient permissions.' },
    });

    expect(await executeTasks('test-token', { planner: true })).toBe(
      'Error: Insufficient permissions.',
    );
  });

  it('caps Planner results at count', async () => {
    mockGraphFetch.mockResolvedValue({
      ok: true,
      data: {
        value: [
          { id: 'a', title: 'A', percentComplete: 0 },
          { id: 'b', title: 'B', percentComplete: 0 },
        ],
      },
    });

    const result = await executeTasks('test-token', { planner: true, count: 1 });

    expect(result).toContain('A');
    expect(result).not.toContain('B');
  });

  it('clamps count to the documented range', async () => {
    mockGraphFetch.mockResolvedValue({ ok: true, data: { value: [] } });

    await executeTasks('test-token', { count: 999 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=50'),
      'test-token',
      expect.any(Object),
    );

    await executeTasks('test-token', { count: 0 });
    expect(mockGraphFetch).toHaveBeenCalledWith(
      expect.stringContaining('$top=25'),
      'test-token',
      expect.any(Object),
    );
  });
});
