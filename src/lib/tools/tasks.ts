import { z } from 'zod';
import { graphFetch } from '../graph.js';

export const tasksToolDefinition = {
  name: 'ms_tasks',
  title: 'To Do & Planner',
  description:
    'Read Microsoft To Do and Planner tasks. Without parameters lists To Do task lists; with list_id returns the tasks in that list; with planner=true returns Planner tasks assigned to the user.',
  inputSchema: z.object({
    list_id: z.string().optional().describe('To Do list ID to read its tasks'),
    planner: z
      .boolean()
      .optional()
      .describe('Return Planner tasks assigned to the user instead of To Do lists'),
    include_completed: z
      .boolean()
      .optional()
      .describe('Include completed tasks (default false \u2014 only open tasks are returned)'),
    count: z.int().min(1).max(50).optional().describe('Max results to return (1-50, default 25)'),
  }),
  annotations: {
    title: 'To Do & Planner',
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: true,
  },
};

interface TodoList {
  id?: string;
  displayName?: string;
  wellknownListName?: string;
  isShared?: boolean;
}

interface TodoListsResponse {
  value: TodoList[];
}

interface TodoTask {
  id?: string;
  title?: string;
  status?: string;
  importance?: string;
  createdDateTime?: string;
  dueDateTime?: { dateTime?: string };
  completedDateTime?: { dateTime?: string };
  body?: { content?: string };
}

interface TodoTasksResponse {
  value: TodoTask[];
}

interface PlannerTask {
  id?: string;
  title?: string;
  planId?: string;
  bucketId?: string;
  percentComplete?: number;
  priority?: number;
  dueDateTime?: string;
  createdDateTime?: string;
}

interface PlannerTasksResponse {
  value: PlannerTask[];
}

/**
 * Formats a To Do list into readable text.
 */
function formatTodoList(list: TodoList): string {
  const lines: string[] = [];
  lines.push(`## ${list.displayName || 'Unnamed List'}${list.isShared ? ' (shared)' : ''}`);
  if (list.id) {
    lines.push(`List ID: ${list.id}`);
  }
  if (list.wellknownListName && list.wellknownListName !== 'none') {
    lines.push(`Well-known list: ${list.wellknownListName}`);
  }
  return lines.join('\n');
}

/**
 * Formats a To Do task into readable text.
 */
function formatTodoTask(task: TodoTask): string {
  const lines: string[] = [];
  const done = task.status === 'completed';
  lines.push(`## ${done ? '[x]' : '[ ]'} ${task.title || 'Untitled task'}`);
  if (task.status) {
    lines.push(`Status: ${task.status}`);
  }
  if (task.importance && task.importance !== 'normal') {
    lines.push(`Importance: ${task.importance}`);
  }
  if (task.dueDateTime?.dateTime) {
    lines.push(`Due: ${new Date(task.dueDateTime.dateTime).toLocaleString()}`);
  }
  if (task.completedDateTime?.dateTime) {
    lines.push(`Completed: ${new Date(task.completedDateTime.dateTime).toLocaleString()}`);
  }
  const body = (task.body?.content || '').trim();
  if (body) {
    lines.push('');
    lines.push(body);
  }
  if (task.id) {
    lines.push(`Task ID: ${task.id}`);
  }
  return lines.join('\n');
}

/**
 * Formats a Planner task into readable text.
 */
function formatPlannerTask(task: PlannerTask): string {
  const lines: string[] = [];
  const pct = task.percentComplete ?? 0;
  lines.push(`## ${pct === 100 ? '[x]' : '[ ]'} ${task.title || 'Untitled task'}`);
  lines.push(`Progress: ${pct}%`);
  if (task.dueDateTime) {
    lines.push(`Due: ${new Date(task.dueDateTime).toLocaleString()}`);
  }
  if (task.planId) {
    lines.push(`Plan ID: ${task.planId}`);
  }
  if (task.id) {
    lines.push(`Task ID: ${task.id}`);
  }
  return lines.join('\n');
}

/**
 * Reads To Do lists, To Do tasks, or assigned Planner tasks
 * depending on which parameters are provided.
 */
export async function executeTasks(
  token: string,
  args: {
    list_id?: string;
    planner?: boolean;
    include_completed?: boolean;
    count?: number;
  },
): Promise<string> {
  const count = Math.min(Math.max(args.count || 25, 1), 50);

  // Mode 1: Planner tasks assigned to the user
  if (args.planner) {
    const result = await graphFetch<PlannerTasksResponse>('/me/planner/tasks', token, {
      timezone: false,
    });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    let tasks = result.data.value || [];
    if (!args.include_completed) {
      tasks = tasks.filter((t) => (t.percentComplete ?? 0) < 100);
    }
    if (tasks.length === 0) {
      return args.include_completed
        ? 'No Planner tasks assigned to you.'
        : 'No open Planner tasks assigned to you. Pass include_completed to see finished ones.';
    }

    return tasks.slice(0, count).map(formatPlannerTask).join('\n\n');
  }

  // Mode 2: tasks within a To Do list
  if (args.list_id) {
    const filter = args.include_completed
      ? ''
      : `&$filter=${encodeURIComponent("status ne 'completed'")}`;
    const path = `/me/todo/lists/${encodeURIComponent(args.list_id)}/tasks?$top=${count}${filter}`;
    const result = await graphFetch<TodoTasksResponse>(path, token, { timezone: false });

    if (!result.ok) {
      return `Error: ${result.error.message}`;
    }

    const tasks = result.data.value;
    if (!tasks || tasks.length === 0) {
      return args.include_completed
        ? 'No tasks found in this list.'
        : 'No open tasks in this list. Pass include_completed to see finished ones.';
    }

    return tasks.map(formatTodoTask).join('\n\n');
  }

  // Mode 3: To Do lists (default)
  const result = await graphFetch<TodoListsResponse>(`/me/todo/lists?$top=${count}`, token, {
    timezone: false,
  });

  if (!result.ok) {
    return `Error: ${result.error.message}`;
  }

  const lists = result.data.value;
  if (!lists || lists.length === 0) {
    return 'No To Do lists found.';
  }

  return lists.map(formatTodoList).join('\n\n');
}
