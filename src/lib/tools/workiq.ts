import { z } from 'zod';
import { truncate, untrusted, echo } from '../format.js';
import {
  callWorkIqTool,
  listWorkIqTools,
  workIqScope,
  READ_ONLY_TOOLS,
  type WorkIqCall,
} from '../workiq.js';
import type { AuthConfig } from '../../types/tokens.js';

export const workIqToolDefinition = {
  name: 'ms_workiq',
  title: 'Work IQ',
  description:
    'Reach the Microsoft Work IQ agent. Without parameters lists the tools Work IQ exposes; with question asks the agent directly; with entity_urls reads Graph paths using your own M365 permissions, which covers directory and channel lookups the other tools are refused. tool + arguments calls any Work IQ tool by name, including ones that create, update, delete or send. Not read-only.',
  inputSchema: z
    .object({
      question: z.string().min(1).optional().describe('A question to put to the Work IQ agent'),
      entity_urls: z
        .array(z.string().min(1))
        .min(1)
        .max(10)
        .optional()
        .describe('Graph paths to read, e.g. ["/me/manager"] (1-10)'),
      tool: z
        .string()
        .min(1)
        .optional()
        .describe('Name of a Work IQ tool to call directly (see the no-parameter listing)'),
      arguments: z
        .record(z.string(), z.unknown())
        .optional()
        .describe('Arguments for the named tool (requires tool)'),
    })
    .refine((a) => !a.arguments || !!a.tool, {
      message: 'arguments requires tool — name the Work IQ tool to pass them to.',
      path: ['arguments'],
    })
    .refine((a) => [a.question, a.entity_urls, a.tool].filter((v) => v !== undefined).length <= 1, {
      message: 'Pass only one of question, entity_urls or tool.',
      path: ['question'],
    }),
  annotations: {
    title: 'Work IQ',
    // Work IQ's scope carries its whole surface. `ask` delegates to an agent and
    // `tool` reaches create/update/delete/do_action, so this cannot be described as
    // read-only however it is usually called — clients should prompt before it runs.
    readOnlyHint: false,
    destructiveHint: true,
    idempotentHint: false,
    openWorldHint: true,
  },
};

/**
 * Pulls readable text out of a Work IQ tool result.
 *
 * `fetch` returns its payload in structuredContent and leaves content empty, while
 * `ask` does the opposite, so both shapes have to be handled.
 */
function renderResult(result: WorkIqCall): string {
  const texts = (result.content ?? [])
    .map((part) => part.text)
    .filter((text): text is string => !!text);
  if (texts.length > 0) {
    return texts.join('\n\n');
  }
  if (result.structuredContent) {
    return JSON.stringify(result.structuredContent, null, 2);
  }
  return '(no content returned)';
}

interface FetchResult {
  statusCode?: number;
  data?: { value?: unknown[] } & Record<string, unknown>;
}

/** Formats the per-path results of the `fetch` tool. */
function renderFetch(paths: string[], result: WorkIqCall): string {
  const results = (result.structuredContent?.['results'] as FetchResult[]) ?? [];
  if (results.length === 0) {
    // Work IQ answered but returned no per-path result. Dumping the empty envelope
    // reads as data; saying nothing came back is what the caller needs to know.
    return `Work IQ returned no results for ${paths.map((p) => `\`${echo(p)}\``).join(', ')}.`;
  }

  return paths
    .map((path, index) => {
      const entry = results[index];
      if (!entry) {
        return `## ${echo(path)}\n\nNo result returned.`;
      }
      const status = entry.statusCode ?? 0;
      const header = `## ${echo(path)}\n\nHTTP ${status}`;
      if (status !== 200) {
        return `${header}\n\nWork IQ could not read this path.`;
      }
      const body = JSON.stringify(entry.data ?? {}, null, 2);
      return `${header}\n\n${untrusted(`Graph ${echo(path)} via Work IQ`, truncate(body, 4000))}`;
    })
    .join('\n\n---\n\n');
}

/**
 * Lists, asks, reads Graph paths, or calls a named Work IQ tool.
 */
export async function executeWorkIq(
  config: AuthConfig,
  args: {
    question?: string;
    entity_urls?: string[];
    tool?: string;
    arguments?: Record<string, unknown>;
  },
): Promise<string> {
  if (!workIqScope()) {
    return [
      'Work IQ is not enabled.',
      'Set MS365_MCP_WORKIQ_SCOPE to the Work IQ scope your app registration has consented, ' +
        'for example api://workiq.svc.cloud.microsoft/WorkIQAgent.Ask.',
      'It is off by default because that scope carries tools that create, update, delete and send.',
    ].join('\n\n');
  }

  try {
    // Mode 1: read Graph paths with the user's own permissions
    if (args.entity_urls) {
      const result = await callWorkIqTool(config, 'fetch', { entityUrls: args.entity_urls });
      return renderFetch(args.entity_urls, result);
    }

    // Mode 2: ask the agent
    if (args.question) {
      const result = await callWorkIqTool(config, 'ask', { question: args.question });
      return untrusted('Work IQ agent response', truncate(renderResult(result), 6000));
    }

    // Mode 3: call a named tool directly
    if (args.tool) {
      const result = await callWorkIqTool(config, args.tool, args.arguments ?? {});
      return untrusted(`Work IQ ${echo(args.tool)} result`, truncate(renderResult(result), 6000));
    }

    // Mode 4: what is available (default)
    const tools = await listWorkIqTools(config);
    if (tools.length === 0) {
      return 'Work IQ returned no tools.';
    }

    const lines = tools.map((tool) => {
      const kind = READ_ONLY_TOOLS.includes(tool.name) ? 'read' : 'CHANGES DATA';
      const description = tool.description ? ` — ${truncate(tool.description, 160)}` : '';
      return `- \`${tool.name}\` (${kind})${description}`;
    });

    return [
      `## Work IQ tools (${tools.length})`,
      lines.join('\n'),
      'Call one with tool and arguments. Tools marked CHANGES DATA can create, update, ' +
        'delete or send on your behalf; `ask` and `call_function` delegate to an agent and ' +
        'may do the same.',
    ].join('\n\n');
  } catch (error) {
    return `Error: ${error instanceof Error ? error.message : String(error)}`;
  }
}
