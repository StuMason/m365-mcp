/**
 * Writes dist/build-info.json so ms_server_info can identify the exact build.
 *
 * The version alone cannot distinguish two builds of the same unreleased version,
 * which makes "is this the fix I just made?" unanswerable during a review cycle.
 * Git details are best-effort: a published package has no repository.
 */
import { execFileSync } from 'node:child_process';
import { writeFileSync, mkdirSync } from 'node:fs';

function git(args) {
  try {
    return execFileSync('git', args, {
      encoding: 'utf-8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
  } catch {
    return undefined;
  }
}

const info = {
  builtAt: new Date().toISOString(),
  commit: git(['rev-parse', '--short', 'HEAD']),
  branch: git(['rev-parse', '--abbrev-ref', 'HEAD']),
  dirty: git(['status', '--porcelain']) ? true : false,
};

mkdirSync('dist', { recursive: true });
writeFileSync('dist/build-info.json', JSON.stringify(info, null, 2));
