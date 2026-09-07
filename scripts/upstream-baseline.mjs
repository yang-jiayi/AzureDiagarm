import { execFileSync } from 'node:child_process';
import { writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

export const baselinePath = '.github/upstream-baseline.json';
export const upstreamRepository = 'https://github.com/Arturo-Quiroga-MSFT/azure-architecture-diagram-builder.git';
const isCommit = (value) => typeof value === 'string' && /^[0-9a-f]{40}$/.test(value);
const git = (...args) => execFileSync('git', args, {
  encoding: 'utf8',
  stdio: ['ignore', 'pipe', 'pipe'],
}).trim();

export function readApprovedBaseline(runGit = git) {
  let record;
  try {
    record = JSON.parse(runGit('show', `HEAD:${baselinePath}`));
  } catch {
    throw new Error('Missing or invalid committed upstream baseline; review .github/upstream-baseline.json.');
  }
  if (record?.repository !== upstreamRepository || !isCommit(record?.commit)) {
    throw new Error('Invalid upstream baseline repository or full commit SHA; refusing source selection.');
  }
  try {
    if (runGit('cat-file', '-t', record.commit) !== 'commit') throw new Error();
    runGit('merge-base', '--is-ancestor', record.commit, 'HEAD');
  } catch {
    throw new Error('The approved upstream baseline is not a commit in this checkout ancestry; refusing source selection.');
  }
  return record.commit;
}

export function recordSelectedBaseline(commit, runGit = git, write = writeFileSync) {
  if (!isCommit(commit)) throw new Error('The selected upstream SHA must be a full commit.');
  if (runGit('remote', 'get-url', 'upstream') !== upstreamRepository) {
    throw new Error('The selected upstream remote does not match the approved repository.');
  }
  if (runGit('rev-parse', '--verify', 'refs/remotes/upstream/main^{commit}') !== commit) {
    throw new Error('The selected SHA is not the fetched upstream/main commit.');
  }
  let mergeHead;
  try {
    mergeHead = runGit('rev-parse', '--quiet', '--verify', 'MERGE_HEAD');
  } catch (error) {
    if (!error || typeof error !== 'object' || !('status' in error) || error.status !== 1) {
      throw error;
    }
    // An already-integrated upstream tip can still pass explicit sync review.
    runGit('merge-base', '--is-ancestor', commit, 'HEAD');
  }
  if (mergeHead !== undefined && mergeHead !== commit) {
    throw new Error('The pending merge does not match the selected upstream commit.');
  }
  write(baselinePath, `${JSON.stringify({ repository: upstreamRepository, commit }, null, 2)}\n`, 'utf8');
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    if (process.argv.length === 2) {
      console.log(readApprovedBaseline());
    } else if (process.argv.length === 4 && process.argv[2] === '--record') {
      recordSelectedBaseline(process.argv[3]);
    } else {
      throw new Error('Usage: node scripts/upstream-baseline.mjs [--record <selected-upstream-sha>]');
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : 'Upstream provenance verification failed.');
    process.exitCode = 1;
  }
}
