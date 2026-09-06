import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const productionRef = 'refs/heads/main';
const isCommit = (value) => typeof value === 'string' && /^[0-9a-f]{40}$/.test(value);

export async function verifyDeploymentSource(
  { ref, repository, expectedCommit, checkedOutCommit, token },
  fetchMain = fetch,
) {
  if (ref !== productionRef) {
    throw new Error('Production deployment requires refs/heads/main; refusing deployment.');
  }
  if (!isCommit(expectedCommit)) {
    throw new Error('The validated deployment commit is missing or malformed; refusing deployment.');
  }
  if (!isCommit(checkedOutCommit) || checkedOutCommit !== expectedCommit) {
    throw new Error('The checkout does not match the validated deployment commit; refusing deployment.');
  }
  if (
    typeof repository !== 'string' ||
    !/^[A-Za-z0-9_-]+\/[A-Za-z0-9_.-]+$/.test(repository) ||
    ['.', '..'].includes(repository.split('/')[1])
  ) {
    throw new Error('The GitHub repository is missing or malformed; refusing deployment.');
  }
  if (typeof token !== 'string' || !token.trim()) {
    throw new Error('A read-only GitHub token is required to verify current main; refusing deployment.');
  }

  let response;
  try {
    response = await fetchMain(
      `https://api.github.com/repos/${repository}/git/ref/heads/main`,
      {
        headers: {
          Accept: 'application/vnd.github+json',
          Authorization: `Bearer ${token}`,
          'User-Agent': 'AzureDiagarm-release-source-guard',
          'X-GitHub-Api-Version': '2022-11-28',
          'Cache-Control': 'no-cache',
        },
        cache: 'no-store',
        redirect: 'error',
        signal: AbortSignal.timeout(15_000),
      },
    );
  } catch {
    throw new Error('Unable to read current main from GitHub; refusing deployment.');
  }
  if (!response.ok) {
    throw new Error(`GitHub main lookup failed (HTTP ${response.status}); refusing deployment.`);
  }

  let main;
  try {
    main = await response.json();
  } catch {
    throw new Error('GitHub returned an invalid main ref response; refusing deployment.');
  }
  if (
    main?.ref !== productionRef ||
    main?.object?.type !== 'commit' ||
    !isCommit(main?.object?.sha)
  ) {
    throw new Error('GitHub returned a missing or malformed main commit; refusing deployment.');
  }
  if (main.object.sha !== expectedCommit) {
    throw new Error(
      `Validated commit ${expectedCommit} is no longer current main (${main.object.sha}); refusing deployment. Use a release from current main, not a historical run.`,
    );
  }

  return main.object.sha;
}

const invokedPath = process.argv[1] ? resolve(process.argv[1]) : '';
if (invokedPath === fileURLToPath(import.meta.url)) {
  try {
    let checkedOutCommit;
    try {
      checkedOutCommit = execFileSync('git', ['rev-parse', '--verify', 'HEAD^{commit}'], {
        encoding: 'utf8',
        stdio: ['ignore', 'pipe', 'pipe'],
      }).trim();
    } catch {
      throw new Error('Unable to read the checked-out commit; refusing deployment.');
    }
    const commit = await verifyDeploymentSource({
      ref: process.env.GITHUB_REF,
      repository: process.env.GITHUB_REPOSITORY,
      expectedCommit: process.env.EXPECTED_COMMIT,
      checkedOutCommit,
      token: process.env.GITHUB_TOKEN,
    });
    console.log(`Verified production source ${commit} is current main.`);
  } catch (error) {
    console.error(error instanceof Error ? error.message : 'Production source verification failed.');
    process.exitCode = 1;
  }
}
