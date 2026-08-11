/**
 * Re-checks every Microsoft Learn link in src/ against the live site.
 *
 * Deliberately NOT part of `npm test`: it needs network access and Learn's
 * availability is not something a build should depend on. Run it when
 * refreshing content, and fold anything it reports as DEAD into the
 * RETIRED_PATHS table in tests/docLinks.test.ts so the offline guard catches a
 * regression. Anything reported as ERROR is a network problem on this machine,
 * not a dead link — never add those to the table.
 *
 *   npm run verify:doc-links
 */
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join, relative, sep } from 'node:path';

const REPO = fileURLToPath(new URL('../', import.meta.url));
const SCANNED_DIRS = ['src', 'scripts'];
const LINK_PATTERN = /https:\/\/learn\.microsoft\.com\/[^\s'"`)\\]+/g;

function collectSourceFiles(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectSourceFiles(full, out);
    else if (/\.(tsx?|[cm]?js)$/.test(entry.name)) out.push(full);
  }
  return out;
}

interface Candidate {
  url: string;
  where: string;
}

const candidates = new Map<string, Candidate>();
for (const dir of SCANNED_DIRS) {
  for (const file of collectSourceFiles(join(REPO, dir))) {
    const where = relative(REPO, file).split(sep).join('/');
    readFileSync(file, 'utf8')
      .split('\n')
      .forEach((text, index) => {
        for (const match of text.matchAll(LINK_PATTERN)) {
          const url = match[0].replace(/[.,;:]+$/, '');
          // Skip template literals: the slug is only known at runtime.
          if (url.includes('${')) continue;
          if (!candidates.has(url)) candidates.set(url, { url, where: `${where}:${index + 1}` });
        }
      });
  }
}

/**
 * Learn 302s every locale-neutral URL to /en-us/ before it can tell us whether
 * the page exists, so ask for the locale directly and treat any further
 * redirect as a real move.
 */
function withLocale(url: string): string {
  return /learn\.microsoft\.com\/[a-z]{2}-[a-z]{2}\//.test(url)
    ? url
    : url.replace('learn.microsoft.com/', 'learn.microsoft.com/en-us/');
}

let dead = 0;
let unreachable = 0;

for (const candidate of [...candidates.values()].sort((a, b) => a.url.localeCompare(b.url))) {
  const target = withLocale(candidate.url);
  let status: number | null = null;
  let location: string | null = null;
  let transportError: string | null = null;
  try {
    const response = await fetch(target, { redirect: 'manual', signal: AbortSignal.timeout(30_000) });
    status = response.status;
    location = response.headers.get('location');
  } catch (error) {
    transportError = error instanceof Error ? error.message : 'request failed';
  }

  if (status === 200) {
    console.log(`  ok    ${candidate.url}`);
    continue;
  }

  // A timeout, DNS failure or proxy block says nothing about the link. Keeping
  // it distinct matters because the reported paths get added to the offline
  // RETIRED_PATHS table, and a transient failure must not ban a live page.
  if (transportError) {
    unreachable += 1;
    console.error(`  ERROR ${candidate.url} could not be reached (${transportError}) at ${candidate.where}`);
    continue;
  }

  dead += 1;
  console.error(`  DEAD  ${candidate.url} (${status}${location ? ` -> ${location}` : ''}) at ${candidate.where}`);
}

console.log(
  `\n${candidates.size} link(s) checked, ${dead} dead`
  + (unreachable > 0 ? `, ${unreachable} unreachable (network, not the link)` : '')
  + '.',
);
if (dead > 0 || unreachable > 0) process.exitCode = 1;
