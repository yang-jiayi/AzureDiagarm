/**
 * Re-checks every Microsoft Learn link in src/ against the live site.
 *
 * Deliberately NOT part of `npm test`: it needs network access and Learn's
 * availability is not something a build should depend on. Run it when
 * refreshing content, and fold anything it reports into the RETIRED_PATHS
 * table in tests/docLinks.test.ts so the offline guard catches a regression.
 *
 *   npm run verify:doc-links
 */
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join, relative, sep } from 'node:path';

const SRC = fileURLToPath(new URL('../src/', import.meta.url));
const LINK_PATTERN = /https:\/\/learn\.microsoft\.com\/[^\s'"`)\\]+/g;

function collectSourceFiles(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectSourceFiles(full, out);
    else if (/\.tsx?$/.test(entry.name)) out.push(full);
  }
  return out;
}

interface Candidate {
  url: string;
  where: string;
}

const candidates = new Map<string, Candidate>();
for (const file of collectSourceFiles(SRC)) {
  const where = relative(SRC, file).split(sep).join('/');
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

let failures = 0;

for (const candidate of [...candidates.values()].sort((a, b) => a.url.localeCompare(b.url))) {
  const target = withLocale(candidate.url);
  let status: number | string;
  let location: string | null = null;
  try {
    const response = await fetch(target, { redirect: 'manual', signal: AbortSignal.timeout(30_000) });
    status = response.status;
    location = response.headers.get('location');
  } catch (error) {
    status = error instanceof Error ? error.message : 'request failed';
  }

  if (status === 200) {
    console.log(`  ok   ${candidate.url}`);
    continue;
  }

  failures += 1;
  const detail = location ? ` -> ${location}` : '';
  console.error(`  DEAD ${candidate.url} (${status}${detail}) at ${candidate.where}`);
}

console.log(`\n${candidates.size} link(s) checked, ${failures} failing.`);
if (failures > 0) process.exitCode = 1;
