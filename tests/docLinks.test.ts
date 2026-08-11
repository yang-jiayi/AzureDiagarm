import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { join, relative, sep } from 'node:path';

/**
 * Microsoft Learn keeps retired URLs alive with a redirect for a while and
 * then lets them 404. Every reference here is a link the user is invited to
 * click from a WAF finding or an exported report, so a dead one fails
 * silently and lands on the customer rather than on us.
 *
 * `npm run verify:doc-links` re-checks every URL against the live site. This
 * test is the offline guard that stops a known-dead path from coming back.
 */

const SRC = fileURLToPath(new URL('../src/', import.meta.url));

function collectSourceFiles(dir: string, out: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) collectSourceFiles(full, out);
    else if (/\.tsx?$/.test(entry.name)) out.push(full);
  }
  return out;
}

const LINK_PATTERN = /https:\/\/learn\.microsoft\.com\/[^\s'"`)\\]+/g;

interface FoundLink {
  path: string;
  where: string;
}

function collectLinks(): FoundLink[] {
  const found: FoundLink[] = [];
  for (const file of collectSourceFiles(SRC)) {
    const where = relative(SRC, file).split(sep).join('/');
    readFileSync(file, 'utf8')
      .split('\n')
      .forEach((text, index) => {
        for (const match of text.matchAll(LINK_PATTERN)) {
          found.push({
            path: match[0].replace(/[.,;:]+$/, '').replace('https://learn.microsoft.com', ''),
            where: `${where}:${index + 1}`,
          });
        }
      });
  }
  return found;
}

/**
 * Paths confirmed dead or permanently moved against the live site. The value
 * is the replacement, so the assertion tells the next author what to write
 * instead of sending them back to a search engine.
 */
const RETIRED_PATHS: Record<string, string> = {
  '/azure/architecture/framework': '/azure/well-architected/',
  '/azure/well-architected/reliability/backup-and-recovery':
    '/azure/well-architected/reliability/disaster-recovery',
  '/azure/well-architected/reliability/data-management':
    '/azure/well-architected/reliability/redundancy',
  '/azure/well-architected/reliability/regions-availability-zones':
    '/azure/well-architected/design-guides/regions-availability-zones',
  '/azure/well-architected/reliability/handle-transient-faults':
    '/azure/well-architected/design-guides/handle-transient-faults',
};

test('no source file references a retired Microsoft Learn path', () => {
  const links = collectLinks();
  assert.ok(links.length > 0, 'The link scanner found nothing, so it is no longer scanning anything.');

  const offenders = links.flatMap((link) =>
    Object.entries(RETIRED_PATHS)
      .filter(([retired]) => link.path === retired || link.path.startsWith(`${retired}/`))
      .map(([, replacement]) => `${link.where} uses ${link.path} - use ${replacement} instead`),
  );

  assert.deepEqual(offenders, [], `Retired documentation links:\n${offenders.join('\n')}`);
});

test('user-facing WAF reference links are locale neutral and absolute', () => {
  const source = readFileSync(join(SRC, 'data', 'wafRules.ts'), 'utf8');
  const referenceUrls = [...source.matchAll(/referenceUrl:\s*'([^']+)'/g)].map((match) => match[1]);
  assert.ok(referenceUrls.length > 0, 'wafRules.ts exposes no referenceUrl, so the parser is stale.');

  for (const url of referenceUrls) {
    assert.ok(
      url.startsWith('https://'),
      `${url} must be absolute and https so it survives export to PDF, PPTX and Markdown.`,
    );
    assert.ok(
      !/learn\.microsoft\.com\/[a-z]{2}-[a-z]{2}\//.test(url),
      `${url} pins a locale, which forces Japanese readers onto the English page.`,
    );
  }
});

test('the generated WAF pillar links resolve to real pillar landing pages', () => {
  const source = readFileSync(join(SRC, 'services', 'wafPatternDetector.ts'), 'utf8');
  const slugBlock = source.match(/const slug: Record<WafPillar, string> = \{([\s\S]*?)\};/);
  assert.ok(slugBlock, 'pillarReference no longer builds its URL from a slug map.');

  const slugs = [...slugBlock[1].matchAll(/'([a-z-]+)',/g)].map((match) => match[1]).sort();
  assert.deepEqual(
    slugs,
    ['cost-optimization', 'operational-excellence', 'performance-efficiency', 'reliability', 'security'],
    'The pillar slugs must match the five /azure/well-architected/<pillar>/ landing pages.',
  );
});
