import { spawnSync } from 'node:child_process';

import { verifyImageSizeSecurityPatch } from './patch-image-size.mjs';

const patchedAdvisories = new Set([
  'https://github.com/advisories/GHSA-5p2g-fcmc-qvqq',
  'https://github.com/advisories/GHSA-w3rx-r6r6-pgpr',
]);
const blockingSeverities = new Set(['high', 'critical']);
const npmExecutable = process.env.npm_execpath;
if (!npmExecutable) {
  throw new Error('Run this audit through `npm run audit:production`.');
}

const result = spawnSync(
  process.execPath,
  [npmExecutable, 'audit', '--omit=dev', '--audit-level=high', '--json'],
  {
    cwd: process.cwd(),
    encoding: 'utf8',
    maxBuffer: 10 * 1024 * 1024,
  },
);

if (result.error) {
  throw result.error;
}

let report;
try {
  report = JSON.parse(result.stdout);
} catch {
  const diagnostic = result.stderr || result.stdout || 'npm audit returned no diagnostic output.';
  throw new Error(`Unable to parse npm audit output:\n${diagnostic}`);
}
if (report.auditReportVersion !== 2 || report.error) {
  const diagnostic = report.error?.summary ?? report.error?.message ?? 'Unexpected npm audit report.';
  throw new Error(`npm audit did not complete successfully: ${diagnostic}`);
}

const vulnerabilities = report.vulnerabilities ?? {};
const imageSizeFinding = vulnerabilities['image-size'];
const pptxFinding = vulnerabilities.pptxgenjs;
const imageSizeOnlyHasPatchedAdvisories =
  Array.isArray(imageSizeFinding?.via) &&
  imageSizeFinding.via.length === patchedAdvisories.size &&
  imageSizeFinding.via.every(
    (entry) =>
      typeof entry === 'object' &&
      entry !== null &&
      typeof entry.url === 'string' &&
      patchedAdvisories.has(entry.url),
  );
const pptxOnlyInheritsImageSizeFinding =
  Array.isArray(pptxFinding?.via) &&
  pptxFinding.via.length === 1 &&
  pptxFinding.via[0] === 'image-size';

const ignoredPackages = new Set();
if (imageSizeOnlyHasPatchedAdvisories && (!pptxFinding || pptxOnlyInheritsImageSizeFinding)) {
  const patchFailures = verifyImageSizeSecurityPatch();
  if (patchFailures.length > 0) {
    throw new Error(
      `The image-size advisories cannot be accepted because the local safeguards are missing from: ${patchFailures.join(', ')}`,
    );
  }
  ignoredPackages.add('image-size');
  if (pptxFinding) {
    ignoredPackages.add('pptxgenjs');
  }
}

const remainingFindings = Object.entries(vulnerabilities).filter(
  ([name, finding]) =>
    !ignoredPackages.has(name) &&
    finding &&
    blockingSeverities.has(finding.severity),
);

if (remainingFindings.length > 0) {
  for (const [name, finding] of remainingFindings) {
    console.error(`${name}: ${finding.severity}`);
    for (const via of finding.via ?? []) {
      console.error(`  - ${typeof via === 'string' ? via : `${via.title} (${via.url})`}`);
    }
  }
  process.exit(result.status || 1);
}

if (ignoredPackages.size > 0) {
  console.log(
    'Accepted only GHSA-5p2g-fcmc-qvqq and GHSA-w3rx-r6r6-pgpr after verifying the local parser safeguards.',
  );
}
console.log('No unmitigated high or critical production dependency vulnerabilities found.');
