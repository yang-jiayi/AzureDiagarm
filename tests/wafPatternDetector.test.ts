import test from 'node:test';
import assert from 'node:assert/strict';
import { detectWafPatterns } from '../src/services/wafPatternDetector.ts';

/**
 * The palette, the icon catalog and every AI prompt emit the official display
 * name ("Azure Kubernetes Service"), while the rule tables were authored with
 * the short form ("Kubernetes Service"). If the two ever stop resolving to the
 * same key the security rules go silently dead and the report shows a clean
 * bill of health for the highest-risk services.
 */
const svc = (name: string, type: string, category = 'compute') => ({ name, type, category });

test('service-specific rules fire for the canonical Azure-prefixed names', () => {
  const cases: Array<{ type: string; expect: string }> = [
    { type: 'Azure Kubernetes Service', expect: 'aks-' },
    { type: 'Azure SQL Database', expect: 'sql-' },
    { type: 'Azure Functions', expect: 'func-' },
    { type: 'Azure Cache for Redis', expect: 'redis-' },
    { type: 'Azure Container Apps', expect: 'aca-' },
  ];
  for (const item of cases) {
    const result = detectWafPatterns([svc('node-1', item.type)], []);
    const ids = result.findings.map((finding) => finding.ruleId);
    assert.ok(
      ids.some((id) => id.startsWith(item.expect)),
      `${item.type} produced no ${item.expect}* finding (got: ${ids.join(', ') || 'none'})`,
    );
  }
});

test('the short and Azure-prefixed spellings produce identical findings', () => {
  const long = detectWafPatterns([svc('a', 'Azure Kubernetes Service')], []);
  const short = detectWafPatterns([svc('a', 'Kubernetes Service')], []);
  assert.deepEqual(
    long.findings.map((f) => f.ruleId).sort(),
    short.findings.map((f) => f.ruleId).sort(),
  );
});

test('an internet-facing app with no WAF is still flagged behind Front Door', () => {
  const result = detectWafPatterns(
    [svc('edge', 'Azure Front Door', 'networking'), svc('web', 'App Service')],
    [],
  );
  assert.ok(
    result.findings.some((finding) => /waf/i.test(finding.ruleId)),
    'Front Door presence must not silently clear the missing-WAF advisory',
  );
});
