import assert from 'node:assert/strict';
import test from 'node:test';

import { hardenArchitecture } from '../dist/hardener.js';
import {
  detectWafPatterns,
  getWafRules,
  groupFindingsByPillar,
} from '../dist/wafDetector.js';

const vulnerableServices = [
  { name: 'Web', type: 'Azure Static Web Apps' },
  { name: 'API', type: 'App Service' },
  { name: 'Database', type: 'SQL Database' },
];

const vulnerableConnections = [
  { from: 'Web', to: 'Database', label: 'direct data access' },
  { from: 'Web', to: 'API', label: 'api' },
  { from: 'API', to: 'Database', label: 'queries' },
];

test('WAF detector identifies concrete topology and service risks', () => {
  const result = detectWafPatterns(vulnerableServices, vulnerableConnections);

  assert.ok(result.patternsDetected.includes('single-region'));
  assert.ok(result.patternsDetected.includes('single-database'));
  assert.ok(result.patternsDetected.includes('no-monitoring'));
  assert.ok(result.patternsDetected.includes('no-identity'));
  assert.ok(result.patternsDetected.includes('no-waf'));
  assert.ok(result.patternsDetected.includes('direct-db-access'));
  assert.ok(result.patternFindings.length > 0);
  assert.ok(result.serviceFindings.some((finding) => finding.resources?.includes('API')));
  assert.ok(result.score < 100);
});

test('architecture hardener improves the score and removes direct database access', () => {
  const result = hardenArchitecture(vulnerableServices, vulnerableConnections);
  const serviceTypes = new Set(result.services.map((service) => service.type));

  assert.ok(result.after.score > result.before.score);
  assert.ok(serviceTypes.has('Microsoft Entra ID'));
  assert.ok(serviceTypes.has('Azure Front Door'));
  assert.ok(serviceTypes.has('Web Application Firewall'));
  assert.ok(serviceTypes.has('Application Insights'));
  assert.ok(serviceTypes.has('Azure Monitor'));
  assert.ok(serviceTypes.has('Backup'));
  assert.equal(
    result.connections.some((connection) => (
      connection.from === 'Web' && connection.to === 'Database'
    )),
    false,
  );

  const repeated = hardenArchitecture(result.services, result.connections, result.groups);
  assert.equal(repeated.changes.length, 0);
  assert.deepEqual(repeated.after.patternsDetected, result.after.patternsDetected);
});

test('WAF rules and grouped findings preserve pillar classification', () => {
  const result = detectWafPatterns(vulnerableServices, vulnerableConnections);
  const grouped = groupFindingsByPillar(result.findings);
  const securityRules = getWafRules('Security');

  assert.ok(securityRules.length > 0);
  assert.ok(grouped.Security.length > 0);
  assert.equal(
    Object.values(grouped).reduce((total, findings) => total + findings.length, 0),
    result.findings.length,
  );
});
