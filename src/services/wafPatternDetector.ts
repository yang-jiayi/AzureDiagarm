// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * WAF Pattern Detector
 * 
 * Performs deterministic, rule-based analysis of an Azure architecture
 * (services, connections, groups) to detect common architectural patterns
 * and anti-patterns. Returns instant findings without any LLM calls.
 * 
 * This runs in milliseconds and provides a baseline set of findings that
 * can be combined with LLM-generated contextual analysis.
 */

import {
  WafRule,
  ARCHITECTURE_PATTERN_RULES,
  SERVICE_SPECIFIC_RULES,
  WAF_RULE_ENRICHMENTS,
  type WafPillar,
  type Severity,
} from '../data/wafRules';
import type { ValidationFinding } from './architectureValidator';
import { resolveServiceIconMapping } from '../data/serviceIconMapping';

interface ServiceInput {
  name: string;
  type: string;
  category?: string;
  description?: string;
}

interface ConnectionInput {
  from: string;
  to: string;
  label?: string;
}

interface GroupInput {
  name: string;
  services?: string[];
}

interface PatternDetectionResult {
  findings: ValidationFinding[];           // All findings combined
  patternFindings: ValidationFinding[];    // Architecture-level pattern findings only
  serviceFindings: ValidationFinding[];    // Per-service best-practice findings
  patternsDetected: string[];
  serviceRulesApplied: number;
  patternRulesApplied: number;
  elapsedMs: number;
}

function remediationSteps(recommendation: string): string[] {
  return recommendation
    .split(/(?<=[.!?])\s+/)
    .map((step) => step.trim())
    .filter(Boolean);
}

function pillarReference(pillar: WafPillar): string {
  const slug: Record<WafPillar, string> = {
    Reliability: 'reliability',
    Security: 'security',
    'Cost Optimization': 'cost-optimization',
    'Operational Excellence': 'operational-excellence',
    'Performance Efficiency': 'performance-efficiency',
  };
  return `https://learn.microsoft.com/azure/well-architected/${slug[pillar]}/`;
}

function findingFromRule(
  rule: WafRule,
  resources: string[],
  evidence: string[],
): ValidationFinding {
  const enrichment = WAF_RULE_ENRICHMENTS[rule.id];
  return {
    severity: rule.severity,
    category: rule.category,
    issue: rule.issue,
    recommendation: rule.recommendation,
    resources,
    ruleId: rule.id,
    source: 'rule-based',
    evidence: enrichment?.evidence || evidence,
    remediation: enrichment?.remediation || remediationSteps(rule.recommendation),
    referenceUrl: enrichment?.referenceUrl || pillarReference(rule.pillar),
    applyAction: enrichment?.applyAction || {
      type: rule.appliesTo[0] === '*' ? 'regenerate' : 'configure',
      label: rule.appliesTo[0] === '*'
        ? 'Apply this architecture change'
        : `Review ${rule.category} configuration`,
    },
  };
}

// ---------------------------------------------------------------------------
// Service category helpers
// ---------------------------------------------------------------------------

const DATABASE_TYPES = typeSet([
  'sql database', 'azure cosmos db', 'postgresql', 'mysql',
  'azure database for postgresql', 'azure database for mysql',
  'cosmos db', 'cosmosdb', 'redis cache', 'azure cache for redis',
]);

const COMPUTE_TYPES = typeSet([
  'app service', 'functions', 'azure functions', 'virtual machines',
  'kubernetes service', 'azure kubernetes service', 'container apps',
  'azure container apps', 'container instances',
]);

const FRONTEND_TYPES = typeSet([
  'static web apps', 'azure static web apps', 'cdn',
  'content delivery network', 'azure front door',
]);

const CACHE_TYPES = typeSet([
  'redis cache', 'azure cache for redis', 'cdn', 'content delivery network',
]);

const MONITORING_TYPES = typeSet([
  'azure monitor', 'application insights', 'log analytics',
  'app insights',
]);

const IDENTITY_TYPES = typeSet([
  'microsoft entra id', 'entra id', 'azure ad',
  'azure active directory',
]);

const WAF_TYPES = typeSet([
  'web application firewall', 'waf', 'azure waf',
]);

const KEY_VAULT_TYPES = typeSet([
  'key vault', 'azure key vault',
]);

const BACKUP_TYPES = typeSet([
  'azure backup', 'backup', 'recovery services',
]);

const API_GATEWAY_TYPES = typeSet([
  'api management', 'apim', 'azure api management',
  'application gateway', 'azure front door',
]);

/**
 * Canonical service key. The catalog, the palette and every AI prompt emit the
 * official display name ("Azure Kubernetes Service"), while the rule tables
 * were written against the short form ("Kubernetes Service"). Every spelling is
 * first resolved through the shared service catalog — which knows renames such
 * as "Redis Cache" → "Azure Cache for Redis" — and the vendor prefix is then
 * folded away, so both forms land on the same key and the rules actually fire.
 */
function normalizeType(type: string): string {
  const resolved = resolveServiceIconMapping(type)?.serviceName ?? type;
  return resolved
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^(?:microsoft|azure)\s+(?=.)/, '');
}

/** Build a lookup whose members go through the same canonicalisation. */
function typeSet(values: string[]): Set<string> {
  return new Set(values.map(normalizeType));
}

function isType(value: string, ...candidates: string[]): boolean {
  const normalized = normalizeType(value);
  return candidates.some((candidate) => normalizeType(candidate) === normalized);
}

function hasServiceOfType(services: ServiceInput[], typeSet: Set<string>): boolean {
  return services.some(s => typeSet.has(normalizeType(s.type)));
}

function getServicesOfType(services: ServiceInput[], typeSet: Set<string>): ServiceInput[] {
  return services.filter(s => typeSet.has(normalizeType(s.type)));
}

// ---------------------------------------------------------------------------
// Pattern detection functions
// ---------------------------------------------------------------------------

function detectPatterns(
  services: ServiceInput[],
  connections: ConnectionInput[],
): string[] {
  const patterns: string[] = [];

  // Single region — we can't truly detect multi-region from the diagram,
  // but if there's no Traffic Manager / Front Door with multiple backends,
  // it's likely single-region
  const hasGlobalLB = services.some(s => isType(s.type, 'Azure Traffic Manager', 'Azure Front Door'));
  if (!hasGlobalLB && services.length >= 3) {
    patterns.push('single-region');
  }

  // Single database — only one DB service with no replication hint
  const databases = getServicesOfType(services, DATABASE_TYPES);
  if (databases.length === 1) {
    patterns.push('single-database');
  }

  // No caching layer
  if (!hasServiceOfType(services, CACHE_TYPES)) {
    const hasCompute = hasServiceOfType(services, COMPUTE_TYPES);
    const hasDB = databases.length > 0;
    if (hasCompute && hasDB) {
      patterns.push('no-cache');
    }
  }

  // No monitoring
  if (!hasServiceOfType(services, MONITORING_TYPES)) {
    patterns.push('no-monitoring');
  }

  // No identity provider
  if (!hasServiceOfType(services, IDENTITY_TYPES)) {
    patterns.push('no-identity');
  }

  // No WAF — only applies if there are public-facing services
  const hasFrontend = hasServiceOfType(services, FRONTEND_TYPES);
  const hasWebApp = services.some(s => isType(s.type, 'App Service', 'Azure Static Web Apps'));
  if ((hasFrontend || hasWebApp) && !hasServiceOfType(services, WAF_TYPES)) {
    // Application Gateway and Front Door *can* carry a WAF policy, but the
    // diagram cannot prove one is attached — so still advise verifying it
    // rather than silently clearing an internet-facing app.
    patterns.push('no-waf');
  }

  // Direct frontend-to-database access
  const frontendNames = new Set(
    services
      .filter(s => FRONTEND_TYPES.has(normalizeType(s.type)))
      .map(s => s.name.toLowerCase())
  );
  const dbNames = new Set(
    databases.map(s => s.name.toLowerCase())
  );
  for (const conn of connections) {
    if (frontendNames.has(conn.from.toLowerCase()) && dbNames.has(conn.to.toLowerCase())) {
      patterns.push('direct-db-access');
      break;
    }
  }

  // No Key Vault
  if (!hasServiceOfType(services, KEY_VAULT_TYPES) && services.length >= 4) {
    patterns.push('no-key-vault');
  }

  // No backup
  if (!hasServiceOfType(services, BACKUP_TYPES) && databases.length > 0) {
    patterns.push('no-backup');
  }

  // No API gateway — only when there are multiple backend services
  const computeServices = getServicesOfType(services, COMPUTE_TYPES);
  if (computeServices.length >= 2 && !hasServiceOfType(services, API_GATEWAY_TYPES)) {
    patterns.push('no-api-gateway');
  }

  return patterns;
}

// ---------------------------------------------------------------------------
// Main detection entry point
// ---------------------------------------------------------------------------

/**
 * Run rule-based WAF validation against an architecture.
 * Returns deterministic findings in milliseconds (no LLM calls).
 */
export function detectWafPatterns(
  services: ServiceInput[],
  connections: ConnectionInput[],
  _groups?: GroupInput[],
): PatternDetectionResult {
  const startTime = performance.now();
  const patternFindings: ValidationFinding[] = [];
  const serviceFindings: ValidationFinding[] = [];
  const patternsDetected: string[] = [];
  let serviceRulesApplied = 0;
  let patternRulesApplied = 0;

  // 1. Detect architecture-wide patterns
  const patterns = detectPatterns(services, connections);
  patternsDetected.push(...patterns);

  // Map detected patterns to findings
  for (const rule of ARCHITECTURE_PATTERN_RULES) {
    if (rule.pattern && patterns.includes(rule.pattern)) {
      patternRulesApplied++;
      const resources = rule.appliesTo[0] === '*'
        ? getAffectedResources(rule, services, patterns)
        : rule.appliesTo;
      patternFindings.push(findingFromRule(
        rule,
        resources,
        [`Detected topology pattern "${rule.pattern}" affecting ${resources.join(', ') || 'the architecture'}.`],
      ));
    }
  }

  // 2. Apply per-service rules
  for (const service of services) {
    const serviceType = normalizeType(service.type);
    const applicableRules = SERVICE_SPECIFIC_RULES.filter(rule =>
      rule.appliesTo.some(t => normalizeType(t) === serviceType)
    );

    for (const rule of applicableRules) {
      serviceRulesApplied++;
      serviceFindings.push(findingFromRule(
        rule,
        [service.name],
        [
          `${service.name} (${service.type}) is present, but a diagram cannot verify whether the ${rule.category.toLowerCase()} configuration is enabled.`,
        ],
      ));
    }
  }

  const findings = [...patternFindings, ...serviceFindings];
  const elapsedMs = Math.round(performance.now() - startTime);

  console.log(`⚡ WAF pattern detection: ${findings.length} findings in ${elapsedMs}ms`);
  console.log(`  📋 Patterns: ${patternsDetected.join(', ') || 'none'}`);
  console.log(`  🔧 Rules applied: ${patternRulesApplied} pattern + ${serviceRulesApplied} service-specific`);

  return {
    findings,
    patternFindings,
    serviceFindings,
    patternsDetected,
    serviceRulesApplied,
    patternRulesApplied,
    elapsedMs,
  };
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Determine which services are affected by an architecture-wide pattern.
 */
function getAffectedResources(
  rule: WafRule,
  services: ServiceInput[],
  _patterns: string[],
): string[] {
  switch (rule.pattern) {
    case 'single-database':
      return getServicesOfType(services, DATABASE_TYPES).map(s => s.name);

    case 'no-cache':
      return [
        ...getServicesOfType(services, COMPUTE_TYPES).map(s => s.name),
        ...getServicesOfType(services, DATABASE_TYPES).map(s => s.name),
      ];

    case 'no-monitoring':
      return services.map(s => s.name); // Affects everything

    case 'no-identity':
      return getServicesOfType(services, COMPUTE_TYPES).map(s => s.name);

    case 'no-waf':
      return services
        .filter(s => FRONTEND_TYPES.has(normalizeType(s.type)) || isType(s.type, 'App Service'))
        .map(s => s.name);

    case 'direct-db-access':
      return [
        ...services.filter(s => FRONTEND_TYPES.has(normalizeType(s.type))).map(s => s.name),
        ...getServicesOfType(services, DATABASE_TYPES).map(s => s.name),
      ];

    case 'no-key-vault':
      return getServicesOfType(services, COMPUTE_TYPES).map(s => s.name);

    case 'no-backup':
      return getServicesOfType(services, DATABASE_TYPES).map(s => s.name);

    case 'no-api-gateway':
      return getServicesOfType(services, COMPUTE_TYPES).map(s => s.name);

    default:
      return services.map(s => s.name);
  }
}

/**
 * Calculate a preliminary score based on rule-based findings alone.
 * This gives an instant approximation before the LLM refines it.
 */
export function calculatePreliminaryScore(findings: ValidationFinding[]): number {
  if (findings.length === 0) return 95;

  // Point deductions by severity
  const deductions: Record<Severity, number> = {
    critical: 12,
    high: 7,
    medium: 3,
    low: 1,
  };

  let totalDeduction = 0;
  for (const finding of findings) {
    totalDeduction += deductions[finding.severity] || 2;
  }

  // Cap at a minimum score of 10
  return Math.max(10, 100 - totalDeduction);
}

/**
 * Group findings by WAF pillar for display.
 */
export function groupFindingsByPillar(
  findings: ValidationFinding[],
): Record<WafPillar, ValidationFinding[]> {
  // Map finding categories back to pillars using the rules
  const categoryToPillar = new Map<string, WafPillar>();
  for (const rule of [...ARCHITECTURE_PATTERN_RULES, ...SERVICE_SPECIFIC_RULES]) {
    categoryToPillar.set(rule.category, rule.pillar);
  }

  const grouped: Record<WafPillar, ValidationFinding[]> = {
    'Reliability': [],
    'Security': [],
    'Cost Optimization': [],
    'Operational Excellence': [],
    'Performance Efficiency': [],
  };

  for (const finding of findings) {
    const pillar = categoryToPillar.get(finding.category) || 'Operational Excellence';
    grouped[pillar].push(finding);
  }

  return grouped;
}
