// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Architecture Validator Agent
 * Uses GPT-5-2 to validate architecture against Azure Well-Architected Framework
 * Provides recommendations for reliability, security, performance, cost optimization, and operational excellence
 */

import { ModelType, ReasoningEffort } from '../stores/modelSettingsStore';
import { detectWafPatterns, calculatePreliminaryScore } from './wafPatternDetector';
import { getKnowledgeBaseStats } from '../data/wafRules';
import { scoreToBand } from './wafMaturity';
import { trackAIModelUsage } from './telemetryService';
import {
  buildRequestBody,
  parseApiResponse,
  callAzureOpenAIProxy,
  createOpenAIProxyError,
  getApiFormatLabel,
} from './apiHelper';
import type { Language } from '../i18n/LanguageContext';
import { getPromptLanguageInstruction } from '../i18n/localization';
import {
  resolveAIModelRuntime,
  type RuntimeModelOverride,
} from './aiModelRuntime';
import { safeParseModelJson } from './aiRetry';

export interface ValidationModelOverride extends RuntimeModelOverride {
  model: ModelType;
  reasoningEffort: ReasoningEffort;
}

// Token usage metrics returned from Azure OpenAI API
export interface AIMetrics {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  elapsedTimeMs: number;
  model?: string;
}

interface CallResult {
  content: string;
  metrics: AIMetrics;
}

async function callAzureOpenAI(messages: any[], maxTokens: number = 8000, modelOverride?: ValidationModelOverride): Promise<CallResult> {
  const runtime = resolveAIModelRuntime('validation', modelOverride);
  console.log(`🌐 Calling AI model service with ${runtime.displayName} | API: ${getApiFormatLabel(runtime.apiFormat)}`);
  
  // Start timing
  const startTime = performance.now();

  // Build request body using the appropriate API format
  const effectiveMaxTokens = Math.min(maxTokens, runtime.maxCompletionTokens);
  const requestBody = buildRequestBody({
    deployment: runtime.deployment,
    messages,
    maxTokens: effectiveMaxTokens,
    apiFormat: runtime.apiFormat,
    isReasoning: runtime.isReasoning,
    reasoningEffort: runtime.reasoningEffort,
  });
  
  console.log(`🤖 Using ${runtime.displayName}${runtime.isReasoning ? ` (reasoning: ${runtime.reasoningEffort})` : ''} | max_tokens: ${effectiveMaxTokens} | API: ${getApiFormatLabel(runtime.apiFormat)}`);

  // Client-side timeout keeps the documented chain intact: client 225s >
  // proxy 210s > Front Door 240s. Without it this path could hang until the
  // platform kills it with an opaque error.
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 225000);
  let proxyResult;
  try {
    proxyResult = await callAzureOpenAIProxy({
      apiFormat: runtime.apiFormat,
      deployment: runtime.deployment,
      body: requestBody,
      byo: runtime.byo,
      signal: controller.signal,
    });
  } catch (error: any) {
    if (error?.name === 'AbortError') {
      throw new Error('The AI provider is taking too long to respond. Please try again.');
    }
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
  
  // Calculate elapsed time
  const elapsedTimeMs = Math.round(performance.now() - startTime);

  if (!proxyResult.ok) {
    console.error('❌ Azure OpenAI API error:', {
      status: proxyResult.status,
      code: proxyResult.error?.code,
      source: proxyResult.error?.source,
      requestId: proxyResult.error?.requestId,
      upstreamRequestId: proxyResult.error?.upstreamRequestId,
    });
    throw createOpenAIProxyError(proxyResult);
  }
  
  // Parse response using the appropriate API format
  const parsed = parseApiResponse(proxyResult.data, runtime.apiFormat);
  const metrics: AIMetrics = {
    promptTokens: parsed.promptTokens,
    completionTokens: parsed.completionTokens,
    totalTokens: parsed.totalTokens,
    elapsedTimeMs,
    model: runtime.displayName,
  };
  
  const content = parsed.content;
  
  console.log('📦 API Response:', content.length, 'chars |',
    `Tokens: ${metrics.promptTokens} in → ${metrics.completionTokens} out (${metrics.totalTokens} total) |`,
    `Time: ${(metrics.elapsedTimeMs / 1000).toFixed(2)}s`);
  
  // Track model usage telemetry
  trackAIModelUsage({
    model: runtime.telemetryModel,
    operation: 'architecture_validation',
    reasoningEffort: runtime.isReasoning ? runtime.reasoningEffort : undefined,
    promptTokens: metrics.promptTokens,
    completionTokens: metrics.completionTokens,
    totalTokens: metrics.totalTokens,
    elapsedTimeMs: metrics.elapsedTimeMs,
  });
  
  return { content, metrics };
}

export interface ValidationResult {
  score: number; // 0-100
  pillar: 'Reliability' | 'Security' | 'Cost Optimization' | 'Operational Excellence' | 'Performance Efficiency';
  findings: ValidationFinding[];
}

export interface ValidationFinding {
  severity: 'critical' | 'high' | 'medium' | 'low';
  category: string;
  issue: string;
  recommendation: string;
  resources?: string[];
  ruleId?: string;
  source?: 'rule-based' | 'ai-analysis';
  evidence?: string[];
  remediation?: string[];
  referenceUrl?: string;
  applyAction?: {
    type: 'add-service' | 'regenerate' | 'configure';
    label: string;
    serviceType?: string;
  };
}

export interface ArchitectureValidation {
  overallScore: number;
  summary: string;
  pillars: ValidationResult[];
  quickWins: ValidationFinding[];
  timestamp: string;
  metrics?: AIMetrics;
  modelUsed?: string;
  diagramImageDataUrl?: string;
}

const FINDING_SEVERITIES = new Set<ValidationFinding['severity']>([
  'critical',
  'high',
  'medium',
  'low',
]);
const APPLY_ACTION_TYPES = new Set<NonNullable<ValidationFinding['applyAction']>['type']>([
  'add-service',
  'regenerate',
  'configure',
]);

function boundedString(value: unknown, maxLength: number): string {
  return typeof value === 'string' ? value.trim().slice(0, maxLength) : '';
}

function boundedStringArray(value: unknown, maxItems: number, maxLength: number): string[] {
  if (!Array.isArray(value)) return [];
  return value
    .map((item) => boundedString(item, maxLength))
    .filter(Boolean)
    .slice(0, maxItems);
}

function recommendationSteps(recommendation: string): string[] {
  return recommendation
    .split(/(?<=[.!?。！？])\s*/)
    .map((step) => step.trim())
    .filter(Boolean)
    .slice(0, 8);
}

function normalizeFinding(
  value: unknown,
  localByRuleId: Map<string, ValidationFinding>,
  serviceNames: Set<string>,
): ValidationFinding | null {
  if (!value || typeof value !== 'object') return null;
  const raw = value as Record<string, unknown>;
  const ruleId = boundedString(raw.ruleId, 100);
  const local = ruleId ? localByRuleId.get(ruleId) : undefined;
  const severity = FINDING_SEVERITIES.has(raw.severity as ValidationFinding['severity'])
    ? raw.severity as ValidationFinding['severity']
    : local?.severity || 'medium';
  const category = boundedString(raw.category, 160) || local?.category || 'Architecture';
  const issue = boundedString(raw.issue, 1200) || local?.issue || '';
  const recommendation = boundedString(raw.recommendation, 2400)
    || local?.recommendation
    || '';
  if (!issue || !recommendation) return null;

  const resources = boundedStringArray(raw.resources, 50, 200)
    .filter((resource) => serviceNames.has(resource));
  const evidence = boundedStringArray(raw.evidence, 8, 800);
  const remediation = boundedStringArray(raw.remediation, 8, 800);
  const source = raw.source === 'rule-based' || raw.source === 'ai-analysis'
    ? raw.source
    : local
      ? 'rule-based'
      : 'ai-analysis';
  const rawAction = raw.applyAction && typeof raw.applyAction === 'object'
    ? raw.applyAction as Record<string, unknown>
    : null;
  const actionType = rawAction && APPLY_ACTION_TYPES.has(
    rawAction.type as NonNullable<ValidationFinding['applyAction']>['type'],
  )
    ? rawAction.type as NonNullable<ValidationFinding['applyAction']>['type']
    : local?.applyAction?.type || 'regenerate';
  const actionLabel = boundedString(rawAction?.label, 160)
    || local?.applyAction?.label
    || 'Apply through AI refinement';
  const serviceType = boundedString(rawAction?.serviceType, 160)
    || local?.applyAction?.serviceType;
  const rawReference = boundedString(raw.referenceUrl, 500);
  const referenceUrl = rawReference.startsWith('https://learn.microsoft.com/')
    ? rawReference
    : local?.referenceUrl;

  return {
    severity,
    category,
    issue,
    recommendation,
    resources: resources.length > 0 ? resources : local?.resources,
    ruleId: ruleId || local?.ruleId,
    source,
    evidence: evidence.length > 0
      ? evidence
      : local?.evidence || [issue],
    remediation: remediation.length > 0
      ? remediation
      : local?.remediation || recommendationSteps(recommendation),
    referenceUrl,
    applyAction: {
      type: actionType,
      label: actionLabel,
      ...(serviceType ? { serviceType } : {}),
    },
  };
}

/**
 * Validate architecture against Azure Well-Architected Framework.
 * 
 * Hybrid approach:
 *   1. Run instant, deterministic rule-based checks (~1ms)
 *   2. Send the pre-computed findings + architecture to the LLM for
 *      contextual refinement, scoring, and additional insights
 * 
 * This is ~3-5x faster than sending everything to the LLM from scratch
 * because the LLM prompt is smaller and more focused.
 */
export async function validateArchitecture(
  services: Array<{ name: string; type: string; category: string; description?: string }>,
  connections: Array<{ from: string; to: string; label: string }>,
  groups?: Array<{ name: string; services?: string[] }>,
  architectureDescription?: string,
  modelOverride?: ValidationModelOverride,
  language: Language = 'en',
): Promise<ArchitectureValidation> {
  const runtime = resolveAIModelRuntime('validation', modelOverride);

  console.log(`🔍 Starting hybrid WAF validation with ${runtime.displayName}...`);

  // ── Phase 1: Instant local rule-based analysis ──────────────────────
  const localResult = detectWafPatterns(services, connections, groups);
  // Service-specific checks describe configuration that a topology diagram
  // cannot prove. Score only observed topology gaps so adding a service does
  // not incorrectly lower the architecture score for invisible settings.
  const preliminaryScore = calculatePreliminaryScore(localResult.patternFindings);
  const kbStats = getKnowledgeBaseStats();
  
  console.log(`⚡ Phase 1 (local): ${localResult.findings.length} findings, preliminary score ${preliminaryScore}/100 (${localResult.elapsedMs}ms)`);
  console.log(`  📚 Knowledge base: ${kbStats.totalRules} rules across ${kbStats.servicesCovered} services`);

  // ── Phase 2: LLM contextual refinement ──────────────────────────────
  // Build architecture context
  const servicesList = services.map(s => `- ${s.name} (${s.type})`).join('\n');
  const connectionsList = connections.map(c => `- ${c.from} → ${c.to}: ${c.label}`).join('\n');
  const groupsList = groups ? groups.map(g => `- ${g.name}`).join('\n') : 'No groups';
  const serviceNamesList = services.map(s => s.name);

  // Only send architecture-level pattern findings to the LLM (not per-service
  // best-practice rules, which are generic and would overwhelm the prompt)
  const patternFindingsSummary = localResult.patternFindings.length > 0
    ? localResult.patternFindings.map(f =>
        [
          `- Rule ${f.ruleId || 'local'} [${f.severity.toUpperCase()}] ${f.category}: ${f.issue}`,
          f.resources?.length ? `  Affects: ${f.resources.join(', ')}` : '',
          f.evidence?.length ? `  Evidence: ${f.evidence.join(' | ')}` : '',
          f.remediation?.length ? `  Remediation: ${f.remediation.join(' | ')}` : '',
        ].filter(Boolean).join('\n')
      ).join('\n')
    : 'No architecture-level anti-patterns detected.';

  const patternsNote = localResult.patternsDetected.length > 0
    ? `Detected topology patterns: ${localResult.patternsDetected.join(', ')}`
    : 'No common anti-patterns detected.';

  const systemPrompt = `You are an Azure Well-Architected Framework expert. Your role is to review Azure architectures and provide actionable recommendations across the five pillars:

1. **Reliability** - Resiliency, availability, disaster recovery
2. **Security** - Identity, data protection, network security
3. **Cost Optimization** - Right-sizing, reserved instances, consumption patterns
4. **Operational Excellence** - Monitoring, automation, DevOps practices
5. **Performance Efficiency** - Scaling, caching, optimization

A topology pre-scan detected these architecture-level patterns to consider:
${patternsNote}
${patternFindingsSummary}

Use these patterns as hints — validate whether they apply in context, dismiss any that don't, and add your own architecture-specific findings.

${getPromptLanguageInstruction(language)}

SCORING GUIDANCE:
- Score the architecture based on what IS present, not what COULD be added
- A well-connected architecture with appropriate services should score 60-80
- Only score below 50 for architectures with critical gaps (no auth, no monitoring, single points of failure)
- Findings are improvement suggestions, not reasons to penalize the score severely
- Each finding must include concrete "evidence" from the diagram, ordered "remediation" steps, and a "source" field: "rule-based" (from pre-scan) or "ai-analysis" (your addition)
- Preserve the "ruleId" for pre-scan findings. For AI findings, omit ruleId.
- Use "applyAction" to describe how the diagram can be improved: type is "add-service", "regenerate", or "configure"; include a concise label and optional exact Azure serviceType.
- Do not claim a runtime setting is disabled when the diagram cannot show it. State that the setting is unverified and requires deployment review.

Return ONLY valid JSON (no markdown) with this structure:
{
  "overallScore": 75,
  "summary": "Brief 2-3 sentence overall assessment",
  "pillars": [
    {
      "score": 80,
      "pillar": "Reliability",
      "findings": [
        {
          "severity": "high",
          "category": "High Availability",
          "issue": "...",
          "recommendation": "...",
          "resources": ["service-name-1"],
          "ruleId": "arch-no-monitoring",
          "source": "rule-based",
          "evidence": ["No Azure Monitor or Log Analytics service is shown."],
          "remediation": ["Add Azure Monitor.", "Connect the application telemetry.", "Create health alerts."],
          "referenceUrl": "https://learn.microsoft.com/azure/well-architected/",
          "applyAction": {
            "type": "add-service",
            "label": "Add observability services",
            "serviceType": "Azure Monitor"
          }
        }
      ]
    }
  ],
  "quickWins": [
    {
      "severity": "medium",
      "category": "Cost Optimization",
      "issue": "...",
      "recommendation": "...",
      "resources": ["Azure Functions"],
      "source": "ai-analysis",
      "evidence": ["The diagram shows ..."],
      "remediation": ["First action", "Second action"],
      "applyAction": {
        "type": "regenerate",
        "label": "Apply through AI refinement"
      }
    }
  ]
}`;

  const userPrompt = `Review this Azure architecture:

**Architecture Description:**
${architectureDescription || 'Not provided'}

**Services (${services.length}):**
${servicesList}

**Connections (${connections.length}):**
${connectionsList}

**Logical Groups:**
${groupsList}

IMPORTANT: In the "resources" arrays, use EXACTLY the service names as listed above (e.g., "${serviceNamesList.slice(0, 3).join('", "')}"). Do not rename or rephrase them.

Provide a comprehensive Well-Architected Framework assessment with actionable recommendations.`;

  try {
    console.log('📤 Phase 2: Sending focused validation to Azure OpenAI...');
    const { content, metrics } = await callAzureOpenAI([
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ], 8000, modelOverride);

    console.log('✅ Hybrid validation response received:', content.length, 'characters');

    // Parse JSON response. safeParseModelJson tolerates ```json fences / prose
    // and raises a typed, localised error (never the raw parser message).
    const validation: ArchitectureValidation = safeParseModelJson<ArchitectureValidation>(content, {
      context: 'architecture validation',
    });
    
    // Validate response structure. A score of 0 is a legitimate result, so the
    // checks must be type-based rather than truthiness-based.
    if (
      typeof validation.overallScore !== 'number'
      || !Number.isFinite(validation.overallScore)
      || !Array.isArray(validation.pillars)
      || typeof validation.summary !== 'string'
    ) {
      console.error('❌ Invalid response structure:', validation);
      throw new Error('Response missing required fields');
    }

    const localByRuleId = new Map(
      localResult.findings
        .filter((finding) => finding.ruleId)
        .map((finding) => [finding.ruleId as string, finding]),
    );
    const exactServiceNames = new Set(serviceNamesList);
    validation.overallScore = Math.max(0, Math.min(100, Math.round(validation.overallScore)));
    validation.pillars = validation.pillars
      .filter((pillar: unknown) => !!pillar && typeof pillar === 'object')
      .map((pillar: any) => ({
        ...pillar,
        score: Number.isFinite(Number(pillar.score))
          ? Math.max(0, Math.min(100, Math.round(Number(pillar.score))))
          : 0,
        findings: (Array.isArray(pillar.findings) ? pillar.findings : [])
          .map((finding: unknown) => normalizeFinding(
            finding,
            localByRuleId,
            exactServiceNames,
          ))
          .filter((finding: ValidationFinding | null): finding is ValidationFinding => Boolean(finding)),
      }));
    validation.quickWins = (Array.isArray(validation.quickWins) ? validation.quickWins : [])
      .map((finding: unknown) => normalizeFinding(
        finding,
        localByRuleId,
        exactServiceNames,
      ))
      .filter((finding: ValidationFinding | null): finding is ValidationFinding => Boolean(finding));
    
    validation.timestamp = new Date().toISOString();
    validation.metrics = metrics;
    validation.modelUsed = runtime.displayName
      + (runtime.isReasoning ? ` (${runtime.reasoningEffort})` : '');

    // Attach hybrid metadata
    (validation as any).hybridMetadata = {
      localFindings: localResult.findings.length,
      patternsDetected: localResult.patternsDetected,
      localElapsedMs: localResult.elapsedMs,
      preliminaryScore,
      kbRulesUsed: kbStats.totalRules,
    };

    console.log('🎯 Hybrid validation complete. Overall score:', validation.overallScore);
    console.log('📊 Pillars analyzed:', validation.pillars.length);
    console.log('⚡ Quick wins identified:', validation.quickWins.length);
    console.log(`🔬 Hybrid breakdown: ${localResult.findings.length} local + LLM refinement in ${(metrics.elapsedTimeMs / 1000).toFixed(2)}s total`);

    return validation;

  } catch (error) {
    console.error('❌ Architecture validation failed:', error);
    throw error;
  }
}

/**
 * Format validation results for display
 */
export function formatValidationReport(validation: ArchitectureValidation): string {
  const date = new Date(validation.timestamp).toLocaleString();
  
  let report = `# 🔍 Azure Architecture Validation Report\n\n`;
  report += `**Generated:** ${date}\n\n`;
  
  // Add architecture diagram image reference if available
  if (validation.diagramImageDataUrl) {
    const imageFilename = `architecture-validation-${new Date(validation.timestamp).getTime()}-diagram.png`;
    report += `## 🖼️ Architecture Diagram\n\n`;
    report += `![Architecture Diagram](./${imageFilename})\n\n`;
  }
  
  report += `---\n\n`;
  
  // Executive Summary
  report += `## 📊 Executive Summary\n\n`;
  const overallBand = scoreToBand(validation.overallScore);
  report += `### Overall Maturity: ${overallBand.label}\n\n`;
  report += `_Numeric signal: ${validation.overallScore}/100 — a diagram-only, design-time heuristic, not a deployed-environment audit._\n\n`;
  
  const scoreColor = validation.overallScore >= 80 ? '🟢' : validation.overallScore >= 60 ? '🟡' : '🔴';
  report += `${scoreColor} **Assessment:** ${validation.summary}\n\n`;
  
  // Pillar Maturity at a Glance
  report += `### Pillar Maturity at a Glance\n\n`;
  report += `| Pillar | Maturity | Score |\n`;
  report += `|--------|----------|-------|\n`;
  validation.pillars.forEach(pillar => {
    const band = scoreToBand(pillar.score);
    report += `| ${pillar.pillar} | ${band.label} | ${pillar.score}/100 |\n`;
  });
  report += `\n---\n\n`;
  
  // Detailed Findings by Pillar
  report += `## 🏗️ Detailed Assessment by Pillar\n\n`;
  
  validation.pillars.forEach((pillar, index) => {
    const band = scoreToBand(pillar.score);
    report += `### ${index + 1}. ${pillar.pillar} — ${band.label} (${pillar.score}/100)\n\n`;
    
    if (pillar.findings.length === 0) {
      report += `✅ No critical findings for this pillar.\n\n`;
      return;
    }
    
    pillar.findings.forEach((finding) => {
      const emoji = {
        critical: '🔴',
        high: '🟠',
        medium: '🟡',
        low: '🟢'
      }[finding.severity];
      
      report += `${emoji} **${finding.category}** [${finding.severity.toUpperCase()}]\n\n`;
      report += `**Issue:**  \n${finding.issue}\n\n`;
      report += `**Recommendation:**  \n${finding.recommendation}\n\n`;
      if (finding.evidence && finding.evidence.length > 0) {
        report += `**Diagram Evidence:**\n`;
        finding.evidence.forEach(item => {
          report += `- ${item}\n`;
        });
        report += `\n`;
      }
      if (finding.remediation && finding.remediation.length > 0) {
        report += `**Remediation Steps:**\n`;
        finding.remediation.forEach((step, stepIndex) => {
          report += `${stepIndex + 1}. ${step}\n`;
        });
        report += `\n`;
      }
      if (finding.resources && finding.resources.length > 0) {
        report += `**Affected Resources:**\n`;
        finding.resources.forEach(resource => {
          report += `- ${resource}\n`;
        });
        report += `\n`;
      }
      if (finding.referenceUrl) {
        report += `**Reference:** ${finding.referenceUrl}\n\n`;
      }
      report += `**Source:** ${finding.source === 'rule-based' ? `Deterministic rule${finding.ruleId ? ` (${finding.ruleId})` : ''}` : 'AI contextual analysis'}\n\n`;
      report += `---\n\n`;
    });
  });
  
  // Quick Wins Section
  if (validation.quickWins.length > 0) {
    report += `## ⚡ Quick Wins - Immediate Action Items\n\n`;
    report += `These are high-impact, low-effort improvements you can implement right away:\n\n`;
    
    validation.quickWins.forEach((win, index) => {
      report += `### ${index + 1}. ${win.category}\n\n`;
      report += `${win.recommendation}\n\n`;
      if (win.remediation && win.remediation.length > 0) {
        win.remediation.forEach((step, stepIndex) => {
          report += `${stepIndex + 1}. ${step}\n`;
        });
        report += `\n`;
      }
    });
  }
  
  // Footer
  report += `---\n\n`;
  report += `## 📚 Additional Resources\n\n`;
  report += `- [Azure Well-Architected Framework](https://learn.microsoft.com/azure/architecture/framework/)\n`;
  report += `- [Azure Architecture Center](https://learn.microsoft.com/azure/architecture/)\n`;
  report += `- [Azure Security Benchmark](https://learn.microsoft.com/security/benchmark/azure/)\n\n`;
  
  report += `---\n\n`;
  report += `*Report generated by Microsoft Product Architecture Diagram Builder*  \n`;
  report += `*Powered by ${validation.modelUsed || 'Azure OpenAI'} and Azure Well-Architected Framework*  \n`;
  report += `*Generated: ${new Date(validation.timestamp).toLocaleString()}*\n`;
  
  return report;
}
