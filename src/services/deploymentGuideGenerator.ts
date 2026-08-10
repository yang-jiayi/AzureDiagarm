// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Deployment Guide Generator Agent
 * Uses GPT-5-2 to generate comprehensive deployment documentation
 * Includes prerequisites, step-by-step instructions, configuration, and troubleshooting
 */

import JSZip from 'jszip';
import { generateModelFilename } from '../utils/modelNaming';
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
import { searchMicrosoftDocs, renderGroundingBlock, DocSource } from './docsGroundingService';
import { resolveAIModelRuntime } from './aiModelRuntime';
import { safeParseModelJson } from './aiRetry';

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

async function callAzureOpenAI(messages: any[], maxTokens: number = 10000): Promise<CallResult> {
  const runtime = resolveAIModelRuntime('deploymentGuide');
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
  let content = parsed.content;
  const metrics: AIMetrics = {
    promptTokens: parsed.promptTokens,
    completionTokens: parsed.completionTokens,
    totalTokens: parsed.totalTokens,
    elapsedTimeMs,
    model: runtime.displayName,
  };
  
  console.log('📦 API Response:', content.length, 'chars |',
    `Tokens: ${metrics.promptTokens} in → ${metrics.completionTokens} out (${metrics.totalTokens} total) |`,
    `Time: ${(metrics.elapsedTimeMs / 1000).toFixed(2)}s`);
  
  // Track model usage telemetry
  trackAIModelUsage({
    model: runtime.telemetryModel,
    operation: 'deployment_guide',
    reasoningEffort: runtime.isReasoning ? runtime.reasoningEffort : undefined,
    promptTokens: metrics.promptTokens,
    completionTokens: metrics.completionTokens,
    totalTokens: metrics.totalTokens,
    elapsedTimeMs: metrics.elapsedTimeMs,
  });
  
  return { content, metrics };
}

export interface DeploymentStep {
  step: number;
  title: string;
  description: string;
  commands?: string[];
  azurePortalSteps?: string[];
  notes?: string[];
}

export interface BicepModule {
  name: string;
  description: string;
  filename: string;
  content: string;
}

export interface DeploymentGuide {
  title: string;
  overview: string;
  prerequisites: string[];
  estimatedTime: string;
  deploymentSteps: DeploymentStep[];
  configuration: {
    section: string;
    settings: Array<{ name: string; value: string; description: string }>;
  }[];
  postDeployment: string[];
  troubleshooting: Array<{ issue: string; solution: string }>;
  estimatedCost: string;
  timestamp: string;
  bicepTemplates?: BicepModule[];
  metrics?: AIMetrics;
  /** Official Microsoft Learn pages used to ground this guide (Phase 1). */
  groundingSources?: DocSource[];
}

/**
 * Generate comprehensive deployment guide for the architecture
 */
export async function generateDeploymentGuide(
  services: Array<{ name: string; type: string; category: string; description?: string }>,
  connections: Array<{ from: string; to: string; label: string; type?: string }>,
  _groups?: Array<{ name: string; services?: string[] }>,
  architectureDescription?: string,
  estimatedCost?: number,
  language: Language = 'en',
): Promise<DeploymentGuide> {
  const runtime = resolveAIModelRuntime('deploymentGuide');

  console.log(`📋 Generating deployment guide with ${runtime.displayName}...`);

  // Phase 1 grounding: pull current, citable Microsoft Learn docs for the top
  // services so the guide reflects up-to-date API versions, CLI flags, and
  // Bicep schemas. Best-effort — if it fails we generate ungrounded.
  const topServiceNames = services.slice(0, 6).map((s) => s.name).join(', ');
  const groundingQuery = `Deploy ${topServiceNames} to Azure using Bicep and Azure CLI`;
  let groundingSources: DocSource[] = [];
  try {
    groundingSources = await searchMicrosoftDocs(groundingQuery, 6);
    console.log(`📚 Grounding: ${groundingSources.length} Microsoft Learn source(s)`);
  } catch (e) {
    console.warn('⚠️ Docs grounding unavailable, proceeding ungrounded:', e);
  }
  const groundingBlock = renderGroundingBlock(groundingSources);

  // Build architecture context (limit to prevent token overflow)
  const servicesList = services.slice(0, 25).map(s => 
    `- ${s.name} (${s.type})`
  ).join('\n');
  
  const connectionsList = connections.slice(0, 20).map(c => 
    `- ${c.from} → ${c.to}: ${c.label}`
  ).join('\n');

  const systemPrompt = `You are an Azure DevOps expert. Create comprehensive deployment guides with Bicep templates.

Generate deployment documentation with:
1. Prerequisites - Azure CLI, permissions, tools
2. Deployment Steps - Azure CLI commands with notes
3. Configuration - Environment variables, settings
4. Post-Deployment - Validation, monitoring
5. Troubleshooting - Common issues and solutions
6. Bicep Templates - Infrastructure as Code files

${getPromptLanguageInstruction(language)}

**Bicep Guidelines:**
- Create main.bicep + module files for each service
- Use parameters for environment, location, naming
- Include outputs for endpoints and connection strings
- Add resource tags and comments

**Grounding:**
- If GROUNDING SOURCES (Microsoft Learn) are provided in the request, prefer their guidance for API versions, CLI flags, and Bicep resource schemas.
- Reference the relevant source URLs inline in step notes (e.g. "See: <url>") where they informed a command or template.
- Never invent URLs; only cite URLs that appear in the provided sources.

Return ONLY valid JSON:
{
  "title": "Deploy [Name] to Azure",
  "overview": "Brief description",
  "prerequisites": ["Azure subscription", "Azure CLI 2.50+"],
  "estimatedTime": "45-60 minutes",
  "deploymentSteps": [{"step": 1, "title": "Setup", "description": "...", "commands": ["az login"], "notes": ["tip"]}],
  "configuration": [{"section": "Settings", "settings": [{"name": "KEY", "value": "val", "description": "desc"}]}],
  "postDeployment": ["Verify resources"],
  "troubleshooting": [{"issue": "Error X", "solution": "Do Y"}],
  "estimatedCost": "$100-200/month",
  "bicepTemplates": [{"name": "Main", "description": "Orchestration", "filename": "main.bicep", "content": "// Bicep code here"}]
}`;

  const userPrompt = `Create a deployment guide for this Azure architecture:

**Description:** ${architectureDescription || 'Azure architecture'}

**Services (${services.length}):**
${servicesList}

**Connections (${connections.length}):**
${connectionsList}

${estimatedCost ? `**Est. Monthly Cost:** $${estimatedCost.toFixed(2)}` : ''}

${groundingBlock ? `${groundingBlock}

` : ''}Generate a deployment guide with Azure CLI commands and Bicep templates for this architecture.`;

  try {
    const { content, metrics } = await callAzureOpenAI([
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ], 32000);

    // Handle empty response
    if (!content || content.length === 0) {
      throw new Error('Empty response from API. The architecture may be too complex. Try with fewer services.');
    }

    // Parse JSON response. safeParseModelJson tolerates ```json fences / prose
    // and raises a typed, localised error (never the raw parser message).
    const guide: DeploymentGuide = safeParseModelJson<DeploymentGuide>(content, {
      context: 'deployment guide',
    });
    guide.timestamp = new Date().toISOString();
    guide.metrics = metrics;

    // Models occasionally omit optional sections. Normalize the collections the
    // guide viewer and exporters iterate over so a partial response degrades
    // into a shorter guide instead of throwing on `.length`/`.map`.
    if (!Array.isArray(guide.deploymentSteps)) guide.deploymentSteps = [];
    if (!Array.isArray(guide.configuration)) guide.configuration = [];
    if (!Array.isArray(guide.troubleshooting)) guide.troubleshooting = [];
    if (!Array.isArray(guide.prerequisites)) guide.prerequisites = [];
    if (!Array.isArray(guide.postDeployment)) guide.postDeployment = [];
    if (guide.deploymentSteps.length === 0) {
      throw new Error('The deployment guide response contained no deployment steps. Please try again.');
    }

    if (groundingSources.length > 0) {
      guide.groundingSources = groundingSources.map((s) => ({ title: s.title, url: s.url }));
    }

    console.log('📋 Deployment guide generated');
    console.log('📝 Steps:', guide.deploymentSteps.length);
    console.log('⚙️ Configuration sections:', guide.configuration.length);
    console.log('🔧 Troubleshooting tips:', guide.troubleshooting.length);
    console.log('📦 Bicep templates:', guide.bicepTemplates?.length || 0);

    return guide;

  } catch (error) {
    console.error('❌ Deployment guide generation failed:', error);
    throw error;
  }
}

/**
 * Format deployment guide as markdown
 */
export function formatDeploymentGuide(guide: DeploymentGuide): string {
  let md = `# ${guide.title}\n\n`;
  
  md += `## Overview\n\n${guide.overview}\n\n`;
  
  md += `**Estimated Time:** ${guide.estimatedTime}\n\n`;
  md += `**Estimated Cost:** ${guide.estimatedCost}\n\n`;
  
  md += `## Prerequisites\n\n`;
  guide.prerequisites.forEach(prereq => {
    md += `- ${prereq}\n`;
  });
  md += `\n`;
  
  md += `## Deployment Steps\n\n`;
  guide.deploymentSteps.forEach(step => {
    md += `### Step ${step.step}: ${step.title}\n\n`;
    md += `${step.description}\n\n`;
    
    if (step.commands && step.commands.length > 0) {
      md += `**Commands:**\n\`\`\`bash\n${step.commands.join('\n')}\n\`\`\`\n\n`;
    }
    
    if (step.azurePortalSteps && step.azurePortalSteps.length > 0) {
      md += `**Azure Portal:**\n`;
      step.azurePortalSteps.forEach(portalStep => {
        md += `- ${portalStep}\n`;
      });
      md += `\n`;
    }
    
    if (step.notes && step.notes.length > 0) {
      md += `**Notes:**\n`;
      step.notes.forEach(note => {
        md += `- 💡 ${note}\n`;
      });
      md += `\n`;
    }
  });
  
  md += `## Configuration\n\n`;
  guide.configuration.forEach(section => {
    md += `### ${section.section}\n\n`;
    md += `| Setting | Value | Description |\n`;
    md += `|---------|-------|-------------|\n`;
    section.settings.forEach(setting => {
      md += `| \`${setting.name}\` | ${setting.value} | ${setting.description} |\n`;
    });
    md += `\n`;
  });
  
  md += `## Post-Deployment Validation\n\n`;
  guide.postDeployment.forEach(task => {
    md += `- [ ] ${task}\n`;
  });
  md += `\n`;
  
  md += `## Troubleshooting\n\n`;
  guide.troubleshooting.forEach(item => {
    md += `**Issue:** ${item.issue}\n\n`;
    md += `**Solution:** ${item.solution}\n\n`;
  });
  
  if (guide.groundingSources && guide.groundingSources.length > 0) {
    md += `## References (Microsoft Learn)\n\n`;
    guide.groundingSources.forEach(src => {
      md += `- [${src.title}](${src.url})\n`;
    });
    md += `\n`;
  }
  
  md += `---\n\n`;
  md += `*Generated: ${new Date(guide.timestamp).toLocaleString()}*\n`;
  
  return md;
}

/**
 * Export deployment guide as downloadable file
 */
export function downloadDeploymentGuide(guide: DeploymentGuide) {
  const markdown = formatDeploymentGuide(guide);
  const blob = new Blob([markdown], { type: 'text/markdown' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = generateModelFilename('deployment-guide', 'md');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Download a single Bicep template
 */
export function downloadBicepTemplate(template: BicepModule) {
  const blob = new Blob([template.content], { type: 'text/plain' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = template.filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Download all Bicep templates as a ZIP file
 */
export async function downloadAllBicepTemplates(guide: DeploymentGuide) {
  if (!guide.bicepTemplates || guide.bicepTemplates.length === 0) {
    console.warn('No Bicep templates available to download');
    return;
  }

  const zip = new JSZip();
  
  // Add README with instructions
  const readme = `# Bicep Templates for ${guide.title}
Generated: ${new Date(guide.timestamp).toLocaleString()}

## Deployment Instructions

1. Review and customize parameters in main.bicep
2. Deploy with Azure CLI:
   \`\`\`bash
   az login
   az group create --name <rg-name> --location <location>
   az deployment group create --resource-group <rg-name> --template-file main.bicep
   \`\`\`

## Files Included
${guide.bicepTemplates.map(t => `- ${t.filename}: ${t.description}`).join('\n')}
`;
  zip.file('README.md', readme);

  // Add each Bicep template
  guide.bicepTemplates.forEach((template) => {
    zip.file(template.filename, template.content);
  });

  // Generate and download ZIP
  const blob = await zip.generateAsync({ type: 'blob' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = generateModelFilename('bicep-templates', 'zip');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
