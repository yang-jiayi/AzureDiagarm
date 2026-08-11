// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import type { Edge, Node } from 'reactflow';
import { resolveServiceIconMapping } from '../data/serviceIconMapping';
import { lookupServiceMeta } from './armExtractor';
import type { IaCFormat } from './azureOpenAI';

export type StarterTemplateFormat = 'bicep' | 'terraform';
export type DriftAction = 'create' | 'update' | 'delete' | 'replace' | 'no-op' | 'other';
export type MatchConfidence = 'exact' | 'normalized' | 'approximate' | 'unmapped';
export type ResourceApproximation = 'exact' | 'type-only' | 'expression' | 'unmapped';

export interface IaCBaselineSourceFile {
  filename: string;
  size: number;
}

export interface IaCBaselineResource {
  id: string;
  logicalName: string;
  resourceName: string | null;
  providerType: string;
  mappedService: string | null;
  category: string | null;
  sourceFile: string;
  origin: IaCFormat;
  approximation: ResourceApproximation;
  notes?: string;
}

export interface IaCBaseline {
  version: 1;
  capturedAt: string;
  format: IaCFormat;
  formatLabel: string;
  sourceFiles: IaCBaselineSourceFile[];
  resources: IaCBaselineResource[];
  resourceCount: number;
  unmappedCount: number;
  warnings: string[];
}

export interface DiagramResourceDescriptor {
  id: string;
  label: string;
  serviceName: string;
  providerType: string | null;
  mappedService: string | null;
  category: string | null;
  approximation: ResourceApproximation;
}

export interface MatchedIaCResource {
  baseline: IaCBaselineResource;
  diagram: DiagramResourceDescriptor;
  confidence: MatchConfidence;
  reason: string;
}

export interface IaCComparisonReport {
  matched: MatchedIaCResource[];
  sourceOnly: IaCBaselineResource[];
  diagramOnly: DiagramResourceDescriptor[];
  approximateMatches: number;
}

export interface DriftChangeSummary {
  id: string;
  address: string;
  action: DriftAction;
  actionText: string;
  resourceType: string | null;
  resourceName: string | null;
  provider: 'azure' | 'terraform';
  rawActions: string[];
}

export interface DriftPlanSummary {
  kind: 'azure-what-if' | 'terraform-plan';
  importedAt: string;
  sourceFile: string;
  changeCounts: Record<DriftAction, number>;
  changes: DriftChangeSummary[];
}

export interface StarterTemplate {
  format: StarterTemplateFormat;
  fileName: string;
  content: string;
  supportedResourceCount: number;
  todoCount: number;
  /** Deployment-ordering clauses derived from the diagram's connections. */
  dependencyCount: number;
}

interface BaselineBuildInput {
  format: IaCFormat;
  files: Array<{ name: string; text: string; size?: number }>;
  importedAt?: string;
}

interface ServiceResolution {
  providerType: string | null;
  mappedService: string | null;
  category: string | null;
}

interface ComparableResource {
  id: string;
  kind: 'baseline' | 'diagram';
  serviceKey: string | null;
  providerKey: string | null;
  nameKey: string | null;
  mappedService: string | null;
  providerType: string | null;
}

interface CandidateMatch {
  baseline: IaCBaselineResource;
  diagram: DiagramResourceDescriptor;
  confidence: MatchConfidence;
  reason: string;
  score: number;
}

interface TerraformTypeMapping {
  armType: string;
}

interface ExportResourceDescriptor {
  nodeId: string;
  label: string;
  canonicalService: string | null;
  providerType: string | null;
  mappedService: string | null;
}

const ARM_CHILD_TYPES_TO_FOLD = new Set([
  'microsoft.storage/storageaccounts/blobservices',
  'microsoft.storage/storageaccounts/fileservices',
  'microsoft.storage/storageaccounts/queueservices',
  'microsoft.storage/storageaccounts/tableservices',
  'microsoft.network/virtualnetworks/subnets',
  'microsoft.web/sites/config',
  'microsoft.sql/servers/firewallrules',
  'microsoft.sql/servers/databases/transparentdataencryption',
]);

const TERRAFORM_TYPE_MAP: Record<string, TerraformTypeMapping> = {
  azurerm_api_management: { armType: 'microsoft.apimanagement/service' },
  azurerm_app_configuration: { armType: 'microsoft.appconfiguration/configurationstores' },
  azurerm_app_service: { armType: 'microsoft.web/sites' },
  azurerm_app_service_plan: { armType: 'microsoft.web/serverfarms' },
  azurerm_application_gateway: { armType: 'microsoft.network/applicationgateways' },
  azurerm_application_insights: { armType: 'microsoft.insights/components' },
  azurerm_bastion_host: { armType: 'microsoft.network/bastionhosts' },
  azurerm_cognitive_account: { armType: 'microsoft.cognitiveservices/accounts' },
  azurerm_container_app: { armType: 'microsoft.app/containerapps' },
  azurerm_container_app_environment: { armType: 'microsoft.app/managedenvironments' },
  azurerm_container_registry: { armType: 'microsoft.containerregistry/registries' },
  azurerm_cosmosdb_account: { armType: 'microsoft.documentdb/databaseaccounts' },
  azurerm_eventgrid_topic: { armType: 'microsoft.eventgrid/topics' },
  azurerm_eventhub_namespace: { armType: 'microsoft.eventhub/namespaces' },
  azurerm_linux_function_app: { armType: 'microsoft.web/sites' },
  azurerm_linux_web_app: { armType: 'microsoft.web/sites' },
  azurerm_key_vault: { armType: 'microsoft.keyvault/vaults' },
  azurerm_kubernetes_cluster: { armType: 'microsoft.containerservice/managedclusters' },
  azurerm_lb: { armType: 'microsoft.network/loadbalancers' },
  azurerm_log_analytics_workspace: { armType: 'microsoft.operationalinsights/workspaces' },
  azurerm_managed_disk: { armType: 'microsoft.compute/disks' },
  azurerm_mysql_flexible_server: { armType: 'microsoft.dbformysql/flexibleservers' },
  azurerm_mssql_database: { armType: 'microsoft.sql/servers/databases' },
  azurerm_mssql_server: { armType: 'microsoft.sql/servers' },
  azurerm_postgresql_flexible_server: { armType: 'microsoft.dbforpostgresql/flexibleservers' },
  azurerm_private_dns_zone: { armType: 'microsoft.network/privatednszones' },
  azurerm_private_endpoint: { armType: 'microsoft.network/privateendpoints' },
  azurerm_public_ip: { armType: 'microsoft.network/publicipaddresses' },
  azurerm_redis_cache: { armType: 'microsoft.cache/redis' },
  azurerm_search_service: { armType: 'microsoft.search/searchservices' },
  azurerm_service_plan: { armType: 'microsoft.web/serverfarms' },
  azurerm_servicebus_namespace: { armType: 'microsoft.servicebus/namespaces' },
  azurerm_sql_database: { armType: 'microsoft.sql/servers/databases' },
  azurerm_sql_server: { armType: 'microsoft.sql/servers' },
  azurerm_static_site: { armType: 'microsoft.web/staticsites' },
  azurerm_storage_account: { armType: 'microsoft.storage/storageaccounts' },
  azurerm_synapse_workspace: { armType: 'microsoft.synapse/workspaces' },
  azurerm_user_assigned_identity: { armType: 'microsoft.managedidentity/userassignedidentities' },
  azurerm_virtual_machine: { armType: 'microsoft.compute/virtualmachines' },
  azurerm_virtual_network: { armType: 'microsoft.network/virtualnetworks' },
  azurerm_windows_function_app: { armType: 'microsoft.web/sites' },
  azurerm_windows_web_app: { armType: 'microsoft.web/sites' },
};

const BICEP_SUPPORTED_TYPES = new Set([
  'microsoft.app/containerapps',
  'microsoft.app/managedenvironments',
  'microsoft.apimanagement/service',
  'microsoft.cognitiveservices/accounts',
  'microsoft.containerregistry/registries',
  'microsoft.documentdb/databaseaccounts',
  'microsoft.insights/components',
  'microsoft.keyvault/vaults',
  'microsoft.operationalinsights/workspaces',
  'microsoft.search/searchservices',
  'microsoft.servicebus/namespaces',
  'microsoft.sql/servers',
  'microsoft.sql/servers/databases',
  'microsoft.storage/storageaccounts',
  'microsoft.web/serverfarms',
  'microsoft.web/sites',
  'microsoft.network/virtualnetworks',
  'microsoft.managedidentity/userassignedidentities',
  'microsoft.web/staticsites',
  'microsoft.dbforpostgresql/flexibleservers',
  'microsoft.dbformysql/flexibleservers',
  'microsoft.eventhub/namespaces',
]);

/**
 * The table above has to stay in step with the switch in bicepStubForResource:
 * a type listed there but missing a stub silently degrades to a TODO comment.
 * Exported so the unit test can assert the two never drift apart.
 */
export function listBicepSupportedTypes(): string[] {
  return [...BICEP_SUPPORTED_TYPES].sort();
}

function normalizeToken(value: string | null | undefined): string | null {
  if (!value) return null;
  const normalized = value
    .toLowerCase()
    .replace(/\bazure\b/g, '')
    .replace(/\bmicrosoft\b/g, '')
    .replace(/\bservice\b/g, '')
    .replace(/[^a-z0-9]+/g, '');
  return normalized || null;
}

function normalizeProviderType(value: string | null | undefined): string | null {
  if (!value) return null;
  const normalized = value.toLowerCase().trim();
  return normalized || null;
}

function looksLikeProviderType(value: string): boolean {
  return /^microsoft\./i.test(value) || /^azurerm_/i.test(value);
}

function formatIaCLabel(format: IaCFormat): string {
  switch (format) {
    case 'arm':
      return 'ARM JSON';
    case 'bicep':
      return 'Bicep';
    case 'terraform-hcl':
      return 'Terraform HCL';
    case 'terraform-state':
      return 'Terraform State';
  }
}

function pad2(value: number): string {
  return value.toString().padStart(2, '0');
}

function makeStableId(prefix: string, logicalName: string, sourceFile: string, index: number): string {
  const stem = `${prefix}-${sourceFile}-${logicalName}-${index}`
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
  return stem || `${prefix}-${index + 1}`;
}

function mapTerraformType(terraformType: string): ServiceResolution {
  const key = terraformType.trim().toLowerCase();
  const mapping = TERRAFORM_TYPE_MAP[key];
  if (!mapping) {
    return { providerType: key, mappedService: null, category: null };
  }
  return resolveFromProviderOrService(mapping.armType);
}

function resolveFromProviderOrService(value: string, kind?: string): ServiceResolution {
  const trimmed = value.trim();
  if (!trimmed) {
    return { providerType: null, mappedService: null, category: null };
  }
  if (/^azurerm_/i.test(trimmed)) {
    return mapTerraformType(trimmed);
  }
  if (/^microsoft\./i.test(trimmed)) {
    const meta = lookupServiceMeta(trimmed, kind || '');
    return {
      providerType: trimmed.toLowerCase(),
      mappedService: meta?.name || null,
      category: meta?.category || null,
    };
  }
  const resolved = resolveServiceIconMapping(trimmed);
  if (resolved) {
    return {
      providerType: null,
      mappedService: resolved.serviceName,
      category: resolved.mapping.category,
    };
  }
  return { providerType: null, mappedService: trimmed, category: null };
}

function resourceNameKey(
  logicalName: string,
  resourceName: string | null,
  serviceName: string | null,
): string | null {
  const candidate = normalizeToken(resourceName || logicalName);
  const genericService = normalizeToken(serviceName);
  if (!candidate) return null;
  if (genericService && candidate === genericService) return null;
  return candidate;
}

function toComparableBaseline(resource: IaCBaselineResource): ComparableResource {
  return {
    id: resource.id,
    kind: 'baseline',
    serviceKey: normalizeToken(resource.mappedService),
    providerKey: normalizeProviderType(resource.providerType),
    nameKey: resourceNameKey(resource.logicalName, resource.resourceName, resource.mappedService),
    mappedService: resource.mappedService,
    providerType: resource.providerType,
  };
}

function toComparableDiagram(resource: DiagramResourceDescriptor): ComparableResource {
  return {
    id: resource.id,
    kind: 'diagram',
    serviceKey: normalizeToken(resource.mappedService),
    providerKey: normalizeProviderType(resource.providerType),
    nameKey: resourceNameKey(resource.label, null, resource.mappedService),
    mappedService: resource.mappedService,
    providerType: resource.providerType,
  };
}

function classifyMatch(
  baseline: ComparableResource,
  diagram: ComparableResource,
): { confidence: MatchConfidence; score: number; reason: string } | null {
  const sameService = Boolean(
    baseline.serviceKey && diagram.serviceKey && baseline.serviceKey === diagram.serviceKey,
  );
  const sameProvider = Boolean(
    baseline.providerKey && diagram.providerKey && baseline.providerKey === diagram.providerKey,
  );
  const sameName = Boolean(
    baseline.nameKey && diagram.nameKey && baseline.nameKey === diagram.nameKey,
  );
  const partialName = Boolean(
    baseline.nameKey
      && diagram.nameKey
      && (
        baseline.nameKey.includes(diagram.nameKey)
        || diagram.nameKey.includes(baseline.nameKey)
      ),
  );

  if (!sameService && !sameProvider && !sameName) return null;

  if ((sameService || sameProvider) && sameName) {
    return {
      confidence: sameProvider ? 'exact' : 'normalized',
      score: sameProvider ? 140 : 130,
      reason: sameProvider
        ? 'Matched by provider type and normalized resource name.'
        : 'Matched by normalized Azure service alias and resource name.',
    };
  }
  if (sameService || sameProvider) {
    return {
      confidence: 'approximate',
      score: sameProvider ? 100 : 90,
      reason: sameProvider
        ? 'Matched by provider type; the diagram node does not expose a distinct resource name.'
        : 'Matched by Azure service family only; review the resource name manually.',
    };
  }
  if (sameName || partialName) {
    return {
      confidence: 'unmapped',
      score: partialName ? 40 : 50,
      reason: 'Matched by a normalized name hint only because the resource type could not be mapped.',
    };
  }
  return null;
}

function compareText(left: string, right: string): number {
  return left.localeCompare(right, 'en', { sensitivity: 'base' });
}

function topLevelPropertyValue(block: string, propertyNames: string[], separator: ':' | '='): string | null {
  const lines = block.split(/\r?\n/);
  let depth = 0;
  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) {
      depth += braceDelta(rawLine);
      continue;
    }
    if (depth === 0) {
      for (const propertyName of propertyNames) {
        const matcher = separator === ':'
          ? new RegExp(`^${propertyName}\\s*:\\s*(.+)$`)
          : new RegExp(`^${propertyName}\\s*=\\s*(.+)$`);
        const match = line.match(matcher);
        if (match) {
          return match[1].trim().replace(/,$/, '').trim();
        }
      }
    }
    depth += braceDelta(rawLine);
  }
  return null;
}

function braceDelta(value: string): number {
  let delta = 0;
  let inSingle = false;
  let inDouble = false;
  for (let index = 0; index < value.length; index += 1) {
    const ch = value[index];
    const next = value[index + 1];
    if (!inDouble && ch === '\'' && value[index - 1] !== '\\') {
      inSingle = !inSingle;
      continue;
    }
    if (!inSingle && ch === '"' && value[index - 1] !== '\\') {
      inDouble = !inDouble;
      continue;
    }
    if (inSingle || inDouble) continue;
    if (ch === '/' && next === '/') break;
    if (ch === '#') break;
    if (ch === '{') delta += 1;
    if (ch === '}') delta -= 1;
  }
  return delta;
}

function findMatchingBrace(text: string, openingBraceIndex: number): number {
  let depth = 0;
  let inSingle = false;
  let inDouble = false;
  let inLineComment = false;
  let inBlockComment = false;

  for (let index = openingBraceIndex; index < text.length; index += 1) {
    const ch = text[index];
    const next = text[index + 1];

    if (inLineComment) {
      if (ch === '\n') inLineComment = false;
      continue;
    }
    if (inBlockComment) {
      if (ch === '*' && next === '/') {
        inBlockComment = false;
        index += 1;
      }
      continue;
    }
    if (!inDouble && ch === '\'' && text[index - 1] !== '\\') {
      inSingle = !inSingle;
      continue;
    }
    if (!inSingle && ch === '"' && text[index - 1] !== '\\') {
      inDouble = !inDouble;
      continue;
    }
    if (inSingle || inDouble) continue;
    if (ch === '/' && next === '/') {
      inLineComment = true;
      index += 1;
      continue;
    }
    if (ch === '/' && next === '*') {
      inBlockComment = true;
      index += 1;
      continue;
    }
    if (ch === '#') {
      inLineComment = true;
      continue;
    }
    if (ch === '{') {
      depth += 1;
      continue;
    }
    if (ch === '}') {
      depth -= 1;
      if (depth === 0) return index;
    }
  }
  return -1;
}

function literalStringValue(rawValue: string | null): string | null {
  if (!rawValue) return null;
  const value = rawValue.trim();
  if (value.startsWith("'")) {
    const match = value.match(/^'([^']*)'$/);
    if (!match || match[1].includes('${')) return null;
    return match[1].trim() || null;
  }
  if (value.startsWith('"')) {
    const match = value.match(/^"([^"]*)"$/);
    if (!match || match[1].includes('${')) return null;
    return match[1].trim() || null;
  }
  return null;
}

function buildArmParameterMap(template: unknown): Record<string, string> {
  const output: Record<string, string> = {};
  if (!isRecord(template) || !isRecord(template.parameters)) return output;
  for (const [key, value] of Object.entries(template.parameters)) {
    if (isRecord(value) && typeof value.defaultValue === 'string' && value.defaultValue.trim()) {
      output[key] = value.defaultValue.trim();
    }
  }
  return output;
}

function cleanArmName(rawName: string, paramMap: Record<string, string>): { logicalName: string; resourceName: string | null; approximation: ResourceApproximation } {
  const trimmed = rawName.trim();
  if (!trimmed) {
    return { logicalName: 'resource', resourceName: null, approximation: 'expression' };
  }
  if (!trimmed.startsWith('[')) {
    return { logicalName: trimmed, resourceName: trimmed.split('/')[0] || trimmed, approximation: 'exact' };
  }
  const paramRefs = [...trimmed.matchAll(/parameters\(\s*'([^']+)'\s*\)/gi)].map((match) => match[1]);
  const preferred = paramRefs.find((token) => /name$/i.test(token)) || paramRefs[0];
  if (preferred && paramMap[preferred]) {
    const resolved = paramMap[preferred].split('/')[0].trim();
    return { logicalName: preferred, resourceName: resolved || null, approximation: 'expression' };
  }
  const quoted = [...trimmed.matchAll(/'([^']+)'/g)]
    .map((match) => match[1])
    .find((token) => !/^microsoft\./i.test(token));
  if (quoted) {
    return { logicalName: quoted.trim(), resourceName: quoted.split('/')[0].trim(), approximation: 'expression' };
  }
  return { logicalName: trimmed, resourceName: null, approximation: 'expression' };
}

function flattenArmResources(
  resources: unknown[],
  sourceFile: string,
  paramMap: Record<string, string>,
  output: IaCBaselineResource[],
  seen: { value: number },
): void {
  for (const resource of resources) {
    if (!isRecord(resource) || typeof resource.type !== 'string') continue;
    const providerType = resource.type.toLowerCase();
    if (ARM_CHILD_TYPES_TO_FOLD.has(providerType)) continue;
    const kind = typeof resource.kind === 'string' ? resource.kind : '';
    const service = resolveFromProviderOrService(resource.type, kind);
    const nameInfo = cleanArmName(typeof resource.name === 'string' ? resource.name : 'resource', paramMap);
    output.push({
      id: makeStableId('arm', nameInfo.logicalName, sourceFile, seen.value),
      logicalName: nameInfo.logicalName,
      resourceName: nameInfo.resourceName,
      providerType,
      mappedService: service.mappedService,
      category: service.category,
      sourceFile,
      origin: 'arm',
      approximation: service.mappedService ? nameInfo.approximation : 'unmapped',
      notes: service.mappedService ? undefined : 'Unmapped ARM resource type.',
    });
    seen.value += 1;
    if (Array.isArray(resource.resources)) {
      flattenArmResources(resource.resources, sourceFile, paramMap, output, seen);
    }
  }
}

function parseArmResources(files: BaselineBuildInput['files'], warnings: string[]): IaCBaselineResource[] {
  const output: IaCBaselineResource[] = [];
  const seen = { value: 0 };
  for (const file of files) {
    try {
      const parsed = JSON.parse(file.text);
      if (!Array.isArray(parsed.resources)) continue;
      flattenArmResources(parsed.resources, file.name, buildArmParameterMap(parsed), output, seen);
    } catch (error) {
      warnings.push(`Could not parse ARM JSON from ${file.name}: ${error instanceof Error ? error.message : 'unknown error'}`);
    }
  }
  return output;
}

function parseBicepResources(files: BaselineBuildInput['files'], warnings: string[]): IaCBaselineResource[] {
  const output: IaCBaselineResource[] = [];
  const seen = { value: 0 };
  const pattern = /^\s*resource\s+([A-Za-z_][\w]*)\s+'([^']+)'(?:\s+existing)?\s*=\s*\{/gm;
  for (const file of files) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(file.text)) !== null) {
      const openingBrace = file.text.indexOf('{', match.index);
      if (openingBrace < 0) continue;
      const closingBrace = findMatchingBrace(file.text, openingBrace);
      const body = closingBrace > openingBrace
        ? file.text.slice(openingBrace + 1, closingBrace)
        : '';
      if (closingBrace < 0) {
        warnings.push(`Skipped an unterminated Bicep resource block in ${file.name}.`);
      }
      const declaredType = match[2].split('@')[0].trim();
      const service = resolveFromProviderOrService(declaredType);
      const nameValue = literalStringValue(topLevelPropertyValue(body, ['name'], ':'));
      output.push({
        id: makeStableId('bicep', match[1], file.name, seen.value),
        logicalName: match[1],
        resourceName: nameValue,
        providerType: declaredType.toLowerCase(),
        mappedService: service.mappedService,
        category: service.category,
        sourceFile: file.name,
        origin: 'bicep',
        approximation: nameValue
          ? (service.mappedService ? 'exact' : 'unmapped')
          : (service.mappedService ? 'type-only' : 'unmapped'),
        notes: nameValue
          ? undefined
          : 'The Bicep resource name uses an expression or could not be inferred deterministically.',
      });
      seen.value += 1;
      if (closingBrace >= 0) {
        pattern.lastIndex = closingBrace + 1;
      }
    }
  }
  return output;
}

function parseTerraformHclResources(files: BaselineBuildInput['files'], warnings: string[]): IaCBaselineResource[] {
  const output: IaCBaselineResource[] = [];
  const seen = { value: 0 };
  const pattern = /^\s*resource\s+"([^"]+)"\s+"([^"]+)"\s*\{/gm;
  for (const file of files) {
    let match: RegExpExecArray | null;
    while ((match = pattern.exec(file.text)) !== null) {
      const openingBrace = file.text.indexOf('{', match.index);
      if (openingBrace < 0) continue;
      const closingBrace = findMatchingBrace(file.text, openingBrace);
      const body = closingBrace > openingBrace
        ? file.text.slice(openingBrace + 1, closingBrace)
        : '';
      if (closingBrace < 0) {
        warnings.push(`Skipped an unterminated Terraform block in ${file.name}.`);
      }
      const service = mapTerraformType(match[1]);
      const nameValue = literalStringValue(topLevelPropertyValue(body, ['name'], '='));
      output.push({
        id: makeStableId('tf', `${match[1]}.${match[2]}`, file.name, seen.value),
        logicalName: `${match[1]}.${match[2]}`,
        resourceName: nameValue,
        providerType: match[1].toLowerCase(),
        mappedService: service.mappedService,
        category: service.category,
        sourceFile: file.name,
        origin: 'terraform-hcl',
        approximation: nameValue
          ? (service.mappedService ? 'exact' : 'unmapped')
          : (service.mappedService ? 'type-only' : 'unmapped'),
        notes: nameValue
          ? undefined
          : 'The Terraform resource name uses an expression or could not be inferred deterministically.',
      });
      seen.value += 1;
      if (closingBrace >= 0) {
        pattern.lastIndex = closingBrace + 1;
      }
    }
  }
  return output;
}

function parseTerraformStateResources(files: BaselineBuildInput['files'], warnings: string[]): IaCBaselineResource[] {
  const output: IaCBaselineResource[] = [];
  const seen = { value: 0 };
  for (const file of files) {
    try {
      const parsed = JSON.parse(file.text);
      const resources = Array.isArray(parsed.resources) ? parsed.resources : [];
      for (const resource of resources) {
        if (!isRecord(resource) || typeof resource.type !== 'string' || typeof resource.name !== 'string') continue;
        const resourceType = resource.type;
        const resourceName = resource.name;
        const service = mapTerraformType(resourceType);
        const instances = Array.isArray(resource.instances) ? resource.instances : [];
        if (instances.length === 0) {
          output.push({
            id: makeStableId('tfstate', `${resourceType}.${resourceName}`, file.name, seen.value),
            logicalName: `${resourceType}.${resourceName}`,
            resourceName: null,
            providerType: resourceType.toLowerCase(),
            mappedService: service.mappedService,
            category: service.category,
            sourceFile: file.name,
            origin: 'terraform-state',
            approximation: service.mappedService ? 'type-only' : 'unmapped',
          });
          seen.value += 1;
          continue;
        }

        instances.forEach((instance, instanceIndex) => {
          const attributes = isRecord(instance)
            && isRecord(instance.attributes)
            ? instance.attributes
            : {};
          const indexKey = isRecord(instance) && instance.index_key !== undefined
            ? `[${String(instance.index_key)}]`
            : instances.length > 1
              ? `[${instanceIndex}]`
              : '';
          output.push({
            id: makeStableId('tfstate', `${resourceType}.${resourceName}${indexKey}`, file.name, seen.value),
            logicalName: `${resourceType}.${resourceName}${indexKey}`,
            resourceName: typeof attributes.name === 'string' ? attributes.name : null,
            providerType: resourceType.toLowerCase(),
            mappedService: service.mappedService,
            category: service.category,
            sourceFile: file.name,
            origin: 'terraform-state',
            approximation: typeof attributes.name === 'string'
              ? (service.mappedService ? 'exact' : 'unmapped')
              : (service.mappedService ? 'type-only' : 'unmapped'),
          });
          seen.value += 1;
        });
      }
    } catch (error) {
      warnings.push(`Could not parse Terraform state JSON from ${file.name}: ${error instanceof Error ? error.message : 'unknown error'}`);
    }
  }
  return output;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

export function buildIaCBaseline(input: BaselineBuildInput): IaCBaseline {
  const warnings: string[] = [];
  const capturedAt = input.importedAt || new Date().toISOString();
  let resources: IaCBaselineResource[] = [];

  switch (input.format) {
    case 'arm':
      resources = parseArmResources(input.files, warnings);
      break;
    case 'bicep':
      resources = parseBicepResources(input.files, warnings);
      break;
    case 'terraform-hcl':
      resources = parseTerraformHclResources(input.files, warnings);
      break;
    case 'terraform-state':
      resources = parseTerraformStateResources(input.files, warnings);
      break;
  }

  resources.sort((left, right) => (
    compareText(left.sourceFile, right.sourceFile)
    || compareText(left.providerType, right.providerType)
    || compareText(left.logicalName, right.logicalName)
  ));

  return {
    version: 1,
    capturedAt,
    format: input.format,
    formatLabel: formatIaCLabel(input.format),
    sourceFiles: input.files.map((file) => ({
      filename: file.name,
      size: file.size ?? file.text.length,
    })),
    resources,
    resourceCount: resources.length,
    unmappedCount: resources.filter((resource) => !resource.mappedService).length,
    warnings,
  };
}

export function restoreIaCBaseline(value: unknown): IaCBaseline | null {
  if (!isRecord(value) || value.version !== 1) return null;
  if (
    typeof value.capturedAt !== 'string'
    || typeof value.format !== 'string'
    || !Array.isArray(value.sourceFiles)
    || !Array.isArray(value.resources)
    || !Array.isArray(value.warnings)
  ) {
    return null;
  }

  const format = value.format;
  if (
    format !== 'arm'
    && format !== 'bicep'
    && format !== 'terraform-hcl'
    && format !== 'terraform-state'
  ) {
    return null;
  }

  const sourceFiles: IaCBaselineSourceFile[] = [];
  for (const sourceFile of value.sourceFiles) {
    if (!isRecord(sourceFile)) return null;
    if (
      typeof sourceFile.filename !== 'string'
      || typeof sourceFile.size !== 'number'
      || !Number.isFinite(sourceFile.size)
      || sourceFile.size < 0
    ) {
      return null;
    }
    sourceFiles.push({ filename: sourceFile.filename, size: sourceFile.size });
  }

  const resources: IaCBaselineResource[] = [];
  for (const resource of value.resources) {
    if (!isRecord(resource)) return null;
    if (
      typeof resource.id !== 'string'
      || typeof resource.logicalName !== 'string'
      || typeof resource.providerType !== 'string'
      || typeof resource.sourceFile !== 'string'
      || typeof resource.origin !== 'string'
      || typeof resource.approximation !== 'string'
    ) {
      return null;
    }
    if (
      resource.origin !== 'arm'
      && resource.origin !== 'bicep'
      && resource.origin !== 'terraform-hcl'
      && resource.origin !== 'terraform-state'
    ) {
      return null;
    }
    if (
      resource.approximation !== 'exact'
      && resource.approximation !== 'type-only'
      && resource.approximation !== 'expression'
      && resource.approximation !== 'unmapped'
    ) {
      return null;
    }
    if (resource.resourceName !== null && resource.resourceName !== undefined && typeof resource.resourceName !== 'string') return null;
    if (resource.mappedService !== null && resource.mappedService !== undefined && typeof resource.mappedService !== 'string') return null;
    if (resource.category !== null && resource.category !== undefined && typeof resource.category !== 'string') return null;
    if (resource.notes !== undefined && typeof resource.notes !== 'string') return null;

    resources.push({
      id: resource.id,
      logicalName: resource.logicalName,
      resourceName: typeof resource.resourceName === 'string' ? resource.resourceName : null,
      providerType: resource.providerType,
      mappedService: typeof resource.mappedService === 'string' ? resource.mappedService : null,
      category: typeof resource.category === 'string' ? resource.category : null,
      sourceFile: resource.sourceFile,
      origin: resource.origin,
      approximation: resource.approximation,
      notes: typeof resource.notes === 'string' ? resource.notes : undefined,
    });
  }

  return {
    version: 1,
    capturedAt: value.capturedAt,
    format,
    formatLabel: typeof value.formatLabel === 'string' ? value.formatLabel : formatIaCLabel(format),
    sourceFiles,
    resources,
    resourceCount: typeof value.resourceCount === 'number' && Number.isFinite(value.resourceCount)
      ? value.resourceCount
      : resources.length,
    unmappedCount: typeof value.unmappedCount === 'number' && Number.isFinite(value.unmappedCount)
      ? value.unmappedCount
      : resources.filter((resource) => !resource.mappedService).length,
    warnings: value.warnings.filter((warning): warning is string => typeof warning === 'string'),
  };
}

export function describeDiagramResources(nodes: Node[]): DiagramResourceDescriptor[] {
  return nodes
    .filter((node) => node.type === 'azureNode')
    .map((node) => {
      const label = typeof node.data?.label === 'string' ? node.data.label : node.id;
      const serviceName = typeof node.data?.serviceName === 'string'
        ? node.data.serviceName
        : label;
      const primaryResolution = resolveFromProviderOrService(serviceName);
      const secondaryResolution = primaryResolution.mappedService
        ? primaryResolution
        : resolveFromProviderOrService(label);
      return {
        id: node.id,
        label,
        serviceName,
        providerType: secondaryResolution.providerType,
        mappedService: secondaryResolution.mappedService,
        category: secondaryResolution.category,
        approximation: secondaryResolution.mappedService ? 'type-only' : 'unmapped',
      };
    });
}

export function compareDiagramToBaseline(nodes: Node[], baseline: IaCBaseline | null): IaCComparisonReport | null {
  if (!baseline) return null;
  const diagramResources = describeDiagramResources(nodes);
  const candidateMatches: CandidateMatch[] = [];

  for (const baselineResource of baseline.resources) {
    const comparableBaseline = toComparableBaseline(baselineResource);
    for (const diagramResource of diagramResources) {
      const comparableDiagram = toComparableDiagram(diagramResource);
      const match = classifyMatch(comparableBaseline, comparableDiagram);
      if (!match) continue;
      candidateMatches.push({
        baseline: baselineResource,
        diagram: diagramResource,
        confidence: match.confidence,
        reason: match.reason,
        score: match.score,
      });
    }
  }

  candidateMatches.sort((left, right) => (
    right.score - left.score
    || compareText(left.baseline.id, right.baseline.id)
    || compareText(left.diagram.id, right.diagram.id)
  ));

  const matchedBaseline = new Set<string>();
  const matchedDiagram = new Set<string>();
  const matched: MatchedIaCResource[] = [];

  for (const candidate of candidateMatches) {
    if (matchedBaseline.has(candidate.baseline.id) || matchedDiagram.has(candidate.diagram.id)) continue;
    matchedBaseline.add(candidate.baseline.id);
    matchedDiagram.add(candidate.diagram.id);
    matched.push({
      baseline: candidate.baseline,
      diagram: candidate.diagram,
      confidence: candidate.confidence,
      reason: candidate.reason,
    });
  }

  matched.sort((left, right) => (
    compareText(left.baseline.sourceFile, right.baseline.sourceFile)
    || compareText(left.baseline.logicalName, right.baseline.logicalName)
  ));

  return {
    matched,
    sourceOnly: baseline.resources.filter((resource) => !matchedBaseline.has(resource.id)),
    diagramOnly: diagramResources.filter((resource) => !matchedDiagram.has(resource.id)),
    approximateMatches: matched.filter((resource) => resource.confidence !== 'exact' && resource.confidence !== 'normalized').length,
  };
}

function normalizeDriftAction(action: DriftAction): string {
  switch (action) {
    case 'create':
      return 'Create';
    case 'update':
      return 'Update';
    case 'delete':
      return 'Delete';
    case 'replace':
      return 'Replace';
    case 'no-op':
      return 'No-op';
    default:
      return 'Review';
  }
}

function emptyDriftCounts(): Record<DriftAction, number> {
  return {
    create: 0,
    update: 0,
    delete: 0,
    replace: 0,
    'no-op': 0,
    other: 0,
  };
}

function terraformActionsToDriftAction(actions: string[]): DriftAction {
  const normalized = actions.map((action) => action.toLowerCase());
  if (normalized.includes('delete') && normalized.includes('create')) return 'replace';
  if (normalized.includes('create')) return 'create';
  if (normalized.includes('update')) return 'update';
  if (normalized.includes('delete')) return 'delete';
  if (normalized.includes('no-op') || normalized.includes('read')) return 'no-op';
  return 'other';
}

function azureChangeTypeToDriftAction(changeType: string): DriftAction {
  switch (changeType.toLowerCase()) {
    case 'create':
      return 'create';
    case 'delete':
      return 'delete';
    case 'modify':
    case 'deploy':
      return 'update';
    case 'nochange':
    case 'ignore':
      return 'no-op';
    default:
      return 'other';
  }
}

function stringifyChangeAddress(value: unknown, fallbackType: string | null, fallbackName: string | null): string {
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (fallbackType && fallbackName) return `${fallbackType}/${fallbackName}`;
  return fallbackName || fallbackType || 'resource';
}

export function parseDeploymentPlan(fileName: string, text: string, importedAt = new Date().toISOString()): DriftPlanSummary {
  let parsed: unknown;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    throw new Error(`Invalid JSON in ${fileName}: ${error instanceof Error ? error.message : 'unknown error'}`);
  }

  if (isRecord(parsed) && Array.isArray(parsed.resource_changes)) {
    const counts = emptyDriftCounts();
    const changes = parsed.resource_changes
      .filter((change): change is Record<string, unknown> => isRecord(change))
      .map((change, index) => {
        const actions = isRecord(change.change) && Array.isArray(change.change.actions)
          ? change.change.actions.filter((value): value is string => typeof value === 'string')
          : [];
        const action = terraformActionsToDriftAction(actions);
        counts[action] += 1;
        const resourceType = typeof change.type === 'string' ? change.type : null;
        const resourceName = typeof change.name === 'string' ? change.name : null;
        return {
          id: `tf-plan-${index + 1}`,
          address: typeof change.address === 'string'
            ? change.address
            : stringifyChangeAddress(null, resourceType, resourceName),
          action,
          actionText: normalizeDriftAction(action),
          resourceType,
          resourceName,
          provider: 'terraform' as const,
          rawActions: actions,
        };
      });
    return {
      kind: 'terraform-plan',
      importedAt,
      sourceFile: fileName,
      changeCounts: counts,
      changes,
    };
  }

  const azureChanges = isRecord(parsed) && Array.isArray(parsed.changes)
    ? parsed.changes
    : isRecord(parsed)
      && isRecord(parsed.properties)
      && Array.isArray(parsed.properties.changes)
        ? parsed.properties.changes
        : null;

  if (azureChanges) {
    const counts = emptyDriftCounts();
    const changes = azureChanges
      .filter((change): change is Record<string, unknown> => isRecord(change))
      .map((change, index) => {
        const after = isRecord(change.after) ? change.after : null;
        const before = isRecord(change.before) ? change.before : null;
        const resourceType = typeof change.resourceType === 'string'
          ? change.resourceType
          : typeof after?.resourceType === 'string'
            ? after.resourceType
            : typeof before?.resourceType === 'string'
              ? before.resourceType
              : null;
        const resourceName = typeof change.name === 'string'
          ? change.name
          : typeof after?.name === 'string'
            ? after.name
            : typeof before?.name === 'string'
              ? before.name
              : null;
        const rawType = typeof change.changeType === 'string' ? change.changeType : 'Review';
        const action = azureChangeTypeToDriftAction(rawType);
        counts[action] += 1;
        return {
          id: `azure-whatif-${index + 1}`,
          address: stringifyChangeAddress(change.resourceId, resourceType, resourceName),
          action,
          actionText: normalizeDriftAction(action),
          resourceType: resourceType ? resourceType.toLowerCase() : null,
          resourceName,
          provider: 'azure' as const,
          rawActions: [rawType],
        };
      });
    return {
      kind: 'azure-what-if',
      importedAt,
      sourceFile: fileName,
      changeCounts: counts,
      changes,
    };
  }

  throw new Error('Unrecognized plan format. Expected Azure what-if JSON or Terraform show -json output.');
}

function slugToken(value: string): string {
  const normalized = value.toLowerCase().replace(/[^a-z0-9]+/g, '');
  return normalized || 'resource';
}

function bicepNameExpression(prefixToken: string, index: number, maxLength?: number): string {
  const expression = `toLower(replace('\${namePrefix}-${prefixToken}-${pad2(index)}', '-', ''))`;
  return maxLength ? `take(${expression}, ${maxLength})` : expression;
}

function bicepDisplayNameExpression(prefixToken: string, index: number): string {
  return `toLower('\${namePrefix}-${prefixToken}-${pad2(index)}')`;
}

function terraformLiteralName(prefixToken: string, index: number, maxLength?: number): string {
  const base = `${prefixToken}${pad2(index)}`;
  const expression = `replace(lower("\${var.name_prefix}-${base}"), "-", "")`;
  return maxLength ? `substr(${expression}, 0, ${maxLength})` : expression;
}

function terraformDisplayName(prefixToken: string, index: number): string {
  return `lower("\${var.name_prefix}-${prefixToken}-${pad2(index)}")`;
}

function exportResourceDescriptors(nodes: Node[]): ExportResourceDescriptor[] {
  return describeDiagramResources(nodes).map((resource) => {
    const providerType = resource.providerType
      || (looksLikeProviderType(resource.serviceName) ? resource.serviceName.toLowerCase() : null);
    const mappedFromProvider = providerType ? resolveFromProviderOrService(providerType) : null;
    const canonicalService = mappedFromProvider?.mappedService || resource.mappedService;
    return {
      nodeId: resource.id,
      label: resource.label,
      canonicalService,
      providerType,
      mappedService: resource.mappedService,
    };
  });
}

function guessArmType(resource: ExportResourceDescriptor): string | null {
  if (resource.providerType && resource.providerType.startsWith('microsoft.')) return resource.providerType;
  if (resource.providerType && resource.providerType.startsWith('azurerm_')) {
    return TERRAFORM_TYPE_MAP[resource.providerType]?.armType || null;
  }
  if (!resource.canonicalService) return null;
  const canonical = normalizeToken(resource.canonicalService);
  const candidate = Object.values(TERRAFORM_TYPE_MAP)
    .map((mapping) => mapping.armType)
    .find((armType) => normalizeToken(lookupServiceMeta(armType)?.name || null) === canonical);
  if (candidate) return candidate;

  const directMap: Record<string, string> = {
    'appservice': 'microsoft.web/sites',
    'appserviceplan': 'microsoft.web/serverfarms',
    'azureopenai': 'microsoft.cognitiveservices/accounts',
    'azureaisearch': 'microsoft.search/searchservices',
    'containerapps': 'microsoft.app/containerapps',
    'containerappsenvironment': 'microsoft.app/managedenvironments',
    'cosmosdb': 'microsoft.documentdb/databaseaccounts',
    'eventhubs': 'microsoft.eventhub/namespaces',
    'loganalytics': 'microsoft.operationalinsights/workspaces',
    'managedidentity': 'microsoft.managedidentity/userassignedidentities',
    'postgresql': 'microsoft.dbforpostgresql/flexibleservers',
    'mysql': 'microsoft.dbformysql/flexibleservers',
    'servicebus': 'microsoft.servicebus/namespaces',
    'sqlserver': 'microsoft.sql/servers',
    'sqldatabase': 'microsoft.sql/servers/databases',
    'storageaccount': 'microsoft.storage/storageaccounts',
    'staticwebapps': 'microsoft.web/staticsites',
    'virtualnetwork': 'microsoft.network/virtualnetworks',
  };
  return canonical ? directMap[canonical] || null : null;
}

function bicepStubForResource(
  armType: string,
  resource: ExportResourceDescriptor,
  index: number,
): string | null {
  const symbol = `${slugToken((resource.canonicalService || armType).split('/').pop() || 'resource')}${pad2(index)}`;
  switch (armType) {
    case 'microsoft.storage/storageaccounts':
      return [
        `resource ${symbol} 'Microsoft.Storage/storageAccounts@2023-05-01' = {`,
        `  name: ${bicepNameExpression('stg', index, 24)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard_LRS'`,
        `  }`,
        `  kind: 'StorageV2'`,
        `  properties: {`,
        `    accessTier: 'Hot'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.network/virtualnetworks':
      return [
        `resource ${symbol} 'Microsoft.Network/virtualNetworks@2023-11-01' = {`,
        `  name: ${bicepDisplayNameExpression('vnet', index)}`,
        `  location: location`,
        `  properties: {`,
        `    addressSpace: {`,
        `      addressPrefixes: [`,
        `        '10.${index}.0.0/16'`,
        `      ]`,
        `    }`,
        `    subnets: [`,
        `      {`,
        `        name: 'default'`,
        `        properties: {`,
        `          addressPrefix: '10.${index}.0.0/24'`,
        `        }`,
        `      }`,
        `    ]`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.keyvault/vaults':
      return [
        `resource ${symbol} 'Microsoft.KeyVault/vaults@2023-07-01' = {`,
        `  name: ${bicepDisplayNameExpression('kv', index)}`,
        `  location: location`,
        `  properties: {`,
        `    tenantId: subscription().tenantId`,
        `    sku: {`,
        `      family: 'A'`,
        `      name: 'standard'`,
        `    }`,
        `    enableRbacAuthorization: true`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.containerregistry/registries':
      return [
        `resource ${symbol} 'Microsoft.ContainerRegistry/registries@2023-07-01' = {`,
        `  name: ${bicepNameExpression('acr', index, 50)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Basic'`,
        `  }`,
        `  properties: {`,
        `    adminUserEnabled: false`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.operationalinsights/workspaces':
      return [
        `resource ${symbol} 'Microsoft.OperationalInsights/workspaces@2023-09-01' = {`,
        `  name: ${bicepDisplayNameExpression('log', index)}`,
        `  location: location`,
        `  properties: {`,
        `    sku: {`,
        `      name: 'PerGB2018'`,
        `    }`,
        `    retentionInDays: 30`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.insights/components':
      return [
        `resource ${symbol} 'Microsoft.Insights/components@2020-02-02' = {`,
        `  name: ${bicepDisplayNameExpression('appi', index)}`,
        `  location: location`,
        `  kind: 'web'`,
        `  properties: {`,
        `    Application_Type: 'web'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.servicebus/namespaces':
      return [
        `resource ${symbol} 'Microsoft.ServiceBus/namespaces@2023-01-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('sb', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard'`,
        `    tier: 'Standard'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.eventhub/namespaces':
      return [
        `resource ${symbol} 'Microsoft.EventHub/namespaces@2023-01-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('eh', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard'`,
        `    tier: 'Standard'`,
        `    capacity: 1`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.documentdb/databaseaccounts':
      return [
        `resource ${symbol} 'Microsoft.DocumentDB/databaseAccounts@2024-05-15' = {`,
        `  name: ${bicepNameExpression('cosmos', index, 44)}`,
        `  location: location`,
        `  kind: 'GlobalDocumentDB'`,
        `  properties: {`,
        `    databaseAccountOfferType: 'Standard'`,
        `    locations: [`,
        `      {`,
        `        locationName: location`,
        `        failoverPriority: 0`,
        `      }`,
        `    ]`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.search/searchservices':
      return [
        `resource ${symbol} 'Microsoft.Search/searchServices@2023-11-01' = {`,
        `  name: ${bicepNameExpression('search', index, 60)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'basic'`,
        `  }`,
        `  properties: {`,
        `    hostingMode: 'default'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.managedidentity/userassignedidentities':
      return [
        `resource ${symbol} 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {`,
        `  name: ${bicepDisplayNameExpression('id', index)}`,
        `  location: location`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/staticsites':
      return [
        `resource ${symbol} 'Microsoft.Web/staticSites@2023-12-01' = {`,
        `  name: ${bicepDisplayNameExpression('swa', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard'`,
        `    tier: 'Standard'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/serverfarms':
      return [
        `resource ${symbol} 'Microsoft.Web/serverfarms@2023-12-01' = {`,
        `  name: ${bicepDisplayNameExpression('plan', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'B1'`,
        `    tier: 'Basic'`,
        `  }`,
        `  kind: 'linux'`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/sites':
      return [
        `resource ${symbol} 'Microsoft.Web/sites@2023-12-01' = {`,
        `  name: ${bicepDisplayNameExpression('app', index)}`,
        `  location: location`,
        `  kind: 'app,linux'`,
        `  properties: {`,
        `    serverFarmId: '/subscriptions/<subscription-id>/resourceGroups/\${resourceGroup().name}/providers/Microsoft.Web/serverfarms/<app-service-plan-name>'`,
        `    httpsOnly: true`,
        `    siteConfig: {`,
        `      linuxFxVersion: 'DOTNETCORE|8.0'`,
        `    }`,
        `  }`,
        `  tags: commonTags`,
        `  // TODO: point serverFarmId at an explicit App Service Plan resource if you customize the layout.`,
        `}`,
      ].join('\n');
    case 'microsoft.app/managedenvironments':
      return [
        `resource ${symbol} 'Microsoft.App/managedEnvironments@2024-03-01' = {`,
        `  name: ${bicepDisplayNameExpression('cae', index)}`,
        `  location: location`,
        `  properties: {`,
        `    zoneRedundant: false`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.app/containerapps':
      return [
        `resource ${symbol} 'Microsoft.App/containerApps@2024-03-01' = {`,
        `  name: ${bicepDisplayNameExpression('ca', index)}`,
        `  location: location`,
        `  properties: {`,
        `    managedEnvironmentId: '/subscriptions/<subscription-id>/resourceGroups/\${resourceGroup().name}/providers/Microsoft.App/managedEnvironments/<container-app-environment-name>'`,
        `    configuration: {`,
        `      ingress: {`,
        `        external: true`,
        `        targetPort: 80`,
        `      }`,
        `    }`,
        `    template: {`,
        `      containers: [`,
        `        {`,
        `          name: 'app'`,
        `          image: 'mcr.microsoft.com/azuredocs/containerapps-helloworld:latest'`,
        `          resources: {`,
        `            cpu: json('0.5')`,
        `            memory: '1Gi'`,
        `          }`,
        `        }`,
        `      ]`,
        `    }`,
        `  }`,
        `  tags: commonTags`,
        `  // TODO: wire managedEnvironmentId to a shared Container Apps environment if needed.`,
        `}`,
      ].join('\n');
    case 'microsoft.sql/servers':
      return [
        `resource ${symbol} 'Microsoft.Sql/servers@2023-08-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('sql', index)}`,
        `  location: location`,
        `  properties: {`,
        `    administratorLogin: 'sqladminuser'`,
        `    administratorLoginPassword: sqlAdministratorPassword`,
        `  }`,
        `  tags: commonTags`,
        `  // Supply sqlAdministratorPassword through a secure parameter file or Key Vault reference.`,
        `}`,
      ].join('\n');
    case 'microsoft.sql/servers/databases':
      return [
        `resource ${symbol} 'Microsoft.Sql/servers/databases@2023-08-01-preview' = {`,
        `  name: '\${namePrefix}-sql-${pad2(index)}/appdb${pad2(index)}'`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Basic'`,
        `    tier: 'Basic'`,
        `  }`,
        `  tags: commonTags`,
        `  // TODO: ensure the parent SQL server path resolves to a real server before deployment.`,
        `}`,
      ].join('\n');
    case 'microsoft.dbforpostgresql/flexibleservers':
      return [
        `resource ${symbol} 'Microsoft.DBforPostgreSQL/flexibleServers@2023-06-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('pg', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard_B1ms'`,
        `    tier: 'Burstable'`,
        `  }`,
        `  properties: {`,
        `    version: '15'`,
        `    storage: {`,
        `      storageSizeGB: 32`,
        `    }`,
        `    administratorLogin: 'pgadmin'`,
        `    administratorLoginPassword: postgresqlAdministratorPassword`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.dbformysql/flexibleservers':
      return [
        `resource ${symbol} 'Microsoft.DBforMySQL/flexibleServers@2023-06-30' = {`,
        `  name: ${bicepDisplayNameExpression('mysql', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Standard_B1ms'`,
        `    tier: 'Burstable'`,
        `  }`,
        `  properties: {`,
        `    version: '8.0.21'`,
        `    administratorLogin: 'mysqladmin'`,
        `    administratorLoginPassword: mysqlAdministratorPassword`,
        `    storage: {`,
        `      storageSizeGB: 20`,
        `    }`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.apimanagement/service':
      return [
        `resource ${symbol} 'Microsoft.ApiManagement/service@2023-09-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('apim', index)}`,
        `  location: location`,
        `  sku: {`,
        `    name: 'Consumption'`,
        `    capacity: 0`,
        `  }`,
        `  properties: {`,
        `    publisherEmail: 'owner@example.com'`,
        `    publisherName: 'Architecture Owner'`,
        `  }`,
        `  tags: commonTags`,
        `}`,
      ].join('\n');
    case 'microsoft.cognitiveservices/accounts':
      return [
        `resource ${symbol} 'Microsoft.CognitiveServices/accounts@2024-06-01-preview' = {`,
        `  name: ${bicepDisplayNameExpression('ai', index)}`,
        `  location: location`,
        `  kind: 'OpenAI'`,
        `  sku: {`,
        `    name: 'S0'`,
        `  }`,
        `  properties: {`,
        `    customSubDomainName: ${bicepNameExpression('ai', index, 24)}`,
        `  }`,
        `  tags: commonTags`,
        `  // TODO: change kind/SKU if this node represents a non-OpenAI Cognitive Services account.`,
        `}`,
      ].join('\n');
    default:
      return null;
  }
}

function terraformStubForResource(
  armType: string,
  resource: ExportResourceDescriptor,
  index: number,
): string | null {
  const name = `${slugToken((resource.canonicalService || armType).split('/').pop() || 'resource')}_${pad2(index)}`;
  switch (armType) {
    case 'microsoft.storage/storageaccounts':
      return [
        `resource "azurerm_storage_account" "${name}" {`,
        `  name                     = ${terraformLiteralName('stg', index, 24)}`,
        `  resource_group_name      = azurerm_resource_group.main.name`,
        `  location                 = azurerm_resource_group.main.location`,
        `  account_tier             = "Standard"`,
        `  account_replication_type = "LRS"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.network/virtualnetworks':
      return [
        `resource "azurerm_virtual_network" "${name}" {`,
        `  name                = ${terraformDisplayName('vnet', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  address_space       = ["10.${index}.0.0/16"]`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.keyvault/vaults':
      return [
        `resource "azurerm_key_vault" "${name}" {`,
        `  name                = ${terraformDisplayName('kv', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  tenant_id           = data.azurerm_client_config.current.tenant_id`,
        `  sku_name            = "standard"`,
        `  purge_protection_enabled = true`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.containerregistry/registries':
      return [
        `resource "azurerm_container_registry" "${name}" {`,
        `  name                = ${terraformDisplayName('acr', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku                 = "Basic"`,
        `  admin_enabled       = false`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.operationalinsights/workspaces':
      return [
        `resource "azurerm_log_analytics_workspace" "${name}" {`,
        `  name                = ${terraformDisplayName('log', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku                 = "PerGB2018"`,
        `  retention_in_days   = 30`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.insights/components':
      return [
        `resource "azurerm_application_insights" "${name}" {`,
        `  name                = ${terraformDisplayName('appi', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  application_type    = "web"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.servicebus/namespaces':
      return [
        `resource "azurerm_servicebus_namespace" "${name}" {`,
        `  name                = ${terraformDisplayName('sb', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku                 = "Standard"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.eventhub/namespaces':
      return [
        `resource "azurerm_eventhub_namespace" "${name}" {`,
        `  name                = ${terraformDisplayName('eh', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku                 = "Standard"`,
        `  capacity            = 1`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.documentdb/databaseaccounts':
      return [
        `resource "azurerm_cosmosdb_account" "${name}" {`,
        `  name                = ${terraformDisplayName('cosmos', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  offer_type          = "Standard"`,
        `  kind                = "GlobalDocumentDB"`,
        ``,
        `  consistency_policy {`,
        `    consistency_level = "Session"`,
        `  }`,
        ``,
        `  geo_location {`,
        `    location          = azurerm_resource_group.main.location`,
        `    failover_priority = 0`,
        `  }`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.search/searchservices':
      return [
        `resource "azurerm_search_service" "${name}" {`,
        `  name                = ${terraformDisplayName('search', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku                 = "basic"`,
        `  replica_count       = 1`,
        `  partition_count     = 1`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.managedidentity/userassignedidentities':
      return [
        `resource "azurerm_user_assigned_identity" "${name}" {`,
        `  name                = ${terraformDisplayName('id', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/staticsites':
      return [
        `resource "azurerm_static_web_app" "${name}" {`,
        `  name                = ${terraformDisplayName('swa', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  sku_tier            = "Standard"`,
        `  sku_size            = "Standard"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/serverfarms':
      return [
        `resource "azurerm_service_plan" "${name}" {`,
        `  name                = ${terraformDisplayName('plan', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  os_type             = "Linux"`,
        `  sku_name            = "B1"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.web/sites':
      return [
        `resource "azurerm_linux_web_app" "${name}" {`,
        `  name                = ${terraformDisplayName('app', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  service_plan_id     = "/subscriptions/<subscription-id>/resourceGroups/\${var.resource_group_name}/providers/Microsoft.Web/serverfarms/<app-service-plan-name>"`,
        ``,
        `  site_config {`,
        `    always_on = true`,
        `  }`,
        ``,
        `  tags = local.common_tags`,
        `  # TODO: point service_plan_id at the correct App Service Plan when you finalize the design.`,
        `}`,
      ].join('\n');
    case 'microsoft.app/managedenvironments':
      return [
        `resource "azurerm_container_app_environment" "${name}" {`,
        `  name                       = ${terraformDisplayName('cae', index)}`,
        `  resource_group_name        = azurerm_resource_group.main.name`,
        `  location                   = azurerm_resource_group.main.location`,
        `  log_analytics_workspace_id = "/subscriptions/<subscription-id>/resourceGroups/\${var.resource_group_name}/providers/Microsoft.OperationalInsights/workspaces/<log-analytics-workspace-name>"`,
        ``,
        `  tags = local.common_tags`,
        `  # TODO: replace the placeholder workspace resource ID or wire in a real workspace resource.`,
        `}`,
      ].join('\n');
    case 'microsoft.app/containerapps':
      return [
        `resource "azurerm_container_app" "${name}" {`,
        `  name                         = ${terraformDisplayName('ca', index)}`,
        `  resource_group_name          = azurerm_resource_group.main.name`,
        `  container_app_environment_id = "/subscriptions/<subscription-id>/resourceGroups/\${var.resource_group_name}/providers/Microsoft.App/managedEnvironments/<container-app-environment-name>"`,
        `  revision_mode                = "Single"`,
        ``,
        `  template {`,
        `    container {`,
        `      name   = "app"`,
        `      image  = "mcr.microsoft.com/azuredocs/containerapps-helloworld:latest"`,
        `      cpu    = 0.5`,
        `      memory = "1Gi"`,
        `    }`,
        `  }`,
        ``,
        `  ingress {`,
        `    external_enabled = true`,
        `    target_port      = 80`,
        `    traffic_weight {`,
        `      latest_revision = true`,
        `      percentage      = 100`,
        `    }`,
        `  }`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.sql/servers':
      return [
        `resource "azurerm_mssql_server" "${name}" {`,
        `  name                         = ${terraformDisplayName('sql', index)}`,
        `  resource_group_name          = azurerm_resource_group.main.name`,
        `  location                     = azurerm_resource_group.main.location`,
        `  version                      = "12.0"`,
        `  administrator_login          = "sqladminuser"`,
        `  administrator_login_password = var.sql_administrator_password`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.sql/servers/databases':
      return [
        `resource "azurerm_mssql_database" "${name}" {`,
        `  name      = "appdb${pad2(index)}"`,
        `  server_id = "/subscriptions/<subscription-id>/resourceGroups/\${var.resource_group_name}/providers/Microsoft.Sql/servers/<sql-server-name>"`,
        `  sku_name  = "Basic"`,
        ``,
        `  tags = local.common_tags`,
        `  # TODO: ensure server_id points at the intended SQL server.`,
        `}`,
      ].join('\n');
    case 'microsoft.dbforpostgresql/flexibleservers':
      return [
        `resource "azurerm_postgresql_flexible_server" "${name}" {`,
        `  name                   = ${terraformDisplayName('pg', index)}`,
        `  resource_group_name    = azurerm_resource_group.main.name`,
        `  location               = azurerm_resource_group.main.location`,
        `  sku_name               = "B_Standard_B1ms"`,
        `  version                = "15"`,
        `  storage_mb             = 32768`,
        `  administrator_login    = "pgadmin"`,
        `  administrator_password = var.postgresql_administrator_password`,
        `  zone                   = "1"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.dbformysql/flexibleservers':
      return [
        `resource "azurerm_mysql_flexible_server" "${name}" {`,
        `  name                   = ${terraformDisplayName('mysql', index)}`,
        `  resource_group_name    = azurerm_resource_group.main.name`,
        `  location               = azurerm_resource_group.main.location`,
        `  sku_name               = "B_Standard_B1ms"`,
        `  version                = "8.0.21"`,
        `  storage {`,
        `    size_gb = 20`,
        `  }`,
        `  administrator_login    = "mysqladmin"`,
        `  administrator_password = var.mysql_administrator_password`,
        `  zone                   = "1"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.apimanagement/service':
      return [
        `resource "azurerm_api_management" "${name}" {`,
        `  name                = ${terraformDisplayName('apim', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  publisher_name      = "Architecture Owner"`,
        `  publisher_email     = "owner@example.com"`,
        `  sku_name            = "Consumption_0"`,
        ``,
        `  tags = local.common_tags`,
        `}`,
      ].join('\n');
    case 'microsoft.cognitiveservices/accounts':
      return [
        `resource "azurerm_cognitive_account" "${name}" {`,
        `  name                = ${terraformDisplayName('ai', index)}`,
        `  resource_group_name = azurerm_resource_group.main.name`,
        `  location            = azurerm_resource_group.main.location`,
        `  kind                = "OpenAI"`,
        `  sku_name            = "S0"`,
        `  custom_subdomain_name = ${terraformLiteralName('ai', index, 24)}`,
        ``,
        `  tags = local.common_tags`,
        `  # TODO: change kind and network settings if this represents another Cognitive Services workload.`,
        `}`,
      ].join('\n');
    default:
      return null;
  }
}

interface DatabasePasswordParameter {
  bicepName: string;
  terraformName: string;
  description: string;
}

const DATABASE_PASSWORD_PARAMETERS: Record<string, DatabasePasswordParameter> = {
  'microsoft.sql/servers': {
    bicepName: 'sqlAdministratorPassword',
    terraformName: 'sql_administrator_password',
    description: 'Administrator password for generated Azure SQL servers.',
  },
  'microsoft.dbforpostgresql/flexibleservers': {
    bicepName: 'postgresqlAdministratorPassword',
    terraformName: 'postgresql_administrator_password',
    description: 'Administrator password for generated Azure Database for PostgreSQL servers.',
  },
  'microsoft.dbformysql/flexibleservers': {
    bicepName: 'mysqlAdministratorPassword',
    terraformName: 'mysql_administrator_password',
    description: 'Administrator password for generated Azure Database for MySQL servers.',
  },
};

function requiredDatabasePasswordParameters(
  resources: ExportResourceDescriptor[],
): DatabasePasswordParameter[] {
  const requiredTypes = new Set(
    resources
      .map(guessArmType)
      .filter((armType): armType is string => Boolean(
        armType && DATABASE_PASSWORD_PARAMETERS[armType],
      )),
  );
  return Object.entries(DATABASE_PASSWORD_PARAMETERS)
    .filter(([armType]) => requiredTypes.has(armType))
    .map(([, parameter]) => parameter);
}

function bicepHeader(passwordParameters: DatabasePasswordParameter[]): string {
  const lines = [
    `targetScope = 'resourceGroup'`,
    ``,
    `// Design-time starter generated by Microsoft Product Architecture Diagram Builder.`,
    `// Review every placeholder, SKU, identity assignment, and dependency before running what-if or any deployment.`,
    `// This template is intentionally conservative and must not be applied without human review.`,
    ``,
    `@description('Deployment location for starter resources.')`,
    `param location string = resourceGroup().location`,
    ``,
    `@description('Short prefix used to derive starter resource names.')`,
    `@minLength(3)`,
    `@maxLength(12)`,
    `param namePrefix string = 'starter'`,
    ``,
    `@description('Environment tag value for the starter deployment.')`,
    `param environment string = 'dev'`,
  ];
  passwordParameters.forEach((parameter) => {
    lines.push(
      ``,
      `@description('${parameter.description}')`,
      `@secure()`,
      `param ${parameter.bicepName} string`,
    );
  });
  lines.push(
    ``,
    `var commonTags = {`,
    `  generatedBy: 'Microsoft Product Architecture Diagram Builder'`,
    `  starterTemplate: 'true'`,
    `  environment: environment`,
    `}`,
  );
  return lines.join('\n');
}

function terraformHeader(passwordParameters: DatabasePasswordParameter[]): string {
  const lines = [
    `terraform {`,
    `  required_version = ">= 1.5.0"`,
    `  required_providers {`,
    `    azurerm = {`,
    `      source  = "hashicorp/azurerm"`,
    `      version = "~> 4.0"`,
    `    }`,
    `  }`,
    `}`,
    ``,
    `provider "azurerm" {`,
    `  features {}`,
    `}`,
    ``,
    `# Design-time starter generated by Microsoft Product Architecture Diagram Builder.`,
    `# Review every placeholder, SKU, identity assignment, and dependency before planning or applying.`,
    ``,
    `variable "location" {`,
    `  type    = string`,
    `  default = "eastus"`,
    `}`,
    ``,
    `variable "resource_group_name" {`,
    `  type    = string`,
    `  default = "rg-azure-diagram-starter"`,
    `}`,
    ``,
    `variable "name_prefix" {`,
    `  type    = string`,
    `  default = "starter"`,
    `}`,
  ];
  passwordParameters.forEach((parameter) => {
    lines.push(
      ``,
      `variable "${parameter.terraformName}" {`,
      `  description = "${parameter.description}"`,
      `  type        = string`,
      `  sensitive   = true`,
      `}`,
    );
  });
  lines.push(
    ``,
    `locals {`,
    `  common_tags = {`,
    `    generated_by    = "Microsoft Product Architecture Diagram Builder"`,
    `    starter_template = "true"`,
    `  }`,
    `}`,
    ``,
    `data "azurerm_client_config" "current" {}`,
    ``,
    `resource "azurerm_resource_group" "main" {`,
    `  name     = var.resource_group_name`,
    `  location = var.location`,
    ``,
    `  tags = local.common_tags`,
    `}`,
  );
  return lines.join('\n');
}

/**
 * Reads back the address the emitted block declares, so the dependency
 * clauses reference exactly what was written rather than a second, parallel
 * derivation of the naming rules that could drift away from it.
 *
 * Every stub emits exactly one top-level resource, and its declaration is
 * always the first line of the block.
 */
function declaredAddress(block: string, format: StarterTemplateFormat): string | null {
  const firstLine = block.split('\n', 1)[0];
  if (format === 'bicep') {
    return /^resource\s+([A-Za-z_][A-Za-z0-9_]*)\s/.exec(firstLine)?.[1] ?? null;
  }
  const match = /^resource\s+"([^"]+)"\s+"([^"]+)"/.exec(firstLine);
  return match ? `${match[1]}.${match[2]}` : null;
}

/**
 * The arrow the user sees is not always `source -> target`: an edge with
 * `direction: 'reverse'` keeps its stored tuple and only moves the arrowhead,
 * so following the tuple would emit the ordering backwards relative to what
 * was drawn. This mirrors normalizeDirectedAdjacency in layoutPresets.ts,
 * which is the codebase's existing reader of the same field.
 *
 * An arrow drawn from A to B means A talks to B, so B has to exist before A
 * can be deployed against it: the arrow's origin depends on its head. A
 * bidirectional edge contributes both orderings and lets the cycle breaker
 * below choose one, rather than silently privileging the stored tuple.
 *
 * Diagrams routinely contain cycles — two services shown calling each other,
 * or a bidirectional link — and both Bicep and Terraform reject a dependency
 * cycle outright. Emitting one would leave the user with a template that
 * cannot deploy at all, which is strictly worse than emitting none, so back
 * edges found during a depth-first walk are dropped. The walk runs over the
 * caller's emission order, which is already sorted, so the edge that gets
 * dropped is stable between runs.
 */
function resolveDependencies(
  edges: Edge[],
  addressByNodeId: Map<string, string>,
): Map<string, string[]> {
  const order = [...addressByNodeId.keys()];
  const rank = new Map(order.map((nodeId, index) => [nodeId, index]));

  const candidates = new Map<string, Set<string>>(order.map((nodeId) => [nodeId, new Set<string>()]));
  const addCandidate = (dependent: string, dependency: string) => {
    if (dependent === dependency) return;
    if (!addressByNodeId.has(dependent) || !addressByNodeId.has(dependency)) return;
    candidates.get(dependent)!.add(dependency);
  };

  for (const edge of edges) {
    if (!edge.source || !edge.target) continue;
    const direction = (edge.data as { direction?: string } | undefined)?.direction ?? 'forward';
    if (direction === 'reverse') {
      addCandidate(edge.target, edge.source);
    } else if (direction === 'bidirectional') {
      addCandidate(edge.source, edge.target);
      addCandidate(edge.target, edge.source);
    } else {
      addCandidate(edge.source, edge.target);
    }
  }

  const sortedTargets = (nodeId: string) => (
    [...candidates.get(nodeId)!].sort((left, right) => (rank.get(left)! - rank.get(right)!))
  );

  const WHITE = 0;
  const GRAY = 1;
  const BLACK = 2;
  const color = new Map(order.map((nodeId) => [nodeId, WHITE]));
  const accepted = new Map<string, string[]>(order.map((nodeId) => [nodeId, []]));

  for (const root of order) {
    if (color.get(root) !== WHITE) continue;
    const stack: Array<{ nodeId: string; pending: string[] }> = [
      { nodeId: root, pending: sortedTargets(root) },
    ];
    color.set(root, GRAY);
    while (stack.length > 0) {
      const frame = stack[stack.length - 1];
      const next = frame.pending.shift();
      if (next === undefined) {
        color.set(frame.nodeId, BLACK);
        stack.pop();
        continue;
      }
      // A gray successor is still on the stack, so this edge closes a cycle.
      if (color.get(next) === GRAY) continue;
      accepted.get(frame.nodeId)!.push(next);
      if (color.get(next) === WHITE) {
        color.set(next, GRAY);
        stack.push({ nodeId: next, pending: sortedTargets(next) });
      }
    }
  }

  const resolved = new Map<string, string[]>();
  for (const [nodeId, targets] of accepted) {
    if (targets.length === 0) continue;
    resolved.set(nodeId, targets.map((target) => addressByNodeId.get(target)!));
  }
  return resolved;
}

/**
 * Bicep's analyzer cannot infer a lower bound through `take(...)`, so it warns
 * BCP334 on every generated resource name even though the shortest possible
 * value is namePrefix (>= 3, enforced by the parameter) plus a suffix of at
 * least five characters — comfortably above every target minimum. Left alone,
 * a "starter" template greets the user with three warnings on first build.
 */
function withNameLengthSuppression(block: string): string {
  return block
    .split('\n')
    .flatMap((line) => (
      /^ {2}name: take\(/.test(line)
        ? ['  #disable-next-line BCP334 // namePrefix is length-constrained, so the name is never too short', line]
        : [line]
    ))
    .join('\n');
}

/**
 * Every stub is assembled from an array whose last element is exactly the
 * closing brace, so the clause can be placed last — where both languages
 * conventionally put it — without having to match braces through the nested
 * property bags each stub contains. If that ever stops holding, fall back to
 * the declaration line, which both languages also accept.
 */
function withDependencyClause(
  block: string,
  dependencies: string[],
  format: StarterTemplateFormat,
): string {
  if (dependencies.length === 0) return block;
  const lines = block.split('\n');
  const clause = format === 'bicep'
    ? `  dependsOn: [${dependencies.map((address) => `\n    ${address}`).join('')}\n  ]`
    : `  depends_on = [${dependencies.map((address) => `\n    ${address},`).join('')}\n  ]`;
  const closingIndex = lines[lines.length - 1].trim() === '}' ? lines.length - 1 : 1;
  lines.splice(closingIndex, 0, clause);
  return lines.join('\n');
}

export function buildStarterTemplate(
  nodes: Node[],
  format: StarterTemplateFormat,
  edges: Edge[] = [],
): StarterTemplate {
  const resources = exportResourceDescriptors(nodes)
    .sort((left, right) => (
      compareText(left.canonicalService || left.label, right.canonicalService || right.label)
      || compareText(left.label, right.label)
      || compareText(left.nodeId, right.nodeId)
    ));

  const passwordParameters = requiredDatabasePasswordParameters(resources);
  const sections: string[] = [
    format === 'bicep'
      ? bicepHeader(passwordParameters)
      : terraformHeader(passwordParameters),
  ];
  let supportedResourceCount = 0;
  let todoCount = 0;

  // First pass: emit every block and record the address it declares. A
  // resource that fell through to a TODO comment has no address, so nothing
  // can depend on it and it cannot depend on anything.
  const emitted: Array<{ nodeId: string; title: string; block: string }> = [];
  const addressByNodeId = new Map<string, string>();

  resources.forEach((resource, index) => {
    const armType = guessArmType(resource);
    const block = armType
      ? format === 'bicep'
        ? bicepStubForResource(armType, resource, index + 1)
        : terraformStubForResource(armType, resource, index + 1)
      : null;
    const title = resource.canonicalService || resource.label;

    const address = block ? declaredAddress(block, format) : null;
    if (!armType || !block || !address) {
      const commentPrefix = format === 'bicep' ? '//' : '#';
      sections.push([
        `${commentPrefix} TODO: Model unsupported service "${title}"`,
        `${commentPrefix}       Diagram node: ${resource.label}`,
        `${commentPrefix}       Detected provider/service: ${resource.providerType || resource.mappedService || 'unmapped'}`,
      ].join('\n'));
      todoCount += 1;
      return;
    }

    emitted.push({
      nodeId: resource.nodeId,
      title,
      block: format === 'bicep' ? withNameLengthSuppression(block) : block,
    });
    addressByNodeId.set(resource.nodeId, address);
    supportedResourceCount += 1;
  });

  // Second pass: now that every address is known, translate the diagram's
  // connections into deployment ordering.
  const dependencies = resolveDependencies(edges, addressByNodeId);
  let dependencyCount = 0;
  for (const item of emitted) {
    const addresses = dependencies.get(item.nodeId) ?? [];
    dependencyCount += addresses.length;
    sections.push([
      format === 'bicep' ? `// ${item.title}` : `# ${item.title}`,
      withDependencyClause(item.block, addresses, format),
    ].join('\n'));
  }

  if (resources.length === 0) {
    const commentPrefix = format === 'bicep' ? '//' : '#';
    sections.push(`${commentPrefix} TODO: Add Azure service nodes to the diagram before exporting a starter template.`);
    todoCount += 1;
  }

  return {
    format,
    fileName: format === 'bicep' ? 'main.bicep' : 'main.tf',
    content: sections.join('\n\n'),
    supportedResourceCount,
    todoCount,
    dependencyCount,
  };
}
