import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const policyAssignmentName = 'Azure_Security_Baseline';
const policyAssignmentScope = '/providers/Microsoft.Management/managementGroups/02310481-6752-4d53-8fc4-f597a8e047f4';
const policyAssignmentId = `${policyAssignmentScope}/providers/Microsoft.Authorization/policyAssignments/${policyAssignmentName}`;
const policySetId = '/providers/Microsoft.Authorization/policySetDefinitions/1f3afdf9-d0c9-4c3d-847f-89da613e70a8';
const storageExemptionName = 'AzureDiagarm-Storage-NSP';
const preservedRuntimeNames = [
  'MCP_ENABLED',
  'ALLOW_BYO_AI_ENDPOINTS',
  'AI_BUDGET_STORE',
  'AI_DAILY_TOKEN_BUDGET',
  'AI_MAX_CONCURRENT_REQUESTS',
  'AZURE_TABLES_BUDGET_ENDPOINT',
];

function requireValue(value, label) {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`${label} is required.`);
  }
  return value.trim();
}

function verifyHttpsEndpoint(value, name) {
  let endpoint;
  try {
    endpoint = new URL(value);
  } catch {
    throw new Error(`${name} must be a valid HTTPS service endpoint.`);
  }
  if (
    endpoint.protocol !== 'https:'
    || !endpoint.hostname.includes('.')
    || endpoint.username
    || endpoint.password
    || endpoint.search
    || endpoint.hash
    || !['', '/'].includes(endpoint.pathname)
  ) {
    throw new Error(`${name} must be a valid HTTPS service endpoint.`);
  }
  return value.trim();
}

function validKeyVaultResourceId(value) {
  return /^\/subscriptions\/[^/]+\/resourceGroups\/[^/]+\/providers\/Microsoft\.KeyVault\/vaults\/[^/]+$/i.test(value);
}

function accessStoreIdentity(store) {
  return store.type === 'key-vault'
    ? store.value.toLowerCase()
    : new URL(store.value).origin.toLowerCase();
}

function isRelevantStorageReference(reference) {
  const normalized = reference.toLowerCase();
  return normalized.includes('storageaccount')
    && (
      normalized.includes('restrictnetworkaccessusingvirtualnetworkrules')
      || normalized.includes('useaprivatelinkconnection')
    );
}

export function verifyAccessStore(env) {
  const keyVaultId = typeof env.AZURE_ACCESS_KEY_VAULT_RESOURCE_ID === 'string'
    ? env.AZURE_ACCESS_KEY_VAULT_RESOURCE_ID.trim()
    : '';
  const tablesEndpoint = typeof env.AZURE_TABLES_ACCESS_ENDPOINT === 'string'
    ? env.AZURE_TABLES_ACCESS_ENDPOINT.trim()
    : '';
  if (Boolean(keyVaultId) === Boolean(tablesEndpoint)) {
    throw new Error('Exactly one access-list store must be configured.');
  }
  if (keyVaultId) {
    if (!validKeyVaultResourceId(keyVaultId)) {
      throw new Error('AZURE_ACCESS_KEY_VAULT_RESOURCE_ID must identify a Key Vault resource.');
    }
    return { type: 'key-vault', value: keyVaultId };
  }
  return {
    type: 'tables',
    value: verifyHttpsEndpoint(tablesEndpoint, 'AZURE_TABLES_ACCESS_ENDPOINT'),
  };
}

export function verifyRuntimeConfiguration(env) {
  const accessStore = verifyAccessStore(env);
  if (env.MCP_ENABLED !== 'true') {
    throw new Error('MCP_ENABLED must be explicitly true to preserve production behavior.');
  }
  if (env.ALLOW_BYO_AI_ENDPOINTS !== 'false') {
    throw new Error('ALLOW_BYO_AI_ENDPOINTS must be explicitly false to preserve production policy.');
  }
  if (env.AI_BUDGET_STORE !== 'table') {
    throw new Error('AI_BUDGET_STORE must be explicitly table to preserve production behavior.');
  }
  for (const name of ['AI_DAILY_TOKEN_BUDGET', 'AI_MAX_CONCURRENT_REQUESTS']) {
    const value = requireValue(env[name], name);
    if (!/^[1-9]\d*$/.test(value) || !Number.isSafeInteger(Number(value))) {
      throw new Error(`${name} must be a positive safe integer.`);
    }
  }
  const budgetEndpoint = verifyHttpsEndpoint(
    requireValue(env.AZURE_TABLES_BUDGET_ENDPOINT, 'AZURE_TABLES_BUDGET_ENDPOINT'),
    'AZURE_TABLES_BUDGET_ENDPOINT',
  );
  return { accessStore, budgetEndpoint };
}

export function verifyPreservedRuntimeConfiguration(env, currentEnv) {
  verifyRuntimeConfiguration(env);
  if (!Array.isArray(currentEnv)) {
    throw new Error('The current Container App environment could not be verified.');
  }
  const entries = new Map();
  for (const entry of currentEnv) {
    if (!entry || typeof entry.name !== 'string') continue;
    if (entries.has(entry.name)) {
      throw new Error(`The current Container App has duplicate ${entry.name} settings.`);
    }
    entries.set(entry.name, entry.secretRef ? undefined : entry.value);
  }
  const actual = Object.fromEntries([...entries].map(([name, value]) => [name, value]));
  const expectedStore = verifyAccessStore(env);
  const actualStore = verifyAccessStore(actual);
  if (
    actualStore.type !== expectedStore.type
    || accessStoreIdentity(actualStore) !== accessStoreIdentity(expectedStore)
  ) {
    throw new Error('The requested access-list store differs from the current production store.');
  }
  for (const name of preservedRuntimeNames) {
    if (
      typeof env[name] !== 'string'
      || env[name] === ''
      || actual[name] !== env[name]
    ) {
      throw new Error(`The requested ${name} does not preserve the current production value.`);
    }
  }
}

export function verifyNoReplicaReduction(currentMaxReplicas, desiredMaxReplicas) {
  const parseCount = value => {
    const text = String(value);
    if (!/^[1-9]\d*$/.test(text)) return Number.NaN;
    return Number(text);
  };
  const current = parseCount(currentMaxReplicas);
  const desired = parseCount(desiredMaxReplicas);
  if (
    !Number.isSafeInteger(current)
    || current < 1
    || !Number.isSafeInteger(desired)
    || desired < 1
    || desired < current
  ) {
    throw new Error('The deployment cannot reduce or ambiguously verify the current maximum replica count.');
  }
}

export function verifyAfdCustomDomain(domains, host) {
  if (!Array.isArray(domains) || typeof host !== 'string' || !host.trim()) {
    throw new Error('The approved Front Door custom domain could not be verified.');
  }
  const normalizedHost = host.trim().toLowerCase();
  const approved = domains.filter(domain => (
    typeof domain?.id === 'string'
    && domain.id.trim() !== ''
    && typeof domain.hostName === 'string'
    && domain.hostName.trim().toLowerCase() === normalizedHost
    && domain.domainValidationState === 'Approved'
    && domain.provisioningState === 'Succeeded'
    && domain.tlsSettings?.minimumTlsVersion === 'TLS12'
    && ['CustomerCertificate', 'ManagedCertificate'].includes(
      domain.tlsSettings?.certificateType,
    )
  ));
  if (approved.length !== 1) {
    throw new Error('The Front Door custom domain is missing or does not meet the approved TLS requirements.');
  }
  return approved[0].id;
}

function targetPolicyAssignment(assignments) {
  if (!Array.isArray(assignments)) {
    throw new Error('The target policy assignments could not be verified.');
  }
  const expectedPolicySetId = policySetId.toLowerCase();
  const matches = assignments.filter(assignment => (
    typeof assignment?.name === 'string'
    && assignment.name.toLowerCase() === policyAssignmentName.toLowerCase()
    && typeof (assignment.properties?.policyDefinitionId ?? assignment.policyDefinitionId) === 'string'
    && (assignment.properties?.policyDefinitionId ?? assignment.policyDefinitionId)
      .trim().replace(/\/+$/, '').toLowerCase() === expectedPolicySetId
    && typeof assignment.id === 'string'
    && assignment.id.trim().toLowerCase() === policyAssignmentId.toLowerCase()
    && typeof (assignment.properties?.scope ?? assignment.scope) === 'string'
    && (assignment.properties?.scope ?? assignment.scope)
      .trim().toLowerCase() === policyAssignmentScope.toLowerCase()
  ));
  if (matches.length !== 1) {
    throw new Error('The expected Azure Security Baseline policy assignment is missing or ambiguous.');
  }
  return matches[0];
}

export function policySetNameForAssignments(assignments) {
  const assignment = targetPolicyAssignment(assignments);
  const definitionId = assignment.properties?.policyDefinitionId ?? assignment.policyDefinitionId;
  return definitionId.split('/').filter(Boolean).at(-1);
}

export function verifyStoragePolicyExemption({
  assignments,
  policySet,
  exemption,
  storageScope,
  now = Date.now(),
}) {
  const assignment = targetPolicyAssignment(assignments);
  const assignmentProperties = assignment.properties ?? assignment;
  const assignedPolicySetId = assignmentProperties.policyDefinitionId.trim().replace(/\/+$/, '');
  const returnedPolicySetId = typeof policySet?.id === 'string'
    ? policySet.id.trim().replace(/\/+$/, '')
    : '';
  if (returnedPolicySetId.toLowerCase() !== assignedPolicySetId.toLowerCase()) {
    throw new Error('The assigned storage policy initiative definition could not be verified.');
  }

  const policyDefinitions = policySet?.properties?.policyDefinitions ?? policySet?.policyDefinitions;
  const applicableReferences = (Array.isArray(policyDefinitions)
    ? policyDefinitions
    : [])
    .map(item => item?.policyDefinitionReferenceId)
    .filter(reference => (
      typeof reference === 'string'
      && isRelevantStorageReference(reference)
    ));
  const normalizedApplicable = applicableReferences.map(reference => reference.toLowerCase());
  if (
    applicableReferences.length === 0
    || new Set(normalizedApplicable).size !== normalizedApplicable.length
  ) {
    throw new Error('The target policy initiative has no unambiguous applicable storage references.');
  }
  if (typeof storageScope !== 'string' || !storageScope.trim()) {
    throw new Error('The storage policy exemption must be resource-scoped and Mitigated.');
  }
  const exemptionProperties = exemption?.properties ?? exemption ?? {};
  const expectedExemptionId = `${storageScope.trim().replace(/\/+$/, '')}/providers/Microsoft.Authorization/policyExemptions/${storageExemptionName}`;
  if (
    typeof exemption?.id !== 'string'
    || exemption.id.trim().toLowerCase() !== expectedExemptionId.toLowerCase()
    || exemptionProperties.exemptionCategory !== 'Mitigated'
    || typeof exemptionProperties.policyAssignmentId !== 'string'
    || exemptionProperties.policyAssignmentId.trim().toLowerCase() !== assignment.id.trim().toLowerCase()
  ) {
    throw new Error('The storage policy exemption must be resource-scoped and Mitigated.');
  }
  if (typeof exemptionProperties.expiresOn !== 'string' || !exemptionProperties.expiresOn.trim()) {
    throw new Error('The storage policy exemption must have an expiry.');
  }
  const expiresAt = Date.parse(exemptionProperties.expiresOn);
  if (!Number.isFinite(expiresAt) || expiresAt <= now) {
    throw new Error('The storage policy exemption must have a valid future expiry.');
  }
  const exemptionReferences = Array.isArray(exemptionProperties.policyDefinitionReferenceIds)
    ? exemptionProperties.policyDefinitionReferenceIds
      .filter(reference => typeof reference === 'string')
      .map(reference => reference.toLowerCase())
    : [];
  const missing = normalizedApplicable.filter(reference => !exemptionReferences.includes(reference));
  if (missing.length) {
    throw new Error('The storage policy exemption is missing references from the target initiative.');
  }
  return applicableReferences;
}

function readJsonFile(path, name) {
  if (typeof path !== 'string' || !path.trim()) {
    throw new Error(`${name} file is required.`);
  }
  try {
    return JSON.parse(readFileSync(path, 'utf8'));
  } catch {
    throw new Error(`${name} file could not be read as JSON.`);
  }
}

function runCli() {
  const mode = process.argv[2];
  if (mode === 'runtime-config') {
    verifyRuntimeConfiguration(process.env);
    console.info('Production runtime configuration validated.');
    return;
  }
  if (mode === 'preserve-runtime') {
    verifyPreservedRuntimeConfiguration(
      process.env,
      readJsonFile(process.env.CURRENT_ENV_FILE, 'Current Container App environment'),
    );
    console.info('Production runtime settings match the current configuration.');
    return;
  }
  if (mode === 'afd-domain') {
    const id = verifyAfdCustomDomain(
      readJsonFile(process.env.AFD_CUSTOM_DOMAIN_FILE, 'Front Door custom domain'),
      process.env.APP_HOST,
    );
    console.info(id);
    return;
  }
  if (mode === 'policy-set-name') {
    console.info(policySetNameForAssignments(
      readJsonFile(process.env.POLICY_ASSIGNMENTS_FILE, 'Policy assignments'),
    ));
    return;
  }
  if (mode === 'storage-policy-exemption') {
    verifyStoragePolicyExemption({
      assignments: readJsonFile(process.env.POLICY_ASSIGNMENTS_FILE, 'Policy assignments'),
      policySet: readJsonFile(process.env.POLICY_SET_FILE, 'Policy initiative definition'),
      exemption: readJsonFile(process.env.POLICY_EXEMPTION_FILE, 'Policy exemption'),
      storageScope: process.env.STORAGE_SCOPE,
    });
    console.info('The storage policy exemption matches the target initiative and scope.');
    return;
  }
  if (mode === 'no-replica-reduction') {
    verifyNoReplicaReduction(process.env.CURRENT_MAX_REPLICAS, process.env.MAX_REPLICAS);
    console.info('The deployment does not reduce the maximum replica count.');
    return;
  }
  throw new Error('A supported deployment preflight mode is required.');
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  try {
    runCli();
  } catch (error) {
    console.error(error instanceof Error ? error.message : 'Deployment preflight failed.');
    process.exitCode = 1;
  }
}
