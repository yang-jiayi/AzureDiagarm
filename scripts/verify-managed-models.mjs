import { execFileSync } from 'node:child_process';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import policy from '../server/astra-policy.js';

export const APPROVED_GPT61_SOL_VERSION = '2026-09-29';
export const APPROVED_CLAUDE_OPUS55_VERSION = '2';
const ARM_API_VERSION = '2025-06-01';
const ACCOUNT_ID_RE = /^\/subscriptions\/[a-f0-9-]{36}\/resourceGroups\/[A-Za-z0-9._()-]+\/providers\/Microsoft\.CognitiveServices\/accounts\/[A-Za-z0-9-]+$/i;

function normalizedEndpoint(value) {
  const parsed = new URL(value);
  return `${parsed.origin}/`;
}

function accountEndpointMatches(account, endpoint) {
  const endpoints = [account?.properties?.endpoint, ...Object.values(account?.properties?.endpoints || {})];
  return endpoints.some(value => {
    try { return normalizedEndpoint(value) === endpoint; } catch { return false; }
  });
}

function resourceIdFor(env, name) {
  const id = env[name];
  if (!ACCOUNT_ID_RE.test(id || '')) {
    throw new Error(`A full ${name} is required for managed model verification.`);
  }
  return id;
}

export function verifyManagedModels(env, read = (url) => JSON.parse(execFileSync('az', [
  'rest', '--method', 'get', '--url', url, '--output', 'json', '--only-show-errors',
], { encoding: 'utf8', timeout: 30_000, stdio: ['ignore', 'pipe', 'pipe'] }))) {
  const models = policy.managedModelsConfiguration(env);
  if (models.configured !== true || !models.sol?.configured || !models.opus?.configured) {
    throw new Error('Both approved managed model configurations are required; no fallback is permitted.');
  }

  const entries = [
    {
      label: 'GPT-6.1 Sol',
      id: resourceIdFor(env, 'AZURE_OPENAI_RESOURCE_ID'),
      config: models.sol,
      expectedFormat: 'OpenAI',
      expectedName: 'gpt-6.1-sol',
      expectedVersion: APPROVED_GPT61_SOL_VERSION,
      expectedKind: 'OpenAI',
    },
    {
      label: 'Claude Opus 5.5',
      id: resourceIdFor(env, 'AZURE_FOUNDRY_RESOURCE_ID'),
      config: models.opus,
      expectedFormat: 'Anthropic',
      expectedName: 'claude-opus-5-5',
      expectedVersion: APPROVED_CLAUDE_OPUS55_VERSION,
      expectedKind: 'AIServices',
    },
  ];

  for (const entry of entries) {
    const accountUrl = `https://management.azure.com${entry.id}?api-version=${ARM_API_VERSION}`;
    const account = read(accountUrl);
    if (
      account?.id?.toLowerCase() !== entry.id.toLowerCase()
      || account?.kind !== entry.expectedKind
      || !accountEndpointMatches(account, entry.config.endpoint)
    ) {
      throw new Error(`${entry.label} endpoint does not belong to the expected Azure account.`);
    }

    const deploymentId = `${entry.id}/deployments/${entry.config.deployment}`;
    const deployment = read(`https://management.azure.com${deploymentId}?api-version=${ARM_API_VERSION}`);
    if (
      deployment?.id?.toLowerCase() !== deploymentId.toLowerCase()
      || deployment?.properties?.provisioningState !== 'Succeeded'
      || deployment?.properties?.model?.format !== entry.expectedFormat
      || deployment?.properties?.model?.name !== entry.expectedName
      || deployment?.properties?.model?.version !== entry.expectedVersion
    ) {
      throw new Error(
        `${entry.label} alias must resolve to a successfully provisioned ${entry.expectedFormat}/${entry.expectedName} model, version ${entry.expectedVersion}. No fallback is permitted.`,
      );
    }
  }
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    verifyManagedModels(process.env);
    console.log('Verified both approved managed model identities and versions through ARM.');
  } catch {
    console.error('Managed model identity verification failed. Check both resource IDs, endpoint ownership, deployment identities/versions and ARM read permissions; no deployment is permitted.');
    process.exitCode = 1;
  }
}
