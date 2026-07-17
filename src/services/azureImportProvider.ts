// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Unified data provider for "Import from Azure".
 *
 * Chooses between two modes transparently:
 *   • Delegated (hosted): VITE_AZURE_AD_CLIENT_ID is set → the browser signs the
 *     user in (MSAL) and calls Azure Resource Manager / Resource Graph DIRECTLY
 *     with the user's token. Everything is RBAC-scoped to that user; the app's
 *     server identity is never involved.
 *   • Server (local / self-host): no client ID → calls the token server's
 *     /api/azure/* routes (gated by AZURE_IMPORT_ENABLED, using the server's
 *     DefaultAzureCredential).
 *
 * Both modes return the same shapes and feed the same Resource Graph adapter.
 */

import {
  listSubscriptions as serverListSubscriptions,
  listResourceGroups as serverListResourceGroups,
  type AzureSubscription,
  type AzureResourceGroup,
} from './azureImport';
import { queryResourceGroupResources as serverQueryResources, type ArgResource } from './resourceGraphAdapter';
import { isDelegatedAuthConfigured, getArmToken, getSignedInName, signIn } from './msalAuth';

const ARM = 'https://management.azure.com';
const GUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const RG_NAME_RE = /^[A-Za-z0-9._()-]{1,90}$/;

export function isDelegatedMode(): boolean {
  return isDelegatedAuthConfigured();
}

function validateSelection(subscriptionId: string, resourceGroup?: string): void {
  if (!GUID_RE.test(subscriptionId)) throw new Error('Invalid Azure subscription ID.');
  if (resourceGroup !== undefined && (!RG_NAME_RE.test(resourceGroup) || resourceGroup.endsWith('.'))) {
    throw new Error('Invalid Azure resource group name.');
  }
}

function armUrl(pathOrUrl: string): string {
  const url = new URL(pathOrUrl, ARM);
  if (url.origin !== ARM) throw new Error('Azure Resource Manager returned an invalid continuation URL.');
  return url.toString();
}

async function armFetchAll(token: string, initialPath: string): Promise<any[]> {
  const values: any[] = [];
  let next: string | null = armUrl(initialPath);
  for (let page = 0; next && page < 100; page += 1) {
    const response: Response = await fetch(next, { headers: { Authorization: `Bearer ${token}` } });
    if (!response.ok) throw new Error(`Azure Resource Manager request failed (${response.status})`);
    const data: { value?: unknown; nextLink?: unknown } = await response.json();
    if (Array.isArray(data.value)) values.push(...data.value);
    next = typeof data.nextLink === 'string' && data.nextLink ? armUrl(data.nextLink) : null;
  }
  if (next) throw new Error('Azure Resource Manager returned too many result pages.');
  return values;
}

// ── Delegated (browser-direct) ────────────────────────────────────────────────

async function delegatedSubscriptions(): Promise<AzureSubscription[]> {
  const token = await getArmToken();
  const values = await armFetchAll(token, '/subscriptions?api-version=2022-12-01');
  return values.map((s: any) => ({ subscriptionId: s.subscriptionId, displayName: s.displayName }));
}

async function delegatedResourceGroups(subscriptionId: string): Promise<AzureResourceGroup[]> {
  validateSelection(subscriptionId);
  const token = await getArmToken();
  const values = await armFetchAll(token, `/subscriptions/${subscriptionId}/resourcegroups?api-version=2021-04-01`);
  return values
    .map((g: any) => ({ name: g.name, location: g.location }))
    .sort((a: AzureResourceGroup, b: AzureResourceGroup) => a.name.localeCompare(b.name));
}

async function delegatedResources(subscriptionId: string, resourceGroup: string): Promise<ArgResource[]> {
  validateSelection(subscriptionId, resourceGroup);
  const token = await getArmToken();
  const query = `Resources | where resourceGroup =~ '${resourceGroup}' | project id, name, type, kind, location, properties | limit 1000`;
  const r = await fetch(`${ARM}/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ subscriptions: [subscriptionId], query, options: { resultFormat: 'objectArray' } }),
  });
  if (!r.ok) throw new Error(`Resource Graph query failed (${r.status})`);
  const data = await r.json();
  return (data.data || []) as ArgResource[];
}

// ── Public unified API ────────────────────────────────────────────────────────

/**
 * Ensure the user is authenticated when in delegated mode; returns the signed-in
 * account name (or undefined in server mode). When not signed in, this starts a
 * full-page sign-in redirect and does not resolve (the app re-opens the import
 * modal when the redirect returns).
 */
export async function ensureSignedIn(): Promise<string | undefined> {
  if (!isDelegatedMode()) return undefined;
  const name = await getSignedInName();
  if (name) return name;
  await signIn(); // navigates away for interactive sign-in
  return undefined;
}

export async function getSubscriptions(): Promise<AzureSubscription[]> {
  return isDelegatedMode() ? delegatedSubscriptions() : serverListSubscriptions();
}

export async function getResourceGroups(subscriptionId: string): Promise<AzureResourceGroup[]> {
  return isDelegatedMode() ? delegatedResourceGroups(subscriptionId) : serverListResourceGroups(subscriptionId);
}

export async function getResources(subscriptionId: string, resourceGroup: string): Promise<ArgResource[]> {
  return isDelegatedMode() ? delegatedResources(subscriptionId, resourceGroup) : serverQueryResources(subscriptionId, resourceGroup);
}

export type { AzureSubscription, AzureResourceGroup };
