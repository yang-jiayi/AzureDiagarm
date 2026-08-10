// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Microsoft Entra ID (MSAL) delegated auth for "Import from Azure".
 *
 * When VITE_AZURE_AD_CLIENT_ID is configured, the browser signs the user in and
 * acquires an Azure Resource Manager token FOR THAT USER, so all Resource Graph
 * calls run with the user's own RBAC — never the app's identity. This is the
 * path used by the hosted multi-tenant app.
 *
 * When it is NOT configured, the app falls back to the server-identity routes
 * (AZURE_IMPORT_ENABLED) for local / self-host use.
 */

import type {
  PublicClientApplication,
  AccountInfo,
  Configuration,
} from '@azure/msal-browser';

// The MSAL SDK is ~90 kB gzip and is only needed once the user actually starts
// a delegated "Import from Azure" sign-in. Load it lazily via dynamic import so
// it never lands in the main bundle. The module-level config helpers below
// (e.g. isDelegatedAuthConfigured) stay SDK-free and synchronous.
type MsalModule = typeof import('@azure/msal-browser');
let msalModulePromise: Promise<MsalModule> | null = null;
function loadMsal(): Promise<MsalModule> {
  if (!msalModulePromise) {
    msalModulePromise = import('@azure/msal-browser');
  }
  return msalModulePromise;
}

const CLIENT_ID = (import.meta.env.VITE_AZURE_AD_CLIENT_ID as string | undefined) || '';
const AUTHORITY = (import.meta.env.VITE_AZURE_AD_AUTHORITY as string | undefined)
  || 'https://login.microsoftonline.com/organizations';
const REDIRECT_URI = (import.meta.env.VITE_AZURE_AD_REDIRECT_URI as string | undefined)
  || (typeof window !== 'undefined' ? window.location.origin : '');
const ARM_SCOPE = (import.meta.env.VITE_ARM_SCOPE as string | undefined)
  || 'https://management.azure.com/user_impersonation';

// Marker so the app can re-open the "Import from Azure" modal after the
// full-page sign-in redirect returns.
const REOPEN_KEY = 'azimp_reopen_after_signin';

/** Whether per-user delegated auth is configured (client ID present). */
export function isDelegatedAuthConfigured(): boolean {
  return CLIENT_ID.length > 0;
}

let pca: PublicClientApplication | null = null;
let initPromise: Promise<void> | null = null;

/** Returns the initialized PublicClientApplication; call ensureInitialized first. */
function requirePca(): PublicClientApplication {
  if (!pca) {
    throw new Error('MSAL is not initialized yet — call ensureInitialized() first.');
  }
  return pca;
}

async function ensureInitialized(): Promise<void> {
  if (!initPromise) {
    initPromise = (async () => {
      const { PublicClientApplication } = await loadMsal();
      const config: Configuration = {
        auth: { clientId: CLIENT_ID, authority: AUTHORITY, redirectUri: REDIRECT_URI },
        cache: { cacheLocation: 'sessionStorage' },
      };
      const p = new PublicClientApplication(config);
      await p.initialize();
      // Consume any pending redirect response (and clear a dangling #code= from
      // the URL) so leftover interaction state can't wedge later popup calls.
      const result = await p.handleRedirectPromise();
      if (result?.account) p.setActiveAccount(result.account);
      pca = p;
    })();
  }
  return initPromise;
}

function activeAccount(): AccountInfo | null {
  const p = requirePca();
  return p.getActiveAccount() || p.getAllAccounts()[0] || null;
}

/** The signed-in user's display name / UPN, or undefined when not signed in. */
export async function getSignedInName(): Promise<string | undefined> {
  await ensureInitialized();
  return activeAccount()?.username;
}

/**
 * Begin an interactive sign-in via full-page redirect. Popup flow is avoided
 * because the Microsoft login page's Cross-Origin-Opener-Policy severs the
 * popup↔opener channel, causing MSAL popup handshakes to time out. Sets a marker
 * so the app can re-open the import modal when the redirect returns. This call
 * navigates away and does not resolve.
 */
export async function signIn(): Promise<void> {
  await ensureInitialized();
  try { sessionStorage.setItem(REOPEN_KEY, '1'); } catch { /* ignore */ }
  try {
    await requirePca().loginRedirect({ scopes: [ARM_SCOPE], redirectUri: REDIRECT_URI });
  } catch (error) {
    try { sessionStorage.removeItem(REOPEN_KEY); } catch { /* ignore */ }
    throw error;
  }
}

/** Whether the app returned from a sign-in redirect that should re-open import. */
export function consumeReopenFlag(): boolean {
  try {
    const v = sessionStorage.getItem(REOPEN_KEY) === '1';
    if (v) sessionStorage.removeItem(REOPEN_KEY);
    return v;
  } catch { return false; }
}

/**
 * Acquire an ARM access token for the signed-in user. Uses a silent refresh
 * when possible; if interaction is required it falls back to a full-page
 * redirect (which navigates away and does not resolve).
 */
export async function getArmToken(): Promise<string> {
  await ensureInitialized();
  const p = requirePca();
  const account = activeAccount();
  if (account) {
    try {
      const res = await p.acquireTokenSilent({ scopes: [ARM_SCOPE], account });
      p.setActiveAccount(res.account);
      return res.accessToken;
    } catch (err) {
      const { InteractionRequiredAuthError } = await loadMsal();
      if (!(err instanceof InteractionRequiredAuthError)) throw err;
    }
  }
  try { sessionStorage.setItem(REOPEN_KEY, '1'); } catch { /* ignore */ }
  try {
    await p.acquireTokenRedirect({ scopes: [ARM_SCOPE], redirectUri: REDIRECT_URI });
  } catch (error) {
    try { sessionStorage.removeItem(REOPEN_KEY); } catch { /* ignore */ }
    throw error;
  }
  // acquireTokenRedirect navigates away; this is effectively unreachable.
  throw new Error('Redirecting to sign in…');
}

/** Sign the current user out of this app's MSAL cache. */
export async function signOut(): Promise<void> {
  await ensureInitialized();
  const p = requirePca();
  const account = activeAccount();
  await p.clearCache(account ? { account } : undefined);
}
