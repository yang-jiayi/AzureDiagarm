// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
function positiveInteger(value, fallback, name) {
  const number = value === undefined || value === '' ? fallback : Number(value);
  if (!Number.isSafeInteger(number) || number < 1) throw new Error(`${name} must be a positive integer.`);
  return number;
}

function deploymentConfig(env = process.env) {
  const mode = env.APP_DEPLOYMENT_MODE || (env.NODE_ENV === 'production' ? '' : 'local');
  if (!['local', 'public'].includes(mode)) {
    throw new Error('Set APP_DEPLOYMENT_MODE to public or local. Production has no implicit local fallback.');
  }
  const store = env.AI_BUDGET_STORE || (mode === 'local'
    ? 'memory'
    : (env.AZURE_TABLES_BUDGET_ENDPOINT || env.AZURE_TABLES_ENDPOINT ? 'table' : 'cosmos'));
  if (!['memory', 'cosmos', 'table'].includes(store) || (mode === 'public' && store === 'memory')) {
    throw new Error('Public AI budgets require AI_BUDGET_STORE=cosmos or table; memory is local-only.');
  }
  let origin = '';
  try {
    const url = new URL(env.PUBLIC_URL);
    if (url.protocol === 'https:' && !url.username && !url.password && !url.search && !url.hash
      && (url.pathname === '/' || url.pathname === '')) origin = url.origin;
  } catch { /* Validated below in public mode. */ }
  if (mode === 'public') {
    const missing = [];
    if (env.ACCESS_CONTROL_ENABLED !== 'true') missing.push('ACCESS_CONTROL_ENABLED=true');
    if (env.EASY_AUTH_ENABLED !== 'true') missing.push('EASY_AUTH_ENABLED=true (verified platform auth)');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(env.ACCESS_ADMIN_EMAIL || '')) missing.push('ACCESS_ADMIN_EMAIL');
    if (!env.AZURE_ACCESS_KEY_VAULT_RESOURCE_ID && !env.AZURE_TABLES_ACCESS_ENDPOINT) missing.push('access-list store');
    if (!origin) missing.push('PUBLIC_URL (HTTPS origin without credentials, query or fragment)');
    if (!/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(env.FRONT_DOOR_ID || '')) missing.push('FRONT_DOOR_ID');
    if (![env.AZURE_OPENAI_ALLOWED_DEPLOYMENTS, env.AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS]
      .some(list => (list || '').split(',').some(value => value.trim()))) {
      missing.push('AZURE_OPENAI_ALLOWED_DEPLOYMENTS or AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS');
    }
    if (String(env.AZURE_IMPORT_ENABLED || '').toLowerCase() === 'true') missing.push('AZURE_IMPORT_ENABLED must be false on public deployments');
    if (missing.length) throw new Error(`Public deployment is not safe to start: ${missing.join(', ')}.`);
  }
  if (store === 'cosmos' && !env.AZURE_COSMOS_ENDPOINT) throw new Error('Cosmos AI budgets require AZURE_COSMOS_ENDPOINT.');
  if (store === 'table' && !(env.AZURE_TABLES_BUDGET_ENDPOINT || env.AZURE_TABLES_ENDPOINT)) {
    throw new Error('Table AI budgets require AZURE_TABLES_BUDGET_ENDPOINT or AZURE_TABLES_ENDPOINT.');
  }
  return {
    mode, store, origin,
    dailyTokens: positiveInteger(env.AI_DAILY_TOKEN_BUDGET, 250_000, 'AI_DAILY_TOKEN_BUDGET'),
    concurrency: positiveInteger(env.AI_MAX_CONCURRENT_REQUESTS, 2, 'AI_MAX_CONCURRENT_REQUESTS'),
    retentionDays: positiveInteger(env.FEEDBACK_RETENTION_DAYS, 30, 'FEEDBACK_RETENTION_DAYS'),
    legacyRetentionEnabled: env.FEEDBACK_LEGACY_RETENTION_ENABLED === 'true',
  };
}

function createOriginGuard(config) {
  return (req, res, next) => {
    if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) return next();
    let url;
    try { url = new URL(req.get('origin')); } catch { /* Deny absent origins publicly. */ }
    const allowed = config.mode === 'public'
      ? req.get('origin') === config.origin
      : (!req.get('origin') || (['http:', 'https:'].includes(url?.protocol)
        && ['localhost', '127.0.0.1', '[::1]'].includes(url?.hostname)));
    if (!allowed) return res.status(403).json({ error: 'A same-origin request is required.' });
    return next();
  };
}

function assertTrustedAuth(auth) {
  const validation = auth?.globalValidation;
  const aad = auth?.identityProviders?.azureActiveDirectory;
  const registration = aad?.registration;
  const excluded = validation?.excludedPaths;
  if (auth?.platform?.enabled !== true || auth?.httpSettings?.requireHttps !== true
    || validation?.unauthenticatedClientAction !== 'RedirectToLoginPage'
    || validation?.redirectToProvider !== 'azureactivedirectory'
    || !Array.isArray(excluded) || excluded.length !== 1 || excluded[0] !== '/healthz'
    || aad?.enabled === false
    || !/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(registration?.clientId || '')
    || !/^https:\/\/login\.microsoftonline\.com\/[a-f0-9-]{36}\/v2\.0$/i.test(registration?.openIdIssuer || '')
    || !aad?.validation?.allowedAudiences?.includes(registration.clientId)) {
    throw new Error('Verify single-tenant Easy Auth, HTTPS, audience validation, and /healthz as the only excluded path before deployment.');
  }
}

if (require.main === module) {
  try {
    if (process.argv.includes('--verify-auth')) {
      assertTrustedAuth(JSON.parse(require('fs').readFileSync(0, 'utf8')));
    } else {
      deploymentConfig();
    }
  } catch (error) { console.error(error.message); process.exitCode = 1; }
}
module.exports = { deploymentConfig, createOriginGuard, assertTrustedAuth };
