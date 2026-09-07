// Names only: remove stale model configuration without reading or logging values.
const retiredModels = [
  'GPT51', 'GPT52', 'GPT52CODEX', 'GPT53CODEX', 'GPT54', 'GPT54MINI',
  'GPT56SOL', 'GPT56TERRA', 'GPT56LUNA', 'DEEPSEEK', 'DEEPSEEK_V4_PRO',
  'GROK4FAST', 'GROK43', 'MISTRALLARGE3', 'KIMIK25', 'KIMIK27CODE',
];
const names = [
  'AZURE_OPENAI_API_KEY', 'AZURE_OPENAI_API_VERSION',
  'AZURE_FOUNDRY_ENDPOINT', 'AZURE_FOUNDRY_API_KEY',
  'AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS', 'AZURE_FOUNDRY_RESOURCE_ID',
  'AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5',
  'VITE_AZURE_FOUNDRY_ENDPOINT', 'VITE_AZURE_FOUNDRY_API_KEY',
  'VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5',
  'AZURE_OPENAI_DEPLOYMENT', 'AZURE_OPENAI_DEPLOYMENT_NAME', 'VITE_AZURE_OPENAI_DEPLOYMENT',
  ...retiredModels.flatMap(model => [
    `AZURE_OPENAI_DEPLOYMENT_${model}`, `VITE_AZURE_OPENAI_DEPLOYMENT_${model}`,
  ]),
];
const args = process.argv.slice(2);
if (args.length > 1 || (args.length === 1 && args[0] !== '--rollback')) {
  console.error('Usage: node scripts/retired-ai-environment.mjs [--rollback]');
  process.exitCode = 1;
} else {
  // BYO is an independently approved runtime policy, never a retired setting.
  console.log(names.join('\n'));
}
