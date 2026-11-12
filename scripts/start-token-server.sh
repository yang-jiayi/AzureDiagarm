#!/usr/bin/env bash
# Dev helper: starts both the speech token server and the Vite dev server.
# Reads AZURE_SPEECH_REGION and AZURE_SPEECH_RESOURCE_ID from .env if present.
#
# Usage:
#   ./scripts/start-token-server.sh        # run both (token server + vite)
#   npm run dev:avatar                     # same, via npm
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$SCRIPT_DIR/.."

# Load .env if present
if [ -f "$ROOT/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  source "$ROOT/.env"
  set +a
fi

: "${AZURE_SPEECH_REGION:?AZURE_SPEECH_REGION must be set (in .env or environment)}"
: "${AZURE_SPEECH_RESOURCE_ID:?AZURE_SPEECH_RESOURCE_ID must be set (in .env or environment)}"

# Bridge public VITE_ routing values to the server-side names used by the
# /api/openai proxy. Credentials are never bridged from the browser config.
export AZURE_OPENAI_ENDPOINT="${AZURE_OPENAI_ENDPOINT:-${VITE_AZURE_OPENAI_ENDPOINT:-}}"
export AZURE_OPENAI_DEPLOYMENT_GPT61SOL="${AZURE_OPENAI_DEPLOYMENT_GPT61SOL:-${VITE_AZURE_OPENAI_DEPLOYMENT_GPT61SOL:-}}"
export AZURE_FOUNDRY_ENDPOINT="${AZURE_FOUNDRY_ENDPOINT:-${VITE_AZURE_FOUNDRY_ENDPOINT:-}}"
export AZURE_FOUNDRY_DEPLOYMENT_OPUS55="${AZURE_FOUNDRY_DEPLOYMENT_OPUS55:-${VITE_AZURE_FOUNDRY_DEPLOYMENT_OPUS55:-}}"
export ALLOW_BYO_AI_ENDPOINTS="${ALLOW_BYO_AI_ENDPOINTS:-false}"

echo "[token-server] Starting on 127.0.0.1:3001 (region=$AZURE_SPEECH_REGION)"
echo "[token-server] Managed AI: GPT-6.1 Sol + Claude Opus 5.5; BYO connections: $ALLOW_BYO_AI_ENDPOINTS"
exec node "$ROOT/server/token-server.js"
