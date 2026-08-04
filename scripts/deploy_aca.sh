#!/bin/bash
#
# Azure Container App Deployment Script (Configurable)
# =====================================================
#
# Generic deployment script that reads ALL configuration from your .env file.
# No hardcoded resource names — clone the repo, fill in .env, and deploy.
#
# Prerequisites:
# 1. Azure CLI installed and authenticated (az login)
# 2. .env file in project root with required variables (see .env.example)
# 3. An Azure Container Registry (ACR) already created
# 4. An Azure Container Apps environment and app already created
# 5. ACR admin credentials configured on the Container App
#
# Required .env variables:
#   ACR_NAME          - Azure Container Registry name (e.g. myacr123)
#   ACA_APP_NAME      - Container App name (e.g. azure-diagram-builder)
#   RESOURCE_GROUP    - Resource group containing the ACA app
#   IMAGE_NAME        - Docker image name (e.g. azure-diagram-builder)
#
#   VITE_AZURE_OPENAI_ENDPOINT       - Azure OpenAI endpoint URL (when using GPT/partner models)
#   VITE_AZURE_FOUNDRY_ENDPOINT      - Microsoft Foundry endpoint URL (when using Claude)
#   VITE_*_DEPLOYMENT_*              - Model deployment names (at least one)
#   AZURE_OPENAI_API_KEY             - Optional server-side fallback key
#
# Usage:
#   chmod +x scripts/deploy_aca.sh
#   ./scripts/deploy_aca.sh
#

set -euo pipefail

# ─── Load .env ───────────────────────────────────────────────────────
ENV_FILE="$(dirname "$0")/../.env"
if [[ ! -f "$ENV_FILE" ]]; then
    echo "❌ .env file not found at $ENV_FILE"
    echo "   Copy .env.example to .env and fill in your values."
    exit 1
fi

set -a
source "$ENV_FILE"
set +a

# ─── Validate required variables ────────────────────────────────────
MISSING=()
for var in ACR_NAME ACA_APP_NAME RESOURCE_GROUP IMAGE_NAME; do
    if [[ -z "${!var:-}" ]]; then
        MISSING+=("$var")
    fi
done

if [[ ${#MISSING[@]} -gt 0 ]]; then
    echo "❌ Missing required .env variables:"
    for v in "${MISSING[@]}"; do
        echo "   - $v"
    done
    exit 1
fi

# Check that at least one model deployment is configured
MODEL_COUNT=0
OPENAI_DEPLOYMENTS=()
for var in VITE_AZURE_OPENAI_DEPLOYMENT_GPT51 VITE_AZURE_OPENAI_DEPLOYMENT_GPT52 \
           VITE_AZURE_OPENAI_DEPLOYMENT_GPT52CODEX VITE_AZURE_OPENAI_DEPLOYMENT_GPT53CODEX \
           VITE_AZURE_OPENAI_DEPLOYMENT_GPT54 VITE_AZURE_OPENAI_DEPLOYMENT_GPT54MINI \
           VITE_AZURE_OPENAI_DEPLOYMENT_GPT56SOL VITE_AZURE_OPENAI_DEPLOYMENT_GPT56TERRA \
           VITE_AZURE_OPENAI_DEPLOYMENT_GPT56LUNA VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK \
           VITE_AZURE_OPENAI_DEPLOYMENT_DEEPSEEK_V4_PRO VITE_AZURE_OPENAI_DEPLOYMENT_GROK4FAST \
           VITE_AZURE_OPENAI_DEPLOYMENT_GROK43 VITE_AZURE_OPENAI_DEPLOYMENT_MISTRALLARGE3 \
           VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK25 VITE_AZURE_OPENAI_DEPLOYMENT_KIMIK27CODE; do
    if [[ -n "${!var:-}" ]]; then
        MODEL_COUNT=$((MODEL_COUNT + 1))
        OPENAI_DEPLOYMENTS+=("${!var}")
    fi
done
FOUNDRY_DEPLOYMENTS=()
if [[ -n "${VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5:-}" ]]; then
    MODEL_COUNT=$((MODEL_COUNT + 1))
    FOUNDRY_DEPLOYMENTS+=("$VITE_AZURE_FOUNDRY_DEPLOYMENT_CLAUDE_OPUS5")
fi

if [[ $MODEL_COUNT -eq 0 ]]; then
    echo "❌ No model deployments configured. Set at least one VITE_*_DEPLOYMENT_* in .env"
    exit 1
fi
if [[ ${#OPENAI_DEPLOYMENTS[@]} -gt 0 && -z "${VITE_AZURE_OPENAI_ENDPOINT:-}" ]]; then
    echo "❌ VITE_AZURE_OPENAI_ENDPOINT is required for configured Azure OpenAI deployments"
    exit 1
fi
if [[ ${#FOUNDRY_DEPLOYMENTS[@]} -gt 0 && -z "${VITE_AZURE_FOUNDRY_ENDPOINT:-}" ]]; then
    echo "❌ VITE_AZURE_FOUNDRY_ENDPOINT is required for configured Foundry deployments"
    exit 1
fi

# ─── Build arguments ────────────────────────────────────────────────
# Collect all VITE_ variables as --build-arg flags into a bash array
# (array avoids eval pitfalls when values contain quotes, $, spaces, etc.)
#
# IMPORTANT — App Insights connection string workaround:
#   VITE_APPINSIGHTS_CONNECTION_STRING contains semicolons (;) which break
#   `az acr build --build-arg`. ACR Tasks forwards build args to a remote
#   Docker agent via shell commands, and semicolons are interpreted as
#   command separators ("docker build requires exactly 1 argument" error).
#
#   Workaround: extract that one value into .env.appinsights (gitignored,
#   NOT in .dockerignore). The Dockerfile COPYs it and `source`s it in the
#   same RUN layer as `npm run build` so Vite embeds it via import.meta.env.
SOURCE_DIR="$(cd "$(dirname "$0")/.." && pwd)"
APPINSIGHTS_FILE="$SOURCE_DIR/.env.appinsights"
: > "$APPINSIGHTS_FILE"
trap 'rm -f "$APPINSIGHTS_FILE"' EXIT

BUILD_ARGS=()
while IFS='=' read -r key value; do
    if [[ "$key" == VITE_* && -n "$value" ]]; then
        # Strip surrounding quotes if present in .env
        value="${value%\"}"
        value="${value#\"}"
        value="${value%\'}"
        value="${value#\'}"
        # Route App Insights conn string through file workaround
        if [[ "$key" == "VITE_APPINSIGHTS_CONNECTION_STRING" ]]; then
            echo "$key=$value" > "$APPINSIGHTS_FILE"
            continue
        fi
        if [[ "$key" =~ (_API_KEY|_SECRET|_TOKEN|_PASSWORD|_CONNECTION_STRING)$ ]]; then
            echo "⚠️  Skipping sensitive build variable: $key" >&2
            continue
        fi
        BUILD_ARGS+=(--build-arg "$key=$value")
    fi
done < <(grep -v '^#' "$ENV_FILE" | grep -v '^\s*$')

if [[ -n "${FRONT_DOOR_ID:-}" ]]; then
    BUILD_ARGS+=(--build-arg "FRONT_DOOR_ID=$FRONT_DOOR_ID")
fi

IMAGE_TAG="${IMAGE_TAG:-$(date -u +%Y%m%d%H%M%S)-$(git -C "$SOURCE_DIR" rev-parse --short HEAD 2>/dev/null || echo local)}"
ACR_IMAGE="$ACR_NAME.azurecr.io/$IMAGE_NAME:$IMAGE_TAG"

echo "🔨 Building image in ACR: $ACR_NAME"
echo "   Image: $IMAGE_NAME:$IMAGE_TAG"
echo "   Models configured: $MODEL_COUNT"
echo "   Source: $SOURCE_DIR"
echo "   Build args: ${#BUILD_ARGS[@]} VITE_* values via --build-arg"
if [[ -s "$APPINSIGHTS_FILE" ]]; then
    echo "   App Insights: routed via .env.appinsights (semicolon workaround)"
fi
echo ""

# ─── Build in ACR ────────────────────────────────────────────────────
az acr build \
    --registry "$ACR_NAME" \
    --image "$IMAGE_NAME:$IMAGE_TAG" \
    "${BUILD_ARGS[@]}" \
    "$SOURCE_DIR"

# ─── Get ACA FQDN ───────────────────────────────────────────────────
FQDN=$(az containerapp show \
    -g "$RESOURCE_GROUP" \
    -n "$ACA_APP_NAME" \
    --query 'properties.configuration.ingress.fqdn' -o tsv 2>/dev/null || echo "")

# ─── Update Container App ───────────────────────────────────────────
echo ""
echo "🚀 Updating Container App: $ACA_APP_NAME"

bash "$SOURCE_DIR/scripts/ensure-containerapp-probes.sh" \
    "$RESOURCE_GROUP" \
    "$ACA_APP_NAME" \
    "${FRONT_DOOR_ID:-}"

PUBLIC_URL="${PUBLIC_URL:-https://$FQDN}"
RUNTIME_ENV_VARS=(
    "PUBLIC_URL=$PUBLIC_URL"
    "MCP_ENABLED=${MCP_ENABLED:-false}"
    "MCP_HTTP_STATELESS=${MCP_HTTP_STATELESS:-true}"
    "MCP_HTTP_MAX_IN_FLIGHT=${MCP_HTTP_MAX_IN_FLIGHT:-20}"
    "MCP_SESSION_MAX=${MCP_SESSION_MAX:-100}"
    "MCP_SESSION_IDLE_SECONDS=${MCP_SESSION_IDLE_SECONDS:-1800}"
    "MCP_SESSION_TTL_SECONDS=${MCP_SESSION_TTL_SECONDS:-7200}"
    "MCP_SESSION_GC_SECONDS=${MCP_SESSION_GC_SECONDS:-60}"
)
REMOVE_ENV_VARS=()

if [[ ${#OPENAI_DEPLOYMENTS[@]} -gt 0 ]]; then
    RUNTIME_ENV_VARS+=("AZURE_OPENAI_ENDPOINT=${AZURE_OPENAI_ENDPOINT:-$VITE_AZURE_OPENAI_ENDPOINT}")
    RUNTIME_ENV_VARS+=("AZURE_OPENAI_ALLOWED_DEPLOYMENTS=$(IFS=,; echo "${OPENAI_DEPLOYMENTS[*]}")")
else
    REMOVE_ENV_VARS+=("AZURE_OPENAI_ENDPOINT" "AZURE_OPENAI_ALLOWED_DEPLOYMENTS")
fi
if [[ ${#FOUNDRY_DEPLOYMENTS[@]} -gt 0 ]]; then
    RUNTIME_ENV_VARS+=("AZURE_FOUNDRY_ENDPOINT=${AZURE_FOUNDRY_ENDPOINT:-$VITE_AZURE_FOUNDRY_ENDPOINT}")
    RUNTIME_ENV_VARS+=("AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS=$(IFS=,; echo "${FOUNDRY_DEPLOYMENTS[*]}")")
else
    REMOVE_ENV_VARS+=("AZURE_FOUNDRY_ENDPOINT" "AZURE_FOUNDRY_ALLOWED_DEPLOYMENTS")
fi

for var in AZURE_CLIENT_ID FEEDBACK_EMAIL_ENDPOINT FEEDBACK_EMAIL_SENDER \
           FEEDBACK_EMAIL_RECIPIENT AZURE_TABLES_ENDPOINT AZURE_TABLES_FEEDBACK_TABLE \
           AZURE_TABLES_RATE_LIMIT_TABLE \
           AZURE_COSMOS_ENDPOINT COSMOS_DATABASE_ID COSMOS_CONTAINER_ID \
           COSMOS_FEEDBACK_CONTAINER_ID AZURE_BLOB_ENDPOINT \
           AZURE_BLOB_DIAGRAMS_CONTAINER AZURE_SPEECH_REGION AZURE_SPEECH_RESOURCE_ID; do
    if [[ -n "${!var:-}" ]]; then
        RUNTIME_ENV_VARS+=("$var=${!var}")
    else
        REMOVE_ENV_VARS+=("$var")
    fi
done

if [[ -n "${AZURE_TABLES_ENDPOINT:-}" && -z "${AZURE_TABLES_RATE_LIMIT_TABLE:-}" ]]; then
    RUNTIME_ENV_VARS+=("AZURE_TABLES_RATE_LIMIT_TABLE=ratelimit")
    FILTERED_REMOVE_ENV_VARS=()
    for var in "${REMOVE_ENV_VARS[@]}"; do
        [[ "$var" == "AZURE_TABLES_RATE_LIMIT_TABLE" ]] || FILTERED_REMOVE_ENV_VARS+=("$var")
    done
    REMOVE_ENV_VARS=("${FILTERED_REMOVE_ENV_VARS[@]}")
fi

OPENAI_KEY="${AZURE_OPENAI_API_KEY:-${VITE_AZURE_OPENAI_API_KEY:-}}"
if [[ -n "$OPENAI_KEY" ]]; then
    az containerapp secret set \
        --name "$ACA_APP_NAME" \
        --resource-group "$RESOURCE_GROUP" \
        --secrets "azure-openai-api-key=$OPENAI_KEY" \
        --output none
    RUNTIME_ENV_VARS+=("AZURE_OPENAI_API_KEY=secretref:azure-openai-api-key")
else
    REMOVE_ENV_VARS+=("AZURE_OPENAI_API_KEY")
fi

FOUNDRY_KEY="${AZURE_FOUNDRY_API_KEY:-${VITE_AZURE_FOUNDRY_API_KEY:-}}"
if [[ -n "$FOUNDRY_KEY" ]]; then
    az containerapp secret set \
        --name "$ACA_APP_NAME" \
        --resource-group "$RESOURCE_GROUP" \
        --secrets "azure-foundry-api-key=$FOUNDRY_KEY" \
        --output none
    RUNTIME_ENV_VARS+=("AZURE_FOUNDRY_API_KEY=secretref:azure-foundry-api-key")
else
    REMOVE_ENV_VARS+=("AZURE_FOUNDRY_API_KEY")
fi

UPDATE_ARGS=(
    --name "$ACA_APP_NAME" \
    --resource-group "$RESOURCE_GROUP" \
    --image "$ACR_IMAGE" \
    --set-env-vars "${RUNTIME_ENV_VARS[@]}" \
    --revision-suffix "v$(date -u +%s)" \
    --min-replicas "${MIN_REPLICAS:-1}" \
    --max-replicas "${MAX_REPLICAS:-2}" \
    --scale-rule-name http \
    --scale-rule-http-concurrency "${HTTP_SCALE_CONCURRENCY:-50}"
)
if [[ ${#REMOVE_ENV_VARS[@]} -gt 0 ]]; then
    UPDATE_ARGS+=(--remove-env-vars "${REMOVE_ENV_VARS[@]}")
fi

az containerapp update "${UPDATE_ARGS[@]}"

echo ""
echo "✅ Deployment complete!"
if [[ -n "$FQDN" ]]; then
    echo "   App URL:        https://$FQDN"
    echo "   MCP endpoint:   https://$FQDN/mcp           (streamable HTTP + SSE)"
    echo "   MCP health:     https://$FQDN/mcp/healthz"
fi
