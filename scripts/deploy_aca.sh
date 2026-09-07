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
#   AZURE_OPENAI_ENDPOINT           - Approved Azure OpenAI account endpoint
#   AZURE_OPENAI_RESOURCE_ID        - Account identity for read-only model verification
#   AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA - Actual approved Astra deployment alias
#   ALLOW_BYO_AI_ENDPOINTS          - Optional admin opt-in (default false)
#   Runtime Azure OpenAI access uses the existing managed identity.
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

: "${AZURE_OPENAI_ENDPOINT:?Set AZURE_OPENAI_ENDPOINT}"
: "${AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA:?Set the approved Astra deployment alias}"
export AZURE_OPENAI_ALLOWED_DEPLOYMENTS="${AZURE_OPENAI_ALLOWED_DEPLOYMENTS:-$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA}"
export VITE_AZURE_OPENAI_ENDPOINT="$AZURE_OPENAI_ENDPOINT"
export VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA="$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA"
export ALLOW_BYO_AI_ENDPOINTS="${ALLOW_BYO_AI_ENDPOINTS:-false}"

# Every cloud update is public mode. Validate the same controls used at runtime,
# rather than building an image that can never safely become ready.
export APP_DEPLOYMENT_MODE=public
export EASY_AUTH_ENABLED=true
export ACCESS_CONTROL_ENABLED="${ACCESS_CONTROL_ENABLED:-true}"
export AI_BUDGET_STORE="${AI_BUDGET_STORE:-$([[ -n "${AZURE_TABLES_BUDGET_ENDPOINT:-${AZURE_TABLES_ENDPOINT:-}}" ]] && echo table || echo cosmos)}"
node "$(dirname "$0")/../server/deployment-security.js"
node "$(dirname "$0")/verify-astra-deployment.mjs"
az containerapp auth show --name "$ACA_APP_NAME" --resource-group "$RESOURCE_GROUP" --output json \
    | node "$(dirname "$0")/../server/deployment-security.js" --verify-auth
ORIGIN_FQDN="$(az containerapp show --name "$ACA_APP_NAME" --resource-group "$RESOURCE_GROUP" --query properties.configuration.ingress.fqdn -o tsv)"
ORIGIN_STATUS="$(curl --silent --output /dev/null --write-out '%{http_code}' --connect-timeout 10 --max-time 30 \
    --header "X-Azure-FDID: $FRONT_DOOR_ID" "https://${ORIGIN_FQDN}/healthz" || true)"
if [[ "$ORIGIN_STATUS" != "403" ]]; then
    echo "❌ The origin must reject a direct request even when the Front Door ID is spoofed." >&2
    exit 1
fi

# ─── Build arguments ────────────────────────────────────────────────
# Collect only supported public VITE_ variables as --build-arg flags into a bash array
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

BUILD_ARGS=(
    --build-arg "FRONT_DOOR_ID=$FRONT_DOOR_ID"
    --build-arg "VITE_AZURE_OPENAI_ENDPOINT=$AZURE_OPENAI_ENDPOINT"
    --build-arg "VITE_AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA"
)
while IFS='=' read -r key value; do
    if [[ "$key" == VITE_* && -n "$value" ]]; then
        case "$key" in
          VITE_APPINSIGHTS_CONNECTION_STRING|VITE_SPEECH_REGION|VITE_AZURE_AD_CLIENT_ID|VITE_AZURE_AD_AUTHORITY|VITE_ARM_SCOPE|VITE_FEEDBACK_CONTACT_ENABLED) ;;
          *) continue ;;
        esac
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

IMAGE_TAG="${IMAGE_TAG:-$(date -u +%Y%m%d%H%M%S)-$(git -C "$SOURCE_DIR" rev-parse --short HEAD 2>/dev/null || echo local)}"
ACR_IMAGE="$ACR_NAME.azurecr.io/$IMAGE_NAME:$IMAGE_TAG"

echo "🔨 Building image in ACR: $ACR_NAME"
echo "   Image: $IMAGE_NAME:$IMAGE_TAG"
echo "   Managed model: GPT-6 Astra only; BYO connections: $ALLOW_BYO_AI_ENDPOINTS"
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
    "APP_DEPLOYMENT_MODE=public"
    "EASY_AUTH_ENABLED=true"
    "ACCESS_CONTROL_ENABLED=$ACCESS_CONTROL_ENABLED"
    "FRONT_DOOR_ID=$FRONT_DOOR_ID"
    "AI_BUDGET_STORE=$AI_BUDGET_STORE"
    "AI_DAILY_TOKEN_BUDGET=${AI_DAILY_TOKEN_BUDGET:-250000}"
    "AI_MAX_CONCURRENT_REQUESTS=${AI_MAX_CONCURRENT_REQUESTS:-2}"
    "FEEDBACK_RETENTION_DAYS=${FEEDBACK_RETENTION_DAYS:-30}"
    "FEEDBACK_LEGACY_RETENTION_ENABLED=${FEEDBACK_LEGACY_RETENTION_ENABLED:-false}"
    "AZURE_IMPORT_ENABLED=false"
    "MCP_ENABLED=${MCP_ENABLED:-false}"
    "MCP_HTTP_STATELESS=${MCP_HTTP_STATELESS:-true}"
    "MCP_HTTP_MAX_IN_FLIGHT=${MCP_HTTP_MAX_IN_FLIGHT:-20}"
    "MCP_SESSION_MAX=${MCP_SESSION_MAX:-100}"
    "MCP_SESSION_IDLE_SECONDS=${MCP_SESSION_IDLE_SECONDS:-1800}"
    "MCP_SESSION_TTL_SECONDS=${MCP_SESSION_TTL_SECONDS:-7200}"
    "MCP_SESSION_GC_SECONDS=${MCP_SESSION_GC_SECONDS:-60}"
)
retired_ai_env="$(node "$SOURCE_DIR/scripts/retired-ai-environment.mjs")"
[[ -n "$retired_ai_env" ]]
mapfile -t REMOVE_ENV_VARS <<< "$retired_ai_env"
RUNTIME_ENV_VARS+=(
    "AZURE_OPENAI_ENDPOINT=$AZURE_OPENAI_ENDPOINT"
    "AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA=$AZURE_OPENAI_DEPLOYMENT_GPT6ASTRA"
    "AZURE_OPENAI_ALLOWED_DEPLOYMENTS=$AZURE_OPENAI_ALLOWED_DEPLOYMENTS"
    "ALLOW_BYO_AI_ENDPOINTS=$ALLOW_BYO_AI_ENDPOINTS"
)

for var in ACCESS_ADMIN_EMAIL AZURE_ACCESS_KEY_VAULT_RESOURCE_ID AZURE_TABLES_ACCESS_ENDPOINT \
           AZURE_TABLES_BUDGET_ENDPOINT AZURE_TABLES_BUDGET_TABLE COSMOS_BUDGET_CONTAINER_ID \
           AZURE_CLIENT_ID FEEDBACK_EMAIL_ENDPOINT FEEDBACK_EMAIL_SENDER \
           FEEDBACK_EMAIL_RECIPIENT FEEDBACK_CONTACT_ENABLED \
           AZURE_TABLES_ENDPOINT AZURE_TABLES_FEEDBACK_TABLE \
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

node "$SOURCE_DIR/scripts/verify-astra-deployment.mjs"
az containerapp update "${UPDATE_ARGS[@]}"

echo ""
echo "✅ Deployment complete!"
if [[ -n "$FQDN" ]]; then
    echo "   App URL:        https://$FQDN"
    echo "   MCP endpoint:   https://$FQDN/mcp           (streamable HTTP + SSE)"
    echo "   MCP health:     https://$FQDN/mcp/healthz"
fi
