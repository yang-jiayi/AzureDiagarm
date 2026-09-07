#!/bin/bash

# Fetch Azure pricing for multiple regions
# Publishes a refreshed snapshot only after every download and compaction succeeds.

set -Eeuo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

# Target regions - 9 regions for the Azure Architecture Diagram Builder
# HERO: primary showcase regions (East US 2, Australia East, Japan East)
# HUB:  regional coverage regions
DEFAULT_REGIONS=("eastus2" "swedencentral" "westeurope" "canadacentral" "brazilsouth" "australiaeast" "southeastasia" "mexicocentral" "japaneast")
if [ "$#" -gt 0 ]; then
  REGIONS=("$@")
else
  REGIONS=("${DEFAULT_REGIONS[@]}")
fi

# =============================================================================
# COMPREHENSIVE SERVICE LIST - 62+ Services
# =============================================================================
SERVICES=(
  # Compute
  "Azure App Service"
  "Virtual Machines"
  "Azure Kubernetes Service"
  "Azure Container Apps"
  "Container Instances"
  "Container Registry"
  "Functions"
  "Logic Apps"
  
  # Databases
  "Azure Cosmos DB"
  "SQL Database"
  "Azure Database for PostgreSQL"
  "Azure Database for MySQL"
  "Azure Cache for Redis"
  "Redis Cache"
  
  # Storage
  "Storage"
  "Azure Data Lake Storage"
  
  # Networking
  "Application Gateway"
  "Azure Front Door Service"
  "Azure Service Bus"
  "Event Hubs"
  "Azure Event Hubs"
  "Service Bus"
  
  # Analytics
  "Azure Data Factory"
  "Azure Synapse Analytics"
  "Stream Analytics"
  "Azure Machine Learning"
  "Microsoft Fabric"
  
  # AI & Cognitive Services
  "Cognitive Services"
  "Azure OpenAI Service"
  "Azure AI Document Intelligence"
  "Azure AI Language"
  "Azure AI Speech"
  "Azure AI Vision"
  "Azure AI Translator"
  "Azure Cognitive Search"
  "Azure API for FHIR"
  "Computer Vision"
  "Form Recognizer"
  "Speech Services"
  "Text Analytics"
  "Translator"
  "Foundry Models"
  "Foundry Tools"
  
  # Monitoring & Management
  "Application Insights"
  "Azure Monitor"
  "Log Analytics"
  "Key Vault"
  "Azure Key Vault"
  "API Management"
  
  # CDN & Edge
  "Content Delivery Network"
  "Azure CDN"
  "CDN"
  "Static Web Apps"
  
  # IoT
  "Azure IoT Hub"
  "Azure IoT Central"
  "IoT Hub"
  "IoT Central"
  "Digital Twins"
  
  # Security
  "Microsoft Defender for Cloud"
  "Microsoft Purview"
  "Azure Sentinel"
  
  # Integration
  "Azure SignalR Service"
  "SignalR"
  "Notification Hubs"
  "Event Grid"
  "Azure Event Grid"
  
  # Backup & Recovery
  "Backup"
  "Azure Backup"
  "Site Recovery"
  
  # Developer Tools
  "Azure DevOps"
  "Azure Automation"
  
  # Networking (Additional)
  "VPN Gateway"
  "Virtual Network"
  "Load Balancer"
  "Azure Load Balancer"
  "Traffic Manager"
  "Azure Traffic Manager"
  "ExpressRoute"
  "Network Watcher"
  "Azure Firewall"
)

# Global services (no region-specific pricing - copy to all regions)
GLOBAL_SERVICES=(
  "Azure Front Door Service"
  "Content Delivery Network"
  "CDN"
  "Static Web Apps"
  "Azure DevOps"
)

# Stage on the same filesystem, keeping unrequested regions and service files.
OUTPUT_DIR="$SOURCE_DIR/public/pricing/regions"
PRICING_TS_FILE="$SOURCE_DIR/src/data/azurePricing.ts"
STAGE_DIR="$SOURCE_DIR/.pricing-refresh"
for tool in curl jq node; do
  command -v "$tool" >/dev/null || { echo "Required tool not found: $tool" >&2; exit 1; }
done
for region in "${REGIONS[@]}"; do
  [[ "$region" =~ ^[a-z0-9]+$ ]] || { echo "Invalid region: $region" >&2; exit 1; }
done
[[ -d "$OUTPUT_DIR" && -f "$PRICING_TS_FILE" ]] || {
  echo "The existing pricing snapshot and freshness source are required." >&2
  exit 1
}
mkdir "$STAGE_DIR" || {
  echo "Pricing refresh workspace already exists; check for another refresh or recovery data: $STAGE_DIR" >&2
  exit 1
}
publication_started=false
publication_complete=false
cleanup() {
  local status=$?
  local restored=true
  set +e
  if [[ "$publication_started" == true && "$publication_complete" != true ]]; then
    if [[ -d "$STAGE_DIR/previous-regions" ]]; then
      rm -rf "$OUTPUT_DIR" && mv "$STAGE_DIR/previous-regions" "$OUTPUT_DIR" || restored=false
    fi
    cp "$STAGE_DIR/previous-azurePricing.ts" "$PRICING_TS_FILE" || restored=false
  fi
  if [[ "$restored" == true ]]; then
    if ! rm -rf "$STAGE_DIR"; then
      echo "Pricing refresh cleanup failed; inspect the remaining workspace: $STAGE_DIR" >&2
      status=1
    fi
  else
    echo "Pricing rollback needs manual recovery from $STAGE_DIR" >&2
    status=1
  fi
  exit "$status"
}
trap cleanup EXIT
trap 'exit 130' INT
trap 'exit 143' TERM
cp -a "$OUTPUT_DIR" "$STAGE_DIR/regions"
mkdir "$STAGE_DIR/global"

echo "🌍 Fetching Azure pricing for ${#REGIONS[@]} regions..."
echo "📦 Services: ${#SERVICES[@]}"
echo "📁 Output directory: $OUTPUT_DIR"
echo ""

fetch_response() {
  local filter="$1"
  local output_file="$2"
  local attempt
  for attempt in 1 2 3 4 5; do
    if curl --fail --silent --show-error --connect-timeout 15 --max-time 90 \
      -G "https://prices.azure.com/api/retail/prices" \
      --data-urlencode "api-version=2023-01-01-preview" \
      --data-urlencode "\$filter=$filter" \
      --data-urlencode "\$top=1000" \
      -o "$output_file" \
      && jq -e '
        type == "object" and
        (.BillingCurrency | type) == "string" and
        (.Items | type) == "array" and
        all(.Items[];
          type == "object" and
          (.serviceName | type) == "string" and
          (.unitOfMeasure | type) == "string" and
          ((.retailPrice // .unitPrice) | type) == "number"
        )
      ' "$output_file" >/dev/null 2>&1; then
      echo "    ✓ Downloaded $(jq '.Items | length' "$output_file") items"
      return 0
    fi
    echo "    ⚠ attempt $attempt: invalid/rate-limited response, retrying..."
    sleep 5
  done
  echo "    ✗ Failed to download valid pricing; the existing snapshot will be preserved." >&2
  return 1
}

fetch_pricing() {
  local service="$1"
  local region="$2"
  echo "  Fetching: $service in $region..."
  fetch_response "serviceName eq '$service' and armRegionName eq '$region' and priceType eq 'Consumption'" "$3"
}

fetch_global_pricing() {
  local service="$1"
  echo "  Fetching: $service (global)..."
  fetch_response "serviceName eq '$service' and priceType eq 'Consumption'" "$2"
}

# Iterate through regions and services
for region in "${REGIONS[@]}"; do
  echo ""
  echo "=== Region: $region ==="
  
  region_dir="$STAGE_DIR/regions/$region"
  mkdir -p "$region_dir"
  
  for service in "${SERVICES[@]}"; do
    # Create safe filename
    safe_name=$(echo "$service" | tr ' ' '_' | tr '[:upper:]' '[:lower:]')
    output_file="$region_dir/${safe_name}.json"
    
    fetch_pricing "$service" "$region" "$output_file"
    
    # Small delay to avoid rate limiting
    sleep 0.5
  done
done

# Fetch global services (only once, not per-region)
if [ ${#GLOBAL_SERVICES[@]} -gt 0 ]; then
  echo ""
  echo "=== Global Services (copying to all regions) ==="
  
  for service in "${GLOBAL_SERVICES[@]}"; do
    # Create safe filename
    safe_name=$(echo "$service" | tr ' ' '_' | tr '[:upper:]' '[:lower:]')
    temp_file="$STAGE_DIR/global/${safe_name}.json"
    
    fetch_global_pricing "$service" "$temp_file"
    
    # Copy to all regions
    for region in "${REGIONS[@]}"; do
      region_dir="$STAGE_DIR/regions/$region"
      cp "$temp_file" "$region_dir/${safe_name}.json"
    done
    
    # Small delay to avoid rate limiting
    sleep 0.5
  done
  
fi

echo ""
echo "📊 Summary by region:"
for region in "${REGIONS[@]}"; do
  region_dir="$STAGE_DIR/regions/$region"
  total_items=0
  
  for file in "$region_dir"/*.json; do
    if [ -f "$file" ]; then
      count=$(jq '.Items | length' "$file")
      total_items=$((total_items + count))
    fi
  done
  
  echo "  $region: $total_items total pricing items"
done

# Compact the freshly-downloaded dumps in place so the browser downloads far
# less pricing JSON. regionalPricingService.ts expands them at runtime, so the
# parsed pricing tiers are identical (see tests/pricing-prep.test.ts).
echo ""
echo "🗜️  Compacting pricing data..."
node "$SCRIPT_DIR/prep-pricing-data.mjs" --dir "$STAGE_DIR/regions"

TODAY=$(date +%Y-%m-%d)
cp "$PRICING_TS_FILE" "$STAGE_DIR/previous-azurePricing.ts"
node --input-type=module - "$STAGE_DIR/previous-azurePricing.ts" "$STAGE_DIR/azurePricing.ts" "$TODAY" <<'NODE'
import { readFileSync, writeFileSync } from 'node:fs';
const [source, destination, today] = process.argv.slice(2);
const text = readFileSync(source, 'utf8');
const pattern = /export const PRICING_DATA_AS_OF = '[0-9-]*';/g;
if ([...text.matchAll(pattern)].length !== 1) {
  throw new Error('Expected exactly one pricing freshness declaration; snapshot unchanged.');
}
writeFileSync(destination, text.replace(pattern, `export const PRICING_DATA_AS_OF = '${today}';`));
NODE

cmp -s "$PRICING_TS_FILE" "$STAGE_DIR/previous-azurePricing.ts" || {
  echo "Pricing source changed during publication preparation; snapshot unchanged." >&2
  exit 1
}

# Keep both originals until the complete snapshot and its date are published.
publication_started=true
mv "$OUTPUT_DIR" "$STAGE_DIR/previous-regions"
mv "$STAGE_DIR/regions" "$OUTPUT_DIR"
mv "$STAGE_DIR/azurePricing.ts" "$PRICING_TS_FILE"
publication_complete=true
echo ""
echo "✅ Pricing snapshot published in: $OUTPUT_DIR"
echo "🗓️  Stamped PRICING_DATA_AS_OF = $TODAY in azurePricing.ts"
