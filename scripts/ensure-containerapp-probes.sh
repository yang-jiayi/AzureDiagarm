#!/usr/bin/env bash

set -Eeuo pipefail

if [[ $# -lt 2 || $# -gt 3 ]]; then
  echo "Usage: $0 <resource-group> <container-app> [front-door-id]" >&2
  exit 2
fi

resource_group="$1"
container_app="$2"
front_door_id="${3:-${FRONT_DOOR_ID:-}}"
current_file="$(mktemp)"
patch_file="$(mktemp)"
trap 'rm -f "$current_file" "$patch_file"' EXIT

read -r -d '' probes_json <<'JSON' || true
[
  {
    "type": "Startup",
    "httpGet": { "path": "/healthz", "port": 80, "scheme": "HTTP" },
    "initialDelaySeconds": 1,
    "periodSeconds": 5,
    "timeoutSeconds": 3,
    "failureThreshold": 30
  },
  {
    "type": "Liveness",
    "httpGet": { "path": "/healthz", "port": 80, "scheme": "HTTP" },
    "initialDelaySeconds": 15,
    "periodSeconds": 15,
    "timeoutSeconds": 5,
    "failureThreshold": 3
  },
  {
    "type": "Readiness",
    "httpGet": { "path": "/healthz", "port": 80, "scheme": "HTTP" },
    "initialDelaySeconds": 3,
    "periodSeconds": 5,
    "timeoutSeconds": 3,
    "failureThreshold": 3,
    "successThreshold": 1
  }
]
JSON

if [[ -n "$front_door_id" ]]; then
  probes_json="$(
    jq -c --arg frontDoorId "$front_door_id" '
      map(
        .httpGet.httpHeaders = [
          {
            name: "X-Azure-FDID",
            value: $frontDoorId
          }
        ]
      )
    ' <<< "$probes_json"
  )"
fi

has_required_probes() {
  jq -e --arg frontDoorId "$front_door_id" '
    (.properties.template.containers[0].probes // []) as $probes |
    ($probes | length) == 3 and
    any($probes[];
      .type == "Startup" and
      .httpGet.path == "/healthz" and
      .httpGet.port == 80 and
      (
        $frontDoorId == "" or
        any(.httpGet.httpHeaders[]?; .name == "X-Azure-FDID" and .value == $frontDoorId)
      )
    ) and
    any($probes[];
      .type == "Liveness" and
      .httpGet.path == "/healthz" and
      .httpGet.port == 80 and
      (
        $frontDoorId == "" or
        any(.httpGet.httpHeaders[]?; .name == "X-Azure-FDID" and .value == $frontDoorId)
      )
    ) and
    any($probes[];
      .type == "Readiness" and
      .httpGet.path == "/healthz" and
      .httpGet.port == 80 and
      (
        $frontDoorId == "" or
        any(.httpGet.httpHeaders[]?; .name == "X-Azure-FDID" and .value == $frontDoorId)
      )
    )
  ' "$1" >/dev/null
}

az containerapp show \
  --resource-group "$resource_group" \
  --name "$container_app" \
  --output json > "$current_file"

if has_required_probes "$current_file"; then
  echo "Container App health probes are already configured."
  exit 0
fi

jq --argjson probes "$probes_json" '
  {
    location,
    properties: {
      template: (
        .properties.template
        | del(.revisionSuffix, .customMetricsSettings, .serviceBinds)
        | .containers |= map(
            del(.imageType)
            | if .resources then .resources |= del(.ephemeralStorage) else . end
          )
        | if .scale then .scale |= del(.cooldownPeriod, .pollingInterval) else . end
        | .containers[0].probes = $probes
        | walk(
            if type == "object"
            then with_entries(select(.value != null))
            else .
            end
          )
      )
    }
  }
' "$current_file" > "$patch_file"

az containerapp update \
  --resource-group "$resource_group" \
  --name "$container_app" \
  --yaml "$patch_file" \
  --output none

for attempt in $(seq 1 60); do
  az containerapp show \
    --resource-group "$resource_group" \
    --name "$container_app" \
    --output json > "$current_file"

  latest_revision="$(jq -r '.properties.latestRevisionName // ""' "$current_file")"
  ready_revision="$(jq -r '.properties.latestReadyRevisionName // ""' "$current_file")"
  if has_required_probes "$current_file" \
    && [[ -n "$latest_revision" && "$latest_revision" == "$ready_revision" ]]; then
    echo "Container App health probes are configured and ready."
    exit 0
  fi

  sleep 5
done

echo "Container App health probes did not become ready within five minutes." >&2
exit 1
