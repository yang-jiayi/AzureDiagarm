#!/bin/bash
#
# Compatibility wrapper for updates to an existing secured Container App.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [[ "${1:-}" == "--provision" ]]; then
  echo "Provisioning through the legacy azd path is retired." >&2
  echo "Use the protected AzureDiagarm sync/deploy workflow for production." >&2
  exit 1
fi

exec "$SCRIPT_DIR/deploy_aca.sh"
