#!/bin/bash
#
# Compatibility wrapper. Infrastructure provisioning is handled by `azd up`;
# updates to an existing Container App use the configurable deploy_aca.sh.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

if [[ "${1:-}" == "--provision" ]]; then
  exec azd up
fi

exec "$SCRIPT_DIR/deploy_aca.sh"
