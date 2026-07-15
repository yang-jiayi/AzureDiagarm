#!/bin/bash
#
# Backward-compatible entry point for updating an existing Container App.
# All resource names, runtime settings, and the unique image tag are handled by
# deploy_aca.sh from values in the project .env file.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
exec "$SCRIPT_DIR/deploy_aca.sh"
