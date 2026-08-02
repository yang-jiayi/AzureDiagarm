#!/bin/sh

set -eu

cat >&2 <<'EOF'
The generic azd deployment path is retired in this secured AzureDiagarm fork.

It does not configure the required Azure Front Door route, Easy Auth enterprise
application assignment, Conditional Access boundaries, or direct-origin
isolation. Use the protected "AzureDiagarm sync and deploy" GitHub Actions
workflow for production updates.
EOF

exit 1
