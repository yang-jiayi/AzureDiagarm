#!/usr/bin/env bash
set -euo pipefail
echo "The legacy bundled web/MCP provisioning experiment is retired." >&2
echo "Use the standalone mcp-server/Dockerfile for deterministic MCP tools, or the protected release workflow for the secured GPT-6 Astra application." >&2
exit 1
