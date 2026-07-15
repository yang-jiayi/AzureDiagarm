#!/bin/sh
set -eu
# Start background services, then run nginx in the foreground.
#
# 1. Speech token server (port 3001)
#    If AZURE_SPEECH_REGION is not set the token server logs a warning and
#    /api/speech-token returns 503 — the avatar "Present" button is hidden.
#
# 2. MCP HTTP server (port 3030, internal — exposed via nginx at /mcp)
#    Streamable HTTP transport for MCP clients (M365 Copilot, hosted agents,
#    Azure SRE Agent, VS Code with remote MCP). Health probe: GET /healthz.
#    Set MCP_AUTH_TOKEN on the Container App to require `Authorization: Bearer
#    <token>` on /mcp (recommended for any public ingress). If unset, /mcp is open.
node /srv/token-server/token-server.js &
TOKEN_SERVER_PID=$!

MCP_SERVER_PID=''
if [ -n "${MCP_AUTH_TOKEN:-}" ]; then
  MCP_HTTP_HOST=127.0.0.1 \
  MCP_HTTP_PORT="${MCP_HTTP_PORT:-3030}" \
  MCP_HTTP_PATH="${MCP_HTTP_PATH:-/mcp}" \
  MCP_AUTH_TOKEN="$MCP_AUTH_TOKEN" \
    node /srv/mcp-server/dist/index.js --http &
  MCP_SERVER_PID=$!
else
  echo "MCP_AUTH_TOKEN is not configured; the public /mcp endpoint is disabled." >&2
fi

nginx -g "daemon off;" &
NGINX_PID=$!

cleanup() {
  trap - INT TERM
  PIDS="$TOKEN_SERVER_PID $NGINX_PID"
  if [ -n "$MCP_SERVER_PID" ]; then PIDS="$PIDS $MCP_SERVER_PID"; fi
  kill $PIDS 2>/dev/null || true
  wait $PIDS 2>/dev/null || true
}
trap cleanup INT TERM

is_mcp_alive() {
  [ -z "$MCP_SERVER_PID" ] || kill -0 "$MCP_SERVER_PID" 2>/dev/null
}

while kill -0 "$TOKEN_SERVER_PID" 2>/dev/null \
  && is_mcp_alive \
  && kill -0 "$NGINX_PID" 2>/dev/null; do
  sleep 5
done

echo "A required process exited; stopping the container so the platform can restart it." >&2
cleanup
exit 1
