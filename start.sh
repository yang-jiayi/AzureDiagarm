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
if [ "${MCP_ENABLED:-false}" = "true" ] || [ -n "${MCP_AUTH_TOKEN:-}" ]; then
  MCP_HTTP_HOST=127.0.0.1 \
  MCP_HTTP_PORT="${MCP_HTTP_PORT:-3030}" \
  MCP_HTTP_PATH="${MCP_HTTP_PATH:-/mcp}" \
  MCP_AUTH_TOKEN="${MCP_AUTH_TOKEN:-}" \
    node /srv/mcp-server/dist/index.js --http &
  MCP_SERVER_PID=$!
  if [ -z "${MCP_AUTH_TOKEN:-}" ]; then
    echo "MCP is enabled without an internal token; the reverse proxy must enforce authentication." >&2
  fi
else
  echo "MCP is disabled. Set MCP_ENABLED=true behind an authenticated reverse proxy or configure MCP_AUTH_TOKEN." >&2
fi

nginx -g "daemon off;" &
NGINX_PID=$!

cleanup_started=0
cleanup() {
  if [ "$cleanup_started" -eq 1 ]; then return; fi
  cleanup_started=1
  trap - INT TERM

  # Stop accepting new public requests first, then let each backend drain
  # in-flight work within the Container Apps termination window.
  kill -QUIT "$NGINX_PID" 2>/dev/null || true
  kill -TERM "$TOKEN_SERVER_PID" 2>/dev/null || true
  if [ -n "$MCP_SERVER_PID" ]; then
    kill -TERM "$MCP_SERVER_PID" 2>/dev/null || true
  fi

  remaining=25
  while [ "$remaining" -gt 0 ]; do
    alive=0
    kill -0 "$TOKEN_SERVER_PID" 2>/dev/null && alive=1
    kill -0 "$NGINX_PID" 2>/dev/null && alive=1
    if [ -n "$MCP_SERVER_PID" ]; then
      kill -0 "$MCP_SERVER_PID" 2>/dev/null && alive=1
    fi
    if [ "$alive" -eq 0 ]; then break; fi
    sleep 1
    remaining=$((remaining - 1))
  done

  for pid in "$TOKEN_SERVER_PID" "$NGINX_PID" ${MCP_SERVER_PID:+"$MCP_SERVER_PID"}; do
    if kill -0 "$pid" 2>/dev/null; then
      echo "Process $pid exceeded the graceful shutdown window; forcing exit." >&2
      kill -KILL "$pid" 2>/dev/null || true
    fi
  done
  wait "$TOKEN_SERVER_PID" 2>/dev/null || true
  wait "$NGINX_PID" 2>/dev/null || true
  if [ -n "$MCP_SERVER_PID" ]; then
    wait "$MCP_SERVER_PID" 2>/dev/null || true
  fi
}

handle_shutdown_signal() {
  cleanup
  exit 0
}
trap handle_shutdown_signal INT TERM

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
