#!/bin/bash
# Entrypoint for m365-mcp MCP server.
# Ensures node_modules exist, then launches the Node.js server.

set -e

if [ -z "${CLAUDE_PLUGIN_ROOT}" ]; then
    echo "[m365-mcp] ERROR: CLAUDE_PLUGIN_ROOT is not set" >&2
    exit 1
fi

if [ ! -d "${CLAUDE_PLUGIN_ROOT}/node_modules" ]; then
    bash "${CLAUDE_PLUGIN_ROOT}/scripts/install.sh"
fi

# Load .env for tenant/client IDs
if [ -f "${CLAUDE_PLUGIN_ROOT}/.env" ]; then
    set -a
    # shellcheck disable=SC1091
    source "${CLAUDE_PLUGIN_ROOT}/.env"
    set +a
fi

export M365_MCP_TOKEN_CACHE="${HOME}/.m365-mcp-token.json"

exec node "${CLAUDE_PLUGIN_ROOT}/index.js" "$@"
