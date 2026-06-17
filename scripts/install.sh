#!/bin/bash
# Install Node.js dependencies for m365-mcp.

set -e

if [ -z "${CLAUDE_PLUGIN_ROOT}" ]; then
    echo "[m365-mcp] ERROR: CLAUDE_PLUGIN_ROOT is not set" >&2
    exit 1
fi

cd "${CLAUDE_PLUGIN_ROOT}"
echo "[m365-mcp] Installing dependencies..."
npm install --omit=dev
echo "[m365-mcp] Ready"
