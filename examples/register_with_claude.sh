#!/usr/bin/env bash
# Register the OnlyOffice MCP server with Claude Code.
#
# Run from anywhere after `pip install -e /home/kali/onlyoffice-mcp` (or pipx
# install). The `--scope user` flag makes the server available across all your
# Claude Code projects; drop it to register only for the current working dir.

set -euo pipefail

PROJECT_DIR="${PROJECT_DIR:-$(cd "$(dirname "$0")/.." && pwd)}"

echo "[+] Registering onlyoffice MCP with Claude Code..."
claude mcp add --scope user onlyoffice -- python -m onlyoffice_mcp

echo ""
echo "[+] Verifying registration:"
claude mcp list | grep -E "(onlyoffice|name)" || true

echo ""
echo "[+] Done. In a Claude Code session, try:"
echo "    'Create a Word document at /tmp/test.docx with a Heading 1 saying Hello and a paragraph.'"
echo ""
echo "    'Use docbuilder_status to check if ONLYOFFICE Document Builder is installed.'"
