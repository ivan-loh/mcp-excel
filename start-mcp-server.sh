#!/bin/bash
# Script to start MCP Excel server for Claude Code

cd /Users/ivan/Projects/mcp-excel
source venv/bin/activate
exec python -m mcp_excel.server --path examples/