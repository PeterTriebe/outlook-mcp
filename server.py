"""
Outlook MCP Server – Access Outlook Desktop calendar & mail via COM.
No cloud permissions or admin consent required.
"""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_client import OutlookClient
from tools import register_all_tools

# ── FastMCP instance ─────────────────────────────────────────────────

mcp = FastMCP(
    "Outlook Desktop",
    instructions=(
        "MCP server for local Outlook Desktop – calendar events, e-mail, "
        "meetings, contacts, tasks. No Azure AD / admin consent required."
    ),
)

# ── Lazy singleton client ────────────────────────────────────────────

_client: OutlookClient | None = None


def _get_client() -> OutlookClient:
    global _client
    if _client is None:
        _client = OutlookClient()
    return _client


# ── Register all tools ───────────────────────────────────────────────

register_all_tools(mcp, _get_client)

# ── Entry point ──────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run(transport="stdio")
