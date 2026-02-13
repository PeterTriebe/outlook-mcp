"""
MCP tool definitions – register all tools on a FastMCP instance.
"""

from tools.calendar import register_calendar_tools
from tools.email import register_email_tools
from tools.contacts import register_contacts_tools
from tools.tasks import register_tasks_tools


def register_all_tools(mcp, get_client):
    """Register all Outlook tools on the given FastMCP server."""
    register_calendar_tools(mcp, get_client)
    register_email_tools(mcp, get_client)
    register_contacts_tools(mcp, get_client)
    register_tasks_tools(mcp, get_client)
