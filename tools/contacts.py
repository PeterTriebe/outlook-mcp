"""
MCP tools – Contacts.
"""

from __future__ import annotations

import json


def register_contacts_tools(mcp, get_client):
    """Register all contact tools."""

    @mcp.tool()
    def search_contacts(query: str, max_results: int = 20) -> str:
        """Search Outlook contacts by name, email, or company.

        Args:
            query: Text to search for in contact name, email, or company.
            max_results: Maximum number of results (default 20).
        """
        contacts = get_client().search_contacts(query, max_results=max_results)
        if not contacts:
            return f"No contacts found matching '{query}'."
        return json.dumps(contacts, indent=2, ensure_ascii=False)

    @mcp.tool()
    def get_contact_details(name_search: str) -> str:
        """Get full details of a contact by name search.

        Args:
            name_search: Partial name to find the contact.
        """
        contact = get_client().get_contact_details(name_search)
        if not contact:
            return f"No contact found matching '{name_search}'."
        return json.dumps(contact, indent=2, ensure_ascii=False)

    @mcp.tool()
    def create_contact(
        full_name: str,
        email: str = "",
        business_phone: str = "",
        mobile_phone: str = "",
        company: str = "",
        job_title: str = "",
        notes: str = "",
    ) -> str:
        """Create a new contact in Outlook.

        Args:
            full_name: Full name of the contact.
            email: Email address.
            business_phone: Business phone number.
            mobile_phone: Mobile phone number.
            company: Company name.
            job_title: Job title.
            notes: Additional notes.
        """
        contact = get_client().create_contact(
            full_name=full_name, email=email, business_phone=business_phone,
            mobile_phone=mobile_phone, company=company, job_title=job_title, notes=notes,
        )
        return f"Contact created:\n{json.dumps(contact, indent=2, ensure_ascii=False)}"
