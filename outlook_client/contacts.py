"""
Outlook COM client – contact operations.
"""

from __future__ import annotations

from outlook_client._helpers import contact_to_dict


class ContactsMixin:
    """Contacts methods – mixed into OutlookClient."""

    def _contacts_folder(self):
        """Return the default Contacts folder (olFolderContacts = 10)."""
        return self._ns.GetDefaultFolder(10)

    # ── Queries ──────────────────────────────────────────────────────

    def search_contacts(self, query: str, max_results: int = 20) -> list[dict]:
        """Search contacts by name, email, or company."""
        contacts = self._contacts_folder()
        items = contacts.Items

        # Try COM-level filter first for performance
        needle = query.replace("'", "''")
        try:
            restriction = (
                f"@SQL=(\"urn:schemas:contacts:cn\" LIKE '%{needle}%'"
                f" OR \"urn:schemas:contacts:email1\" LIKE '%{needle}%'"
                f" OR \"urn:schemas:contacts:o\" LIKE '%{needle}%')"
            )
            filtered = items.Restrict(restriction)
            results: list[dict] = []
            for i, contact in enumerate(filtered):
                if i >= max_results:
                    break
                results.append(contact_to_dict(contact))
            return results
        except Exception:
            pass

        # Fallback: Python-level search
        needle_lower = query.lower()
        results = []
        for contact in items:
            if len(results) >= max_results:
                break
            try:
                name = str(getattr(contact, "FullName", "") or "").lower()
                email = str(getattr(contact, "Email1Address", "") or "").lower()
                company = str(getattr(contact, "CompanyName", "") or "").lower()
                if needle_lower in name or needle_lower in email or needle_lower in company:
                    results.append(contact_to_dict(contact))
            except Exception:
                continue
        return results

    def get_contact_details(self, name_search: str) -> dict | None:
        """Get full details of a contact by name search."""
        contacts = self._contacts_folder()
        items = contacts.Items
        needle = name_search.lower()
        for contact in items:
            try:
                name = str(getattr(contact, "FullName", "") or "").lower()
                if needle in name:
                    return contact_to_dict(contact, full=True)
            except Exception:
                continue
        return None

    # ── Mutations ────────────────────────────────────────────────────

    def create_contact(
        self,
        full_name: str,
        email: str = "",
        business_phone: str = "",
        mobile_phone: str = "",
        company: str = "",
        job_title: str = "",
        notes: str = "",
    ) -> dict:
        """Create a new contact."""
        contact = self._app.CreateItem(2)  # olContactItem
        contact.FullName = full_name
        if email:
            contact.Email1Address = email
        if business_phone:
            contact.BusinessTelephoneNumber = business_phone
        if mobile_phone:
            contact.MobileTelephoneNumber = mobile_phone
        if company:
            contact.CompanyName = company
        if job_title:
            contact.JobTitle = job_title
        if notes:
            contact.Body = notes
        contact.Save()
        return contact_to_dict(contact)
