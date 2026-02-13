"""
MCP tools – Email.
"""

from __future__ import annotations

import json


def register_email_tools(mcp, get_client):
    """Register all email tools."""

    @mcp.tool()
    def list_unread_emails(max_results: int = 10) -> str:
        """List unread emails from the Inbox.

        Args:
            max_results: Maximum number of emails to return (default 10).
        """
        mails = get_client().list_unread_emails(max_results=max_results)
        if not mails:
            return "No unread emails."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def list_recent_emails(max_results: int = 10) -> str:
        """List the most recent emails from the Inbox.

        Args:
            max_results: Maximum number of emails to return (default 10).
        """
        mails = get_client().list_recent_emails(max_results=max_results)
        if not mails:
            return "No emails found."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def search_emails(query: str, max_results: int = 10) -> str:
        """Search emails by subject.

        Args:
            query: Search text to look for in email subjects.
            max_results: Maximum number of results (default 10).
        """
        mails = get_client().search_emails(query=query, max_results=max_results)
        if not mails:
            return f"No emails found matching '{query}'."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def search_emails_full(
        query: str,
        max_results: int = 20,
        days_back: int = 365,
        search_subject: bool = True,
        search_body: bool = True,
        search_sender: bool = True,
        include_body: bool = False,
    ) -> str:
        """Search emails across subject, body, and sender name with date range filter.

        Use this for comprehensive email search when the simple subject search is not enough.
        Can search in subject, email body text, and sender name simultaneously.

        Args:
            query: Text to search for in emails.
            max_results: Maximum number of results to return (default 20).
            days_back: How many days back to search (default 365).
            search_subject: Search in email subject (default True).
            search_body: Search in email body text (default True).
            search_sender: Search in sender name (default True).
            include_body: Include the full email body in results (default False).
        """
        mails = get_client().search_emails_full(
            query=query, max_results=max_results, days_back=days_back,
            search_subject=search_subject, search_body=search_body,
            search_sender=search_sender, include_body=include_body,
        )
        if not mails:
            return f"No emails found matching '{query}'."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def read_email(subject_search: str) -> str:
        """Read the full body of an email by searching its subject.

        Args:
            subject_search: Partial subject text to find the email.
        """
        mail = get_client().get_email_body(subject_search)
        if not mail:
            return f"No email found matching '{subject_search}'."
        return json.dumps(mail, indent=2, ensure_ascii=False)

    @mcp.tool()
    def send_email(to: str, subject: str, body: str, cc: str = "") -> str:
        """Send an email via Outlook.

        Args:
            to: Recipient email address(es), separated by semicolons.
            subject: Email subject.
            body: Email body text.
            cc: Optional CC recipients, separated by semicolons.
        """
        return get_client().send_email(to=to, subject=subject, body=body, cc=cc)

    @mcp.tool()
    def reply_to_email(
        subject_search: str, body: str, reply_all: bool = False,
    ) -> str:
        """Reply to an email found by subject search.

        Args:
            subject_search: Partial subject text to find the email.
            body: Reply body text.
            reply_all: If True, reply to all recipients (default False).
        """
        return get_client().reply_to_email(subject_search, body, reply_all=reply_all)

    @mcp.tool()
    def forward_email(subject_search: str, to: str, body: str = "") -> str:
        """Forward an email to another recipient.

        Args:
            subject_search: Partial subject text to find the email.
            to: Recipient email address(es), separated by semicolons.
            body: Optional additional text to include above the forwarded message.
        """
        return get_client().forward_email(subject_search, to, body=body)

    @mcp.tool()
    def delete_email(subject_search: str) -> str:
        """Delete an email by searching its subject.

        Args:
            subject_search: Partial subject text to find the email to delete.
        """
        return get_client().delete_email(subject_search)

    @mcp.tool()
    def move_email(subject_search: str, folder_name: str) -> str:
        """Move an email to a different Outlook folder.

        Args:
            subject_search: Partial subject text to find the email.
            folder_name: Target folder name (e.g. 'Archive', 'Projects').
        """
        return get_client().move_email(subject_search, folder_name)

    @mcp.tool()
    def mark_email(subject_search: str, read: bool = True) -> str:
        """Mark an email as read or unread.

        Args:
            subject_search: Partial subject text to find the email.
            read: True to mark as read, False to mark as unread (default True).
        """
        return get_client().mark_email(subject_search, read=read)

    @mcp.tool()
    def flag_email(subject_search: str, flag: bool = True) -> str:
        """Flag or unflag an email for follow-up.

        Args:
            subject_search: Partial subject text to find the email.
            flag: True to flag, False to unflag (default True).
        """
        return get_client().flag_email(subject_search, flag=flag)

    @mcp.tool()
    def list_email_folders() -> str:
        """List all email folders with item and unread counts."""
        folders = get_client().list_email_folders()
        if not folders:
            return "No folders found."
        return json.dumps(folders, indent=2, ensure_ascii=False)

    @mcp.tool()
    def list_sent_emails(max_results: int = 10) -> str:
        """List recently sent emails.

        Args:
            max_results: Maximum number of emails to return (default 10).
        """
        mails = get_client().list_sent_emails(max_results=max_results)
        if not mails:
            return "No sent emails found."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def list_drafts(max_results: int = 10) -> str:
        """List draft emails.

        Args:
            max_results: Maximum number of drafts to return (default 10).
        """
        mails = get_client().list_drafts(max_results=max_results)
        if not mails:
            return "No drafts found."
        return json.dumps(mails, indent=2, ensure_ascii=False)

    @mcp.tool()
    def create_draft(to: str, subject: str, body: str, cc: str = "") -> str:
        """Create a draft email (saved but not sent).

        Args:
            to: Recipient email address(es), separated by semicolons.
            subject: Email subject.
            body: Email body text.
            cc: Optional CC recipients, separated by semicolons.
        """
        return get_client().create_draft(to=to, subject=subject, body=body, cc=cc)

    @mcp.tool()
    def list_email_attachments(subject_search: str) -> str:
        """List all attachments of an email found by subject.

        Args:
            subject_search: Partial subject text to find the email.
        """
        attachments = get_client().list_email_attachments(subject_search)
        if attachments is None:
            return f"No email found matching '{subject_search}'."
        if not attachments:
            return "Email has no attachments."
        return json.dumps(attachments, indent=2, ensure_ascii=False)

    @mcp.tool()
    def save_attachment(
        subject_search: str, attachment_name: str, save_path: str,
    ) -> str:
        """Save an email attachment to a local directory.

        Args:
            subject_search: Partial subject text to find the email.
            attachment_name: Partial filename of the attachment to save.
            save_path: Local directory path to save the file to.
        """
        return get_client().save_attachment(subject_search, attachment_name, save_path)
