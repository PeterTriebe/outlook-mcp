"""
Outlook COM client – email operations.
"""

from __future__ import annotations

import datetime as _dt
import os

from outlook_client._helpers import mail_to_dict


class EmailMixin:
    """Email methods – mixed into OutlookClient."""

    def _inbox_folder(self):
        """Return the default Inbox folder (olFolderInbox = 6)."""
        return self._ns.GetDefaultFolder(6)

    # ── Queries ──────────────────────────────────────────────────────

    def list_unread_emails(self, max_results: int = 10) -> list[dict]:
        """Return unread e-mails from the Inbox."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        filtered = items.Restrict("[Unread] = True")

        results: list[dict] = []
        for i, mail in enumerate(filtered):
            if i >= max_results:
                break
            results.append(mail_to_dict(mail))
        return results

    def list_recent_emails(self, max_results: int = 10) -> list[dict]:
        """Return the most recent e-mails from the Inbox."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        results: list[dict] = []
        for i, mail in enumerate(items):
            if i >= max_results:
                break
            results.append(mail_to_dict(mail))
        return results

    def search_emails(self, query: str, max_results: int = 10) -> list[dict]:
        """Search e-mails by subject (fast COM filter)."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        filtered = items.Restrict(
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{query}%'"
        )

        results: list[dict] = []
        for i, mail in enumerate(filtered):
            if i >= max_results:
                break
            results.append(mail_to_dict(mail))
        return results

    def search_emails_full(
        self,
        query: str,
        max_results: int = 20,
        days_back: int = 365,
        search_subject: bool = True,
        search_body: bool = True,
        search_sender: bool = True,
        include_body: bool = False,
    ) -> list[dict]:
        """Search e-mails across subject, body, and sender with date range."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        cutoff = _dt.datetime.now() - _dt.timedelta(days=days_back)
        fmt = "%m/%d/%Y %H:%M"
        date_restriction = f"[ReceivedTime] >= '{cutoff.strftime(fmt)}'"
        filtered = items.Restrict(date_restriction)

        needle = query.lower()
        results: list[dict] = []

        for mail in filtered:
            if len(results) >= max_results:
                break
            try:
                matched = False
                subject = str(getattr(mail, "Subject", "") or "")
                sender = str(getattr(mail, "SenderName", "") or "")
                body_text = ""

                if search_subject and needle in subject.lower():
                    matched = True
                if not matched and search_sender and needle in sender.lower():
                    matched = True
                if not matched and search_body:
                    body_text = str(getattr(mail, "Body", "") or "")
                    if needle in body_text.lower():
                        matched = True

                if matched:
                    d = mail_to_dict(mail)
                    if include_body:
                        d["body"] = body_text or str(getattr(mail, "Body", "") or "")
                    preview = body_text or str(getattr(mail, "Body", "") or "")
                    d["preview"] = preview[:300].strip().replace("\r\n", " ").replace("\n", " ")
                    results.append(d)
            except Exception:
                continue
        return results

    def get_email_body(self, subject_search: str) -> dict | None:
        """Find an email by subject and return full body."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                d = mail_to_dict(mail)
                d["body"] = getattr(mail, "Body", "")
                return d
        return None

    def list_sent_emails(self, max_results: int = 10) -> list[dict]:
        """List recent sent emails."""
        sent = self._ns.GetDefaultFolder(5)  # olFolderSentMail
        items = sent.Items
        items.Sort("[SentOn]", True)
        results: list[dict] = []
        for i, mail in enumerate(items):
            if i >= max_results:
                break
            results.append(mail_to_dict(mail))
        return results

    def list_drafts(self, max_results: int = 10) -> list[dict]:
        """List draft emails."""
        drafts = self._ns.GetDefaultFolder(16)  # olFolderDrafts
        items = drafts.Items
        items.Sort("[LastModificationTime]", True)
        results: list[dict] = []
        for i, mail in enumerate(items):
            if i >= max_results:
                break
            results.append(mail_to_dict(mail))
        return results

    # ── Mutations ────────────────────────────────────────────────────

    def send_email(
        self,
        to: str,
        subject: str,
        body: str,
        cc: str = "",
        is_html: bool = False,
    ) -> str:
        """Compose and send an e-mail."""
        mail = self._app.CreateItem(0)  # olMailItem
        mail.To = to
        mail.Subject = subject
        if is_html:
            mail.HTMLBody = body
        else:
            mail.Body = body
        if cc:
            mail.CC = cc
        mail.Send()
        return f"E-mail sent to {to}: '{subject}'"

    def create_draft(
        self, to: str, subject: str, body: str, cc: str = "", is_html: bool = False,
    ) -> str:
        """Create a draft email (saved, not sent)."""
        mail = self._app.CreateItem(0)  # olMailItem
        mail.To = to
        mail.Subject = subject
        if is_html:
            mail.HTMLBody = body
        else:
            mail.Body = body
        if cc:
            mail.CC = cc
        mail.Save()
        return f"Draft created: '{subject}' to {to}"

    def reply_to_email(
        self, subject_search: str, body: str, reply_all: bool = False,
    ) -> str:
        """Reply (or reply-all) to an email found by subject."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                reply = mail.ReplyAll() if reply_all else mail.Reply()
                reply.Body = body + "\n\n" + reply.Body
                reply.Send()
                mode = "reply-all" if reply_all else "reply"
                return f"Sent {mode} to '{mail.Subject}'"
        return f"No email found matching '{subject_search}'."

    def forward_email(
        self, subject_search: str, to: str, body: str = "",
    ) -> str:
        """Forward an email found by subject."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                fwd = mail.Forward()
                fwd.To = to
                if body:
                    fwd.Body = body + "\n\n" + fwd.Body
                fwd.Send()
                return f"Forwarded '{mail.Subject}' to {to}"
        return f"No email found matching '{subject_search}'."

    def delete_email(self, subject_search: str) -> str:
        """Delete an email found by subject."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                subj = mail.Subject
                mail.Delete()
                return f"Deleted email: '{subj}'"
        return f"No email found matching '{subject_search}'."

    def move_email(self, subject_search: str, folder_name: str) -> str:
        """Move an email to a different folder."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)

        target = self._find_folder(folder_name)
        if not target:
            return f"Folder '{folder_name}' not found."

        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                subj = mail.Subject
                mail.Move(target)
                return f"Moved '{subj}' to '{folder_name}'"
        return f"No email found matching '{subject_search}'."

    def mark_email(self, subject_search: str, read: bool = True) -> str:
        """Mark an email as read or unread."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                mail.UnRead = not read
                mail.Save()
                status = "read" if read else "unread"
                return f"Marked '{mail.Subject}' as {status}"
        return f"No email found matching '{subject_search}'."

    def flag_email(self, subject_search: str, flag: bool = True) -> str:
        """Flag or unflag an email for follow-up."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                mail.FlagStatus = 2 if flag else 0  # olFlagMarked / olNoFlag
                mail.Save()
                status = "flagged" if flag else "unflagged"
                return f"{status.capitalize()} '{mail.Subject}'"
        return f"No email found matching '{subject_search}'."

    # ── Folders ──────────────────────────────────────────────────────

    def list_email_folders(self) -> list[dict]:
        """List all email folders with item counts."""
        root = self._ns.GetDefaultFolder(6).Parent  # Inbox parent = account root
        return self._enumerate_folders(root)

    def _enumerate_folders(self, folder, prefix: str = "") -> list[dict]:
        """Recursively enumerate mail folders."""
        results: list[dict] = []
        try:
            for i in range(1, folder.Folders.Count + 1):
                f = folder.Folders.Item(i)
                path = f"{prefix}/{f.Name}" if prefix else f.Name
                results.append({
                    "name": f.Name,
                    "path": path,
                    "item_count": f.Items.Count,
                    "unread_count": getattr(f, "UnReadItemCount", 0),
                })
                results.extend(self._enumerate_folders(f, path))
        except Exception:
            pass
        return results

    def _find_folder(self, folder_name: str):
        """Find a folder by name (recursive search)."""
        root = self._ns.GetDefaultFolder(6).Parent
        return self._find_folder_recursive(root, folder_name.lower())

    def _find_folder_recursive(self, folder, needle: str):
        """Recursively search for a folder by name."""
        try:
            for i in range(1, folder.Folders.Count + 1):
                f = folder.Folders.Item(i)
                if f.Name.lower() == needle:
                    return f
                result = self._find_folder_recursive(f, needle)
                if result:
                    return result
        except Exception:
            pass
        return None

    # ── Attachments ──────────────────────────────────────────────────

    def list_email_attachments(self, subject_search: str) -> list[dict] | None:
        """List attachments of an email found by subject."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                attachments: list[dict] = []
                try:
                    for i in range(1, mail.Attachments.Count + 1):
                        att = mail.Attachments.Item(i)
                        attachments.append({
                            "name": att.FileName,
                            "size": att.Size,
                            "index": i,
                        })
                except Exception:
                    pass
                return attachments
        return None

    def save_attachment(
        self, subject_search: str, attachment_name: str, save_path: str,
    ) -> str:
        """Save an email attachment to a local directory."""
        inbox = self._inbox_folder()
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
        needle = subject_search.lower()
        att_needle = attachment_name.lower()
        for mail in items:
            if needle in str(getattr(mail, "Subject", "")).lower():
                try:
                    for i in range(1, mail.Attachments.Count + 1):
                        att = mail.Attachments.Item(i)
                        if att_needle in att.FileName.lower():
                            full_path = os.path.join(save_path, att.FileName)
                            att.SaveAsFile(full_path)
                            return f"Saved '{att.FileName}' to {full_path}"
                except Exception as e:
                    return f"Error saving attachment: {e}"
                return f"No attachment matching '{attachment_name}' found."
        return f"No email found matching '{subject_search}'."
