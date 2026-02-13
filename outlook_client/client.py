"""
OutlookClient – composes all domain mixins into a single client.
"""

from __future__ import annotations

import win32com.client
import pythoncom

from outlook_client.calendar import CalendarMixin
from outlook_client.email import EmailMixin
from outlook_client.contacts import ContactsMixin
from outlook_client.tasks import TasksMixin


class OutlookClient(CalendarMixin, EmailMixin, ContactsMixin, TasksMixin):
    """Full Outlook Desktop client via COM – calendar, email, contacts, tasks."""

    def __init__(self) -> None:
        pythoncom.CoInitialize()
        self._app = win32com.client.Dispatch("Outlook.Application")
        self._ns = self._app.GetNamespace("MAPI")
