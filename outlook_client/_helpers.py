"""
Shared helpers for Outlook COM interaction.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any


# ── BusyStatus constants ────────────────────────────────────────────

BUSY_STATUS_MAP = {
    0: "free",
    1: "tentative",
    2: "busy",
    3: "out_of_office",
    4: "working_elsewhere",
}
BUSY_STATUS_REVERSE = {v: k for k, v in BUSY_STATUS_MAP.items()}

IMPORTANCE_MAP = {0: "low", 1: "normal", 2: "high"}
IMPORTANCE_REVERSE = {v: k for k, v in IMPORTANCE_MAP.items()}

TASK_STATUS_MAP = {
    0: "not_started",
    1: "in_progress",
    2: "complete",
    3: "waiting",
    4: "deferred",
}
TASK_STATUS_REVERSE = {v: k for k, v in TASK_STATUS_MAP.items()}


# ── Datetime utilities ──────────────────────────────────────────────


def ensure_datetime(val: Any) -> _dt.datetime | None:
    """Convert a COM date (pywintypes.datetime) or string to a Python datetime.

    pywin32 adds a spurious +00:00 tzinfo to COM VT_DATE values which are
    actually local time.  We strip it so the caller sees plain local times.
    """
    if val is None:
        return None
    if isinstance(val, _dt.datetime):
        return val.replace(tzinfo=None)
    try:
        return _dt.datetime.fromisoformat(str(val)).replace(tzinfo=None)
    except Exception:
        return None


def local_dt_string(d: _dt.datetime) -> str:
    """Format a datetime as a string that Outlook COM interprets as local time.

    pywin32 silently converts timezone-aware datetimes to UTC before handing
    them to COM, but COM VT_DATE has no timezone concept – Outlook then treats
    the resulting value as local time, which shifts the event.  Using a string
    sidesteps this entirely.
    """
    return d.strftime("%m/%d/%Y %I:%M %p")


def date_restriction(start: _dt.datetime, end: _dt.datetime) -> str:
    """Build a DASL date-range restriction string."""
    fmt = "%m/%d/%Y %H:%M"
    return (
        f"[Start] >= '{start.strftime(fmt)}'"
        f" AND [Start] <= '{end.strftime(fmt)}'"
    )


# ── Serialisation helpers ───────────────────────────────────────────


def event_to_dict(
    item: Any,
    include_body: bool = False,
    include_attendees: bool = False,
) -> dict:
    """Serialise an Outlook AppointmentItem to a plain dict."""
    start = ensure_datetime(item.Start)
    end = ensure_datetime(item.End)
    busy = getattr(item, "BusyStatus", 2)
    sensitivity = getattr(item, "Sensitivity", 0)
    d: dict[str, Any] = {
        "subject": getattr(item, "Subject", ""),
        "start": start.isoformat() if start else "",
        "end": end.isoformat() if end else "",
        "location": getattr(item, "Location", ""),
        "organizer": getattr(item, "Organizer", ""),
        "is_recurring": getattr(item, "IsRecurring", False),
        "show_as": BUSY_STATUS_MAP.get(busy, "busy"),
        "is_private": sensitivity == 2,
        "meeting_status": getattr(item, "MeetingStatus", 0),
        "is_online_meeting": bool(
            getattr(item, "Location", "")
            and "teams" in getattr(item, "Location", "").lower()
        ),
    }
    if include_body:
        d["body"] = getattr(item, "Body", "")
    if include_attendees:
        try:
            recipients = item.Recipients
            attendees = []
            for i in range(1, recipients.Count + 1):
                r = recipients.Item(i)
                attendees.append({
                    "name": r.Name,
                    "address": r.Address,
                    "type": (
                        "required" if r.Type == 1
                        else "optional" if r.Type == 2
                        else "resource"
                    ),
                })
            d["attendees"] = attendees
        except Exception:
            d["attendees"] = []
    return d


def mail_to_dict(mail: Any) -> dict:
    """Serialise an Outlook MailItem to a plain dict."""
    received = ensure_datetime(getattr(mail, "ReceivedTime", None))
    return {
        "subject": getattr(mail, "Subject", ""),
        "from": getattr(mail, "SenderName", ""),
        "from_address": getattr(mail, "SenderEmailAddress", ""),
        "received": received.isoformat() if received else "",
        "unread": getattr(mail, "UnRead", False),
        "has_attachments": (
            getattr(mail, "Attachments", None) is not None
            and getattr(mail.Attachments, "Count", 0) > 0
        ),
    }


def contact_to_dict(contact: Any, full: bool = False) -> dict:
    """Serialise an Outlook ContactItem to a plain dict."""
    d: dict[str, Any] = {
        "full_name": getattr(contact, "FullName", ""),
        "email": getattr(contact, "Email1Address", ""),
        "company": getattr(contact, "CompanyName", ""),
        "job_title": getattr(contact, "JobTitle", ""),
        "business_phone": getattr(contact, "BusinessTelephoneNumber", ""),
        "mobile_phone": getattr(contact, "MobileTelephoneNumber", ""),
    }
    if full:
        d.update({
            "email2": getattr(contact, "Email2Address", ""),
            "email3": getattr(contact, "Email3Address", ""),
            "home_phone": getattr(contact, "HomeTelephoneNumber", ""),
            "department": getattr(contact, "Department", ""),
            "office_location": getattr(contact, "OfficeLocation", ""),
            "business_address": getattr(contact, "BusinessAddress", ""),
            "home_address": getattr(contact, "HomeAddress", ""),
            "web_page": getattr(contact, "WebPage", ""),
            "im_address": getattr(contact, "IMAddress", ""),
            "categories": getattr(contact, "Categories", ""),
            "notes": getattr(contact, "Body", ""),
        })
    return d


def task_to_dict(task: Any) -> dict:
    """Serialise an Outlook TaskItem to a plain dict."""
    due = ensure_datetime(getattr(task, "DueDate", None))
    start = ensure_datetime(getattr(task, "StartDate", None))
    # Outlook uses Jan 1 4501 as "no date" sentinel
    if due and due.year > 4000:
        due = None
    if start and start.year > 4000:
        start = None
    return {
        "subject": getattr(task, "Subject", ""),
        "due_date": due.isoformat() if due else "",
        "start_date": start.isoformat() if start else "",
        "status": TASK_STATUS_MAP.get(getattr(task, "Status", 0), "not_started"),
        "percent_complete": getattr(task, "PercentComplete", 0),
        "importance": IMPORTANCE_MAP.get(getattr(task, "Importance", 1), "normal"),
        "complete": getattr(task, "Complete", False),
        "categories": getattr(task, "Categories", ""),
        "reminder": getattr(task, "ReminderSet", False),
    }
