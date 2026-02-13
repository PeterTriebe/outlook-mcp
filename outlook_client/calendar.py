"""
Outlook COM client – calendar operations.
"""

from __future__ import annotations

import datetime as _dt

from outlook_client._helpers import (
    BUSY_STATUS_REVERSE,
    ensure_datetime,
    event_to_dict,
    local_dt_string,
)


class CalendarMixin:
    """Calendar methods – mixed into OutlookClient."""

    def _calendar_folder(self):
        """Return the default Calendar folder (olFolderCalendar = 9)."""
        return self._ns.GetDefaultFolder(9)

    # ── Queries ──────────────────────────────────────────────────────

    def list_events(
        self,
        start: _dt.datetime | None = None,
        end: _dt.datetime | None = None,
        max_results: int = 25,
    ) -> list[dict]:
        """Return calendar events in the given time window."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        if start is None:
            start = _dt.datetime.now()
        if end is None:
            end = start + _dt.timedelta(days=7)

        fmt = "%m/%d/%Y %H:%M"
        restriction = (
            f"[Start] >= '{start.strftime(fmt)}'"
            f" AND [Start] <= '{end.strftime(fmt)}'"
        )
        filtered = items.Restrict(restriction)

        results: list[dict] = []
        for i, item in enumerate(filtered):
            if i >= max_results:
                break
            results.append(event_to_dict(item))
        return results

    def get_next_meeting(self) -> dict | None:
        """Return the very next upcoming event."""
        events = self.list_events(
            start=_dt.datetime.now(),
            end=_dt.datetime.now() + _dt.timedelta(days=30),
            max_results=1,
        )
        return events[0] if events else None

    def get_event_details(
        self, subject_search: str, days_ahead: int = 90,
    ) -> dict | None:
        """Find an event by (partial) subject and return full details."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = _dt.datetime.now()
        end = now + _dt.timedelta(days=days_ahead)
        fmt = "%m/%d/%Y %H:%M"
        restriction = (
            f"[Start] >= '{now.strftime(fmt)}'"
            f" AND [Start] <= '{end.strftime(fmt)}'"
        )
        filtered = items.Restrict(restriction)

        needle = subject_search.lower()
        for item in filtered:
            if needle in str(getattr(item, "Subject", "")).lower():
                return event_to_dict(item, include_body=True, include_attendees=True)
        return None

    def search_events(
        self,
        query: str,
        days_ahead: int = 30,
        max_results: int = 10,
    ) -> list[dict]:
        """Search events by subject, organizer, or attendee name."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = _dt.datetime.now()
        end = now + _dt.timedelta(days=days_ahead)
        fmt = "%m/%d/%Y %H:%M"
        restriction = (
            f"[Start] >= '{now.strftime(fmt)}'"
            f" AND [Start] <= '{end.strftime(fmt)}'"
        )

        # COM-level subject filter first for performance
        needle = query.replace("'", "''")
        subject_restriction = (
            f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{needle}%'"
        )
        try:
            pre_filtered = items.Restrict(restriction)
            subject_matches = pre_filtered.Restrict(subject_restriction)
            results: list[dict] = []
            for i, item in enumerate(subject_matches):
                if i >= max_results:
                    break
                results.append(event_to_dict(item, include_attendees=True))
        except Exception:
            results = []

        # Also search organizer and attendees (Python-level)
        seen = {(r["subject"], r["start"]) for r in results}
        needle_lower = query.lower()
        filtered = items.Restrict(restriction)
        for item in filtered:
            if len(results) >= max_results:
                break
            subj = getattr(item, "Subject", "")
            start_dt = ensure_datetime(item.Start)
            start_iso = start_dt.isoformat() if start_dt else ""
            if (subj, start_iso) in seen:
                continue
            organizer = str(getattr(item, "Organizer", "")).lower()
            if needle_lower in organizer:
                results.append(event_to_dict(item, include_attendees=True))
                seen.add((subj, start_iso))
                continue
            try:
                recipients = item.Recipients
                for ri in range(1, recipients.Count + 1):
                    r = recipients.Item(ri)
                    if needle_lower in r.Name.lower():
                        results.append(event_to_dict(item, include_attendees=True))
                        seen.add((subj, start_iso))
                        break
            except Exception:
                pass
        return results

    # ── Mutations ────────────────────────────────────────────────────

    def create_event(
        self,
        subject: str,
        start: _dt.datetime,
        end: _dt.datetime,
        location: str = "",
        body: str = "",
        is_online: bool = False,
        is_private: bool = False,
        show_as: str = "busy",
    ) -> dict:
        """Create a new calendar event."""
        appt = self._app.CreateItem(1)  # olAppointmentItem
        appt.Subject = subject
        appt.Start = local_dt_string(start)
        appt.End = local_dt_string(end)
        if location:
            appt.Location = location
        if body:
            appt.Body = body
        if is_private:
            appt.Sensitivity = 2  # olPrivate
        appt.BusyStatus = BUSY_STATUS_REVERSE.get(show_as, 2)
        if is_online:
            try:
                appt.MeetingStatus = 1  # olMeeting
            except Exception:
                pass
        appt.Save()
        return event_to_dict(appt)

    def respond_to_event(self, subject_search: str, response: str) -> str:
        """Respond to a meeting (accept / tentative / decline)."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = _dt.datetime.now()
        fmt = "%m/%d/%Y %H:%M"
        restriction = f"[Start] >= '{now.strftime(fmt)}'"
        filtered = items.Restrict(restriction)

        needle = subject_search.lower()
        for item in filtered:
            if needle in str(getattr(item, "Subject", "")).lower():
                resp = response.lower().strip()
                if resp == "accept":
                    item.Respond(3)  # olMeetingAccepted
                elif resp == "tentative":
                    item.Respond(2)  # olMeetingTentative
                elif resp == "decline":
                    item.Respond(4)  # olMeetingDeclined
                else:
                    return f"Unknown response '{response}'. Use accept/tentative/decline."
                return f"Meeting '{item.Subject}' responded with: {resp}"
        return f"No upcoming meeting found matching '{subject_search}'."

    def update_event(
        self,
        subject_search: str,
        new_subject: str | None = None,
        new_start: _dt.datetime | None = None,
        new_end: _dt.datetime | None = None,
        new_location: str | None = None,
        is_private: bool | None = None,
        show_as: str | None = None,
    ) -> dict | None:
        """Find an event by subject and update its properties."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = _dt.datetime.now()
        fmt = "%m/%d/%Y %H:%M"
        restriction = f"[Start] >= '{now.strftime(fmt)}'"
        filtered = items.Restrict(restriction)

        needle = subject_search.lower()
        for item in filtered:
            if needle in str(getattr(item, "Subject", "")).lower():
                if new_subject is not None:
                    item.Subject = new_subject
                if new_start is not None:
                    item.Start = local_dt_string(new_start)
                if new_end is not None:
                    item.End = local_dt_string(new_end)
                if new_location is not None:
                    item.Location = new_location
                if is_private is not None:
                    item.Sensitivity = 2 if is_private else 0
                if show_as is not None:
                    item.BusyStatus = BUSY_STATUS_REVERSE.get(show_as, 2)
                item.Save()
                return event_to_dict(item)
        return None

    def delete_event(self, subject_search: str) -> str:
        """Delete an event by (partial) subject match."""
        cal = self._calendar_folder()
        items = cal.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        now = _dt.datetime.now()
        fmt = "%m/%d/%Y %H:%M"
        restriction = f"[Start] >= '{now.strftime(fmt)}'"
        filtered = items.Restrict(restriction)

        needle = subject_search.lower()
        for item in filtered:
            if needle in str(getattr(item, "Subject", "")).lower():
                subj = item.Subject
                item.Delete()
                return f"Deleted event: '{subj}'"
        return f"No upcoming event found matching '{subject_search}'."
