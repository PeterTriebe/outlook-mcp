"""
MCP tools – Calendar.
"""

from __future__ import annotations

import json
import datetime as dt


def register_calendar_tools(mcp, get_client):
    """Register all calendar tools."""

    @mcp.tool()
    def get_next_meeting() -> str:
        """Get the next upcoming calendar event / meeting."""
        event = get_client().get_next_meeting()
        if not event:
            return "No upcoming meetings found."
        return json.dumps(event, indent=2, ensure_ascii=False)

    @mcp.tool()
    def list_events_today() -> str:
        """List all calendar events for today."""
        now = dt.datetime.now()
        end = now.replace(hour=23, minute=59, second=59)
        events = get_client().list_events(start=now, end=end)
        if not events:
            return "No more events today."
        return json.dumps(events, indent=2, ensure_ascii=False)

    @mcp.tool()
    def list_events(
        start_date: str = "",
        end_date: str = "",
        max_results: int = 25,
    ) -> str:
        """List calendar events in a date range.

        Args:
            start_date: Start date (YYYY-MM-DD or YYYY-MM-DDTHH:MM). Defaults to now.
            end_date: End date (YYYY-MM-DD or YYYY-MM-DDTHH:MM). Defaults to 7 days from start.
            max_results: Maximum number of events to return (default 25).
        """
        start = dt.datetime.fromisoformat(start_date) if start_date else None
        end = dt.datetime.fromisoformat(end_date) if end_date else None
        events = get_client().list_events(start=start, end=end, max_results=max_results)
        if not events:
            return "No events found in the given date range."
        return json.dumps(events, indent=2, ensure_ascii=False)

    @mcp.tool()
    def get_event_details(subject_search: str) -> str:
        """Get full details of a calendar event (attendees, body, etc.) by searching its subject.

        Args:
            subject_search: Partial subject text to search for.
        """
        event = get_client().get_event_details(subject_search)
        if not event:
            return f"No upcoming event found matching '{subject_search}'."
        return json.dumps(event, indent=2, ensure_ascii=False)

    @mcp.tool()
    def search_events(
        query: str,
        days_ahead: int = 30,
        max_results: int = 10,
    ) -> str:
        """Search calendar events by subject, organizer, or attendee name.

        Use this to find events involving a specific person (organizer or attendee),
        or events whose subject matches a keyword.

        Args:
            query: Text to search for in subject, organizer, or attendee names.
            days_ahead: How many days into the future to search (default 30).
            max_results: Maximum number of matching events to return (default 10).
        """
        events = get_client().search_events(query, days_ahead=days_ahead, max_results=max_results)
        if not events:
            return f"No upcoming events found matching '{query}'."
        return json.dumps(events, indent=2, ensure_ascii=False)

    @mcp.tool()
    def create_event(
        subject: str,
        start_time: str,
        end_time: str,
        location: str = "",
        body: str = "",
        is_private: bool = False,
        show_as: str = "busy",
    ) -> str:
        """Create a new calendar event in Outlook.

        Args:
            subject: Event subject / title.
            start_time: Start datetime (YYYY-MM-DDTHH:MM).
            end_time: End datetime (YYYY-MM-DDTHH:MM).
            location: Optional location.
            body: Optional body / description text.
            is_private: Mark the event as private (default False).
            show_as: Status shown in calendar: free, tentative, busy (default), out_of_office, working_elsewhere.
        """
        start = dt.datetime.fromisoformat(start_time).replace(tzinfo=None)
        end = dt.datetime.fromisoformat(end_time).replace(tzinfo=None)
        event = get_client().create_event(
            subject=subject, start=start, end=end, location=location, body=body,
            is_private=is_private, show_as=show_as,
        )
        return f"Event created:\n{json.dumps(event, indent=2, ensure_ascii=False)}"

    @mcp.tool()
    def respond_to_meeting(subject_search: str, response: str) -> str:
        """Respond to a meeting invitation (accept / tentative / decline).

        Args:
            subject_search: Partial subject text to find the meeting.
            response: Response type – one of: accept, tentative, decline.
        """
        return get_client().respond_to_event(subject_search, response)

    @mcp.tool()
    def update_event(
        subject_search: str,
        new_subject: str = "",
        new_start_time: str = "",
        new_end_time: str = "",
        new_location: str = "",
        is_private: bool | None = None,
        show_as: str = "",
    ) -> str:
        """Update an existing calendar event found by subject search.

        Args:
            subject_search: Partial subject text to find the event.
            new_subject: New subject (leave empty to keep current).
            new_start_time: New start datetime YYYY-MM-DDTHH:MM (leave empty to keep current).
            new_end_time: New end datetime YYYY-MM-DDTHH:MM (leave empty to keep current).
            new_location: New location (leave empty to keep current).
            is_private: Set private flag (true/false, omit to keep current).
            show_as: New status: free, tentative, busy, out_of_office, working_elsewhere (leave empty to keep current).
        """
        event = get_client().update_event(
            subject_search=subject_search,
            new_subject=new_subject or None,
            new_start=dt.datetime.fromisoformat(new_start_time).replace(tzinfo=None) if new_start_time else None,
            new_end=dt.datetime.fromisoformat(new_end_time).replace(tzinfo=None) if new_end_time else None,
            new_location=new_location or None,
            is_private=is_private,
            show_as=show_as or None,
        )
        if not event:
            return f"No upcoming event found matching '{subject_search}'."
        return f"Event updated:\n{json.dumps(event, indent=2, ensure_ascii=False)}"

    @mcp.tool()
    def delete_event(subject_search: str) -> str:
        """Delete a calendar event by searching its subject.

        Args:
            subject_search: Partial subject text to find the event to delete.
        """
        return get_client().delete_event(subject_search)
