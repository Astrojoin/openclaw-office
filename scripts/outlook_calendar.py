#!/usr/bin/env python3
"""Outlook calendar operations via Microsoft Graph API.

List calendars, list/view/create/update/delete events.

Usage:
  from outlook_calendar import Calendar

  cal = Calendar()
  events = cal.list_events(days_ahead=7)
  cal.create(subject="Meeting", start="2026-05-16T10:00:00", end="2026-05-16T11:00:00")
"""

from __future__ import annotations

import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from graph_client import GraphClient


class Calendar:
    """Outlook calendar operations via Microsoft Graph API."""

    def __init__(self):
        self.client = GraphClient()

    # ── CALENDARS ───────────────────────────────────────────────────────

    def list_calendars(self) -> list[dict]:
        """List all calendars for the user."""
        items = self.client.get_all("/me/calendars")
        return [{
            "id": c.get("id", ""),
            "name": c.get("name", ""),
            "isDefault": c.get("isDefaultCalendar", False),
            "owner": c.get("owner", {}).get("name", ""),
        } for c in items]

    def get_calendar(self, calendar_id: str) -> dict:
        """Get a specific calendar by ID."""
        c = self.client.get(f"/me/calendars/{calendar_id}")
        return {
            "id": c.get("id", ""),
            "name": c.get("name", ""),
            "isDefault": c.get("isDefaultCalendar", False),
        }

    # ── EVENTS ──────────────────────────────────────────────────────────

    def list_events(self, calendar_id: str | None = None,
                    days_ahead: int = 7, days_back: int = 0,
                    count: int = 25) -> list[dict]:
        """List upcoming events.

        Args:
            calendar_id: Specific calendar ID, or None for default
            days_ahead: Number of days ahead to look
            days_back: Number of days back to look
            count: Max events to return
        """
        now = datetime.now(timezone.utc)
        start = (now - timedelta(days=days_back)).isoformat()
        end = (now + timedelta(days=days_ahead)).isoformat()

        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/calendarview"
        else:
            endpoint = "/me/calendar/calendarview"

        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$top": min(count, 200),
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,organizer,isAllDay,isCancelled,recurrence,responseStatus",
        }

        items = self.client.get_all(endpoint, params=params)
        return [self._summarize(e) for e in items]

    def get_event(self, event_id: str, calendar_id: str | None = None) -> dict:
        """Get full details of an event.

        Args:
            event_id: Event ID
            calendar_id: Calendar ID (optional, defaults to primary)
        """
        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events/{event_id}"
        else:
            endpoint = f"/me/events/{event_id}"

        item = self.client.get(endpoint, params={
            "$select": "id,subject,start,end,body,location,organizer,attendees,isAllDay,isCancelled,recurrence,responseStatus,sensitivity,importance"
        })
        return self._full_event(item)

    # ── CREATE ──────────────────────────────────────────────────────────

    def create(self, subject: str, start: str, end: str,
               body: str = "", body_type: str = "text",
               location: str = "", attendees: list[str] | None = None,
               is_all_day: bool = False, calendar_id: str | None = None,
               recurrence: dict | None = None, sensitivity: str = "normal",
               importance: str = "normal") -> dict:
        """Create a calendar event.

        Args:
            subject: Event title
            start: Start time ISO string (e.g. "2026-05-16T10:00:00")
            end: End time ISO string
            body: Event description
            body_type: "text" or "html"
            location: Location name
            attendees: List of attendee email addresses
            is_all_day: All-day event flag
            calendar_id: Target calendar (optional)
            recurrence: Recurrence pattern dict (optional)
            sensitivity: "normal", "personal", "private", "confidential"
            importance: "low", "normal", "high"
        """
        # Determine timezone from start string or use UTC
        timezone_id = "UTC"
        if start.endswith(("+05:00", "-05:00")) or "America/" in start or "Europe/" in start:
            # Try to extract timezone
            pass
        # Use the user's preferred timezone — default to UTC
        # The Graph API accepts time without zone and uses the user's mailbox timezone
        event_data = {
            "subject": subject,
            "start": {
                "dateTime": start,
                "timeZone": "UTC",
            },
            "end": {
                "dateTime": end,
                "timeZone": "UTC",
            },
            "isAllDay": is_all_day,
            "sensitivity": sensitivity,
            "importance": importance,
        }

        if body:
            event_data["body"] = {
                "contentType": "HTML" if body_type == "html" else "Text",
                "content": body,
            }
        if location:
            event_data["location"] = {"displayName": location}
        if attendees:
            event_data["attendees"] = [
                {
                    "emailAddress": {"address": addr},
                    "type": "required",
                }
                for addr in attendees
            ]
        if recurrence:
            event_data["recurrence"] = recurrence

        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events"
        else:
            endpoint = "/me/events"

        result = self.client.post(endpoint, event_data)
        return self._summarize(result)

    # ── UPDATE ──────────────────────────────────────────────────────────

    def update(self, event_id: str, calendar_id: str | None = None,
               **fields) -> dict:
        """Update an event. Pass fields to update as kwargs.

        Supported: subject, start, end, body, body_type, location, is_all_day,
                   sensitivity, importance, attendees

        For start/end, pass ISO string — this method wraps them in the timeZone dict.
        """
        data = {}

        if "subject" in fields:
            data["subject"] = fields["subject"]

        if "start" in fields:
            data["start"] = {"dateTime": fields["start"], "timeZone": "UTC"}

        if "end" in fields:
            data["end"] = {"dateTime": fields["end"], "timeZone": "UTC"}

        if "body" in fields:
            bt = fields.get("body_type", "text")
            data["body"] = {
                "contentType": "HTML" if bt == "html" else "Text",
                "content": fields["body"],
            }

        if "location" in fields:
            data["location"] = {"displayName": fields["location"]}

        if "is_all_day" in fields:
            data["isAllDay"] = fields["is_all_day"]

        if "sensitivity" in fields:
            data["sensitivity"] = fields["sensitivity"]

        if "importance" in fields:
            data["importance"] = fields["importance"]

        if "attendees" in fields:
            data["attendees"] = [
                {"emailAddress": {"address": addr}, "type": "required"}
                for addr in fields["attendees"]
            ]

        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events/{event_id}"
        else:
            endpoint = f"/me/events/{event_id}"

        result = self.client.patch(endpoint, data)
        return self._summarize(result) if result else {"updated": True}

    # ── DELETE ──────────────────────────────────────────────────────────

    def delete(self, event_id: str, calendar_id: str | None = None) -> bool:
        """Delete a calendar event."""
        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/events/{event_id}"
        else:
            endpoint = f"/me/events/{event_id}"
        return self.client.delete(endpoint)

    # ── RESPOND ─────────────────────────────────────────────────────────

    def accept(self, event_id: str, comment: str = "") -> str:
        """Accept a meeting invitation."""
        self.client.post(f"/me/events/{event_id}/accept", {"comment": comment})
        return "Accepted"

    def decline(self, event_id: str, comment: str = "") -> str:
        """Decline a meeting invitation."""
        self.client.post(f"/me/events/{event_id}/decline", {"comment": comment})
        return "Declined"

    def tentatively_accept(self, event_id: str, comment: str = "") -> str:
        """Tentatively accept a meeting invitation."""
        self.client.post(f"/me/events/{event_id}/tentativelyAccept", {"comment": comment})
        return "Tentatively accepted"

    # ── HELPERS ─────────────────────────────────────────────────────────

    @staticmethod
    def _summarize(event: dict) -> dict:
        start = event.get("start", {})
        end = event.get("end", {})
        organizer = event.get("organizer", {})
        org_email = organizer.get("emailAddress", {})
        response = event.get("responseStatus", {})

        return {
            "id": event.get("id", ""),
            "subject": event.get("subject", "(no title)"),
            "start": start.get("dateTime", ""),
            "end": end.get("dateTime", ""),
            "isAllDay": event.get("isAllDay", False),
            "location": event.get("location", {}).get("displayName", ""),
            "organizer": org_email.get("address", org_email.get("name", "")),
            "isCancelled": event.get("isCancelled", False),
            "response": response.get("response", ""),
        }

    @staticmethod
    def _full_event(event: dict) -> dict:
        result = Calendar._summarize(event)
        result["body"] = event.get("body", {}).get("content", "")
        result["body_type"] = event.get("body", {}).get("contentType", "")
        result["attendees"] = [
            a.get("emailAddress", {}).get("address", "")
            for a in event.get("attendees", [])
        ]
        result["recurrence"] = event.get("recurrence", None)
        result["sensitivity"] = event.get("sensitivity", "normal")
        result["importance"] = event.get("importance", "normal")
        return result


# ── CLI ─────────────────────────────────────────────────────────────────────

def main():
    """Quick test: list upcoming events."""
    cal = Calendar()
    try:
        events = cal.list_events(days_ahead=7)
        print(f"Upcoming events ({len(events)}):")
        for e in events:
            print(f"  📅 [{e['start'][:16]}] {e['subject']} @ {e['location'] or 'no location'}")
        if not events:
            print("  (no events in the next 7 days)")
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
