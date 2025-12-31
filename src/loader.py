"""
Calendar JSON loader.
"""

import json
from pathlib import Path
from typing import TypedDict

from src.config import get_settings, get_excluded

class CalendarEvent(TypedDict):
    start: str
    end: str
    category: str
    title: str
    minutes: int
    all_day: bool
    external_domains: str
    location: str
    recipients: int
    busy_status: int


def load_calendar(path: str | Path | None = None) -> list[CalendarEvent]:
    """Load calendar events from JSON export."""
    if path is None:
        settings = get_settings()
        path = Path(settings["paths"]["calendar_input"])

    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)

    return data["events"]


def filter_excluded(events: list[CalendarEvent]) -> list[CalendarEvent]:
    """Remove events with excluded categories."""
    excluded = get_excluded()
    excluded_cats = {c.upper() for c in excluded["categories"]}

    return [e for e in events if e["category"].upper() not in excluded_cats]


def load_and_filter(path: str | Path | None = None) -> list[CalendarEvent]:
    """Load calendar and filter excluded categories."""
    events = load_calendar(path)
    return filter_excluded(events)
