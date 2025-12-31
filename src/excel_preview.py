"""
Generate Excel preview for approval before SharePoint submission.
"""

import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

from src.config import get_settings, get_category_mapping
from src.loader import load_and_filter
from src.mapper import map_category, detect_client

from datetime import datetime, timedelta


def split_multiday_events(events: list) -> list:
    """Split multi-day all_day events into daily entries (8h per workday)."""
    result = []
    
    for event in events:
        if not event.get("all_day") or event["minutes"] <= 480:
            result.append(event)
            continue
        
        # Parse dates
        start = datetime.strptime(event["start"][:10], "%Y-%m-%d")
        end = datetime.strptime(event["end"][:10], "%Y-%m-%d")
        
        # Create daily entries
        current = start
        while current < end:
            # Skip weekends (Saturday=5, Sunday=6)
            if current.weekday() < 5:
                daily_event = event.copy()
                daily_event["start"] = current.strftime("%Y-%m-%d 09:00")
                daily_event["end"] = current.strftime("%Y-%m-%d 17:00")
                daily_event["minutes"] = 480  # 8 hours
                daily_event["all_day"] = False
                result.append(daily_event)
            current += timedelta(days=1)
    
    return result

def get_week_beginning(date_str: str) -> str:
    """Get Sunday of the week for given date."""
    dt = datetime.strptime(date_str[:10], "%Y-%m-%d")
    days_since_sunday = (dt.weekday() + 1) % 7
    sunday = dt - timedelta(days=days_since_sunday)
    return sunday.strftime("%Y-%m-%d")


def round_hours(hours: float, precision: float = 0.5) -> float:
    """Round hours to nearest precision."""
    return round(hours / precision) * precision

def generate_preview(clients_config: dict, output_path: str | Path | None = None) -> pd.DataFrame:
    """Generate Excel preview from calendar events."""
    from src.project_codes import load_project_codes, match_opportunity_id
    
    settings = get_settings()
    category_mapping = get_category_mapping()
    sales_categories = set(category_mapping.get("sales_categories", []))
    
    if output_path is None:
        output_path = Path(settings["paths"]["excel_preview"])
    
    events = load_and_filter()
    events = split_multiday_events(events)
    
    project_codes = load_project_codes()
    rows = []
    
    for event in events:
        sp_category = map_category(event["category"])
        if not sp_category:
            continue
        
        client = detect_client(
            event,
            clients_config.get("domains", {}),
            clients_config.get("keywords", {})
        )
        
        week = get_week_beginning(event["start"])
        hours = round_hours(event["minutes"] / 60)
        needs_opp_id = sp_category in sales_categories
        
        # Match opportunity ID - for ANY row with client
        opp_id, needs_review = "", False
        if client:
            opp_id, needs_review = match_opportunity_id(client, event["title"], project_codes)

        rows.append({
            "week_beginning": week,
            "category": sp_category,
            "hours": hours,
            "client": client or "",
            "opportunity_id": opp_id,
            "title": event["title"],
            "external_domains": event.get("external_domains", ""),
            "needs_opp_id": bool(client),  # Any row with client needs Opp ID
            "needs_review": needs_review,
            "status": "NEW"
        })
    
    df = pd.DataFrame(rows)
    
    return df

def generate_aggregated_preview(clients_config: dict, output_path: str | Path | None = None) -> pd.DataFrame:
    """Generate aggregated Excel preview with week summaries."""
    from src.aggregator import aggregate_entries, add_week_summaries
    
    settings = get_settings()
    if output_path is None:
        output_path = Path(settings["paths"]["excel_preview"])
    
    df = generate_preview(clients_config, output_path=None)
    agg = aggregate_entries(df)
    agg_with_summary = add_week_summaries(agg)
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    agg_with_summary.to_excel(output_path, index=False)
    
    return agg_with_summary

def generate_final_preview(clients_config: dict, output_path: str | Path | None = None, fill: bool = True) -> pd.DataFrame:
    """Generate final Excel preview with optional gap filling."""
    from src.gap_filler import fill_gaps
    
    agg = generate_aggregated_preview(clients_config, output_path=None)
    
    if fill:
        agg = fill_gaps(agg)
    
    if output_path is None:
        settings = get_settings()
        output_path = Path(settings["paths"]["excel_preview"])
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    agg.to_excel(output_path, index=False)
    
    return agg