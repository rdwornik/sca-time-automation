"""
Generate Excel preview for approval before SharePoint submission.
"""

import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path

from src.config import get_settings, get_category_mapping
from src.loader import load_and_filter
from src.mapper import map_category, detect_client


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
    settings = get_settings()
    category_mapping = get_category_mapping()
    sales_categories = set(category_mapping.get("sales_categories", []))
    
    if output_path is None:
        output_path = Path(settings["paths"]["excel_preview"])
    
    events = load_and_filter()
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
        
        rows.append({
            "week_beginning": week,
            "category": sp_category,
            "hours": hours,
            "client": client or "",
            "opportunity_id": "",  # To fill from project_codes
            "title": event["title"],
            "external_domains": event.get("external_domains", ""),
            "needs_opp_id": needs_opp_id,
            "status": "NEW"
        })
    
    df = pd.DataFrame(rows)
    
    # Save to Excel
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    
    return df