"""
Fill gaps to reach 40h/week by distributing hours proportionally.
"""

import pandas as pd

EXCLUDED_FROM_FILL = {"Time Off", "Travel", ">>> WEEK TOTAL"}
MIN_TRACKED_HOURS = 20.0  # Don't fill weeks with less than this (probably incomplete)


def calculate_gaps(df: pd.DataFrame, target_hours: float = 40.0) -> dict[str, float]:
    """Calculate missing hours per week."""
    gaps = {}
    
    for week in df["week_beginning"].unique():
        week_data = df[
            (df["week_beginning"] == week) & 
            (~df["category"].isin(EXCLUDED_FROM_FILL))
        ]
        
        # Skip weeks with Time Off
        has_time_off = ((df["week_beginning"] == week) & (df["category"] == "Time Off")).any()
        if has_time_off:
            continue
        
        total = week_data["hours"].sum()
        
        # Only fill if above minimum threshold
        if total >= MIN_TRACKED_HOURS and total < target_hours:
            gaps[week] = target_hours - total
    
    return gaps


def fill_gaps(df: pd.DataFrame, target_hours: float = 40.0) -> pd.DataFrame:
    """Distribute missing hours proportionally to existing entries."""
    df = df.copy()
    gaps = calculate_gaps(df, target_hours)
    
    for week, missing in gaps.items():
        mask = (
            (df["week_beginning"] == week) & 
            (~df["category"].isin(EXCLUDED_FROM_FILL))
        )
        week_data = df[mask]
        
        if week_data.empty:
            continue
        
        total_hours = week_data["hours"].sum()
        if total_hours == 0:
            continue
        
        # Distribute and track remainder
        distributed = 0.0
        indices = list(week_data.index)
        
        for i, idx in enumerate(indices):
            proportion = df.loc[idx, "hours"] / total_hours
            additional = missing * proportion
            
            if i == len(indices) - 1:
                # Last entry gets remainder to hit exactly target
                additional = missing - distributed
            
            rounded = round(additional * 2) / 2  # Round to 0.5
            df.loc[idx, "hours"] += rounded
            distributed += rounded
        
        # Update week total
        total_mask = (df["week_beginning"] == week) & (df["category"] == ">>> WEEK TOTAL")
        if total_mask.any():
            new_total = df[mask]["hours"].sum()
            df.loc[total_mask, "hours"] = new_total
            df.loc[total_mask, "comments"] = f"Total: {new_total}h / 40h = {new_total/40*100:.0f}%"
    
    return df