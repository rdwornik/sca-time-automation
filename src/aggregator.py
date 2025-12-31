"""
Aggregate time entries by week + category + client.
"""

import pandas as pd


def aggregate_entries(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aggregate hours by week_beginning + category + client.
    """
    grouped = df.groupby(
        ["week_beginning", "category", "client", "needs_opp_id"],
        as_index=False
    ).agg({
        "hours": "sum",
        "opportunity_id": "first",  # Add this
        "title": lambda x: "; ".join(x.unique()[:5]),
        "external_domains": lambda x: ",".join(set(",".join(x).split(",")) - {""}),
        "needs_review": "max"  # Add this - True if any needs review
    })
    
    grouped = grouped.rename(columns={"title": "comments"})
    grouped["status"] = "NEW"
    grouped = grouped.sort_values(["week_beginning", "category", "client"])
    
    return grouped

def add_week_summaries(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add summary row after each week with total hours.
    """
    rows = []
    
    for week in df["week_beginning"].unique():
        week_data = df[df["week_beginning"] == week]
        
        # Add week entries
        for _, row in week_data.iterrows():
            rows.append(row.to_dict())
        
        # Add summary row
        total_hours = week_data["hours"].sum()
        rows.append({
            "week_beginning": week,
            "category": ">>> WEEK TOTAL",
            "client": "",
            "hours": total_hours,
            "needs_opp_id": False,
            "comments": f"Total: {total_hours}h / 40h = {total_hours/40*100:.0f}%",
            "external_domains": "",
            "status": "---",
            "opportunity_id": ""
        })
    
    return pd.DataFrame(rows)