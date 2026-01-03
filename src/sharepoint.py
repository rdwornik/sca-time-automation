"""
SharePoint Graph API connector for SCA Time Tracker.
"""

import requests
from src.config import get_env, get_settings


def get_graph_url() -> str:
    """Build Graph API URL from config."""
    settings = get_settings()
    site_id = settings["sharepoint"]["site_id"]
    list_id = settings["sharepoint"]["list_id"]
    base_url = settings["sharepoint"]["graph_base_url"]
    return f"{base_url}/sites/{site_id}/lists/{list_id}/items"

# Map our categories to SharePoint valid values
CATEGORY_MAP = {
    "Prep - Demo/ Presentation": "Prep – Demo/ Presentation",
    "Customer - Demo/ Presentation": "Customer – Demo/ Presentation", 
    "Time Off": "Time Off",
    "Admin": "Admin",
    "Support": "Support",
    "Internal Meeting": "Internal Meeting",
    "Training": "Training",
    "Discovery": "Discovery",
    "RFI/RFP/RFQ": "RFI/RFP/RFQ",
    "POC": "POC",
    "Travel": "Travel",
}


def get_access_token() -> str:
    """Get access token from environment or interactive login."""
    token = get_env("GRAPH_ACCESS_TOKEN")
    if not token:
        raise ValueError("GRAPH_ACCESS_TOKEN not set in .env")
    return token


def post_time_entry(entry: dict, access_token: str = None) -> dict:
    """Post single time entry to SharePoint."""
    import math
    
    if access_token is None:
        access_token = get_access_token()
    
    # Map category to SharePoint format
    sp_category = CATEGORY_MAP.get(entry.get("category"), entry.get("category"))
    
    # Clean NaN values
    def clean_value(val):
        if val is None:
            return None
        if isinstance(val, float) and math.isnan(val):
            return None
        return val
    
    fields = {
        "WeekBeginning": entry["week_beginning"],
        "Category": sp_category,
        "Hours": float(entry["hours"]),
    }
    
    # Add optional fields only if not NaN/None
    comments = clean_value(entry.get("comments"))
    if comments:
        fields["Comments"] = str(comments)
    
    opp_id = clean_value(entry.get("opportunity_id"))
    if opp_id:
        fields["OpportunityID"] = str(opp_id)
    
    client = clean_value(entry.get("client"))
    if client:
        fields["AccountName"] = str(client)
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.post(get_graph_url(), headers=headers, json={"fields": fields})
    
    if response.status_code == 201:
        return {"success": True, "data": response.json()}
    else:
        return {"success": False, "error": response.text, "status": response.status_code}

def post_week_entries(df, week: str, access_token: str = None) -> list:
    """Post all entries for a specific week."""
    if access_token is None:
        access_token = get_access_token()
    
    week_data = df[
        (df["week_beginning"] == week) & 
        (df["category"] != ">>> WEEK TOTAL")
    ]
    
    results = []
    for _, row in week_data.iterrows():
        result = post_time_entry(row.to_dict(), access_token)
        results.append({
            "category": row["category"],
            "hours": row["hours"],
            "success": result["success"],
            "error": result.get("error")
        })
        print(f"{'✓' if result['success'] else '✗'} {row['category']}: {row['hours']}h")
    
    return results