import pandas as pd
from src.sharepoint import post_week_entries

df = pd.read_excel("data/output/time_entries_preview.xlsx")

# Get unique weeks (exclude test week)
weeks = df["week_beginning"].unique()
skip_week = "2025-12-07"

print(f"Found {len(weeks)} weeks")
print(f"Skipping: {skip_week}")

total_success = 0
total_failed = 0

for week in sorted(weeks):
    if week == skip_week:
        print(f"\n⏭️  Skipping {week} (already uploaded)")
        continue
    
    print(f"\n📤 Uploading {week}...")
    results = post_week_entries(df, week)
    
    success = sum(1 for r in results if r["success"])
    failed = len(results) - success
    
    total_success += success
    total_failed += failed
    
    print(f"   {success}/{len(results)} entries uploaded")

print(f"\n{'='*40}")
print(f"TOTAL: {total_success} success, {total_failed} failed")