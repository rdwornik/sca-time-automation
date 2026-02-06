#!/usr/bin/env python3
"""
SCA Time Automation - Main CLI entry point

Commands:
  export              Reminder to run VBA export script
  preview             Generate Excel preview with time entries
  preview --no-ai     Generate preview without AI (YAML-based only, faster)
  preview --weeks N   Filter to last N weeks (default: from config)
  upload WEEK         Upload specific week (e.g., "2025-12-07")
  upload --latest     Upload most recent week from preview
  upload --all        Upload all weeks from preview
  status              Show weeks in preview and their upload status
  report              Generate manager report (Weekly Hours + Opportunities)
  report --weeks N    Report for last N weeks (default: from config)
"""

import argparse
import sys
from pathlib import Path

from src.config import get_settings
from src.excel_preview import generate_final_preview
from src.sharepoint import post_week_entries, post_all_weeks

import pandas as pd


def cmd_export():
    """Remind user to run VBA export script."""
    print("=" * 60)
    print("STEP 1: Export Calendar from Outlook")
    print("=" * 60)
    print()
    print("1. Open Outlook")
    print("2. Press Alt+F11 to open VBA Editor")
    print("3. Run the 'ExportCalendarWithExternalDomains' macro")
    print("4. Save export to: data/input/calendar_export.json")
    print()
    print("Then run: python run.py preview")
    print()


def cmd_preview(use_ai: bool = True, weeks_back: int | None = None):
    """Generate Excel preview with time entries."""
    settings = get_settings()

    # Use default from config if not specified
    if weeks_back is None:
        weeks_back = settings.get("report", {}).get("weeks_back", 12)

    # Check AI configuration
    ai_enabled = settings["ai"]["enabled"] and use_ai
    mode = "AI-enabled" if ai_enabled else "YAML-only"

    print(f"Generating preview ({mode}, last {weeks_back} weeks)...")
    print()

    # Generate preview using the complete workflow in excel_preview
    output_path = settings["paths"]["excel_preview"]
    df = generate_final_preview(output_path, fill=True, weeks_back=weeks_back)

    # Count entries (excluding summary rows)
    entry_count = len(df[df["category"] != ">>> WEEK TOTAL"])
    week_count = df[df["category"] != ">>> WEEK TOTAL"]["week_beginning"].nunique()

    print(f"Generated {entry_count} entries across {week_count} weeks")
    print()
    print(f"Preview generated: {output_path}")
    print()
    print("Review the Excel file and then:")
    print("  - Upload all weeks:    python run.py upload --all")
    print("  - Upload latest week:  python run.py upload --latest")
    print("  - Upload specific week: python run.py upload 2025-12-07")
    print()



def cmd_upload(week: str = None, latest: bool = False, all_weeks: bool = False):
    """Upload time entries to SharePoint.

    Args:
        week: Specific week to upload (e.g., "2025-12-07")
        latest: Upload only the most recent week
        all_weeks: Upload all weeks from preview
    """
    settings = get_settings()
    preview_path = settings["paths"]["excel_preview"]

    if not Path(preview_path).exists():
        print(f"Error: Preview file not found: {preview_path}")
        print("Run 'python run.py preview' first")
        sys.exit(1)

    # Load preview Excel
    df = pd.read_excel(preview_path)

    # Get unique weeks (excluding summary rows)
    weeks = df[df["category"] != ">>> WEEK TOTAL"]["week_beginning"].unique()
    weeks = sorted([w for w in weeks if pd.notna(w)])

    if len(weeks) == 0:
        print("No weeks found in preview")
        sys.exit(1)

    # Upload all weeks
    if all_weeks:
        print(f"Uploading all {len(weeks)} weeks from preview...")
        result = post_all_weeks(df)

        print()
        print(f"Upload complete: {result['totals']['success']} successful, {result['totals']['failed']} failed")

        if result['totals']['failed'] > 0:
            print("\nFailed entries by week:")
            for week_key, week_results in result['by_week'].items():
                failed_in_week = [r for r in week_results if not r['success']]
                if failed_in_week:
                    print(f"  [{week_key}]")
                    for r in failed_in_week:
                        print(f"    - {r['category']}: {r.get('error', 'Unknown error')}")
            sys.exit(1)
        return

    # Determine which single week to upload
    if latest:
        target_week = weeks[-1]
        print(f"Uploading latest week: {target_week}")
    elif week:
        if week not in weeks:
            print(f"Error: Week '{week}' not found in preview")
            print(f"Available weeks: {', '.join(weeks)}")
            sys.exit(1)
        target_week = week
        print(f"Uploading week: {target_week}")
    else:
        print("Error: Specify --all, --latest, or provide a week date")
        sys.exit(1)

    # Upload single week
    print()
    results = post_week_entries(df, target_week)

    # Summary
    successful = sum(1 for r in results if r["success"])
    failed = len(results) - successful

    print()
    print(f"Upload complete: {successful} successful, {failed} failed")

    if failed > 0:
        print("\nFailed entries:")
        for r in results:
            if not r["success"]:
                print(f"  - {r['category']}: {r.get('error', 'Unknown error')}")
        sys.exit(1)


def cmd_report(weeks_back: int | None = None):
    """Generate manager report (Weekly Hours + Opportunities)."""
    from scripts.manager_report import generate_manager_report
    generate_manager_report(weeks_back=weeks_back)


def cmd_status():
    """Show weeks in preview and their upload status."""
    settings = get_settings()
    preview_path = settings["paths"]["excel_preview"]

    if not Path(preview_path).exists():
        print(f"Preview file not found: {preview_path}")
        print("Run 'python run.py preview' first")
        return

    # Load preview Excel
    df = pd.read_excel(preview_path)

    # Get weeks and summaries
    weeks_data = []
    for week in df["week_beginning"].unique():
        if pd.isna(week):
            continue

        week_df = df[df["week_beginning"] == week]
        total_row = week_df[week_df["category"] == ">>> WEEK TOTAL"]

        if len(total_row) > 0:
            total_hours = total_row.iloc[0]["hours"]
        else:
            total_hours = week_df[week_df["category"] != ">>> WEEK TOTAL"]["hours"].sum()

        weeks_data.append({
            "week": week,
            "hours": total_hours,
            "entries": len(week_df[week_df["category"] != ">>> WEEK TOTAL"])
        })

    # Sort by week
    weeks_data = sorted(weeks_data, key=lambda x: x["week"])

    print()
    print("=" * 60)
    print("WEEKS IN PREVIEW")
    print("=" * 60)
    print()

    for w in weeks_data:
        status = "✓ Ready" if w["hours"] >= 40 else f"⚠ {w['hours']}h (target: 40h)"
        print(f"{w['week']:>12} | {w['entries']:>2} entries | {w['hours']:>5.1f}h | {status}")

    print()
    print(f"Total weeks: {len(weeks_data)}")
    print()


def main():
    parser = argparse.ArgumentParser(
        description="SCA Time Automation CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # export command
    subparsers.add_parser("export", help="Reminder to run VBA export script")

    # preview command
    preview_parser = subparsers.add_parser("preview", help="Generate Excel preview")
    preview_parser.add_argument("--no-ai", action="store_true", help="Disable AI (faster, YAML-based only)")
    preview_parser.add_argument("--weeks", type=int, default=None, help="Number of weeks back to include (default: from config)")

    # upload command
    upload_parser = subparsers.add_parser("upload", help="Upload time entries to SharePoint")
    upload_parser.add_argument("week", nargs="?", help="Week to upload (YYYY-MM-DD)")
    upload_parser.add_argument("--latest", action="store_true", help="Upload most recent week")
    upload_parser.add_argument("--all", action="store_true", help="Upload all weeks from preview")

    # status command
    subparsers.add_parser("status", help="Show weeks in preview")

    # report command
    report_parser = subparsers.add_parser("report", help="Generate manager report (Weekly Hours + Opportunities)")
    report_parser.add_argument("--weeks", type=int, default=None, help="Number of weeks back to include (default: from config)")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        sys.exit(1)

    try:
        if args.command == "export":
            cmd_export()
        elif args.command == "preview":
            cmd_preview(use_ai=not args.no_ai, weeks_back=args.weeks)
        elif args.command == "upload":
            cmd_upload(week=args.week, latest=args.latest, all_weeks=getattr(args, 'all', False))
        elif args.command == "status":
            cmd_status()
        elif args.command == "report":
            cmd_report(weeks_back=args.weeks)
        else:
            parser.print_help()
            sys.exit(1)
    except KeyboardInterrupt:
        print("\n\nInterrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
