#!/usr/bin/env python3
"""
Generate clients.yaml from project_codes.xlsx

This is an optional helper script that extracts company names from project_codes.xlsx
and creates a basic keyword mapping YAML file. Users can then manually enhance it
with additional keywords, aliases, and domain hints.

Usage:
    python scripts/generate_clients_yaml.py
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.project_codes import load_project_codes
import yaml


def generate_clients_yaml():
    """Generate clients.yaml from project_codes.xlsx"""

    print("Loading project codes...")
    df = load_project_codes()

    # Extract unique companies
    companies = sorted(df["company"].unique())

    print(f"Found {len(companies)} companies")

    # Build clients structure
    clients = {}

    for company in companies:
        # Get project codes for this company
        company_rows = df[df["company"] == company]
        project_codes = company_rows["project_code"].unique().tolist()

        # Create basic entry
        # Users can manually add more keywords, domains, aliases
        clients[company] = {
            "keywords": [company.lower()],  # Basic keyword matching
            "project_codes": project_codes,
            # "domains": [],  # Users can add email domains here
            # "aliases": [],  # Users can add alternative names here
        }

    # Create output structure
    output = {
        "clients": clients,
        "notes": "This file was auto-generated. You can manually add:",
        "enhancement_guide": {
            "keywords": "Add words that appear in meeting titles (e.g., ['wurth', 'wuerth', 'würth'])",
            "domains": "Add email domains (e.g., ['michelin.com', 'michelin.fr'])",
            "aliases": "Add alternative company names (e.g., ['IBM', 'International Business Machines'])"
        }
    }

    # Write to config/clients.yaml
    output_path = Path(__file__).parent.parent / "config" / "clients.yaml"

    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(output, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

    print(f"\n✓ Generated: {output_path}")
    print("\nNext steps:")
    print("1. Review config/clients.yaml")
    print("2. Add email domains for better detection")
    print("3. Add keywords/aliases as needed")
    print("\nExample enhancement:")
    print("  Michelin:")
    print("    keywords: [michelin, bibendum]")
    print("    domains: [michelin.com, michelin.fr, michelin.it]")
    print("    aliases: [Compagnie Générale des Établissements Michelin]")
    print()


if __name__ == "__main__":
    try:
        generate_clients_yaml()
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
