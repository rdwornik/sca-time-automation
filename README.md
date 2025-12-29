# SCA Time Automation

Automates time entry submission to SharePoint SCA Time Tracker from Outlook calendar exports.

## Overview

This tool:
1. Reads calendar events exported from Outlook (JSON)
2. Maps Outlook categories to SharePoint categories
3. Matches events to Project Codes (Opportunity IDs)
4. Generates Excel preview for approval
5. Submits approved entries to SharePoint via Graph API

## Project Structure

```
sca-time-automation/
├── config/
│   ├── settings.yaml           # Paths, processing parameters
│   ├── category_mapping.yaml   # Outlook -> SharePoint category mapping
│   └── excluded.yaml           # Categories/keywords to skip
├── data/
│   ├── input/                  # Calendar JSON, Project Codes Excel
│   └── output/                 # Excel preview for approval
├── src/
│   └── sca_time_automation/
├── scripts/                    # VBA export script, utilities
├── tests/
├── .env                        # API keys (not in repo)
├── .env.example
├── requirements.txt
└── README.md
```

## Setup

```bash
# Clone
git clone https://github.com/YOUR_USERNAME/sca-time-automation.git
cd sca-time-automation

# Virtual environment
python -m venv .venv
.venv\Scripts\Activate.ps1  # Windows PowerShell

# Install dependencies
pip install -r requirements.txt

# Configure
copy .env.example .env
# Edit .env with your Azure/SharePoint credentials
```

## Input Files

Place in `data/input/`:
- `calendar_export.json` - Outlook calendar export (from VBA macro)
- `project_codes.xlsx` - Columns: Company | Project Description | Project Code

## Usage

```bash
# Analyze calendar and generate Excel preview
python -m sca_time_automation.analyze

# Review and edit: data/output/time_entries_preview.xlsx

# Submit approved entries to SharePoint
python -m sca_time_automation.submit
```

## Configuration

### Category Mapping (config/category_mapping.yaml)

Maps your Outlook categories to SharePoint categories. Your Outlook categories remain unchanged.

### Excluded (config/excluded.yaml)

Categories and title keywords to skip (e.g., PERSONAL).

## License

Internal use only.