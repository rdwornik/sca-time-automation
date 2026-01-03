# SCA Time Automation

Automates weekly time entry submission to SharePoint SCA Time Tracker from Outlook calendar exports.

## Overview

This tool streamlines your weekly time tracking workflow:

1. **Export** calendar events from Outlook (VBA script)
2. **Detect** clients from meeting titles and external attendee domains
3. **Map** Outlook categories to SharePoint categories
4. **Match** events to Project Codes (Opportunity IDs) from Excel
5. **Fill gaps** with intelligent autofill to reach 40h target
6. **Generate** Excel preview for manual review and approval
7. **Upload** approved entries to SharePoint via Graph API

## Features

- **AI-Powered Client Detection**: Uses Gemini AI to intelligently match meetings to clients
- **Dual Mode**: Run with AI (accurate) or YAML-only (faster) client detection
- **Smart Overlap Resolution**: Automatically prioritizes customer-facing activities
- **Intelligent Gap Filling**: Distributes missing hours based on your actual work patterns
- **Week Summaries**: Automatic totals and validation
- **Manual Override**: Excel preview allows full editing before upload

## Quick Start

### 1. Prerequisites

- Python 3.12+
- Outlook with VBA access
- SharePoint access with Graph API permissions
- (Optional) Gemini API key for AI-powered client detection

### 2. Installation

```bash
# Clone repository
git clone <repository-url>
cd sca-time-automation

# Create virtual environment
python -m venv .venv

# Activate (Windows)
.venv\Scripts\Activate.ps1  # PowerShell
# OR
.venv\Scripts\activate.bat  # Command Prompt

# Install dependencies
pip install -r requirements.txt
```

### 3. Configuration

#### 3.1 Environment Variables (.env)

Create `.env` file in project root:

```bash
# OneDrive path (for project_codes.xlsx)
ONEDRIVE_PATH=C:\Users\YourName\OneDrive

# SharePoint Graph API (required for upload)
GRAPH_ACCESS_TOKEN=your_access_token_here

# Gemini AI (optional, for intelligent client detection)
GEMINI_API_KEY=your_gemini_api_key_here
```

**How to get Graph Access Token:**
1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Sign in with your JDA365 account
3. Grant permissions: `Sites.ReadWrite.All`, `User.Read`
4. Copy the access token (valid for ~1 hour)

**How to get Gemini API Key:**
1. Go to https://aistudio.google.com/apikey
2. Create a new API key
3. Copy and add to `.env`

#### 3.2 Settings (config/settings.yaml)

Main configuration file (already configured with sensible defaults):

```yaml
paths:
  calendar_input: "data/input/calendar_export.json"
  project_codes: "${ONEDRIVE_PATH}/Projects/_Technical Presales/Projects/Project_Codes.xlsx"
  excel_preview: "data/output/time_entries_preview.xlsx"

processing:
  work_hours_target: 40        # Weekly target
  work_start_hour: 9           # Workday start
  work_end_hour: 17            # Workday end
  hour_rounding: 0.5           # Round to 0.5h increments

sharepoint:
  site_id: "jda365.sharepoint.com,..."
  list_id: "70738fad-ba9a-4e2c-99cf-3adc450f6127"

ai:
  enabled: true                # Set to false to disable AI
  model: "gemini-2.0-flash-exp"
```

#### 3.3 Input Files

Place in `data/input/`:
- **project_codes.xlsx** - Symlink to your OneDrive Project Codes file
  - Columns: `Company | Project Description | Project Code`

### 4. Install VBA Export Script

1. Open Outlook
2. Press `Alt+F11` to open VBA Editor
3. Insert > Module
4. Copy content from `scripts/calendar_export.vbs`
5. Save and close VBA Editor

## Usage

### Weekly Workflow

#### Step 1: Export Calendar

```bash
python run.py export
```

This shows instructions to run the VBA script:
1. Open Outlook
2. Press `Alt+F11`
3. Run `ExportCalendarWithExternalDomains` macro
4. Export saves to `data/input/calendar_export.json`

#### Step 2: Generate Preview

**With AI (recommended, slower but more accurate):**
```bash
python run.py preview
```

**Without AI (faster, keyword matching only):**
```bash
python run.py preview --no-ai
```

This generates `data/output/time_entries_preview.xlsx` with:
- All calendar events mapped to SharePoint categories
- Client detection (AI or keyword-based)
- Opportunity IDs matched from project codes
- Overlap resolution (highest priority activity wins)
- Gap filling to reach 40h target
- Week totals and validation

#### Step 3: Review and Edit

Open `data/output/time_entries_preview.xlsx`:
- ✅ Check client detection accuracy
- ✅ Verify Opportunity IDs
- ✅ Review autofilled entries (marked with `is_autofilled=True`)
- ✅ Edit comments, adjust hours as needed
- ✅ Ensure week totals are correct

#### Step 4: Check Status

```bash
python run.py status
```

Shows all weeks in the preview with totals and validation status.

#### Step 5: Upload to SharePoint

**Upload most recent week:**
```bash
python run.py upload --latest
```

**Upload specific week:**
```bash
python run.py upload 2025-12-07
```

The upload will:
- Post entries to SharePoint Time Tracker
- Show progress and results
- Report any errors

## Configuration Files

### Category Mapping (config/category_mapping.yaml)

Maps Outlook categories to SharePoint categories:

```yaml
mapping:
  PREP: "Prep - Demo/ Presentation"
  CUSTOMER: "Customer - Demo/ Presentation"
  ADMIN: "Admin"
  SUPPORT: "Support"
  # ... more mappings
```

Your Outlook categories keep the `.` prefix - it's stripped automatically during export.

### Excluded Categories (config/excluded.yaml)

Categories to skip (not tracked):

```yaml
categories:
  - PERSONAL
  - PRIVATE
```

## Client Detection Modes

### AI Mode (Default)
- Uses Gemini AI to intelligently match meetings to clients
- Considers external domains (e.g., `michelin.com` → Michelin)
- Understands language hints (Italian titles → Italian clients)
- Recognizes abbreviations and context clues
- Falls back to keyword matching if AI unavailable

### YAML-Only Mode
- Keyword matching from `project_codes.xlsx` company names
- Faster but less accurate
- Use when you don't have Gemini API key
- Use with `--no-ai` flag: `python run.py preview --no-ai`

## Gap Filling

If your calendar has empty time slots, the system automatically generates autofill entries:

- **Distribution**: Based on your actual work patterns for the week
- **Categories**: Only fills safe categories (Prep, Admin, Support, Training, Internal Meeting)
- **Never autofills**: Customer meetings, Discovery, POC, RFI/RFP/RFQ, Travel, Time Off
- **Smart comments**: AI-generated realistic descriptions (or simple fallback)
- **Reviewable**: All autofilled entries marked `is_autofilled=True` in Excel

## Overlap Resolution

When calendar events overlap:
1. System assigns each hour slot to highest priority activity
2. Priority order (high to low):
   - Customer - Demo/Presentation
   - Discovery
   - RFI/RFP/RFQ
   - POC
   - Prep - Demo/Presentation
   - Internal Meeting
   - Training
   - Support
   - Admin
   - Travel
   - Time Off

## Troubleshooting

### "GEMINI_API_KEY not set"
- Add API key to `.env` OR
- Run with `--no-ai` flag: `python run.py preview --no-ai`

### "GRAPH_ACCESS_TOKEN not set"
- Get new token from Graph Explorer (expires hourly)
- Add to `.env`

### "Project codes file not found"
- Ensure `ONEDRIVE_PATH` is correct in `.env`
- Create symlink: `mklink data\input\project_codes.xlsx "path\to\Project_Codes.xlsx"`

### "No events loaded"
- Check calendar export JSON exists: `data/input/calendar_export.json`
- Run VBA export script again
- Verify JSON format

### Week total not 40h
- Review autofilled entries
- Check for Time Off weeks (skipped from autofill)
- Manually adjust hours in Excel preview

## Project Structure

```
sca-time-automation/
├── config/
│   ├── settings.yaml              # Main configuration
│   ├── category_mapping.yaml      # Outlook → SharePoint mapping
│   └── excluded.yaml              # Categories to skip
├── data/
│   ├── input/
│   │   ├── calendar_export.json   # VBA export from Outlook
│   │   └── project_codes.xlsx     # Symlink to OneDrive
│   └── output/
│       └── time_entries_preview.xlsx  # Generated preview
├── scripts/
│   ├── calendar_export.vbs        # Outlook VBA export script
│   └── generate_clients_yaml.py   # Helper to create clients.yaml
├── src/
│   ├── config.py                  # Config and env loader
│   ├── loader.py                  # Calendar JSON loader
│   ├── mapper.py                  # Category & client mapping
│   ├── aggregator.py              # Hour aggregation
│   ├── overlap.py                 # Overlap resolution
│   ├── gap_filler.py              # Smart autofill
│   ├── excel_preview.py           # Excel generation
│   ├── excel_writer.py            # Excel utilities
│   ├── project_codes.py           # Project codes loader
│   ├── gemini_client.py           # Gemini AI client
│   ├── sharepoint.py              # SharePoint Graph API
│   └── text_utils.py              # Text normalization
├── tests/                         # Unit tests
├── run.py                         # Main CLI entry point
├── requirements.txt               # Python dependencies
├── .env                           # Environment variables (not in git)
├── .env.example                   # Environment template
├── README.md                      # This file
└── CLAUDE.md                      # Development guide
```

## Development

See [CLAUDE.md](CLAUDE.md) for:
- Architecture overview
- Coding standards
- How to extend the system
- Key functions and data flow

## Commands Reference

```bash
# Show VBA export instructions
python run.py export

# Generate preview with AI
python run.py preview

# Generate preview without AI (faster)
python run.py preview --no-ai

# Show weeks and status
python run.py status

# Upload latest week
python run.py upload --latest

# Upload specific week
python run.py upload 2025-12-07
```

## License

Internal use only.
