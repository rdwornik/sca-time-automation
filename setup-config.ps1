# setup-config.ps1

# config/settings.yaml - parametry techniczne
@"
# Ścieżki
paths:
  calendar_input: "data/input/calendar_export.json"
  project_codes: "data/input/project_codes.xlsx"
  excel_preview: "data/output/time_entries_preview.xlsx"

# Parametry przetwarzania
processing:
  min_duration_minutes: 15
  hour_rounding: 0.5
  week_start_day: "sunday"
  work_hours_per_day: 8

# SharePoint
sharepoint:
  graph_base_url: "https://graph.microsoft.com/v1.0"
"@ | Set-Content "config\settings.yaml" -Encoding UTF8

# config/category_mapping.yaml - Twoje kategorie Outlook -> SharePoint
@"
# Outlook Category -> SharePoint Category
mapping:
  CUSTOMER PRES/DEMO: "Customer – Demo/ Presentation"
  CUSTOMER PREP: "Customer – Demo/ Presentation"
  PREP: "Prep – Demo/ Presentation"
  RFI/RFP/RFQ: "RFI/RFP/RFQ"
  HOLIDAY: "Time Off"
  INTERNAL MEETING: "Internal Meeting"
  MARKETING: "Support"
  PARTNER SUPPORT: "Support"
  TRAINING: "Training"
  TRAVEL: "Travel"
  ADMIN: "Admin"

# Kategorie wymagające Opportunity ID
sales_categories:
  - "Customer – Demo/ Presentation"
  - "Prep – Demo/ Presentation"
  - "RFI/RFP/RFQ"
  - "Discovery"
  - "POC"
"@ | Set-Content "config\category_mapping.yaml" -Encoding UTF8

# config/excluded.yaml - wykluczenia
@"
# Kategorie do pominięcia (nie trackowane)
categories:
  - "PERSONAL"

# Słowa kluczowe w tytułach do pominięcia (opcjonalne)
title_keywords: []
"@ | Set-Content "config\excluded.yaml" -Encoding UTF8

Write-Host "Config files created in config/"