# setup-structure.ps1
# Uruchom w: C:\Users\1028120\Documents\sca-time-automation

# Katalogi
$dirs = @(
    "src\sca_time_automation",
    "config",
    "data\input",
    "data\output",
    "scripts",
    "tests"
)
foreach ($d in $dirs) { New-Item -ItemType Directory -Path $d -Force | Out-Null }

# .gitignore
@"
__pycache__/
*.pyc
.env
.venv/
data/input/*
data/output/*
!data/**/.gitkeep
*.log
token_cache.json
"@ | Set-Content ".gitignore" -Encoding UTF8

# requirements.txt
@"
requests>=2.31.0
msal>=1.24.0
python-dotenv>=1.0.0
pyyaml>=6.0.1
openpyxl>=3.1.0
pandas>=2.0.0
"@ | Set-Content "requirements.txt" -Encoding UTF8

# .env.example
@"
AZURE_TENANT_ID=
AZURE_CLIENT_ID=
SHAREPOINT_SITE_ID=
SHAREPOINT_LIST_ID=
"@ | Set-Content ".env.example" -Encoding UTF8

# Placeholder __init__.py
"" | Set-Content "src\sca_time_automation\__init__.py" -Encoding UTF8
"" | Set-Content "tests\__init__.py" -Encoding UTF8

# .gitkeep
"" | Set-Content "data\input\.gitkeep"
"" | Set-Content "data\output\.gitkeep"

Write-Host "Done. Run: pip install -r requirements.txt"