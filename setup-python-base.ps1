# setup-python-base.ps1

# src/sca_time_automation/config.py - loader konfiguracji
@"
import os
from pathlib import Path
from dotenv import load_dotenv
import yaml

# Load .env
load_dotenv()

def get_project_root() -> Path:
    return Path(__file__).parent.parent.parent

def load_yaml(filename: str) -> dict:
    config_path = get_project_root() / "config" / filename
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def get_settings() -> dict:
    return load_yaml("settings.yaml")

def get_category_mapping() -> dict:
    return load_yaml("category_mapping.yaml")

def get_excluded() -> dict:
    return load_yaml("excluded.yaml")

def get_env(key: str, default: str = "") -> str:
    return os.getenv(key, default)

# Shortcuts
SETTINGS = None
CATEGORY_MAPPING = None
EXCLUDED = None

def init():
    global SETTINGS, CATEGORY_MAPPING, EXCLUDED
    SETTINGS = get_settings()
    CATEGORY_MAPPING = get_category_mapping()
    EXCLUDED = get_excluded()
"@ | Set-Content "src\sca_time_automation\config.py" -Encoding UTF8

# src/sca_time_automation/__init__.py
@"
__version__ = "0.1.0"
"@ | Set-Content "src\sca_time_automation\__init__.py" -Encoding UTF8

Write-Host "Created: src/sca_time_automation/config.py"