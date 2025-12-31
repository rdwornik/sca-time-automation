"""
Configuration loader for SCA Time Automation.
Loads settings from YAML files and environment variables.
"""

import os
from pathlib import Path
from dotenv import load_dotenv
import yaml

load_dotenv()


def get_project_root() -> Path:
    """Return project root directory."""
    return Path(__file__).parent.parent  # było parent.parent.parent


def load_yaml(filename: str) -> dict:
    """Load YAML config file from config/ directory."""
    config_path = get_project_root() / "config" / filename
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_settings() -> dict:
    """Load settings.yaml."""
    return load_yaml("settings.yaml")


def get_category_mapping() -> dict:
    """Load category_mapping.yaml."""
    return load_yaml("category_mapping.yaml")


def get_excluded() -> dict:
    """Load excluded.yaml."""
    return load_yaml("excluded.yaml")


def get_env(key: str, default: str = "") -> str:
    """Get environment variable."""
    return os.getenv(key, default)
