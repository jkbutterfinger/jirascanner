#!/usr/bin/env python3
"""Zentrale Konfigurationsverwaltung für den Jira/Gira Scanner.

Alle Module (main.py, scanner/, api/, scripts/) laden ihre Konfiguration über
diese Datei, damit Defaults, Umgebungsvariablen und Pfadauflösung überall
identisch funktionieren.
"""

from __future__ import annotations

import copy
import os
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = ROOT / "config.yaml"

# Optionale .env-Unterstützung (python-dotenv ist optional).
try:  # pragma: no cover - nur Komfort
    from dotenv import load_dotenv

    load_dotenv(ROOT / ".env")
except Exception:  # pragma: no cover
    pass


DEFAULTS: Dict[str, Any] = {
    "server": {
        "url": "",
        "protocol": "https",
        "port": 443,
    },
    "api": {
        "token": "",
        "auth_type": "token",
        "cors_origins": ["http://localhost:8000", "http://127.0.0.1:8000"],
    },
    "project": {
        "key": "JRA",
        "name": "Jira Reporting",
        "jql": "",
    },
    "storage": {
        "inventory_path": str(ROOT / "inventory"),
        "inventory_file": "inventory.json",
        "suggestions_file": "suggestions.json",
        "jira_cache": str(ROOT / "jira_data" / "tasks.json"),
    },
    "scripts": {
        "base_path": str(ROOT / "scripts"),
        # Es werden ausschließlich SQL-Dateien gescannt.
        "extensions": [".sql"],
        "include_subdirectories": True,
        "exclude": [],
    },
    "filters": {
        "prefix": "JRA-",
        "min_priority": "medium",
        "statuses": ["open", "in progress", "done"],
        "exclude_keywords": [],
    },
    "performance": {
        "max_results": 1000,
        "timeout": 30,
        "retry_attempts": 3,
    },
    "excel": {
        # Fallback-Ordner nur für leere Berichte ohne Task-Verzeichnis.
        "output_path": str(ROOT / "output"),
        # Ein Excel-Bericht pro Auftrag, abgelegt im Task-Verzeichnis.
        "output_filename": "{JIRA}.xlsx",
        "fallback_filename": "jira_report_{YYYY}_{MM}_{DD}.xlsx",
        "include_headers": True,
        "include_csv": False,
        "datetime_format": "%Y-%m-%d %H:%M:%S",
        "sheet_name": "Jira-Task-Report",
    },
    "logging": {
        "level": "INFO",
        "output_file": str(ROOT / "logs" / "jira_scanner.log"),
        "max_size_mb": 10,
        "backup_count": 3,
    },
}

# Umgebungsvariablen -> Config-Pfad
_ENV_OVERRIDES = {
    "JIRA_URL": ("server", "url"),
    "JIRA_API_TOKEN": ("api", "token"),
    "JIRA_PROJECT_KEY": ("project", "key"),
    "JIRA_SCRIPT_PATH": ("scripts", "base_path"),
    "JIRA_OUTPUT_PATH": ("excel", "output_path"),
    "JIRA_LOG_LEVEL": ("logging", "level"),
}


def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    """Rekursiv zwei Dicts zusammenführen (override gewinnt)."""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if (
            key in result
            and isinstance(result[key], dict)
            and isinstance(value, dict)
        ):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = copy.deepcopy(value)
    return result


def _apply_env(config: Dict[str, Any]) -> Dict[str, Any]:
    for env_name, (section, key) in _ENV_OVERRIDES.items():
        value = os.environ.get(env_name)
        if value:
            config.setdefault(section, {})[key] = value
    return config


def load_config(path: Optional[str | Path] = None) -> Dict[str, Any]:
    """Lädt config.yaml, mischt Defaults und wendet ENV-Overrides an.

    Fehlt die Datei, werden ausschließlich die Defaults verwendet. Das ist
    wichtig, damit das Tool auch ohne gültige Jira-Konfiguration einen leeren
    Report erzeugen kann.
    """
    config_path = Path(path) if path else DEFAULT_CONFIG_PATH
    file_config: Dict[str, Any] = {}

    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as handle:
            loaded = yaml.safe_load(handle) or {}
            if not isinstance(loaded, dict):
                raise ValueError(f"Ungültige Konfiguration in {config_path}")
            file_config = loaded

    return _apply_env(_deep_merge(DEFAULTS, file_config))


def resolve_path(value: str | Path, base: Path = ROOT) -> Path:
    """Löst relative Pfade relativ zum Projektwurzelverzeichnis auf."""
    path = Path(value).expanduser()
    if not path.is_absolute():
        path = base / path
    return path.resolve()


def setup_logging(config: Dict[str, Any], verbose: bool = False) -> None:
    """Konfiguriert Console- und Datei-Logging (idempotent)."""
    import logging
    from logging.handlers import RotatingFileHandler

    level_name = "DEBUG" if verbose else str(config.get("logging", {}).get("level", "INFO"))
    level = getattr(logging, level_name.upper(), logging.INFO)

    root = logging.getLogger("jira_scanner")
    root.setLevel(level)
    root.propagate = False

    if root.handlers:
        for handler in root.handlers:
            handler.setLevel(level)
        return

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler()
    console.setLevel(level)
    console.setFormatter(formatter)
    root.addHandler(console)

    log_cfg = config.get("logging", {})
    log_file = resolve_path(log_cfg.get("output_file", DEFAULTS["logging"]["output_file"]))
    try:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        max_bytes = int(log_cfg.get("max_size_mb", 10)) * 1024 * 1024
        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=max_bytes,
            backupCount=int(log_cfg.get("backup_count", 3)),
            encoding="utf-8",
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        root.addHandler(file_handler)
    except OSError:
        root.warning("Logdatei %s konnte nicht geöffnet werden", log_file)
