#!/usr/bin/env python3
"""Zentrale Konfigurations- und Logging-Schicht des Jira/Gira Scanners.

Alle Module laden ihre Konfiguration über diese Datei, damit Defaults,
Umgebungsvariablen und Pfadauflösung überall identisch funktionieren.
"""

from __future__ import annotations

import copy
import logging
import os
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Dict, Optional

import yaml

ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG_PATH = ROOT / "config.yaml"

# Optionale .env-Unterstützung; python-dotenv ist optional.
try:  # pragma: no cover - reiner Komfort
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
    "filters": {
        "prefix": "JRA-",
        "min_priority": "medium",
        "statuses": ["open", "in progress", "done"],
        "exclude_keywords": [],
    },
    "scripts": {
        "base_path": str(ROOT / "scripts"),
        "extensions": [".sql"],
        "include_subdirectories": True,
        "exclude": [],
    },
    "performance": {
        "max_results": 1000,
        "timeout": 30,
        "retry_attempts": 3,
    },
    "jira": {
        # Jira-Details werden in Blöcken abgefragt. Maximal 100 Keys pro
        # Request, damit der Jira-Server nicht überlastet wird.
        "batch_size": 100,
        "batch_delay_seconds": 1.0,
    },
    "storage": {
        "inventory_path": str(ROOT / "inventory"),
        "inventory_file": "inventory.json",
        "suggestions_file": "suggestions.json",
        "scan_index": str(ROOT / "inventory" / "scan_index.json"),
        "jira_cache": str(ROOT / "jira_data" / "tasks.json"),
    },
    "excel": {
        "output_path": str(ROOT / "output"),
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
    "confluence": {
        "enabled": False,
        "base_url": "",
        "api_token": "",
        "attachments_path": str(ROOT / "jira_data" / "attachments"),
    },
}

# Umgebungsvariable -> Pfad in der Konfiguration (Punktnotation).
ENV_OVERRIDES = {
    "JIRA_URL": ("server", "url"),
    "JIRA_API_TOKEN": ("api", "token"),
    "JIRA_PROJECT_KEY": ("project", "key"),
    "JIRA_SCRIPT_PATH": ("scripts", "base_path"),
    "JIRA_OUTPUT_PATH": ("excel", "output_path"),
    "JIRA_LOG_LEVEL": ("logging", "level"),
}

# Pfade, die relativ zum Projektwurzelverzeichnis aufgelöst werden.
PATH_FIELDS = (
    ("scripts", "base_path"),
    ("excel", "output_path"),
    ("storage", "inventory_path"),
    ("storage", "scan_index"),
    ("storage", "jira_cache"),
    ("logging", "output_file"),
    ("confluence", "attachments_path"),
)


def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> Dict[str, Any]:
    """Mischt ``override`` rekursiv über ``base`` und gibt ein neues Dict zurück."""
    result = copy.deepcopy(base)
    for key, value in (override or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _set_nested(config: Dict[str, Any], path: tuple, value: Any) -> None:
    current = config
    for part in path[:-1]:
        current = current.setdefault(part, {})
    current[path[-1]] = value


def _apply_env_overrides(config: Dict[str, Any]) -> Dict[str, Any]:
    for env_name, path in ENV_OVERRIDES.items():
        value = os.environ.get(env_name)
        if value:
            _set_nested(config, path, value)
    return config


def resolve_path(value: Any, base: Optional[Path] = None) -> Path:
    """Löst einen Pfad relativ zum Projektwurzelverzeichnis auf."""
    path = Path(str(value)).expanduser()
    if not path.is_absolute():
        path = Path(base or ROOT) / path
    return path


def _resolve_config_paths(config: Dict[str, Any]) -> Dict[str, Any]:
    for section, key in PATH_FIELDS:
        section_data = config.get(section)
        if isinstance(section_data, dict) and section_data.get(key):
            section_data[key] = str(resolve_path(section_data[key]))
    return config


def load_config(path: Optional[str] = None) -> Dict[str, Any]:
    """Lädt config.yaml, ergänzt Defaults und wendet Umgebungsvariablen an."""
    config_path = Path(path).expanduser() if path else DEFAULT_CONFIG_PATH
    file_data: Dict[str, Any] = {}
    if config_path.exists():
        try:
            loaded = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        except (OSError, yaml.YAMLError) as exc:
            raise ValueError(
                f"Konfiguration konnte nicht gelesen werden ({config_path}): {exc}"
            ) from exc
        if loaded is not None and not isinstance(loaded, dict):
            raise ValueError(f"Ungültige config.yaml (Dict erwartet): {config_path}")
        file_data = loaded or {}

    config = _deep_merge(DEFAULTS, file_data)
    config = _apply_env_overrides(config)
    config = _resolve_config_paths(config)
    config["_config_path"] = str(config_path)
    return config


def setup_logging(config: Dict[str, Any], verbose: bool = False) -> logging.Logger:
    """Richtet Console- und Datei-Logging ein (idempotent)."""
    log_cfg = config.get("logging", {})
    level_name = "DEBUG" if verbose else str(log_cfg.get("level", "INFO")).upper()
    level = getattr(logging, level_name, logging.INFO)

    logger = logging.getLogger("jira_scanner")
    logger.setLevel(level)
    logger.propagate = False

    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console = logging.StreamHandler()
    console.setLevel(level)
    console.setFormatter(formatter)
    logger.addHandler(console)

    output_file = log_cfg.get("output_file")
    if output_file:
        log_path = resolve_path(output_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = RotatingFileHandler(
            log_path,
            maxBytes=max(1, int(log_cfg.get("max_size_mb", 10))) * 1024 * 1024,
            backupCount=max(0, int(log_cfg.get("backup_count", 3))),
            encoding="utf-8",
        )
        file_handler.setLevel(level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
