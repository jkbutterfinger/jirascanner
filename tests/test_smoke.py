"""Smoke-Tests für den Jira/Gira Scanner.

Ausführen:
    .venv/bin/python -m pytest
"""

from __future__ import annotations

from pathlib import Path
from unittest.mock import patch

from fastapi.testclient import TestClient
from openpyxl import load_workbook

from main import (
    ExcelExporter,
    JiraClient,
    JiraIssue,
    ScriptScanner,
    adf_to_text,
    build_report,
    empty_report_row,
    format_datetime,
)
from scripts.script_inventory import ScriptAnalyzer

BASE_CONFIG = {
    "server": {"url": "https://jira.example.com"},
    "api": {"token": "test-token", "cors_origins": ["http://localhost:8000"]},
    "project": {"key": "CPU", "jql": ""},
    "filters": {"prefix": "CPU-"},
    "performance": {"max_results": 50, "timeout": 1, "retry_attempts": 1},
    "scripts": {
        "base_path": "scripts",
        "extensions": [".py"],
        "include_subdirectories": True,
        "exclude": [],
    },
}


def test_adf_to_text_and_datetime() -> None:
    adf = {
        "type": "doc",
        "content": [
            {
                "type": "paragraph",
                "content": [{"type": "text", "text": "Hallo Welt"}],
            }
        ],
    }
    assert adf_to_text(adf).strip() == "Hallo Welt"
    assert format_datetime("2024-01-15T09:30:00.000+0000") == "2024-01-15 09:30:00"
    assert format_datetime("") == ""


def test_jira_client_parses_issues_and_filters_prefix() -> None:
    raw = [
        {
            "key": "CPU-1001",
            "fields": {
                "summary": "Monthly Sales Dashboard",
                "status": {"name": "Done"},
                "priority": {"name": "High"},
                "reporter": {"displayName": "Max Mustermann"},
                "assignee": {"displayName": "Anna Meier"},
                "created": "2024-01-15T09:30:00.000+0000",
                "resolutiondate": "2024-02-01T10:00:00.000+0000",
                "issuetype": {"name": "Task"},
                "description": {
                    "type": "doc",
                    "content": [
                        {
                            "type": "paragraph",
                            "content": [{"type": "text", "text": "Details"}],
                        }
                    ],
                },
                "labels": ["report"],
            },
        },
        {"key": "OPS-1", "fields": {"summary": "Fremdprojekt"}},
    ]
    client = JiraClient(BASE_CONFIG)
    with patch.object(client, "_fetch_issues", return_value=raw):
        tasks = client.search_tasks()

    assert len(tasks) == 1
    task = tasks[0]
    assert task.key == "CPU-1001"
    assert task.priority == "High"
    assert task.reporter == "Max Mustermann"
    assert task.assignee == "Anna Meier"
    assert task.description.strip() == "Details"
    assert task.labels == ["report"]


def test_jira_client_without_config_returns_empty() -> None:
    config = {**BASE_CONFIG, "server": {"url": "https://your-jira-server.com"}}
    assert JiraClient(config).search_tasks() == []


def _write_script(directory: Path, name: str, content: str) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / name
    path.write_text(content, encoding="utf-8")
    return path


def test_script_scanner_and_report(tmp_path: Path) -> None:
    scripts_dir = tmp_path / "scripts"
    _write_script(
        scripts_dir,
        "report_monthly_sales.py",
        "import pandas\nimport requests\n\n# Umsetzung für CPU-1001\n",
    )
    _write_script(scripts_dir, "unrelated.py", "print('nichts')\n")

    config = {**BASE_CONFIG, "scripts": {**BASE_CONFIG["scripts"], "base_path": str(scripts_dir)}}
    scanner = ScriptScanner(config)
    tasks = [
        JiraIssue(
            key="CPU-1001",
            summary="Monthly Sales Dashboard",
            status="Done",
            priority="High",
            reporter="Max",
            assignee="Anna",
            resolved="2024-02-01T10:00:00.000+0000",
        )
    ]

    rows = build_report(tasks, scanner)
    assert len(rows) == 1
    row = rows[0]
    assert row["Jira-Nummer"] == "CPU-1001"
    assert row["File-Location"].endswith("report_monthly_sales.py")
    assert "pandas" in row["Verwendete-Technologien"]
    assert "requests" in row["Verwendete-Technologien"]
    assert row["File-Datum"]
    assert row["Jahr/Monat/Tag-Abgeschlossen"] == "2024-02-01 10:00:00"


def test_excel_and_csv_export(tmp_path: Path) -> None:
    config = {
        "excel": {
            "output_path": str(tmp_path / "output"),
            "output_filename": "jira_report_{YYYY}_{MM}_{DD}.xlsx",
            "include_csv": True,
            "sheet_name": "Test",
            "datetime_format": "%Y-%m-%d %H:%M:%S",
        }
    }
    exporter = ExcelExporter(config)
    rows = [empty_report_row("CPU-")]
    rows[0]["Jira-Nummer"] = "CPU-1001"

    path = Path(exporter.export(rows))
    assert path.exists()
    assert path.with_suffix(".csv").exists()

    worksheet = load_workbook(path).active
    assert worksheet.cell(1, 1).value == "Jira-Nummer"
    assert worksheet.cell(1, 10).value == "Priorität"
    assert worksheet.cell(2, 1).value == "CPU-1001"
    assert worksheet.cell(3, 1).value == "GESAMT"


def test_script_analyzer(tmp_path: Path) -> None:
    scripts_dir = tmp_path / "scripts"
    _write_script(
        scripts_dir,
        "demo.py",
        "import pandas\nimport requests\n\n# CPU-2002\n# KPI: umsatz\n",
    )
    analyzer = ScriptAnalyzer(extensions=[".py"], exclude=[])
    inventory = analyzer.scan_directory(scripts_dir)
    assert len(inventory) == 1
    assert inventory[0]["jira_refs"] == ["CPU-2002"]
    assert "pandas" in inventory[0]["technologies"]
    assert "api_calls" in inventory[0]


def test_api_endpoints(tmp_path: Path) -> None:
    import api.suggestion_api as api_module

    api_module.store = api_module.SuggestionStore(tmp_path / "suggestions.json")
    client = TestClient(api_module.app)

    health = client.get("/api/health")
    assert health.status_code == 200
    assert health.json()["status"] == "healthy"

    created = client.post(
        "/api/suggestions",
        json={
            "title": "Test",
            "description": "Beschreibung",
            "task_id": "CPU-1001",
            "severity": "high",
        },
    )
    assert created.status_code == 200
    assert created.json()["suggestion"]["id"].startswith("SUG-")

    listed = client.get("/api/suggestions")
    assert listed.status_code == 200
    assert listed.json()["count"] == 1

    gaps = client.get("/api/gaps", params={"task_keys": "CPU-1001"})
    assert gaps.status_code == 200
    assert gaps.json()["count"] == 1

    root = client.get("/")
    assert root.status_code == 200
    assert "Jira/Gira Scanner" in root.text
