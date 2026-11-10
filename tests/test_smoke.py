"""Smoke-Tests für den Jira/Gira Scanner.

Ausführen:
    .venv/bin/python -m pytest
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from unittest.mock import patch

import yaml
from fastapi.testclient import TestClient
from openpyxl import load_workbook

from main import (
    ExcelExporter,
    JiraClient,
    JiraIssue,
    ScriptScanner,
    adf_to_text,
    build_enriched_rows,
    build_report,
    build_scan_index,
    build_scan_rows,
    derive_path_context,
    empty_report_row,
    format_datetime,
    load_scan_index,
    main,
    save_scan_index,
    task_key_from_directory,
)
from scripts.script_inventory import ScriptAnalyzer

BASE_CONFIG = {
    "server": {"url": "https://jira.example.com"},
    "api": {"token": "test-token", "cors_origins": ["http://localhost:8000"]},
    "project": {"key": "JRA", "jql": ""},
    "filters": {"prefix": "JRA-"},
    "performance": {"max_results": 50, "timeout": 1, "retry_attempts": 1},
    "scripts": {
        "base_path": "scripts",
        "extensions": [".sql"],
        "include_subdirectories": True,
        "exclude": [],
    },
}

EXCEL_CONFIG = {
    "output_path": "output",
    "output_filename": "{JIRA}.xlsx",
    "fallback_filename": "jira_report_{YYYY}_{MM}_{DD}.xlsx",
    "include_csv": False,
    "sheet_name": "Test",
    "datetime_format": "%Y-%m-%d %H:%M:%S",
}


def _write_file(directory: Path, name: str, content: str) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / name
    path.write_text(content, encoding="utf-8")
    return path


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
    assert format_datetime("2024-01-15T09:30:00Z") == "2024-01-15 09:30:00"
    assert format_datetime("2024-01-15T09:30:00+02:00") == "2024-01-15 09:30:00"
    assert format_datetime("") == ""


def test_jira_client_parses_issues_and_filters_prefix() -> None:
    raw = [
        {
            "key": "JRA-1001",
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
    assert task.key == "JRA-1001"
    assert task.priority == "High"
    assert task.reporter == "Max Mustermann"
    assert task.assignee == "Anna Meier"
    assert task.description.strip() == "Details"
    assert task.labels == ["report"]


def test_jira_client_without_config_returns_empty() -> None:
    config = {**BASE_CONFIG, "server": {"url": "https://your-jira-server.com"}}
    assert JiraClient(config).search_tasks() == []


def test_task_key_and_path_context_from_directory(tmp_path: Path) -> None:
    task_dir = tmp_path / "2024" / "IT" / "ACME" / "JRA-1001"
    assert task_key_from_directory(task_dir) == "JRA-1001"
    assert task_key_from_directory(tmp_path / "kein-task") == ""

    context = derive_path_context(task_dir)
    assert context == {"YEAR": "2024", "DEPARTMENT": "IT", "CUSTOMER": "ACME"}


def test_script_scanner_uses_directory_structure(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    _write_file(task_dir, "report.sql", "SELECT * FROM sales;\n")
    _write_file(task_dir, "notes.txt", "JRA-9999\n")  # keine SQL-Datei
    _write_file(mount / "2024" / "IT" / "ACME", "loose.sql", "SELECT 1;\n")

    config = {
        **BASE_CONFIG,
        "scripts": {**BASE_CONFIG["scripts"], "base_path": str(mount)},
    }
    scanner = ScriptScanner(config)
    task = JiraIssue(
        key="JRA-1001",
        summary="Monthly Sales Dashboard",
        status="Done",
        priority="High",
        reporter="Max",
        assignee="Anna",
        resolved="2024-02-01T10:00:00.000+0000",
    )

    matches = scanner.get_task_info(task)
    assert len(matches) == 1
    assert matches[0]["path"].endswith("report.sql")
    assert matches[0]["task_key"] == "JRA-1001"
    assert matches[0]["technologies"] == ["SQL"]
    assert scanner.get_task_directory(task) == task_dir

    rows = build_report([task], scanner)
    assert len(rows) == 1
    assert rows[0]["Jira-Nummer"] == "JRA-1001"
    assert rows[0]["File-Location"].endswith("report.sql")
    assert rows[0]["File-Datum"]
    assert rows[0]["Jahr/Monat/Tag-Abgeschlossen"] == "2024-02-01 10:00:00"


def test_export_per_task_into_task_directory(tmp_path: Path) -> None:
    task_dir = tmp_path / "2024" / "IT" / "ACME" / "JRA-1001"
    task_dir.mkdir(parents=True)
    exporter = ExcelExporter({"excel": EXCEL_CONFIG})

    rows = [empty_report_row("JRA-")]
    rows[0]["Jira-Nummer"] = "JRA-1001"
    rows[0]["File-Location"] = str(task_dir / "report.sql")

    path = Path(
        exporter.export_task(
            "JRA-1001",
            rows,
            task_dir,
            sql_stem="report",
            path_context=derive_path_context(task_dir),
        )
    )
    assert path == task_dir / "JRA-1001.xlsx"
    assert path.exists()

    worksheet = load_workbook(path).active
    assert worksheet.cell(1, 1).value == "Jira-Nummer"
    assert worksheet.cell(1, 10).value == "Priorität"
    assert worksheet.cell(1, 11).value == "Jahr"
    assert worksheet.cell(1, 12).value == "Abteilung"
    assert worksheet.cell(2, 1).value == "JRA-1001"
    assert worksheet.cell(3, 1).value == "GESAMT"


def test_export_filename_placeholders(tmp_path: Path) -> None:
    task_dir = tmp_path / "2024" / "IT" / "ACME" / "JRA-1001"
    task_dir.mkdir(parents=True)
    config = {
        "excel": {
            **EXCEL_CONFIG,
            "output_filename": "{YEAR}_{DEPARTMENT}_{CUSTOMER}_{JIRA}.xlsx",
        }
    }
    exporter = ExcelExporter(config)
    path = Path(
        exporter.export_task(
            "JRA-1001",
            [empty_report_row("JRA-")],
            task_dir,
            sql_stem="report",
            path_context=derive_path_context(task_dir),
        )
    )
    assert path.name == "2024_IT_ACME_JRA-1001.xlsx"


def test_fallback_export_for_empty_report(tmp_path: Path) -> None:
    config = {"excel": {**EXCEL_CONFIG, "output_path": str(tmp_path / "output")}}
    exporter = ExcelExporter(config)
    path = Path(exporter.export([empty_report_row("JRA-")]))
    assert path.exists()
    assert path.name.startswith("jira_report_")
    assert path.suffix == ".xlsx"


def test_script_analyzer_only_sql_and_directory_reference(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-2002"
    _write_file(task_dir, "demo.sql", "SELECT 1;\n-- KPI: umsatz\n")
    _write_file(task_dir, "demo.py", "import pandas\n# JRA-9999\n")

    analyzer = ScriptAnalyzer(
        {
            "scripts": {
                "base_path": str(mount),
                "extensions": [".sql"],
                "include_subdirectories": True,
                "exclude": [],
            }
        }
    )
    inventory = analyzer.analyze_directory(mount)
    assert len(inventory) == 1
    assert inventory[0]["name"] == "demo.sql"
    assert inventory[0]["jira_refs"] == ["JRA-2002"]
    assert inventory[0]["task_key"] == "JRA-2002"
    assert "SQL" in inventory[0]["technologies"]
    assert "api_calls" in inventory[0]


def test_fetch_tasks_by_keys_uses_batches_of_max_100() -> None:
    keys = [f"JRA-{index}" for index in range(1, 206)]
    client = JiraClient(BASE_CONFIG)
    calls: list[str] = []

    def fake_fetch(jql: str) -> list[dict]:
        calls.append(jql)
        found = re.findall(r'"([A-Z]+-\d+)"', jql)
        return [{"key": key, "fields": {"summary": f"Task {key}"}} for key in found]

    with patch.object(client, "_fetch_issues", side_effect=fake_fetch), patch(
        "main.time.sleep"
    ) as sleep_mock:
        tasks = client.fetch_tasks_by_keys(keys, batch_size=100, batch_delay=0.5)

    assert len(tasks) == 205
    assert len(calls) == 3
    assert [len(re.findall(r'"([A-Z]+-\d+)"', jql)) for jql in calls] == [
        100,
        100,
        5,
    ]
    assert sleep_mock.call_count == 2
    sleep_mock.assert_called_with(0.5)

    # Größere Werte werden auf das harte Maximum von 100 begrenzt.
    calls.clear()
    with patch.object(client, "_fetch_issues", side_effect=fake_fetch), patch(
        "main.time.sleep"
    ):
        client.fetch_tasks_by_keys(keys, batch_size=500, batch_delay=0)
    assert len(calls) == 3


def test_build_scan_and_enriched_rows(tmp_path: Path) -> None:
    match = {
        "path": str(tmp_path / "2024" / "IT" / "ACME" / "JRA-1001" / "report.sql"),
        "modified_date": "2024-01-10 12:00:00",
        "technologies": ["SQL"],
    }

    scan_rows = build_scan_rows("JRA-1001", [match])
    assert scan_rows[0]["Status"] == "Pending"
    assert scan_rows[0]["Jira-Auftraggeber"] == ""
    assert scan_rows[0]["File-Datum"] == "2024-01-10"
    assert scan_rows[0]["File-Location"].endswith("report.sql")
    # Lokale Spalten werden bereits im Scan-Modus gefüllt.
    assert scan_rows[0]["Jahr"] == "2024"
    assert scan_rows[0]["Abteilung"] == "IT"

    missing = build_enriched_rows("JRA-1001", None, [match])
    assert missing[0]["Status"] == "Missing"
    assert missing[0]["File-Location"].endswith("report.sql")
    assert missing[0]["Jahr"] == "2024"
    assert missing[0]["Abteilung"] == "IT"

    task = JiraIssue(
        key="JRA-1001",
        summary="Monthly Sales",
        status="Done",
        priority="High",
        reporter="Max",
        assignee="Anna",
        resolved="2024-02-01T10:00:00.000+0000",
    )
    enriched = build_enriched_rows("JRA-1001", task, [match])
    assert enriched[0]["Status"] == "Done"
    assert enriched[0]["Jira-Auftraggeber"] == "Max"
    assert enriched[0]["Jahr/Monat/Tag-Abgeschlossen"] == "2024-02-01 10:00:00"
    assert enriched[0]["Jahr"] == "2024"
    assert enriched[0]["Abteilung"] == "IT"


def test_scan_index_roundtrip(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    _write_file(task_dir, "report.sql", "SELECT 1;\n")
    _write_file(mount / "oracle_sql", "list_tables.sql", "SELECT 2;\n")

    config = {
        **BASE_CONFIG,
        "scripts": {**BASE_CONFIG["scripts"], "base_path": str(mount)},
        "storage": {"scan_index": str(tmp_path / "inventory" / "scan_index.json")},
    }
    scanner = ScriptScanner(config)
    data = build_scan_index(scanner)
    path = save_scan_index(config, data)
    loaded = load_scan_index(config)

    assert path.exists()
    assert loaded is not None
    assert loaded["task_count"] == 1
    assert loaded["file_count"] == 2
    assert len(loaded["files"]) == 2
    assert loaded["tasks"]["JRA-1001"]["files"][0]["path"].endswith("report.sql")


def _write_cli_config(tmp_path: Path, mount: Path) -> Path:
    config = {
        "server": {"url": "https://jira.example.com"},
        "api": {"token": "test-token"},
        "project": {"key": "JRA", "jql": ""},
        "filters": {"prefix": "JRA-"},
        "scripts": {
            "base_path": str(mount),
            "extensions": [".sql"],
            "include_subdirectories": True,
            "exclude": [],
        },
        "storage": {"scan_index": str(tmp_path / "inventory" / "scan_index.json")},
        "excel": {**EXCEL_CONFIG, "output_path": str(tmp_path / "output")},
        "logging": {
            "level": "ERROR",
            "output_file": str(tmp_path / "logs" / "test.log"),
        },
        "jira": {"batch_size": 100, "batch_delay_seconds": 0.0},
    }
    path = tmp_path / "config.yaml"
    path.write_text(yaml.safe_dump(config), encoding="utf-8")
    return path


def test_cli_scan_and_enrich_modes(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    _write_file(task_dir, "report.sql", "SELECT 1;\n")
    config_path = _write_cli_config(tmp_path, mount)

    # Phase 1: nur Verzeichnisstruktur, kein Jira-Zugriff.
    assert main(["--mode", "scan", "--config", str(config_path)]) == 0
    reports = list((tmp_path / "output").glob("jira_report_*.xlsx"))
    assert len(reports) == 1
    report = reports[0]
    assert report.exists()
    assert not (task_dir / "JRA-1001.xlsx").exists()
    assert (tmp_path / "inventory" / "scan_index.json").exists()
    scan_sheet = load_workbook(report).active
    assert scan_sheet.cell(2, 9).value == "Pending"
    assert scan_sheet.cell(2, 2).value in (None, "")
    assert scan_sheet.cell(2, 11).value == "2024"
    assert scan_sheet.cell(2, 12).value == "IT"

    # Phase 2: Jira-Details anreichern (Batch-Parameter werden durchgereicht).
    task = JiraIssue(
        key="JRA-1001",
        summary="Monthly Sales",
        status="Done",
        priority="High",
        reporter="Max",
        assignee="Anna",
        resolved="2024-02-01T10:00:00.000+0000",
    )
    with patch.object(
        JiraClient, "fetch_tasks_by_keys", return_value={"JRA-1001": task}
    ) as fetch_mock:
        assert (
            main(
                [
                    "--mode",
                    "enrich",
                    "--config",
                    str(config_path),
                    "--batch-size",
                    "50",
                    "--batch-delay",
                    "0",
                ]
            )
            == 0
        )

    fetch_mock.assert_called_once()
    assert fetch_mock.call_args.kwargs["batch_size"] == 50
    assert fetch_mock.call_args.kwargs["batch_delay"] == 0.0

    enriched_sheet = load_workbook(report).active
    assert enriched_sheet.cell(2, 2).value == "Max"
    assert enriched_sheet.cell(2, 9).value == "Done"
    assert enriched_sheet.cell(2, 10).value == "High"
    assert enriched_sheet.cell(2, 11).value == "2024"
    assert enriched_sheet.cell(2, 12).value == "IT"


def test_scan_phase_filters_foreign_prefix(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    jra_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    ops_dir = mount / "2024" / "IT" / "ACME" / "OPS-1"
    _write_file(jra_dir, "report.sql", "SELECT 1;\n")
    _write_file(ops_dir, "foreign.sql", "SELECT 2;\n")
    config_path = _write_cli_config(tmp_path, mount)

    assert (
        main(
            [
                "--mode",
                "scan",
                "--prefix",
                "JRA-",
                "--config",
                str(config_path),
            ]
        )
        == 0
    )
    reports = list((tmp_path / "output").glob("jira_report_*.xlsx"))
    assert len(reports) == 1
    worksheet = load_workbook(reports[0]).active
    locations = [
        worksheet.cell(row, 7).value for row in range(2, worksheet.max_row)
    ]
    assert any(str(location).endswith("report.sql") for location in locations)
    assert not any("foreign.sql" in str(location) for location in locations)
    assert not (jra_dir / "JRA-1001.xlsx").exists()
    assert not (ops_dir / "OPS-1.xlsx").exists()


def test_enrich_without_jira_keeps_scan_report(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    _write_file(task_dir, "report.sql", "SELECT 1;\n")
    config_path = _write_cli_config(tmp_path, mount)

    data = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    data["server"]["url"] = "https://your-jira-server.com"
    data["api"]["token"] = ""
    config_path.write_text(yaml.safe_dump(data), encoding="utf-8")

    assert main(["--mode", "scan", "--config", str(config_path)]) == 0
    reports = list((tmp_path / "output").glob("jira_report_*.xlsx"))
    assert len(reports) == 1
    report = reports[0]
    assert load_workbook(report).active.cell(2, 9).value == "Pending"

    # Ohne Jira-Konfiguration darf enrich den Scan-Bericht nicht verändern.
    assert main(["--mode", "enrich", "--config", str(config_path)]) == 0
    assert load_workbook(report).active.cell(2, 9).value == "Pending"


def test_aggregate_scan_includes_files_outside_task_dirs(tmp_path: Path) -> None:
    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    _write_file(task_dir, "report.sql", "SELECT 1;\n")
    oracle_dir = mount / "oracle_sql"
    _write_file(oracle_dir, "list_tables.sql", "SELECT * FROM user_tables;\n")
    _write_file(oracle_dir, "list_views.sql", "SELECT * FROM user_views;\n")
    config_path = _write_cli_config(tmp_path, mount)

    assert main(["--mode", "scan", "--config", str(config_path)]) == 0

    reports = list((tmp_path / "output").glob("jira_report_*.xlsx"))
    assert len(reports) == 1
    worksheet = load_workbook(reports[0]).active
    rows = [
        {
            "task": worksheet.cell(row, 1).value,
            "location": worksheet.cell(row, 7).value,
            "year": worksheet.cell(row, 11).value,
            "department": worksheet.cell(row, 12).value,
        }
        for row in range(2, worksheet.max_row)
    ]
    assert len(rows) == 3
    assert {row["location"] for row in rows} == {
        str(task_dir / "report.sql"),
        str(oracle_dir / "list_tables.sql"),
        str(oracle_dir / "list_views.sql"),
    }
    task_row = next(row for row in rows if row["location"].endswith("report.sql"))
    assert task_row["task"] == "JRA-1001"
    assert task_row["year"] == "2024"
    assert task_row["department"] == "IT"
    orphan_rows = [row for row in rows if "oracle_sql" in row["location"]]
    assert all(row["task"] in (None, "") for row in orphan_rows)

    index = json.loads(
        (tmp_path / "inventory" / "scan_index.json").read_text(encoding="utf-8")
    )
    assert index["task_count"] == 1
    assert index["file_count"] == 3
    assert len(index["files"]) == 3


def test_api_endpoints(tmp_path: Path) -> None:
    import api.suggestion_api as api_module

    mount = tmp_path / "mount"
    task_dir = mount / "2024" / "IT" / "ACME" / "JRA-1001"
    script = _write_file(task_dir, "report.sql", "SELECT * FROM sales;\n")

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
            "task_id": "JRA-1001",
            "severity": "high",
        },
    )
    assert created.status_code == 200
    assert created.json()["suggestion"]["id"].startswith("SUG-")

    listed = client.get("/api/suggestions")
    assert listed.status_code == 200
    assert listed.json()["count"] == 1

    with patch.object(api_module, "_script_path", return_value=mount):
        scripts = client.get("/api/scripts")
        assert scripts.status_code == 200
        assert scripts.json()["count"] == 1
        assert scripts.json()["scripts"][0]["task_key"] == "JRA-1001"

        analysis = client.post("/api/scripts/analyze", params={"path": str(script)})
        assert analysis.status_code == 200
        assert analysis.json()["script"]["jira_refs"] == ["JRA-1001"]

        gaps = client.get("/api/gaps", params={"task_keys": "JRA-1001"})
        assert gaps.status_code == 200
        assert gaps.json()["count"] == 1
        assert gaps.json()["gaps"][0]["has_implementation"] is True

        match = client.get("/api/match", params={"task_key": "JRA-1001"})
        assert match.status_code == 200
        assert match.json()["match_score"] == 1.0

    root = client.get("/")
    assert root.status_code == 200
    assert "Jira/Gira Scanner" in root.text
