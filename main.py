#!/usr/bin/env python3
"""Jira/Gira Scanner with Excel Export.

Scannt Jira-Tasks mit einem konfigurierbaren Prefix (Standard: ``CPU-``),
verknüpft sie mit lokalen Skripten und erzeugt einen Excel-Bericht
(optional zusätzlich als CSV).

Die kanonische Implementierung liegt in dieser Datei. Der historische
Einstiegspunkt ``scanner/main_scanner.py`` delegiert nur noch hierher.
"""

from __future__ import annotations

import argparse
import csv
import logging
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import ROOT, load_config, resolve_path, setup_logging

LOGGER = logging.getLogger("jira_scanner")

# Spalten des Excel-Berichts: (Anzeigename, Breite, Umbruch)
COLUMNS: List[tuple[str, int, bool]] = [
    ("Jira-Nummer", 18, False),
    ("Jira-Auftraggeber", 24, False),
    ("Entwickler", 24, False),
    ("Jahr/Monat/Tag-Abgeschlossen", 26, False),
    ("Zusammenfassung-Anforderung", 55, True),
    ("File-Datum", 16, False),
    ("File-Location", 48, True),
    ("Verwendete-Technologien", 45, True),
    ("Status", 14, False),
    ("Priorität", 12, False),
]

PRIORITY_ALIASES = {
    "highest": "Blocker",
    "blocker": "Blocker",
    "critical": "Critical",
    "high": "High",
    "medium": "Medium",
    "low": "Low",
    "lowest": "Low",
}

STATUS_FILLS = {
    "open": "FFEBEE",
    "to do": "FFEBEE",
    "in progress": "FFF3E0",
    "in review": "E3F2FD",
    "done": "E8F5E9",
    "closed": "E8F5E9",
    "resolved": "E8F5E9",
    "missing": "FCE4EC",
}

JIRA_FIELDS = [
    "summary",
    "status",
    "priority",
    "reporter",
    "assignee",
    "created",
    "resolutiondate",
    "issuetype",
    "description",
    "labels",
]


class JiraError(RuntimeError):
    """Fehler bei der Kommunikation mit der Jira-API."""


def adf_to_text(node: Any) -> str:
    """Wandelt Jira-ADF (API v3) rekursiv in lesbaren Text um."""
    if node is None:
        return ""
    if isinstance(node, str):
        return node
    if isinstance(node, list):
        return "".join(adf_to_text(item) for item in node)
    if isinstance(node, dict):
        node_type = node.get("type")
        if node_type == "text":
            return str(node.get("text", ""))
        if node_type == "hardBreak":
            return "\n"
        text = adf_to_text(node.get("content", []))
        if node_type in {"paragraph", "heading", "listItem", "tableRow", "codeBlock"}:
            text += "\n"
        return text
    return str(node)


def format_datetime(value: Any, fmt: str = "%Y-%m-%d %H:%M:%S") -> str:
    """Formatiert ISO-8601-Strings aus Jira; lässt Unbekanntes unverändert."""
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime(fmt)
    text = str(value).strip()
    if not text:
        return ""
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed.strftime(fmt)
    except ValueError:
        return text


def format_date(value: Any) -> str:
    """Formatiert einen Zeitstempel als reines Datum (YYYY-MM-DD)."""
    formatted = format_datetime(value, "%Y-%m-%d")
    return formatted


@dataclass
class JiraIssue:
    """Jira Issue/Task Model."""

    key: str = ""
    issue_type: str = "Task"
    status: str = ""
    priority: str = "Medium"
    summary: str = ""
    description: str = ""
    reporter: str = ""
    assignee: str = ""
    created: str = ""
    resolved: str = ""
    labels: List[str] = field(default_factory=list)


class JiraClient:
    """Jira/Gira REST-API-Client (API v3)."""

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        server = self.config.get("server", {})
        api = self.config.get("api", {})
        project = self.config.get("project", {})
        performance = self.config.get("performance", {})

        self.url = str(server.get("url", "")).rstrip("/")
        self.token = str(api.get("token", "") or "")
        self.project_key = str(project.get("key", "CPU") or "CPU")
        self.custom_jql = str(project.get("jql", "") or "")
        self.timeout = max(1, int(performance.get("timeout", 30)))
        self.retry_attempts = max(1, int(performance.get("retry_attempts", 3)))
        self.max_results = max(1, int(performance.get("max_results", 1000)))

        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})
        if self.token:
            self.session.headers["Authorization"] = f"Bearer {self.token}"

    # ------------------------------------------------------------------ HTTP
    def is_configured(self) -> bool:
        """True, wenn URL und Token gesetzt und keine Platzhalter sind."""
        if not self.url or "your-jira-server" in self.url:
            return False
        if not self.token:
            return False
        return True

    def _get_json(self, path: str, params: Dict[str, Any]) -> Dict[str, Any]:
        last_error: Optional[Exception] = None
        for attempt in range(1, self.retry_attempts + 1):
            try:
                response = self.session.get(
                    f"{self.url}{path}", params=params, timeout=self.timeout
                )
                if response.status_code in (401, 403):
                    raise JiraError(
                        f"Authentifizierung fehlgeschlagen (HTTP {response.status_code}). "
                        "Bitte API-Token prüfen."
                    )
                if response.status_code >= 400:
                    raise JiraError(
                        f"HTTP {response.status_code} bei {path}: {response.text[:200]}"
                    )
                return response.json()
            except JiraError:
                raise
            except (requests.RequestException, ValueError) as exc:
                last_error = exc
                LOGGER.warning(
                    "Jira-Request fehlgeschlagen (Versuch %s/%s): %s",
                    attempt,
                    self.retry_attempts,
                    exc,
                )
        raise JiraError(f"Jira-Request endgültig fehlgeschlagen: {last_error}")

    # --------------------------------------------------------------- Abfragen
    def _fetch_issues(self, jql: str) -> List[Dict[str, Any]]:
        """Holt Issues mit Pagination.

        Bevorzugt den aktuellen Endpoint ``/rest/api/3/search/jql``
        (``nextPageToken``) und fällt bei älteren Servern auf
        ``/rest/api/3/search`` (``startAt``) zurück.
        """
        fields = ",".join(JIRA_FIELDS)
        try:
            return self._fetch_with_token(jql, fields)
        except JiraError as exc:
            LOGGER.info("Neuer Such-Endpoint nicht verfügbar (%s) – Fallback.", exc)
            return self._fetch_with_start_at(jql, fields)

    def _fetch_with_token(self, jql: str, fields: str) -> List[Dict[str, Any]]:
        issues: List[Dict[str, Any]] = []
        next_token: Optional[str] = None
        while len(issues) < self.max_results:
            params: Dict[str, Any] = {
                "jql": jql,
                "maxResults": min(100, self.max_results - len(issues)),
                "fields": fields,
            }
            if next_token:
                params["nextPageToken"] = next_token
            data = self._get_json("/rest/api/3/search/jql", params)
            batch = data.get("issues", []) or []
            issues.extend(batch)
            next_token = data.get("nextPageToken")
            if not batch or not next_token:
                break
        return issues

    def _fetch_with_start_at(self, jql: str, fields: str) -> List[Dict[str, Any]]:
        issues: List[Dict[str, Any]] = []
        start_at = 0
        while len(issues) < self.max_results:
            params = {
                "jql": jql,
                "startAt": start_at,
                "maxResults": min(100, self.max_results - len(issues)),
                "fields": fields,
            }
            data = self._get_json("/rest/api/3/search", params)
            batch = data.get("issues", []) or []
            issues.extend(batch)
            start_at += len(batch)
            total = data.get("total")
            if not batch or (total is not None and start_at >= int(total)):
                break
        return issues

    def search_tasks(self, prefix: Optional[str] = None) -> List[JiraIssue]:
        """Sucht Tasks im Projekt und filtert optional nach Key-Prefix."""
        prefix = prefix or str(self.config.get("filters", {}).get("prefix", "CPU-"))
        if not self.is_configured():
            LOGGER.warning(
                "Jira ist nicht konfiguriert (server.url/api.token). "
                "Es wird ein leerer Bericht erzeugt."
            )
            return []

        jql = self.custom_jql or f'project = "{self.project_key}" ORDER BY created DESC'
        try:
            raw_issues = self._fetch_issues(jql)
        except JiraError as exc:
            LOGGER.error("Jira-Abfrage fehlgeschlagen: %s", exc)
            return []

        tasks: List[JiraIssue] = []
        for issue in raw_issues:
            key = str(issue.get("key", ""))
            if prefix and not key.upper().startswith(prefix.upper()):
                continue
            fields = issue.get("fields", {}) or {}
            task = JiraIssue(
                key=key,
                issue_type=(fields.get("issuetype") or {}).get("name", "Task"),
                status=(fields.get("status") or {}).get("name", ""),
                priority=self._get_priority(
                    (fields.get("priority") or {}).get("name", "")
                ),
                summary=fields.get("summary", "") or "",
                description=adf_to_text(fields.get("description")).strip(),
                reporter=(fields.get("reporter") or {}).get("displayName", ""),
                assignee=(fields.get("assignee") or {}).get("displayName", ""),
                created=fields.get("created", "") or "",
                resolved=fields.get("resolutiondate", "") or "",
                labels=list(fields.get("labels", []) or []),
            )
            tasks.append(task)

        LOGGER.info("%s Task(s) mit Prefix %s gefunden", len(tasks), prefix)
        return tasks

    @staticmethod
    def _get_priority(priority_name: str) -> str:
        name = (priority_name or "").strip()
        if not name:
            return "Medium"
        return PRIORITY_ALIASES.get(name.lower(), name)


class ScriptScanner:
    """Scannt Skripte und extrahiert Implementierungsdetails.

    Die Dateien werden genau einmal gelesen und über die gefundenen
    Jira-Referenzen indexiert. Dadurch entfällt das frühere
    O(Tasks × Scripts)-Verhalten.
    """

    IMPORT_PATTERNS = [
        re.compile(r"^\s*import\s+([A-Za-z_][\w.]*)", re.MULTILINE),
        re.compile(r"^\s*from\s+([A-Za-z_][\w.]*)\s+import\b", re.MULTILINE),
        re.compile(r"""require\(\s*["']([^"']+)["']\s*\)"""),
    ]
    JIRA_REF_PATTERN = re.compile(r"\b([A-Z][A-Z0-9]+-\d+)\b")
    KPI_PATTERN = re.compile(r"\b(kpi|dashboard|report|metric)\w*\b", re.IGNORECASE)

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        scripts_cfg = self.config.get("scripts", {})
        self.base_path = resolve_path(scripts_cfg.get("base_path", ROOT / "scripts"))
        self.extensions = list(scripts_cfg.get("extensions", [".py", ".sh", ".sql"]))
        self.include_subdirectories = bool(
            scripts_cfg.get("include_subdirectories", True)
        )
        self.exclude = {str(name) for name in scripts_cfg.get("exclude", [])}
        self._index: Optional[Dict[str, List[Dict[str, Any]]]] = None
        self._scripts: List[Path] = []

    # ---------------------------------------------------------------- Dateien
    def find_scripts(self) -> List[Path]:
        """Findet alle relevanten Skriptdateien (ohne Duplikate)."""
        if not self.base_path.exists():
            LOGGER.warning("Scripts-Pfad nicht gefunden: %s", self.base_path)
            return []

        found: List[Path] = []
        for extension in self.extensions:
            pattern = f"*{extension}"
            iterator = (
                self.base_path.rglob(pattern)
                if self.include_subdirectories
                else self.base_path.glob(pattern)
            )
            for path in iterator:
                if not path.is_file():
                    continue
                if path.name in self.exclude:
                    continue
                if path not in found:
                    found.append(path)
        self._scripts = sorted(found)
        return self._scripts

    def extract_script_info(self, script_path: Path) -> Dict[str, Any]:
        """Extrahiert Metadaten, Jira-Referenzen und Technologien."""
        info: Dict[str, Any] = {
            "path": str(script_path),
            "name": script_path.name,
            "size_bytes": 0,
            "lines": 0,
            "modified_date": "",
            "jira_refs": [],
            "technologies": [],
            "kpi_reports": [],
            "error": "",
        }

        try:
            stat = script_path.stat()
            info["size_bytes"] = stat.st_size
            info["modified_date"] = datetime.fromtimestamp(stat.st_mtime).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            content = script_path.read_text(encoding="utf-8", errors="ignore")
        except OSError as exc:
            info["error"] = str(exc)
            return info

        info["lines"] = len(content.splitlines())
        info["jira_refs"] = sorted(set(self.JIRA_REF_PATTERN.findall(content)))
        info["technologies"] = self._detect_technologies(content)
        info["kpi_reports"] = sorted(
            {match.lower() for match in self.KPI_PATTERN.findall(content)}
        )
        return info

    def _detect_technologies(self, content: str) -> List[str]:
        names: set[str] = set()
        for pattern in self.IMPORT_PATTERNS:
            for match in pattern.findall(content):
                root = match.split(".")[0].strip()
                if root and not root.startswith((".", "/")):
                    names.add(root)

        stdlib = getattr(sys, "stdlib_module_names", frozenset())
        builtin = set(sys.builtin_module_names)
        return sorted(name for name in names if name not in stdlib and name not in builtin)

    # ---------------------------------------------------------------- Matching
    def build_index(self) -> Dict[str, List[Dict[str, Any]]]:
        """Liest alle Skripte einmal und indexiert sie nach Jira-Referenz."""
        index: Dict[str, List[Dict[str, Any]]] = {}
        for script in self.find_scripts():
            info = self.extract_script_info(script)
            if info.get("error"):
                LOGGER.warning("Skript konnte nicht gelesen werden: %s", info["error"])
            for reference in info.get("jira_refs", []):
                index.setdefault(reference.upper(), []).append(info)
        self._index = index
        LOGGER.info("%s Skript(e) indexiert", len(self._scripts))
        return index

    def get_task_info(self, task: JiraIssue) -> List[Dict[str, Any]]:
        """Liefert alle zum Task passenden Skript-Informationen."""
        if self._index is None:
            self.build_index()
        assert self._index is not None
        return self._index.get(task.key.upper(), [])


class ExcelExporter:
    """Erzeugt den Excel-Bericht (alle 10 Spalten nebeneinander)."""

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        excel_cfg = self.config.get("excel", {})
        self.output_dir = resolve_path(excel_cfg.get("output_path", ROOT / "output"))
        self.filename_pattern = str(
            excel_cfg.get("output_filename", "jira_report_{YYYY}_{MM}_{DD}.xlsx")
        )
        self.datetime_format = str(
            excel_cfg.get("datetime_format", "%Y-%m-%d %H:%M:%S")
        )
        self.include_csv = bool(excel_cfg.get("include_csv", True))
        self.sheet_name = str(excel_cfg.get("sheet_name", "Jira-Task-Report"))
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _filename(self) -> str:
        now = datetime.now()
        return (
            self.filename_pattern.replace("{YYYY}", now.strftime("%Y"))
            .replace("{MM}", now.strftime("%m"))
            .replace("{DD}", now.strftime("%d"))
        )

    def export(self, data: List[Dict[str, Any]]) -> str:
        """Schreibt die Daten als XLSX und optional als CSV."""
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = self.sheet_name[:31]

        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(
            start_color="4472C4", end_color="4472C4", fill_type="solid"
        )
        header_alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        thin = Side(style="thin", color="B0B0B0")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        data_font = Font(size=10)

        for column_index, (title, width, _wrap) in enumerate(COLUMNS, start=1):
            cell = worksheet.cell(row=1, column=column_index, value=title)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = border
            worksheet.column_dimensions[get_column_letter(column_index)].width = width

        for row_index, row in enumerate(data, start=2):
            for column_index, (title, _width, wrap) in enumerate(COLUMNS, start=1):
                value = row.get(title, "")
                cell = worksheet.cell(row=row_index, column=column_index, value=value)
                cell.font = data_font
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="left" if wrap else "center",
                    vertical="center",
                    wrap_text=wrap,
                )
                if title == "Status":
                    fill_color = STATUS_FILLS.get(str(value).strip().lower())
                    if fill_color:
                        cell.fill = PatternFill(
                            start_color=fill_color,
                            end_color=fill_color,
                            fill_type="solid",
                        )
            worksheet.row_dimensions[row_index].height = 30 if any(
                row.get(title) for title, _w, wrap in COLUMNS if wrap
            ) else 18

        total_row = len(data) + 2
        worksheet.cell(row=total_row, column=1, value="GESAMT").font = Font(bold=True)
        worksheet.cell(row=total_row, column=10, value=len(data)).font = Font(bold=True)
        for column_index in range(1, len(COLUMNS) + 1):
            worksheet.cell(row=total_row, column=column_index).border = border

        worksheet.freeze_panes = "A2"
        if data:
            worksheet.auto_filter.ref = (
                f"A1:{get_column_letter(len(COLUMNS))}{len(data) + 1}"
            )

        output_path = self.output_dir / self._filename()
        workbook.save(output_path)
        LOGGER.info("Excel-Bericht gespeichert: %s", output_path)

        if self.include_csv:
            csv_path = output_path.with_suffix(".csv")
            self.export_csv(data, csv_path)
            LOGGER.info("CSV-Bericht gespeichert: %s", csv_path)

        return str(output_path)

    def export_csv(self, data: List[Dict[str, Any]], path: Path) -> str:
        """Schreibt denselben Inhalt als UTF-8-CSV (Semikolon-getrennt)."""
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=[title for title, _width, _wrap in COLUMNS],
                delimiter=";",
                extrasaction="ignore",
            )
            writer.writeheader()
            for row in data:
                writer.writerow(
                    {
                        title: row.get(title, "")
                        for title, _width, _wrap in COLUMNS
                    }
                )
        return str(path)

    def export_simple(self, data: List[Dict[str, Any]]) -> str:
        """Gibt die Tabelle für Debugging/--dry-run auf der Konsole aus."""
        print("-" * 120)
        for row in data:
            print(
                " | ".join(
                    str(row.get(title, ""))[:30] for title, _w, _wrap in COLUMNS
                )
            )
        print("-" * 120)
        return ""


def build_report(tasks: List[JiraIssue], scanner: ScriptScanner) -> List[Dict[str, Any]]:
    """Baut die Excel-Zeilen aus Jira-Tasks und gefundenen Skripten."""
    rows: List[Dict[str, Any]] = []

    for task in tasks:
        matches = scanner.get_task_info(task)
        paths = [match.get("path", "") for match in matches if match.get("path")]
        technologies: List[str] = []
        for match in matches:
            for technology in match.get("technologies", []):
                if technology not in technologies:
                    technologies.append(technology)

        file_dates = [
            format_date(match.get("modified_date"))
            for match in matches
            if match.get("modified_date")
        ]

        rows.append(
            {
                "Jira-Nummer": task.key,
                "Jira-Auftraggeber": task.reporter,
                "Entwickler": task.assignee,
                "Jahr/Monat/Tag-Abgeschlossen": format_datetime(
                    task.resolved
                )
                or "Pending",
                "Zusammenfassung-Anforderung": task.summary,
                "File-Datum": file_dates[0] if file_dates else "",
                "File-Location": "\n".join(paths) if paths else "Not found",
                "Verwendete-Technologien": ", ".join(technologies[:8]),
                "Status": task.status or ("Missing" if not matches else "Open"),
                "Priorität": task.priority,
            }
        )

    return rows


def empty_report_row(prefix: str = "CPU-") -> Dict[str, Any]:
    """Platzhalterzeile, wenn keine Tasks gefunden wurden."""
    return {
        "Jira-Nummer": prefix,
        "Jira-Auftraggeber": "",
        "Entwickler": "",
        "Jahr/Monat/Tag-Abgeschlossen": "Pending",
        "Zusammenfassung-Anforderung": "Keine Tasks gefunden",
        "File-Datum": "",
        "File-Location": "Not found",
        "Verwendete-Technologien": "",
        "Status": "Open",
        "Priorität": "Medium",
    }


def main(argv: Optional[List[str]] = None) -> int:
    """CLI-Einstiegspunkt."""
    parser = argparse.ArgumentParser(
        description="Jira/Gira Scanner – erzeugt einen Excel-Bericht für Tasks mit Prefix."
    )
    parser.add_argument("--config", help="Pfad zur config.yaml")
    parser.add_argument(
        "--prefix",
        help="Key-Prefix (überschreibt filters.prefix, Standard: CPU-)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Nur auf der Konsole ausgeben, keine Datei schreiben",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG-Logging")
    args = parser.parse_args(argv)

    config = load_config(args.config)
    setup_logging(config, verbose=args.verbose)
    prefix = args.prefix or str(config.get("filters", {}).get("prefix", "CPU-"))

    LOGGER.info("=" * 70)
    LOGGER.info("Jira/Gira Scanner gestartet (Prefix: %s)", prefix)
    LOGGER.info("=" * 70)

    client = JiraClient(config)
    scanner = ScriptScanner(config)
    exporter = ExcelExporter(config)

    tasks = client.search_tasks(prefix=prefix)
    if tasks:
        rows = build_report(tasks, scanner)
    else:
        LOGGER.warning("Keine Tasks gefunden – es wird ein leerer Bericht erzeugt.")
        rows = [empty_report_row(prefix)]

    if args.dry_run:
        exporter.export_simple(rows)
        return 0

    try:
        output_path = exporter.export(rows)
    except OSError as exc:
        LOGGER.error("Bericht konnte nicht geschrieben werden: %s", exc)
        return 1

    LOGGER.info("Fertig: %s (%s Zeile(n))", output_path, len(rows))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
