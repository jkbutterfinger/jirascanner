#!/usr/bin/env python3
"""Jira/Gira Scanner with per-task Excel export.

Der Scanner arbeitet in zwei Phasen:

1. ``--mode scan``   liest ausschließlich die Verzeichnisstruktur und erzeugt
   pro Task einen Excel-Bericht mit den lokal verfügbaren Informationen.
2. ``--mode enrich`` reichert die bestehenden Berichte mit Jira-Details an.
   Die Jira-Keys werden in Blöcken von maximal 100 abgefragt, mit einer
   konfigurierbaren Pause zwischen den Blöcken.

``--mode all`` (Default) führt beide Phasen nacheinander aus.

Die Jira-Referenz wird nicht aus dem Dateiinhalt, sondern aus der
Verzeichnisstruktur abgeleitet::

    <base_path>/<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<name>.sql

Für jeden Jira-Task wird ein Excel-Bericht direkt im Task-Verzeichnis
abgelegt::

    <base_path>/<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<Jira-Task>.xlsx
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import load_config, resolve_path, setup_logging

LOGGER = logging.getLogger("jira_scanner")

TASK_KEY_PATTERN = re.compile(r"^[A-Z][A-Z0-9]+-\d+$")
DEFAULT_FIELDS = [
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

# Spalten des Excel-Berichts: (Anzeigename, Breite, Umbruch)
COLUMNS: List[Tuple[str, int, bool]] = [
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
    ("Jahr", 8, False),
    ("Abteilung", 20, False),
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
    "pending": "FFFDE7",
}


# ---------------------------------------------------------------------------
# Allgemeine Helfer
# ---------------------------------------------------------------------------
def _adf_collect(node: Any, parts: List[str]) -> None:
    if node is None:
        return
    if isinstance(node, str):
        parts.append(node)
        return
    if isinstance(node, list):
        for item in node:
            _adf_collect(item, parts)
        return
    if not isinstance(node, dict):
        return

    node_type = node.get("type")
    if node_type == "text":
        parts.append(str(node.get("text", "")))
        return
    if node_type == "hardBreak":
        parts.append("\n")
        return
    if node_type in ("paragraph", "heading", "listItem", "blockquote", "codeBlock"):
        if parts and not parts[-1].endswith("\n"):
            parts.append("\n")
    _adf_collect(node.get("content"), parts)
    if node_type in ("paragraph", "heading", "listItem", "blockquote", "codeBlock"):
        parts.append("\n")


def adf_to_text(value: Any) -> str:
    """Wandelt Atlassian Document Format (ADF) rekursiv in Text um."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    parts: List[str] = []
    _adf_collect(value, parts)
    text = "".join(parts)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def format_datetime(value: Any, fmt: Optional[str] = None) -> str:
    """Formatiert ISO-/Jira-Zeitstempel; gibt bei Fehlern den Rohwert zurück."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        parsed = value
    else:
        text = str(value).strip()
        if not text:
            return ""
        normalized = text.replace("Z", "+00:00")
        normalized = re.sub(r"([+-]\d{2})(\d{2})$", r"\1:\2", normalized)
        parsed = None
        try:
            parsed = datetime.fromisoformat(normalized)
        except ValueError:
            for pattern in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
                try:
                    parsed = datetime.strptime(text, pattern)
                    break
                except ValueError:
                    continue
            if parsed is None:
                return text
    return parsed.strftime(fmt or "%Y-%m-%d %H:%M:%S")


def format_date(value: Any) -> str:
    """Formatiert einen Zeitstempel als reines Datum (YYYY-MM-DD)."""
    return format_datetime(value, "%Y-%m-%d")


def task_key_from_directory(directory: Path) -> str:
    """Leitet den Jira-Key aus dem Verzeichnisnamen ab."""
    candidate = Path(directory).name.strip().upper()
    return candidate if TASK_KEY_PATTERN.match(candidate) else ""


def derive_path_context(task_directory: Path) -> Dict[str, str]:
    """Liefert Jahr/Abteilung/Auftraggeber aus der Verzeichnisstruktur.

    Erwartet ``<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>``. Fehlende
    Ebenen werden als leerer String zurückgegeben.
    """
    parents = list(Path(task_directory).parents)
    return {
        "CUSTOMER": Path(task_directory).parent.name if parents else "",
        "DEPARTMENT": parents[1].name if len(parents) >= 2 else "",
        "YEAR": parents[2].name if len(parents) >= 3 else "",
    }


def local_path_columns(match: Dict[str, Any]) -> Dict[str, str]:
    """Liefert Jahr und Abteilung aus dem Dateipfad (lokale Spalten)."""
    path = str(match.get("path", "") or "")
    if not path:
        return {"Jahr": "", "Abteilung": ""}
    context = derive_path_context(Path(path).parent)
    if not re.fullmatch(r"\d{4}", context["YEAR"]):
        return {"Jahr": "", "Abteilung": ""}
    return {"Jahr": context["YEAR"], "Abteilung": context["DEPARTMENT"]}


# ---------------------------------------------------------------------------
# Jira
# ---------------------------------------------------------------------------
class JiraError(RuntimeError):
    """Fehler beim Zugriff auf die Jira-REST-API."""

    def __init__(self, message: str, status: Optional[int] = None):
        super().__init__(message)
        self.status = status


@dataclass
class JiraIssue:
    """Jira Issue/Task-Modell."""

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

    def to_dict(self) -> Dict[str, Any]:
        return {
            "key": self.key,
            "issue_type": self.issue_type,
            "status": self.status,
            "priority": self.priority,
            "summary": self.summary,
            "description": self.description,
            "reporter": self.reporter,
            "assignee": self.assignee,
            "created": self.created,
            "resolved": self.resolved,
            "labels": list(self.labels),
        }


class JiraClient:
    """Kleiner Jira-REST-API-v3-Client mit Retry, Pagination und Batch-Abfrage."""

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        server = self.config.get("server", {})
        api = self.config.get("api", {})
        performance = self.config.get("performance", {})
        jira_cfg = self.config.get("jira", {})

        raw_url = str(server.get("url", "") or "").strip().rstrip("/")
        if raw_url and not raw_url.startswith(("http://", "https://")):
            protocol = str(server.get("protocol", "https") or "https")
            raw_url = f"{protocol}://{raw_url}"
        if raw_url:
            parsed = urlparse(raw_url)
            port = int(server.get("port", 443) or 0)
            if parsed.port is None and port and port not in (80, 443):
                raw_url = f"{raw_url}:{port}"

        self.base_url = raw_url
        self.token = str(api.get("token", "") or "")
        self.auth_type = str(api.get("auth_type", "token") or "token")
        self.verify_ssl = bool(server.get("verify_ssl", True))
        self.timeout = max(1, int(performance.get("timeout", 30)))
        self.retry_attempts = max(1, int(performance.get("retry_attempts", 3)))
        self.max_results = max(1, int(performance.get("max_results", 1000)))

        # Maximal 100 Keys pro Request, damit der Jira-Server geschont wird.
        self.batch_size = max(1, min(100, int(jira_cfg.get("batch_size", 100))))
        self.batch_delay = max(0.0, float(jira_cfg.get("batch_delay_seconds", 1.0)))

    def is_configured(self) -> bool:
        """Prüft, ob URL und Token gesetzt und plausibel sind."""
        if not self.base_url or not self.token:
            return False
        if "your-jira-server" in self.base_url.lower():
            return False
        return True

    def _headers(self) -> Dict[str, str]:
        if self.auth_type == "basic":
            return {"Accept": "application/json"}
        return {
            "Accept": "application/json",
            "Authorization": f"Bearer {self.token}",
        }

    def _get_json(
        self, path: str, params: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        url = f"{self.base_url}{path}"
        last_error: Optional[JiraError] = None
        for attempt in range(1, self.retry_attempts + 1):
            try:
                response = requests.get(
                    url,
                    headers=self._headers(),
                    params=params,
                    timeout=self.timeout,
                    verify=self.verify_ssl,
                )
            except requests.RequestException as exc:
                last_error = JiraError(f"Jira nicht erreichbar: {exc}")
            else:
                status = response.status_code
                if status in (401, 403):
                    raise JiraError(
                        f"Jira-Zugriff verweigert (HTTP {status}) – Token/Rechte prüfen.",
                        status,
                    )
                if status >= 500:
                    last_error = JiraError(f"Jira-Serverfehler (HTTP {status})", status)
                elif status >= 400:
                    raise JiraError(
                        f"Jira-Request fehlgeschlagen (HTTP {status}): {response.text[:200]}",
                        status,
                    )
                else:
                    try:
                        return response.json()
                    except ValueError as exc:
                        raise JiraError("Ungültige JSON-Antwort von Jira.") from exc

            if attempt < self.retry_attempts:
                time.sleep(min(2 ** attempt, 10))
        raise last_error or JiraError("Unbekannter Jira-Fehler.")

    def _fetch_issues(
        self, jql: str, fields: Optional[List[str]] = None
    ) -> List[Dict[str, Any]]:
        """Lädt Issues mit nextPageToken (neu) oder startAt (ältere Server)."""
        selected = fields or DEFAULT_FIELDS
        try:
            return self._fetch_issues_token(jql, selected)
        except JiraError as exc:
            if exc.status in (400, 404, 405, 410):
                LOGGER.debug(
                    "nextPageToken-Endpunkt nicht verfügbar (HTTP %s) – "
                    "Fallback auf startAt.",
                    exc.status,
                )
                return self._fetch_issues_start_at(jql, selected)
            raise

    def _fetch_issues_token(
        self, jql: str, fields: List[str]
    ) -> List[Dict[str, Any]]:
        issues: List[Dict[str, Any]] = []
        next_token: Optional[str] = None
        while len(issues) < self.max_results:
            params: Dict[str, Any] = {
                "jql": jql,
                "maxResults": min(100, self.max_results - len(issues)),
                "fields": ",".join(fields),
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

    def _fetch_issues_start_at(
        self, jql: str, fields: List[str]
    ) -> List[Dict[str, Any]]:
        issues: List[Dict[str, Any]] = []
        start_at = 0
        while len(issues) < self.max_results:
            params = {
                "jql": jql,
                "startAt": start_at,
                "maxResults": min(100, self.max_results - len(issues)),
                "fields": ",".join(fields),
            }
            data = self._get_json("/rest/api/3/search", params)
            batch = data.get("issues", []) or []
            issues.extend(batch)
            start_at += len(batch)
            total = data.get("total")
            if not batch or (total is not None and start_at >= int(total)):
                break
        return issues

    @staticmethod
    def _get_priority(priority_name: str) -> str:
        name = str(priority_name or "").strip()
        return PRIORITY_ALIASES.get(name.lower(), name or "Medium")

    @classmethod
    def _issue_from_raw(cls, issue: Dict[str, Any]) -> JiraIssue:
        fields = issue.get("fields", {}) or {}
        return JiraIssue(
            key=str(issue.get("key", "")),
            issue_type=(fields.get("issuetype") or {}).get("name", "Task"),
            status=(fields.get("status") or {}).get("name", ""),
            priority=cls._get_priority((fields.get("priority") or {}).get("name", "")),
            summary=fields.get("summary", "") or "",
            description=adf_to_text(fields.get("description")).strip(),
            reporter=(fields.get("reporter") or {}).get("displayName", ""),
            assignee=(fields.get("assignee") or {}).get("displayName", ""),
            created=fields.get("created", "") or "",
            resolved=fields.get("resolutiondate", "") or "",
            labels=list(fields.get("labels", []) or []),
        )

    def search_tasks(self, prefix: Optional[str] = None) -> List[JiraIssue]:
        """Sucht Tasks im Projekt und filtert optional nach Key-Prefix."""
        prefix = prefix or str(self.config.get("filters", {}).get("prefix", "JRA-"))
        if not self.is_configured():
            LOGGER.warning(
                "Jira ist nicht konfiguriert (server.url/api.token) – "
                "es werden keine Tasks geladen."
            )
            return []

        project = self.config.get("project", {})
        jql = str(project.get("jql", "") or "").strip()
        if not jql:
            key = str(project.get("key", "JRA") or "JRA")
            jql = f'project = "{key}" ORDER BY key ASC'

        tasks = [
            self._issue_from_raw(issue)
            for issue in self._fetch_issues(jql)
        ]
        if prefix:
            tasks = [task for task in tasks if task.key.upper().startswith(prefix.upper())]
        LOGGER.info("%s Task(s) mit Prefix %s gefunden", len(tasks), prefix)
        return tasks

    def fetch_tasks_by_keys(
        self,
        keys: List[str],
        batch_size: Optional[int] = None,
        batch_delay: Optional[float] = None,
    ) -> Dict[str, JiraIssue]:
        """Lädt Jira-Details für konkrete Keys in Blöcken (max. 100)."""
        unique_keys: List[str] = []
        seen: set = set()
        for key in keys:
            normalized = str(key).strip().upper()
            if normalized and normalized not in seen:
                seen.add(normalized)
                unique_keys.append(normalized)

        if not unique_keys:
            return {}
        if not self.is_configured():
            LOGGER.warning(
                "Jira ist nicht konfiguriert (server.url/api.token) – "
                "Anreicherung wird übersprungen."
            )
            return {}

        size = self.batch_size if batch_size is None else int(batch_size)
        size = max(1, min(100, size))
        delay = self.batch_delay if batch_delay is None else float(batch_delay)
        delay = max(0.0, delay)

        result: Dict[str, JiraIssue] = {}
        total_batches = (len(unique_keys) + size - 1) // size
        for batch_index, start in enumerate(
            range(0, len(unique_keys), size), start=1
        ):
            chunk = unique_keys[start : start + size]
            jql = "key in (" + ", ".join(f'"{key}"' for key in chunk) + ")"
            LOGGER.info(
                "Jira-Batch %s/%s: %s Key(s) werden geladen",
                batch_index,
                total_batches,
                len(chunk),
            )
            try:
                raw_issues = self._fetch_issues(jql)
            except JiraError as exc:
                LOGGER.error(
                    "Jira-Batch %s/%s fehlgeschlagen: %s",
                    batch_index,
                    total_batches,
                    exc,
                )
                continue

            for issue in raw_issues:
                task = self._issue_from_raw(issue)
                if task.key.upper() in seen:
                    result[task.key.upper()] = task

            if delay and batch_index < total_batches:
                LOGGER.debug("Warte %.1f Sekunde(n) vor dem nächsten Batch", delay)
                time.sleep(delay)

        LOGGER.info(
            "%s von %s Jira-Task(s) angereichert", len(result), len(unique_keys)
        )
        return result


# ---------------------------------------------------------------------------
# SQL-Scan
# ---------------------------------------------------------------------------
class ScriptScanner:
    """Findet SQL-Dateien und indexiert sie nach dem Jira-Task-Verzeichnis."""

    TABLE_PATTERN = re.compile(
        r"\b(?:FROM|JOIN|INTO|UPDATE|DELETE\s+FROM)\s+([A-Za-z_][\w.$]*)",
        re.IGNORECASE,
    )
    KPI_PATTERN = re.compile(
        r"\b(?:kpi|dashboard|report|auswertung|umsatz|gewinn|marge|bestand|"
        r"kunde|lieferant|monat|jahr)\w*",
        re.IGNORECASE,
    )

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        scripts = self.config.get("scripts", {})
        self.base_path = resolve_path(scripts.get("base_path", "scripts"))
        self.extensions = [
            str(ext).lower()
            for ext in (scripts.get("extensions") or [".sql"])
            if str(ext).strip()
        ]
        self.include_subdirectories = bool(scripts.get("include_subdirectories", True))
        self.exclude = [str(item) for item in (scripts.get("exclude") or [])]
        self._scripts: List[Path] = []
        self._files: List[Dict[str, Any]] = []
        self._index: Optional[Dict[str, List[Dict[str, Any]]]] = None

    def find_scripts(self) -> List[Path]:
        """Liefert alle passenden SQL-Dateien (jede Datei genau einmal)."""
        if not self.base_path.exists():
            LOGGER.warning("Scan-Root nicht gefunden: %s", self.base_path)
            self._scripts = []
            return self._scripts

        found: List[Path] = []
        seen: set = set()
        for extension in self.extensions:
            pattern = f"*{extension}"
            iterator = (
                self.base_path.rglob(pattern)
                if self.include_subdirectories
                else self.base_path.glob(pattern)
            )
            for path in iterator:
                if not path.is_file() or path in seen:
                    continue
                relative = path.relative_to(self.base_path).as_posix()
                if path.name in self.exclude or relative in self.exclude:
                    continue
                seen.add(path)
                found.append(path)
        self._scripts = sorted(found)
        return self._scripts

    @staticmethod
    def _detect_technologies(content: str, suffix: str = ".sql") -> List[str]:
        """Erkennt SQL-Technologien und grobe Dialekt-Hinweise."""
        if suffix.lower() != ".sql":
            return []
        upper = content.upper()
        technologies = {"SQL"}
        if "PL/PGSQL" in upper or "LANGUAGE PLPGSQL" in upper:
            technologies.add("PL/pgSQL")
        if "NVARCHAR" in upper or "\nGO\n" in upper or "IDENTITY(" in upper:
            technologies.add("T-SQL")
        if "AUTO_INCREMENT" in upper or "ENGINE=INNODB" in upper:
            technologies.add("MySQL")
        if "ROWNUM" in upper or "CONNECT BY" in upper or "NVL(" in upper:
            technologies.add("Oracle")
        return sorted(technologies)

    def extract_script_info(self, script_path: Path) -> Dict[str, Any]:
        """Extrahiert Metadaten und den Task-Key aus dem Elternverzeichnis."""
        info: Dict[str, Any] = {
            "path": str(script_path),
            "name": script_path.name,
            "size_bytes": 0,
            "lines": 0,
            "modified_date": "",
            "task_key": "",
            "task_dir": str(script_path.parent),
            "jira_refs": [],
            "technologies": [],
            "tables": [],
            "kpi_reports": [],
            "error": "",
        }

        task_key = task_key_from_directory(script_path.parent)
        if task_key:
            info["task_key"] = task_key
            info["jira_refs"] = [task_key]
        else:
            LOGGER.debug(
                "SQL-Datei ohne Jira-Task-Verzeichnis wird ohne Task-Zuordnung "
                "analysiert: %s",
                script_path,
            )

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
        info["technologies"] = self._detect_technologies(content, script_path.suffix)
        info["tables"] = sorted(
            {match.upper() for match in self.TABLE_PATTERN.findall(content)}
        )
        info["kpi_reports"] = sorted(
            {match.lower() for match in self.KPI_PATTERN.findall(content)}
        )
        return info

    def build_index(self) -> Dict[str, List[Dict[str, Any]]]:
        """Liest alle SQL-Dateien einmal und indexiert sie nach Task-Key.

        Alle gefundenen Dateien werden in ``self._files`` gesammelt – auch
        Dateien außerhalb eines Jira-Task-Verzeichnisses. Der Task-Index
        enthält weiterhin nur Dateien mit erkennbarem Jira-Key.
        """
        index: Dict[str, List[Dict[str, Any]]] = {}
        files: List[Dict[str, Any]] = []
        for script in self.find_scripts():
            info = self.extract_script_info(script)
            if info.get("error"):
                LOGGER.warning("SQL-Datei konnte nicht gelesen werden: %s", info["error"])
            files.append(info)
            for reference in info.get("jira_refs", []):
                index.setdefault(reference.upper(), []).append(info)
        self._files = files
        self._index = index
        LOGGER.info(
            "%s SQL-Datei(en) analysiert (%s Task-Verzeichnis(se))",
            len(files),
            len(index),
        )
        return index

    def get_all_files(self) -> List[Dict[str, Any]]:
        """Liefert alle analysierten SQL-Dateien (auch ohne Jira-Task-Key)."""
        if self._index is None:
            self.build_index()
        return list(self._files)

    def get_task_info(self, task: JiraIssue) -> List[Dict[str, Any]]:
        """Liefert alle SQL-Dateien, die im Verzeichnis des Tasks liegen."""
        if self._index is None:
            self.build_index()
        assert self._index is not None
        return self._index.get(task.key.upper(), [])

    def get_task_directory(self, task: JiraIssue) -> Optional[Path]:
        """Liefert das Verzeichnis des Tasks, falls SQL-Dateien existieren."""
        matches = self.get_task_info(task)
        if not matches:
            return None
        return Path(matches[0]["path"]).parent


# ---------------------------------------------------------------------------
# Scan-Index
# ---------------------------------------------------------------------------
def scan_index_path(config: Dict[str, Any]) -> Path:
    """Pfad zur JSON-Ablage der Verzeichnis-/SQL-Inventur."""
    storage = config.get("storage", {})
    value = storage.get("scan_index", "inventory/scan_index.json")
    return resolve_path(value)


def build_scan_index(
    scanner: ScriptScanner,
    index: Optional[Dict[str, List[Dict[str, Any]]]] = None,
) -> Dict[str, Any]:
    """Erzeugt eine serialisierbare Beschreibung der Verzeichnisstruktur."""
    task_index = index if index is not None else scanner.build_index()
    files = scanner.get_all_files()
    tasks: Dict[str, Any] = {}
    for task_key in sorted(task_index):
        matches = task_index[task_key]
        if not matches:
            continue
        directory = Path(matches[0]["path"]).parent
        tasks[task_key] = {
            "directory": str(directory),
            "path_context": derive_path_context(directory),
            "files": matches,
        }

    return {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "base_path": str(scanner.base_path),
        "task_count": len(tasks),
        "file_count": len(files),
        "files": files,
        "tasks": tasks,
    }


def save_scan_index(config: Dict[str, Any], data: Dict[str, Any]) -> Path:
    """Schreibt die Scan-Inventur als JSON."""
    path = scan_index_path(config)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
    LOGGER.info("Scan-Index gespeichert: %s", path)
    return path


def load_scan_index(config: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Lädt eine zuvor erzeugte Scan-Inventur, falls vorhanden."""
    path = scan_index_path(config)
    if not path.exists():
        LOGGER.warning("Scan-Index nicht gefunden: %s", path)
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        LOGGER.warning("Scan-Index konnte nicht gelesen werden (%s): %s", path, exc)
        return None
    if not isinstance(data, dict) or not isinstance(data.get("tasks"), dict):
        LOGGER.warning("Scan-Index hat ein unerwartetes Format: %s", path)
        return None
    return data


def filter_scan_data_by_prefix(
    scan_data: Dict[str, Any], prefix: Optional[str]
) -> Dict[str, Any]:
    """Filtert die Inventur auf Task-Keys mit dem konfigurierten Prefix.

    Dateien ohne Jira-Task-Key (z. B. allgemeine SQL-Sammlungen) bleiben
    erhalten, damit wirklich alle Dateien im Bericht auftauchen. Dateien in
    Verzeichnissen fremder Projekte werden entfernt.
    """
    if not prefix:
        return scan_data

    prefix_upper = prefix.upper()
    all_tasks = scan_data.get("tasks", {})
    kept = {
        key: entry
        for key, entry in all_tasks.items()
        if key.upper().startswith(prefix_upper)
    }
    ignored = sorted(set(all_tasks) - set(kept))
    if ignored:
        LOGGER.info(
            "%s Task(s) außerhalb des Prefix %s werden ignoriert: %s",
            len(ignored),
            prefix,
            ", ".join(ignored),
        )

    all_files = scan_data.get("files")
    if isinstance(all_files, list):
        files = [
            entry
            for entry in all_files
            if not entry.get("task_key")
            or str(entry["task_key"]).upper().startswith(prefix_upper)
        ]
    else:
        files = [match for entry in kept.values() for match in entry.get("files", [])]

    filtered = dict(scan_data)
    filtered["tasks"] = kept
    filtered["files"] = files
    filtered["task_count"] = len(kept)
    filtered["file_count"] = len(files)
    return filtered


# ---------------------------------------------------------------------------
# Excel-Export
# ---------------------------------------------------------------------------
class ExcelExporter:
    """Erzeugt pro Jira-Task einen Excel-Bericht im Task-Verzeichnis."""

    def __init__(self, config: Optional[Dict[str, Any]] = None):
        self.config = config or load_config()
        excel_cfg = self.config.get("excel", {})
        self.output_dir = resolve_path(excel_cfg.get("output_path", "output"))
        self.filename_pattern = str(excel_cfg.get("output_filename", "{JIRA}.xlsx"))
        self.fallback_pattern = str(
            excel_cfg.get("fallback_filename", "jira_report_{YYYY}_{MM}_{DD}.xlsx")
        )
        self.include_csv = bool(excel_cfg.get("include_csv", False))
        self.sheet_name = str(excel_cfg.get("sheet_name", "Jira-Task-Report"))
        self.datetime_format = str(
            excel_cfg.get("datetime_format", "%Y-%m-%d %H:%M:%S")
        )
        self.aggregate_report = bool(excel_cfg.get("aggregate_report", False))
        self.aggregate_filename = str(
            excel_cfg.get(
                "aggregate_filename",
                "jira_report_{YYYY}_{MM}_{DD}.xlsx",
            )
        )

    def _render_pattern(self, pattern: str, values: Dict[str, Any]) -> str:
        now = datetime.now()
        merged: Dict[str, str] = {
            "YYYY": now.strftime("%Y"),
            "MM": now.strftime("%m"),
            "DD": now.strftime("%d"),
        }
        merged.update(
            {key: str(value) for key, value in values.items() if value is not None}
        )

        rendered = pattern
        for key, value in merged.items():
            rendered = rendered.replace("{" + key + "}", value)
        rendered = re.sub(r"\{[A-Z_]+\}", "", rendered)
        rendered = rendered.replace("/", "_").replace("\\", "_").strip()
        if not rendered:
            rendered = "report.xlsx"
        if not rendered.lower().endswith(".xlsx"):
            rendered += ".xlsx"
        return rendered

    def export_task(
        self,
        task_key: str,
        rows: List[Dict[str, Any]],
        directory: Path,
        sql_stem: Optional[str] = None,
        path_context: Optional[Dict[str, str]] = None,
    ) -> str:
        """Schreibt einen Bericht direkt in das Verzeichnis des Tasks."""
        values: Dict[str, Any] = {"JIRA": task_key, "SQL": sql_stem or task_key}
        if path_context:
            values.update(path_context)

        filename = self._render_pattern(self.filename_pattern, values)
        output_path = Path(directory) / filename
        self._write_workbook(rows, output_path)
        LOGGER.info("Excel-Bericht gespeichert: %s", output_path)
        return str(output_path)

    def export(self, data: List[Dict[str, Any]]) -> str:
        """Fallback-Export in ``excel.output_path`` (z. B. leerer Bericht)."""
        filename = self._render_pattern(self.fallback_pattern, {})
        output_path = self.output_dir / filename
        self._write_workbook(data, output_path)
        LOGGER.info("Excel-Bericht gespeichert: %s", output_path)
        return str(output_path)

    def export_aggregate(self, data: List[Dict[str, Any]]) -> str:
        """Schreibt einen einzigen Gesamtbericht nach ``excel.output_path``."""
        filename = self._render_pattern(self.aggregate_filename, {})
        output_path = self.output_dir / filename
        self._write_workbook(data, output_path)
        LOGGER.info("Gesamtbericht gespeichert: %s", output_path)
        return str(output_path)

    def _write_workbook(self, data: List[Dict[str, Any]], output_path: Path) -> None:
        """Erzeugt die Arbeitsmappe inklusive Formatierung und optional CSV."""
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
            worksheet.row_dimensions[row_index].height = (
                30
                if any(row.get(title) for title, _w, wrap in COLUMNS if wrap)
                else 18
            )

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

        output_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(output_path)

        if self.include_csv:
            csv_path = output_path.with_suffix(".csv")
            self.export_csv(data, csv_path)
            LOGGER.info("CSV-Bericht gespeichert: %s", csv_path)

    def export_csv(self, data: List[Dict[str, Any]], path: Path) -> str:
        """Schreibt denselben Inhalt als UTF-8-CSV (Semikolon-getrennt)."""
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(
                handle,
                fieldnames=[title for title, _width, _wrap in COLUMNS],
                delimiter=";",
                extrasaction="ignore",
            )
            writer.writeheader()
            for row in data:
                writer.writerow({key: row.get(key, "") for key in writer.fieldnames})
        return str(path)

    def export_simple(self, data: List[Dict[str, Any]]) -> str:
        """Gibt die Tabelle auf der Konsole aus (Dry-Run)."""
        titles = [title for title, _width, _wrap in COLUMNS]
        print(" | ".join(titles))
        print("-" * 120)
        for row in data:
            print(" | ".join(str(row.get(title, "")) for title in titles))
        print("-" * 120)
        print(f"GESAMT: {len(data)}")
        return ""


# ---------------------------------------------------------------------------
# Zeilenaufbau
# ---------------------------------------------------------------------------
def build_rows_for_task(
    task: JiraIssue, matches: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Baut die Excel-Zeilen eines Tasks (eine Zeile pro SQL-Datei)."""
    resolved = format_datetime(task.resolved) or "Pending"

    if not matches:
        return [
            {
                "Jira-Nummer": task.key,
                "Jira-Auftraggeber": task.reporter,
                "Entwickler": task.assignee,
                "Jahr/Monat/Tag-Abgeschlossen": resolved,
                "Zusammenfassung-Anforderung": task.summary,
                "File-Datum": "",
                "File-Location": "Not found",
                "Verwendete-Technologien": "",
                "Status": task.status or "Missing",
                "Priorität": task.priority,
                "Jahr": "",
                "Abteilung": "",
            }
        ]

    rows: List[Dict[str, Any]] = []
    for match in matches:
        rows.append(
            {
                "Jira-Nummer": task.key,
                "Jira-Auftraggeber": task.reporter,
                "Entwickler": task.assignee,
                "Jahr/Monat/Tag-Abgeschlossen": resolved,
                "Zusammenfassung-Anforderung": task.summary,
                "File-Datum": format_date(match.get("modified_date")),
                "File-Location": match.get("path", ""),
                "Verwendete-Technologien": ", ".join(
                    match.get("technologies", [])[:8]
                ),
                "Status": task.status or "Open",
                "Priorität": task.priority,
                **local_path_columns(match),
            }
        )
    return rows


def build_scan_rows(
    task_key: str, matches: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """Baut Zeilen ausschließlich aus Verzeichnisinformationen (ohne Jira)."""
    if not matches:
        row = empty_report_row(task_key)
        row["Zusammenfassung-Anforderung"] = "Keine SQL-Datei im Task-Verzeichnis"
        row["Status"] = "Pending"
        return [row]

    rows: List[Dict[str, Any]] = []
    for match in matches:
        rows.append(
            {
                "Jira-Nummer": task_key,
                "Jira-Auftraggeber": "",
                "Entwickler": "",
                "Jahr/Monat/Tag-Abgeschlossen": "",
                "Zusammenfassung-Anforderung": "",
                "File-Datum": format_date(match.get("modified_date")),
                "File-Location": match.get("path", ""),
                "Verwendete-Technologien": ", ".join(
                    match.get("technologies", [])[:8]
                ),
                "Status": "Pending",
                "Priorität": "",
                **local_path_columns(match),
            }
        )
    return rows


def build_enriched_rows(
    task_key: str,
    task: Optional[JiraIssue],
    matches: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Baut Zeilen mit Jira-Details; ohne Jira-Treffer bleiben Scan-Daten."""
    if task is None:
        rows = build_scan_rows(task_key, matches)
        for row in rows:
            row["Status"] = "Missing"
        return rows
    return build_rows_for_task(task, matches)


def scan_files_from_data(scan_data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Liefert alle Dateien einer Scan-Inventur (inkl. Dateien ohne Task)."""
    files = scan_data.get("files")
    if isinstance(files, list):
        return files
    return [
        match
        for entry in scan_data.get("tasks", {}).values()
        for match in entry.get("files", [])
    ]


def build_aggregate_scan_rows(
    files: List[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    """Baut einen Gesamtbericht ausschließlich aus Scan-Daten."""
    rows: List[Dict[str, Any]] = []
    for match in files:
        task_key = str(match.get("task_key", "") or "")
        rows.extend(build_scan_rows(task_key, [match]))
    return rows


def build_aggregate_enriched_rows(
    scan_data: Dict[str, Any],
    jira_tasks: Dict[str, JiraIssue],
) -> List[Dict[str, Any]]:
    """Baut einen Gesamtbericht mit optionalen Jira-Details pro Datei."""
    rows: List[Dict[str, Any]] = []
    for match in scan_files_from_data(scan_data):
        task_key = str(match.get("task_key", "") or "").upper()
        task = jira_tasks.get(task_key) if task_key else None
        if task is not None:
            rows.extend(build_rows_for_task(task, [match]))
        elif task_key:
            rows.extend(build_enriched_rows(task_key, None, [match]))
        else:
            rows.extend(build_scan_rows("", [match]))
    return rows


def build_report(
    tasks: List[JiraIssue], scanner: ScriptScanner
) -> List[Dict[str, Any]]:
    """Baut alle Excel-Zeilen aus Jira-Tasks und gefundenen SQL-Dateien."""
    rows: List[Dict[str, Any]] = []
    for task in tasks:
        rows.extend(build_rows_for_task(task, scanner.get_task_info(task)))
    return rows


def empty_report_row(prefix: str = "JRA-") -> Dict[str, Any]:
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
        "Jahr": "",
        "Abteilung": "",
    }


# ---------------------------------------------------------------------------
# Phasen
# ---------------------------------------------------------------------------
def run_scan_phase(
    config: Dict[str, Any],
    exporter: ExcelExporter,
    dry_run: bool = False,
    prefix: Optional[str] = None,
) -> Tuple[Dict[str, Any], List[str], List[Dict[str, Any]]]:
    """Phase 1: Verzeichnisstruktur einlesen und Berichte soweit möglich bauen."""
    scanner = ScriptScanner(config)
    index = scanner.build_index()
    scan_data = filter_scan_data_by_prefix(build_scan_index(scanner, index), prefix)

    written: List[str] = []
    rows_all: List[Dict[str, Any]] = []
    files = scan_files_from_data(scan_data)

    if not files:
        fallback_prefix = prefix or str(
            config.get("filters", {}).get("prefix", "JRA-")
        )
        LOGGER.warning(
            "Keine SQL-Dateien gefunden – es wird ein leerer Fallback-Bericht erzeugt."
        )
        rows_all = [empty_report_row(fallback_prefix)]
        if not dry_run:
            if exporter.aggregate_report:
                written.append(exporter.export_aggregate(rows_all))
            else:
                written.append(exporter.export(rows_all))
        return scan_data, written, rows_all

    if not dry_run:
        save_scan_index(config, scan_data)

    if exporter.aggregate_report:
        rows_all = build_aggregate_scan_rows(files)
        if not dry_run:
            written.append(exporter.export_aggregate(rows_all))
        return scan_data, written, rows_all

    for task_key, entry in scan_data["tasks"].items():
        rows = build_scan_rows(task_key, entry["files"])
        rows_all.extend(rows)
        if dry_run:
            continue
        first_file = Path(entry["files"][0]["path"])
        written.append(
            exporter.export_task(
                task_key,
                rows,
                Path(entry["directory"]),
                sql_stem=first_file.stem,
                path_context=entry.get("path_context"),
            )
        )

    return scan_data, written, rows_all


def run_enrich_phase(
    config: Dict[str, Any],
    exporter: ExcelExporter,
    scan_data: Dict[str, Any],
    batch_size: Optional[int] = None,
    batch_delay: Optional[float] = None,
    dry_run: bool = False,
) -> Tuple[List[str], List[Dict[str, Any]]]:
    """Phase 2: Jira-Details für die gefundenen Task-Keys nachladen."""
    tasks = scan_data.get("tasks", {}) if scan_data else {}
    if not tasks:
        LOGGER.warning("Keine Tasks im Scan-Index – Anreicherung wird übersprungen.")
        return [], []

    client = JiraClient(config)
    if not client.is_configured():
        LOGGER.warning(
            "Jira ist nicht konfiguriert – Anreicherung wird übersprungen; "
            "die Scan-Berichte bleiben unverändert."
        )
        return [], []

    jira_tasks = client.fetch_tasks_by_keys(
        list(tasks.keys()),
        batch_size=batch_size,
        batch_delay=batch_delay,
    )

    written: List[str] = []
    rows_all: List[Dict[str, Any]] = []
    missing: List[str] = []

    for task_key in tasks:
        if task_key.upper() not in jira_tasks:
            missing.append(task_key)

    if exporter.aggregate_report:
        rows_all = build_aggregate_enriched_rows(scan_data, jira_tasks)
        if not dry_run:
            written.append(exporter.export_aggregate(rows_all))
        if missing:
            LOGGER.warning(
                "Für %s Task(s) wurden keine Jira-Details gefunden: %s",
                len(missing),
                ", ".join(missing),
            )
        return written, rows_all

    for task_key, entry in tasks.items():
        task = jira_tasks.get(task_key.upper())
        rows = build_enriched_rows(task_key, task, entry["files"])
        rows_all.extend(rows)
        if dry_run:
            continue
        first_file = Path(entry["files"][0]["path"])
        written.append(
            exporter.export_task(
                task_key,
                rows,
                Path(entry["directory"]),
                sql_stem=first_file.stem,
                path_context=entry.get("path_context"),
            )
        )

    if missing:
        LOGGER.warning(
            "Für %s Task(s) wurden keine Jira-Details gefunden: %s",
            len(missing),
            ", ".join(missing),
        )
    return written, rows_all


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv: Optional[List[str]] = None) -> int:
    """CLI-Einstiegspunkt."""
    parser = argparse.ArgumentParser(
        description=(
            "Jira/Gira Scanner – Phase 1 liest die Verzeichnisstruktur und "
            "erzeugt Excel-Berichte; Phase 2 reichert sie mit Jira-Details an."
        ),
        epilog=(
            "Beispiele: main.py --mode scan | main.py --mode enrich | "
            "main.py --mode all --batch-size 50 --batch-delay 2"
        ),
    )
    parser.add_argument("--config", help="Pfad zur config.yaml")
    parser.add_argument(
        "--mode",
        choices=("scan", "enrich", "all"),
        default="all",
        help=(
            "scan: nur Verzeichnisstruktur ohne Jira; "
            "enrich: Jira-Details nachladen; "
            "all: beides nacheinander (Default)"
        ),
    )
    parser.add_argument(
        "--batch-size",
        type=int,
        help="Jira-Keys pro Request (1-100, Default aus jira.batch_size)",
    )
    parser.add_argument(
        "--batch-delay",
        type=float,
        help="Pause zwischen Jira-Batches in Sekunden",
    )
    parser.add_argument(
        "--prefix",
        help="Key-Prefix (überschreibt filters.prefix, Standard: JRA-)",
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
    prefix = args.prefix or str(config.get("filters", {}).get("prefix", "JRA-"))

    LOGGER.info("=" * 70)
    LOGGER.info(
        "Jira/Gira Scanner gestartet (Modus: %s, Prefix: %s)", args.mode, prefix
    )
    LOGGER.info("=" * 70)

    exporter = ExcelExporter(config)
    scan_data: Optional[Dict[str, Any]] = None
    written: List[str] = []
    all_rows: List[Dict[str, Any]] = []

    try:
        if args.mode in ("scan", "all"):
            scan_data, scan_written, scan_rows = run_scan_phase(
                config, exporter, dry_run=args.dry_run, prefix=prefix
            )
            written.extend(scan_written)
            all_rows.extend(scan_rows)
            LOGGER.info(
                "%s Task-Verzeichnis(se), %s SQL-Datei(en) im Scan",
                scan_data.get("task_count", 0),
                scan_data.get("file_count", 0),
            )
            if args.mode == "scan":
                if args.dry_run:
                    exporter.export_simple(all_rows)
                else:
                    LOGGER.info(
                        "Scan-Phase abgeschlossen: %s Bericht(e) geschrieben",
                        len(written),
                    )
                return 0
            if not scan_data.get("tasks"):
                if args.dry_run:
                    exporter.export_simple(all_rows)
                return 0
        else:
            scan_data = load_scan_index(config)
            if scan_data is None:
                LOGGER.warning(
                    "Kein Scan-Index vorhanden – Verzeichnisstruktur wird "
                    "zuerst eingelesen."
                )
                scan_data, scan_written, scan_rows = run_scan_phase(
                    config, exporter, dry_run=args.dry_run, prefix=prefix
                )
                written.extend(scan_written)
                all_rows.extend(scan_rows)
                if not scan_data.get("tasks"):
                    if args.dry_run:
                        exporter.export_simple(all_rows)
                    return 0

        enrich_written, enrich_rows = run_enrich_phase(
            config,
            exporter,
            scan_data or {},
            batch_size=args.batch_size,
            batch_delay=args.batch_delay,
            dry_run=args.dry_run,
        )
        written.extend(enrich_written)
        all_rows.extend(enrich_rows)
    except OSError as exc:
        LOGGER.error("Bericht konnte nicht geschrieben werden: %s", exc)
        return 1

    if args.dry_run:
        exporter.export_simple(all_rows)
        return 0

    unique_written = list(dict.fromkeys(written))
    if not unique_written:
        LOGGER.warning("Keine Berichte geschrieben.")
        return 0

    LOGGER.info("Fertig: %s Excel-Bericht(e) geschrieben", len(unique_written))
    for path in unique_written:
        LOGGER.info("  %s", path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
