#!/usr/bin/env python3
"""SQL-Inventar für den Jira/Gira Scanner.

Der Analyzer liest ausschließlich ``.sql``-Dateien, leitet den Jira-Task-Key
aus der Verzeichnisstruktur ab und sammelt lokale Metadaten (Technologien,
Tabellen, KPI-Begriffe, Komplexität). Zusätzlich stehen ein einfacher
Task-Mapper und ein Report-Generator für JSON/Markdown bereit.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import load_config, resolve_path, setup_logging
from main import derive_path_context, task_key_from_directory

LOGGER = logging.getLogger("jira_scanner")

TABLE_PATTERN = re.compile(
    r"\b(?:FROM|JOIN|INTO|UPDATE|DELETE\s+FROM)\s+([A-Za-z_][\w.$]*)",
    re.IGNORECASE,
)
API_PATTERN = re.compile(
    r"\b(?:UTL_HTTP|DBMS_CLOUD|APEX_WEB_SERVICE|HTTP\.?GET|HTTP\.?POST)\b",
    re.IGNORECASE,
)
KPI_PATTERN = re.compile(
    r"\b(?:kpi|dashboard|report|auswertung|umsatz|gewinn|marge|bestand|"
    r"kunde|lieferant|monat|jahr)\w*",
    re.IGNORECASE,
)


class ScriptAnalyzer:
    """Analysiert SQL-Dateien und liefert ein serialisierbares Inventar."""

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

    @staticmethod
    def _detect_technologies(content: str, suffix: str = ".sql") -> List[str]:
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

    @staticmethod
    def _complexity(lines: int) -> str:
        if lines < 50:
            return "low"
        if lines < 200:
            return "medium"
        return "high"

    def find_scripts(self, path: Optional[Path] = None) -> List[Path]:
        """Findet alle passenden SQL-Dateien unterhalb von ``path``."""
        root = Path(path) if path else self.base_path
        if not root.exists():
            LOGGER.warning("Scan-Pfad nicht gefunden: %s", root)
            return []

        found: List[Path] = []
        seen: set = set()
        for extension in self.extensions:
            iterator = (
                root.rglob(f"*{extension}")
                if self.include_subdirectories
                else root.glob(f"*{extension}")
            )
            for script in iterator:
                if not script.is_file() or script in seen:
                    continue
                relative = script.relative_to(root).as_posix()
                if script.name in self.exclude or relative in self.exclude:
                    continue
                seen.add(script)
                found.append(script)
        return sorted(found)

    def analyze(self, script_path: Path) -> Dict[str, Any]:
        """Analysiert eine einzelne SQL-Datei."""
        script_path = Path(script_path)
        task_key = task_key_from_directory(script_path.parent)
        info: Dict[str, Any] = {
            "path": str(script_path),
            "name": script_path.name,
            "task_key": task_key,
            "jira_refs": [task_key] if task_key else [],
            "path_context": derive_path_context(script_path.parent),
            "size_bytes": 0,
            "lines": 0,
            "modified_date": "",
            "technologies": [],
            "db_access": [],
            "api_calls": [],
            "kpi_reports": [],
            "complexity": "low",
            "error": "",
        }

        try:
            stat = script_path.stat()
            try:
                content = script_path.read_text(encoding="utf-8")
            except UnicodeDecodeError:
                content = script_path.read_text(encoding="latin-1")
        except OSError as exc:
            info["error"] = str(exc)
            return info

        info["size_bytes"] = stat.st_size
        info["lines"] = len(content.splitlines())
        info["modified_date"] = datetime.fromtimestamp(stat.st_mtime).strftime(
            "%Y-%m-%d %H:%M:%S"
        )
        info["technologies"] = self._detect_technologies(content, script_path.suffix)

        # Strip single-line (--) and multi-line (/* */) comments before parsing
        # tables/KPIs/APIs to prevent false positives in SQL comments.
        clean_content = re.sub(r"/\*[\s\S]*?\*/", "", content)
        clean_content = re.sub(r"--.*", "", clean_content)

        info["db_access"] = sorted(
            {match.upper() for match in TABLE_PATTERN.findall(clean_content)}
        )
        info["api_calls"] = sorted(
            {match.upper() for match in API_PATTERN.findall(clean_content)}
        )
        info["kpi_reports"] = sorted(
            {match.lower() for match in KPI_PATTERN.findall(clean_content)}
        )
        info["complexity"] = self._complexity(info["lines"])
        return info

    def analyze_directory(self, path: Optional[Path] = None) -> List[Dict[str, Any]]:
        """Analysiert alle SQL-Dateien eines Verzeichnisses."""
        results: List[Dict[str, Any]] = []
        for script in self.find_scripts(path):
            info = self.analyze(script)
            if info.get("error"):
                LOGGER.warning(
                    "SQL-Datei konnte nicht gelesen werden (%s): %s",
                    script,
                    info["error"],
                )
            results.append(info)
        return results


class JiraMapper:
    """Ordnet SQL-Inventar-Einträge den Jira-Tasks zu."""

    @staticmethod
    def map(
        scripts: List[Dict[str, Any]],
        tasks: Optional[List[Any]] = None,
    ) -> Dict[str, List[Dict[str, Any]]]:
        task_keys: List[str] = []
        for task in tasks or []:
            if isinstance(task, dict):
                key = str(task.get("key", ""))
            else:
                key = str(getattr(task, "key", ""))
            if key:
                task_keys.append(key.upper())

        mapping: Dict[str, List[Dict[str, Any]]] = {}
        for script in scripts:
            refs = [str(ref).upper() for ref in script.get("jira_refs", [])]
            for reference in refs:
                if task_keys and reference not in task_keys:
                    continue
                mapping.setdefault(reference, []).append(script)
        return mapping


class ReportGenerator:
    """Erzeugt JSON- und Markdown-Reports aus dem SQL-Inventar."""

    @staticmethod
    def to_json(scripts: List[Dict[str, Any]]) -> str:
        return json.dumps(scripts, indent=2, ensure_ascii=False)

    @staticmethod
    def to_markdown(scripts: List[Dict[str, Any]]) -> str:
        lines = [
            "# SQL-Inventar",
            "",
            f"Generiert: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Dateien: {len(scripts)}",
            "",
            "| Datei | Task | Zeilen | Technologien | Komplexität |",
            "|-------|------|-------:|--------------|-------------|",
        ]
        for script in scripts:
            lines.append(
                "| {name} | {task} | {lines} | {tech} | {complexity} |".format(
                    name=script.get("name", ""),
                    task=script.get("task_key", "") or "–",
                    lines=script.get("lines", 0),
                    tech=", ".join(script.get("technologies", [])) or "–",
                    complexity=script.get("complexity", ""),
                )
            )
        return "\n".join(lines) + "\n"


def main(argv: Optional[List[str]] = None) -> int:
    """CLI für das SQL-Inventar."""
    parser = argparse.ArgumentParser(description="SQL-Inventar des Jira/Gira Scanners")
    parser.add_argument("--config", help="Pfad zur config.yaml")
    parser.add_argument("--path", help="Zu scannendes Verzeichnis")
    parser.add_argument("--json", action="store_true", help="JSON ausgeben")
    parser.add_argument("-v", "--verbose", action="store_true", help="DEBUG-Logging")
    args = parser.parse_args(argv)

    config = load_config(args.config)
    setup_logging(config, verbose=args.verbose)
    analyzer = ScriptAnalyzer(config)
    scripts = analyzer.analyze_directory(Path(args.path) if args.path else None)

    if args.json:
        print(ReportGenerator.to_json(scripts))
    else:
        print(ReportGenerator.to_markdown(scripts))
    return 0


if __name__ == "__main__":
    sys.exit(main())
