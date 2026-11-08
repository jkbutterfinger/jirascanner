#!/usr/bin/env python3
"""Script Inventory Scanner – analysiert Reporting-Skripte.

Das Modul wird von der API (``/api/scripts``) und als eigenständiges CLI
verwendet. Es liest Skripte, erkennt Jira-Referenzen, Technologien,
Datenbankzugriffe und KPI-/Dashboard-Begriffe.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

# Direktaufruf "python scripts/script_inventory.py" ermöglichen.
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from config import ROOT, load_config, resolve_path  # noqa: E402

DEFAULT_EXTENSIONS = [".py", ".sh", ".sql", ".js", ".ts"]

IMPORT_PATTERNS = [
    re.compile(r"^\s*import\s+([A-Za-z_][\w.]*)", re.MULTILINE),
    re.compile(r"^\s*from\s+([A-Za-z_][\w.]*)\s+import\b", re.MULTILINE),
    re.compile(r"""require\(\s*["']([^"']+)["']\s*\)"""),
]

DB_PATTERNS = [
    r"(?:mysql|postgres(?:ql)?|sqlite|oracle|mssql)://[^\s\"']+",
    r"DATABASE_URL\s*=\s*[\"']([^\"']+)[\"']",
    r"(?:DB|DATABASE)_[A-Z_]+\s*=\s*[\"']([^\"']+)[\"']",
]

API_PATTERNS = [
    r"\brequests\.(?:get|post|put|patch|delete)\s*\(",
    r"\bhttpx\.(?:get|post|put|patch|delete)\s*\(",
    r"\burllib\.request\b",
    r"\bfetch\s*\(",
    r"\baxios\.(?:get|post|put|patch|delete)\s*\(",
]

KPI_PATTERNS = [
    r"\bkpi[_:]?\s*([^\s,:\n]+)",
    r"\bdashboard[_:]?\s*([^\s,:\n]+)",
    r"\bmetric[_:]?\s*([^\s,:\n]+)",
    r"\breport[_:]?\s*([^\s,:\n]+)",
]


class ScriptAnalyzer:
    """Analysiert Reporting-Skripte und extrahiert Metadaten."""

    def __init__(
        self,
        extensions: Optional[List[str]] = None,
        exclude: Optional[List[str]] = None,
    ):
        self.extensions = extensions or list(DEFAULT_EXTENSIONS)
        self.exclude = set(exclude or ["__init__.py", "script_inventory.py"])
        self.jira_ref_pattern = re.compile(r"\b([A-Z][A-Z0-9]+-\d+)\b")

    def scan_directory(self, path: str | Path) -> List[Dict[str, Any]]:
        """Scannt ein Verzeichnis rekursiv."""
        scripts_path = Path(path)
        if not scripts_path.exists():
            return []

        results: List[Dict[str, Any]] = []
        for file_path in sorted(scripts_path.rglob("*")):
            if not file_path.is_file():
                continue
            if file_path.suffix not in self.extensions:
                continue
            if file_path.name in self.exclude:
                continue
            results.append(self.analyze_file(file_path))
        return results

    def analyze_file(self, file_path: Path) -> Dict[str, Any]:
        """Analysiert eine einzelne Datei."""
        analysis: Dict[str, Any] = {
            "path": str(file_path),
            "name": file_path.name,
            "size_bytes": 0,
            "lines": 0,
            "modified_date": "",
            "jira_refs": [],
            "technologies": [],
            "kpi_tags": [],
            "db_access": [],
            "api_calls": [],
            "complexity": 1,
            "error": "",
        }

        try:
            stat = file_path.stat()
            analysis["size_bytes"] = stat.st_size
            analysis["modified_date"] = datetime.fromtimestamp(stat.st_mtime).strftime(
                "%Y-%m-%d %H:%M:%S"
            )
            content = file_path.read_text(encoding="utf-8", errors="replace")
        except OSError as exc:
            analysis["error"] = str(exc)
            return analysis

        analysis["lines"] = len(content.splitlines())
        analysis["jira_refs"] = sorted(set(self.jira_ref_pattern.findall(content)))
        analysis["technologies"] = self._detect_technologies(content)
        analysis["kpi_tags"] = sorted(
            {
                match.strip().lower()
                for pattern in KPI_PATTERNS
                for match in re.findall(pattern, content, re.IGNORECASE)
                if match.strip()
            }
        )
        analysis["db_access"] = self._collect(DB_PATTERNS, content)
        analysis["api_calls"] = self._collect(API_PATTERNS, content)
        analysis["complexity"] = self.estimate_complexity(content)
        return analysis

    @staticmethod
    def _collect(patterns: List[str], content: str) -> List[str]:
        matches: List[str] = []
        for pattern in patterns:
            try:
                matches.extend(re.findall(pattern, content, re.IGNORECASE))
            except re.error:
                continue
        return sorted({match for match in matches if match})

    @staticmethod
    def _detect_technologies(content: str) -> List[str]:
        import sys

        names: set[str] = set()
        for pattern in IMPORT_PATTERNS:
            for match in pattern.findall(content):
                root = match.split(".")[0].strip()
                if root and not root.startswith((".", "/")):
                    names.add(root)
        stdlib = getattr(sys, "stdlib_module_names", frozenset())
        builtin = set(sys.builtin_module_names)
        return sorted(name for name in names if name not in stdlib and name not in builtin)

    @staticmethod
    def estimate_complexity(code: str) -> int:
        """Grobe Komplexitätsschätzung zwischen 1 und 5."""
        lines = len(code.splitlines())
        imports = len(re.findall(r"^\s*(?:import|from)\s+\w+", code, re.MULTILINE))
        classes = len(re.findall(r"\bclass\s+\w+", code))
        functions = len(re.findall(r"\bdef\s+\w+\s*\(", code))
        score = lines // 50 + imports // 10 + classes + functions // 10
        return min(5, max(1, score))


class JiraMapper:
    """Mapping zwischen Jira-Tasks und Implementierungen."""

    def __init__(self):
        self.tasks: List[Dict[str, Any]] = []
        self.implementations: List[Dict[str, Any]] = []

    def load_jira_tasks(self, tasks_data: List[Dict[str, Any]]) -> None:
        self.tasks = tasks_data

    def load_implementations(self, implementations: List[Dict[str, Any]]) -> None:
        self.implementations = implementations

    def map_requirements(self) -> List[Dict[str, Any]]:
        results: List[Dict[str, Any]] = []
        for task in self.tasks:
            key = str(task.get("key", ""))
            matching = [
                implementation
                for implementation in self.implementations
                if key and key in implementation.get("jira_refs", [])
            ]
            results.append(
                {
                    "task": key,
                    "summary": task.get("summary", ""),
                    "status": task.get("status", ""),
                    "implementation_count": len(matching),
                    "fully_implemented": bool(matching),
                }
            )
        return results


class ReportGenerator:
    """Erzeugt Zusammenfassungen und Gap-Reports."""

    @staticmethod
    def generate_summary(
        inventory: List[Dict[str, Any]], jira_tasks: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        return {
            "total_scripts": len(inventory),
            "total_tasks": len(jira_tasks),
            "scripts_with_jira_refs": sum(
                1 for script in inventory if script.get("jira_refs")
            ),
            "scripts_with_errors": sum(1 for script in inventory if script.get("error")),
            "date": datetime.now().isoformat(timespec="seconds"),
        }

    @staticmethod
    def generate_gap_report(mapping: List[Dict[str, Any]]) -> str:
        gaps = [item for item in mapping if not item.get("fully_implemented")]
        lines = ["# Gap Report", "", f"## Total Gaps: {len(gaps)}", ""]
        for gap in gaps:
            lines.extend(
                [
                    f"### {gap.get('task', '')} - {gap.get('summary', '')}",
                    f"*Status*: {gap.get('status', '')}",
                    f"*Implementation Count*: {gap.get('implementation_count', 0)}",
                    "",
                ]
            )
        return "\n".join(lines)


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Script-Inventar scannen")
    parser.add_argument("--path", help="Zu scannendes Verzeichnis")
    parser.add_argument("--json", action="store_true", help="JSON ausgeben")
    args = parser.parse_args(argv)

    config = load_config()
    scripts_cfg = config.get("scripts", {})
    path = resolve_path(args.path or scripts_cfg.get("base_path", ROOT / "scripts"))

    analyzer = ScriptAnalyzer(
        extensions=scripts_cfg.get("extensions"),
        exclude=scripts_cfg.get("exclude"),
    )
    inventory = analyzer.scan_directory(path)
    summary = ReportGenerator.generate_summary(inventory, [])

    if args.json:
        print(json.dumps({"summary": summary, "scripts": inventory}, indent=2))
    else:
        print(json.dumps(summary, indent=2))
        for script in inventory:
            print(
                f"- {script['name']}: {script['lines']} Zeilen, "
                f"Jira-Refs: {', '.join(script['jira_refs']) or '-'}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
