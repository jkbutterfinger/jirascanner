#!/usr/bin/env python3
"""Kompatibilitäts-Einstiegspunkt für den Jira/Gira Scanner.

Die frühere, eigenständige Implementierung war ein Duplikat von ``main.py``
und enthielt Syntax-/Importfehler. Diese Datei bleibt erhalten, damit
bestehende Aufrufe (``python -m scanner.main_scanner``) weiter funktionieren,
delegiert aber vollständig an die kanonische Implementierung in ``main.py``.
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from main import (  # noqa: E402  (Import nach sys.path-Anpassung)
    ExcelExporter,
    JiraClient,
    JiraIssue,
    ScriptScanner,
    build_report,
    main,
)

__all__ = [
    "JiraClient",
    "JiraIssue",
    "ScriptScanner",
    "ExcelExporter",
    "build_report",
    "main",
]


if __name__ == "__main__":
    raise SystemExit(main())
