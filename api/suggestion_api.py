#!/usr/bin/env python3
"""FastAPI-Service für Task-, SQL- und Suggestion-Daten.

Der Service stellt die fachlichen Endpunkte unter ``/api`` bereit und liefert
das Dashboard unter ``/`` aus. Er nutzt dieselbe Konfigurationsschicht und
dieselben SQL-Analyseklassen wie die CLI, damit Dashboard und Scanner immer
dieselben Daten sehen.

Start:
    python -m api.suggestion_api
    # oder
    uvicorn api.suggestion_api:app --reload
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import uuid
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import uvicorn
from fastapi import APIRouter, FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

# Direktaufruf "python api/suggestion_api.py" ermöglichen.
_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from config import ROOT, load_config, resolve_path  # noqa: E402
from scripts.script_inventory import ScriptAnalyzer  # noqa: E402

LOGGER = logging.getLogger("jira_scanner.api")
CONFIG = load_config()

# --------------------------------------------------------------------------- Models


class JiraTask(BaseModel):
    """Jira-Task-Modell für die JSON-Ablage."""

    key: str = Field(..., description="Task-ID, z. B. JRA-345")
    summary: str = Field(..., description="Task-Zusammenfassung")
    description: str = Field(default="", description="Task-Beschreibung")
    status: str = Field(default="open", description="Aktueller Status")
    type: str = Field(default="story", description="Issue Type")
    project_key: str = Field(default="JRA", description="Projekt-Key")
    priority: str = Field(default="medium", description="Priorität")
    jql: str = Field(default="", description="Zugehörige JQL-Abfrage")
    created: str = Field(default="", description="Erstellungsdatum")
    requirements: List[Dict[str, Any]] = Field(default_factory=list)
    confluence_docs: List[str] = Field(default_factory=list)


class Suggestion(BaseModel):
    """Suggestion-Modell."""

    title: str = Field(..., min_length=1)
    description: str = Field(..., min_length=1)
    task_id: str = Field(..., min_length=1)
    severity: str = Field(default="medium")
    category: str = Field(default="feature")
    file_path: Optional[str] = None


class GapAnalysis(BaseModel):
    """Ergebnis der einfachen Gap-Analyse."""

    task_key: str
    summary: str = ""
    has_implementation: bool
    implementation_path: Optional[str] = None
    missing_requirements: List[str] = Field(default_factory=list)
    completion_percentage: float = 0


# --------------------------------------------------------------------------- Storage


class SuggestionStore:
    """Kleine JSON-Datenbank für Tasks und Suggestions."""

    def __init__(self, db_path: str | Path):
        self.db_path = Path(db_path)
        self.database = self._load_or_create()

    @staticmethod
    def _empty() -> Dict[str, List[Dict[str, Any]]]:
        return {"tasks": [], "suggestions": [], "scripts": []}

    def _load_or_create(self) -> Dict[str, List[Dict[str, Any]]]:
        if self.db_path.exists():
            try:
                with open(self.db_path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)
                if isinstance(data, dict):
                    return {**self._empty(), **data}
            except (OSError, ValueError) as exc:
                LOGGER.warning("Suggestions-DB konnte nicht gelesen werden: %s", exc)
        return self._empty()

    def save(self) -> None:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.db_path, "w", encoding="utf-8") as handle:
            json.dump(self.database, handle, indent=2, ensure_ascii=False)

    def list_suggestions(self) -> List[Dict[str, Any]]:
        return list(self.database.get("suggestions", []))

    def list_tasks(self) -> List[Dict[str, Any]]:
        return list(self.database.get("tasks", []))

    def add_suggestion(self, suggestion: Suggestion) -> Dict[str, Any]:
        entry = {
            "id": f"SUG-{uuid.uuid4().hex[:8]}",
            "title": suggestion.title,
            "description": suggestion.description,
            "task_id": suggestion.task_id,
            "severity": suggestion.severity,
            "category": suggestion.category,
            "file_path": suggestion.file_path,
            "created": datetime.now().isoformat(timespec="seconds"),
            "status": "pending",
        }
        self.database.setdefault("suggestions", []).append(entry)
        self.save()
        return entry

    def add_task(self, task: Dict[str, Any]) -> None:
        self.database.setdefault("tasks", []).append(task)
        self.save()


def _default_db_path() -> Path:
    env_path = os.environ.get("JIRA_SCANNER_DB")
    if env_path:
        return Path(env_path)
    storage = CONFIG.get("storage", {})
    inventory_dir = resolve_path(storage.get("inventory_path", ROOT / "inventory"))
    return inventory_dir / str(storage.get("suggestions_file", "suggestions.json"))


store = SuggestionStore(_default_db_path())


# --------------------------------------------------------------------------- Helpers


def extract_jira_references(text: str) -> List[str]:
    """Extrahiert Jira-Keys (z. B. JRA-1234) aus Text."""
    return re.findall(r"\b([A-Z][A-Z0-9]+-\d+)\b", text or "")


def extract_requirements(text: str) -> List[Dict[str, Any]]:
    """Extrahiert einfache Anforderungen aus Abschnitten wie ``Acceptance:``."""
    requirements: List[Dict[str, Any]] = []
    sections = {
        "Acceptance": ("acceptance", "high"),
        "Technical": ("technical", "medium"),
        "Performance": ("non-functional", "medium"),
    }
    for heading, (kind, priority) in sections.items():
        match = re.search(rf"{heading}:\s*([\s\S]*?)(?:\n[A-Z][a-z]+:|$)", text or "")
        if not match:
            continue
        for line in match.group(1).splitlines():
            line = line.strip(" -\t")
            if line:
                requirements.append(
                    {
                        "id": f"{kind[:3].upper()}-{len(requirements) + 1}",
                        "type": kind,
                        "description": line,
                        "priority": priority,
                    }
                )
    return requirements


def _script_path() -> Path:
    return resolve_path(CONFIG.get("scripts", {}).get("base_path", ROOT / "scripts"))


def _analyzer() -> ScriptAnalyzer:
    return ScriptAnalyzer(CONFIG)


def _cached_tasks() -> List[Dict[str, Any]]:
    cache = resolve_path(
        CONFIG.get("storage", {}).get("jira_cache", ROOT / "jira_data" / "tasks.json")
    )
    if not cache.exists():
        return []
    try:
        with open(cache, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            return list(data.get("tasks", []))
    except (OSError, ValueError) as exc:
        LOGGER.warning("Task-Cache konnte nicht gelesen werden: %s", exc)
    return []


def _index_scripts(scripts: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    """Indexiert SQL-Inventar-Einträge nach Task-Key."""
    index: Dict[str, List[Dict[str, Any]]] = {}
    for script in scripts:
        references = set(script.get("jira_refs", []))
        if script.get("task_key"):
            references.add(script["task_key"])
        for reference in references:
            index.setdefault(str(reference).upper(), []).append(script)
    return index


# --------------------------------------------------------------------------- App


@asynccontextmanager
async def lifespan(_app: FastAPI):
    LOGGER.info("Jira/Gira Scanner API gestartet")
    yield
    LOGGER.info("Jira/Gira Scanner API beendet")


app = FastAPI(
    title="Jira Scanner API",
    description="API für Task-Scanner, SQL-Inventar und Requirements-Matching",
    version="1.3.0",
    lifespan=lifespan,
)

origins = list(CONFIG.get("api", {}).get("cors_origins") or [])
if not origins:
    origins = ["http://localhost:8000", "http://127.0.0.1:8000"]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials="*" not in origins,
    allow_methods=["*"],
    allow_headers=["*"],
)

router = APIRouter(prefix="/api")


@router.get("/tasks")
def get_tasks(
    project: Optional[str] = None,
    jql: Optional[str] = None,
    refresh: bool = False,
) -> Dict[str, Any]:
    """Liefert gecachte Tasks oder lädt sie bei ``refresh=true`` aus Jira."""
    if refresh:
        from main import JiraClient

        config = dict(CONFIG)
        if project:
            config["project"] = {**config.get("project", {}), "key": project}
        if jql:
            config["project"] = {**config.get("project", {}), "jql": jql}
        tasks = JiraClient(config).search_tasks()
        return {
            "tasks": [task.to_dict() for task in tasks],
            "count": len(tasks),
            "source": "jira",
        }

    tasks = store.list_tasks() + _cached_tasks()
    return {"tasks": tasks, "count": len(tasks), "source": "cache"}


@router.post("/tasks")
def create_task(task: JiraTask) -> Dict[str, Any]:
    store.add_task(task.model_dump())
    return {"status": "success", "message": f"Task {task.key} gespeichert"}


@router.get("/scripts")
def get_scripts(path: Optional[str] = None) -> Dict[str, Any]:
    """Scannt das SQL-Verzeichnis und liefert das Inventar."""
    target = resolve_path(path) if path else _script_path()
    scripts = _analyzer().analyze_directory(target)
    return {"scripts": scripts, "count": len(scripts), "path": str(target)}


@router.post("/scripts/analyze")
def analyze_script(path: str) -> Dict[str, Any]:
    """Analysiert eine einzelne SQL-Datei."""
    target = resolve_path(path)
    if not target.exists():
        raise HTTPException(status_code=404, detail=f"Datei nicht gefunden: {target}")
    return {"script": _analyzer().analyze(target)}


@router.get("/suggestions")
def get_suggestions() -> Dict[str, Any]:
    suggestions = store.list_suggestions()
    return {"suggestions": suggestions, "count": len(suggestions)}


@router.post("/suggestions")
def create_suggestion(suggestion: Suggestion) -> Dict[str, Any]:
    entry = store.add_suggestion(suggestion)
    return {"status": "success", "suggestion": entry}


@router.post("/suggestions/batch")
def batch_create_suggestions(suggestions: List[Suggestion]) -> Dict[str, Any]:
    created = [store.add_suggestion(suggestion) for suggestion in suggestions]
    return {"status": "success", "created": len(created), "suggestions": created}


@router.get("/gaps")
def analyze_gaps(
    task_keys: Optional[List[str]] = Query(default=None),
) -> Dict[str, Any]:
    """Vergleicht Tasks mit vorhandenen SQL-Implementierungen."""
    index = _index_scripts(_analyzer().analyze_directory(_script_path()))

    keys = [key.strip().upper() for key in (task_keys or []) if key.strip()]
    if not keys:
        keys = sorted(index)

    gaps: List[Dict[str, Any]] = []
    for key in keys:
        implementations = index.get(key, [])
        gaps.append(
            GapAnalysis(
                task_key=key,
                summary="",
                has_implementation=bool(implementations),
                implementation_path=(
                    implementations[0]["path"] if implementations else None
                ),
                missing_requirements=(
                    [] if implementations else ["Keine Implementierung gefunden"]
                ),
                completion_percentage=100.0 if implementations else 0.0,
            ).model_dump()
        )
    return {"gaps": gaps, "count": len(gaps)}


@router.get("/match")
def match_requirements_and_implementation(
    task_key: str, script_path: Optional[str] = None
) -> Dict[str, Any]:
    """Prüft, ob eine SQL-Datei den Task referenziert."""
    index = _index_scripts(_analyzer().analyze_directory(_script_path()))
    candidates = index.get(task_key.strip().upper(), [])
    if script_path:
        wanted = resolve_path(script_path)
        candidates = [
            script
            for script in candidates
            if Path(script["path"]).resolve() == wanted
        ]
    return {
        "task": task_key,
        "script": candidates[0]["path"] if candidates else script_path,
        "match_score": 1.0 if candidates else 0.0,
        "matching_requirements": [],
        "missing_requirements": [] if candidates else ["Keine Implementierung gefunden"],
        "partially_implemented": False,
    }


@router.get("/health")
def health_check() -> Dict[str, Any]:
    return {
        "status": "healthy",
        "version": app.version,
        "suggestions": len(store.list_suggestions()),
    }


app.include_router(router)


@app.get("/", include_in_schema=False)
def dashboard() -> Any:
    """Liefert das Dashboard aus, wenn vorhanden."""
    path = ROOT / "dashboard" / "dashboard.html"
    if path.exists():
        return FileResponse(path)
    return {
        "service": "Jira/Gira Scanner API",
        "version": app.version,
        "docs": "/docs",
        "endpoints": ["/api/tasks", "/api/scripts", "/api/suggestions", "/api/gaps"],
    }


def main() -> int:
    """Startet den API-Server (auch als Console-Script nutzbar)."""
    uvicorn.run(
        app,
        host=os.environ.get("API_HOST", "127.0.0.1"),
        port=int(os.environ.get("API_PORT", "8000")),
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
