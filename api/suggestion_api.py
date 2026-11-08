#!/usr/bin/env python3
"""
Suggestion System API - Web Service für Requirements Matching und Gap Analysis
Buildet auf FastAPI + SQLAlchemy
"""

import os
import json
from datetime import datetime
from typing import List, Dict, Any
from pydantic import BaseModel, Field
from fastapi import FastAPI, HTTPException, Depends
from fastapi.middleware.cors import CORSMiddleware
import uvicorn

# Models
class JiraTask(BaseModel):
    """Jira Task Model"""
    key: str = Field(..., description="Task ID z.B. PROJ-345")
    summary: str = Field(..., description="Task Zusammenfassung")
    description: str = Field(default="", description="Task Beschreibung")
    status: str = Field(default="open", description="Aktueller Status")
    type: str = Field(default="story", description="Issue Type")
    project_key: str = Field(default="PROJ", description="Projekt Key")
    priority: str = Field(default="medium", description="Priorität")
    jql: str = Field(default="", description="Zugehörige JQL Abfrage")
    created: str = Field(default="", description="Erstellungs-Datum")
    requirements: List[Dict[str, Any]] = Field(default=[], description="Extrahierte Anforderungen")
    confluence_docs: List[str] = Field(default=[], description="Confluence Dokument-Pfade")


class ScriptImplementation(BaseModel):
    """Implementation Script Model"""
    file_path: str = Field(..., description="Pfad zur Script-Datei")
    name: str = Field(..., description="Script Name")
    lines: int = Field(..., description="Anzahl Zeilen")
    size_bytes: int = Field(..., description="Größe in Bytes")
    jira_references: List[str] = Field(default=[], description="Verwiesene Jira-Tasks")
    kpi_tags: List[str] = Field(default=[], description="Erkannte KPI/Dashboard Tags")
    database_connections: List[str] = Field(default=[], description="Datenbank-Verbindungen")
    api_imports: List[str] = Field(default=[], description="Externe API-Aufrufe")
    complexity: int = Field(default=1, description="Geschätzte Komplexität (1-5)")


class Suggestion(BaseModel):
    """Suggestion Model"""
    title: str = Field(..., description="Suggestion Titel")
    description: str = Field(..., description="Beschreibung")
    task_id: str = Field(..., description="Zugehöriger Jira Task")
    severity: str = Field(default="medium", description="Schweregrad")
    category: str = Field(default="feature", description="Kategorie")
    file_path: Optional[str] = Field(default=None, description="Zugehöriges Script")


class GapAnalysis(BaseModel):
    """Gap Analysis Model"""
    task_key: str = Field(..., description="Zu analysierender Task")
    summary: str = Field(..., description="Task Zusammenfassung")
    has_implementation: bool = Field(..., description="Hat Implementation")
    implementation_path: Optional[str] = Field(default=None, description="Pfad zur Implementation")
    missing_requirements: List[str] = Field(default=[], description="Fehlende Anforderungen")
    completion_percentage: float = Field(default=0, description="Abgeschlossener Anteil %")


class Requirement(BaseModel):
    """Requirement Model"""
    id: str = Field(default="", description="Anforderungs-ID")
    task_id: str = Field(..., description="Zugehöriger Task")
    description: str = Field(..., description="Anforderungsbeschreibung")
    type: str = Field(default="functional", description="Typ")  # functional, technical, non-functional
    acceptance_criteria: Optional[str] = Field(default=None, description="Akzeptanzkriterien")
    priority: str = Field(default="medium", description="Priorität")


# Database Interface (vereinfacht)
class SuggestionDB:
    """In-Memory Datenbank für Suggestions"""
    
    def __init__(self, db_path: str):
        self.db_path = db_path
        self.database = self._load_or_create()
    
    def _load_or_create(self) -> Dict:
        """Lade oder erstelle Datenbank"""
        if os.path.exists(self.db_path):
            try:
                with open(self.db_path, 'r') as f:
                    return json.load(f)
            except:
                return defaultdict(lambda: {"tasks": [], "suggestions": [], "scripts": []})
        return defaultdict(lambda: {"tasks": [], "suggestions": [], "scripts": []})
    
    def save(self):
        """Speichere Datenbank"""
        os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
        with open(self.db_path, 'w') as f:
            json.dump(self.database, f, indent=2, default=str)


# FastAPI App Initialisierung
app = FastAPI(
    title="Jira Scanner API",
    description="API für Task Scanner und Requirement Matching",
    version="1.0.0"
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Database Instance
db = SuggestionDB(db_path="/home/user1/work/jirascanner/inventory/suggestions.json")

# Utility Functions
def extract_jira_references(text: str) -> List[str]:
    """Extrahiere Jira-ID aus Text"""
    pattern = r'\b([A-Z]+-\d{3,6})\b'
    return re.findall(pattern, text)


def extract_requirements(text: str) -> List[Dict]:
    """Extrahiere Anforderungen aus Text"""
    requirements = []
    
    # Suche nach "Acceptance:" Pattern
    acceptance_section = re.search(r'Acceptance:[\s\S]*?', text)
    if acceptance_section:
        lines = acceptance_section.group().strip().split('\n')
        for line in lines:
            if line.strip():
                requirements.append({
                    'id': f'ACC-{len(requirements) + 1}',
                    'type': 'acceptance',
                    'description': line.strip(),
                    'priority': 'high'
                })
    
    # Suche nach "Technical:" Pattern
    technical_section = re.search(r'Technical:[\s\S]*?', text)
    if technical_section:
        lines = technical_section.group().strip().split('\n')
        for line in lines:
            if line.strip():
                requirements.append({
                    'id': f'TEC-{len(requirements) + 1}',
                    'type': 'technical',
                    'description': line.strip(),
                    'priority': 'medium'
                })
    
    # Suche nach "Perfomance:" Pattern  
    performance_section = re.search(r'Performance:[\s\S]*?', text)
    if performance_section:
        lines = performance_section.group().strip().split('\n')
        for line in lines:
            if line.strip():
                requirements.append({
                    'id': f'PERF-{len(requirements) + 1}',
                    'type': 'non-functional',
                    'description': line.strip(),
                    'priority': 'medium'
                })
    
    return requirements


@app.get("/")
async def root():
    """API Willkommen"""
    return {
        "service": "Jira/Gira Scanner API",
        "version": "1.0.0",
        "status": "running",
        "endpoints": [
            "/tasks/",
            "/scripts/",
            "/suggestions/",
            "/gaps/"
        ]
    }


@app.get("/tasks")
async def get_tasks(jql: str = ..., project: str = ...):
    """Hole Jira-Tasks"""
    try:
        # Jira API Aufruf würde hier sein
        tasks = []
        
        return tasks
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/tasks")
async def create_task(task: JiraTask):
    """Erstelle neuen Task"""
    # Task würde in Database gespeichert
    return {"status": "success", "message": f"Task {task.key} erstellt"}


@app.get("/scripts")
async def get_scripts(path: str = "/home/user1/work/jirascanner/scripts"):
    """Hole alle Scanner-Scripts"""
    try:
        from scripts.script_inventory import ScriptAnalyzer
        analyzer = ScriptAnalyzer()
        
        # Scan directory
        scripts = analyzer.scan_directory(path)
        return {"scripts": scripts}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/scripts/analyze")
async def analyze_script(path: str, content: str = None):
    """Analysiere einzelnes Script"""
    try:
        from scripts.script_inventory import ScriptAnalyzer
        analyzer = ScriptAnalyzer()
        
        script_info = analyzer.analyze_file(Path(path))
        
        if content and script_info.get('path') == str(path):
            # Update mit übergebenem Inhalt
            script_info['content_length'] = len(content)
        
        return {"script": script_info}
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/suggestions")
async def get_suggestions():
    """Erzie alle Suggestions"""
    if "tasks" not in db.database:
        return {"suggestions": [], "count": 0}
    
    return {
        "suggestions": db.database["tasks"].get("suggestions", []),
        "count": len(db.database["tasks"].get("suggestions", []))
    }


@app.post("/suggestions")
async def create_suggestion(suggestion: Suggestion):
    """Erstelle Suggestion"""
    # Validation
    if not suggestion.title or not suggestion.description:
        raise HTTPException(status_code=400, detail="Title und Description sind required")
    
    # Speichere Suggestion
    db.database["tasks"]["suggestions"].append({
        "data": {
            "id": f"SUG-{datetime.now().timestamp():.0f}",
            "title": suggestion.title,
            "description": suggestion.description,
            "task_id": suggestion.task_id,
            "severity": suggestion.severity,
            "category": suggestion.category,
            "file_path": suggestion.file_path,
            "created": datetime.now().isoformat(),
            "status": "pending"
        }
    })
    db.save()
    
    return {"status": "success", "id": f"SUG-{datetime.now().timestamp():.0f}"}


@app.post("/suggestions/batch")
async def batch_create_suggestions(suggestions: List[Suggestion]):
    """Batch Create Suggestions"""
    results = []
    for suggestion in suggestions:
        result = await create_suggestion(suggestion)
        results.append(result)
    return {"status": "success", "created": len(results)}


@app.get("/gaps")
async def analyze_gaps(task_keys: List[str] = []):
    """Analyse Gaps zwischen Requirements und Implementation"""
    if not task_keys:
        task_keys = [""]
    
    gaps = []
    
    for task_key in task_keys:
        # Gap Analysis würde hier implementiert werden
        # Prüfte Implementation vs Requirements
        
        gap = {}
        if task_key:
            # Gap-Analyse Logik
            pass
        
        gaps.append(gap)
    
    return {"gaps": gaps}


@app.get("/match")
async def match_requirements_and_implementation(task_key: str, script_path: str = None):
    """Match Requirements mit Implementation"""
    # Match-Logik würde hier sein
    
    matches = {
        "task": task_key,
        "script": script_path,
        "match_score": 0.0,
        "matching_requirements": [],
        "missing_requirements": [],
        "partially_implemeted": False
    }
    
    return matches


# Startup Events
@app.on_event("startup")
async def startup():
    """Startup initialisierung"""
    print("Jira/Gira Scanner API starting...")


# Health Check
@app.get("/health")
async def health_check():
    """Health Check"""
    return {"status": "healthy"}


# Run Local
if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)