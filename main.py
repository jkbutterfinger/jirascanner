#!/usr/bin/env python3
"""
Jira/Gira Scanner Main API Entry Point
"""
import os
import sys
from pathlib import Path

# Add the parent directory to the Python path
sys.path.insert(0, str(Path(__file__).parent))

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import logging
from typing import Optional
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Import modules
try:
    from scanner.inventorizer import Inventorizer
    from scanner.scanner import Scanner
    from api.suggestion_api import SuggestionAPI
    from database.repositories import RequirementRepository, ImplementationRepository
    from database.models import Requirement, Implementation
    print("✅ All modules loaded successfully")
except ImportError as e:
    logger.error(f"Failed to import modules: {e}")
    sys.exit(1)


# FastAPI Application
app = FastAPI(
    title="Jira/Gira Scanner API",
    description="An API for scanning Jira tasks and matching them with existing implementations",
    version="1.0.0"
)

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Dependencies
async def check_auth():
    """Authentication chekck - Placeholder for now"""
    pass  # In production, add real auth logic


class JiraProjectConfig(BaseModel):
    """Jira project configuration"""
    server_url: str
    api_token: str
    project_key: str
    space_key: str


class SearchRequest(BaseModel):
    """Search request model"""
    query: str
    max_results: int = 100


class SearchResponse(BaseModel):
    """Search response model"""
    tasks: list
    total_count: int


# Routes
@app.get("/")
async def root():
    """Root endpoint"""
    return {
        "service": "Jira/Gira Scanner API",
        "version": "1.0.0",
        "status": "running"
    }


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy"}


@app.get("/jira/search", response_model=SearchResponse)
async def search_jira_tasks(search_request: SearchRequest):
    """Search Jira tasks based on query"""
    try:
        scanner = Scanner(
            jira_args={
                "server": os.getenv("JIRA_SERVER_URL", os.getenv("GIRA_SERVER_URL")),
                "api_token": os.getenv("JIRA_API_TOKEN", os.getenv("GIRA_API_TOKEN")),
                "project_key": os.getenv("JIRA_PROJECT_KEY", "PROJ")
            }
        )
        
        results = scanner.search_tasks(query=search_request.query, max_results=search_request.max_results)
        
        return {
            "tasks": results,
            "total_count": len(results)
        }
    except Exception as e:
        logger.error(f"Error searching Jira tasks: {e}")
        return SearchResponse(tasks=[], total_count=0), 500


@app.get("/jira/tasks/{jira_task_key}")
async def get_jira_task(jira_task_key: str):
    """Get specific Jira task details"""
    try:
        scanner = Scanner(
            jira_args={
                "server": os.getenv("JIRA_SERVER_URL", os.getenv("GIRA_SERVER_URL")),
                "api_token": os.getenv("JIRA_API_TOKEN", os.getenv("GIRA_API_TOKEN")),
                "project_key": os.getenv("JIRA_PROJECT_KEY", "PROJ")
            }
        )
        
        task = scanner.get_task(task_key=jira_task_key)
        
        return {
            "task": task,
            "status": "success"
        }
    except Exception as e:
        logger.error(f"Error getting Jira task {jira_task_key}: {e}")
        return {"error": str(e), "status": "error"}, 500


@app.get("/scripts/analyze")
async def analyze_scripts():
    """Analyze scripts for requirements"""
    try:
        inventory = Inventorizer(scripts_path="/path/to/scripts")
        results = inventory.scan_all_scripts()
        
        return {
            "analyze_results": results,
            "status": "success"
        }
    except Exception as e:
        logger.error(f"Error analyzing scripts: {e}")
        return {"error": str(e), "status": "error"}, 500


@app.post("/requirements/match")
async def match_requirements_and_implementation(request: MatchRequest):
    """Match requirements with implementations"""
    try:
        # Create scanner with database
        inventory = Inventorizer(scripts_path=request.script_path)
        scanner = Scanner(
            jira_args={
                "server": os.getenv("JIRA_SERVER_URL", os.getenv("GIRA_SERVER_URL")),
                "api_token": os.getenv("JIRA_API_TOKEN", os.getenv("GIRA_API_TOKEN")),
                "project_key": os.getenv("JIRA_PROJECT_KEY", "PROJ")
            }
        )
        
        # Load tasks
        tasks = scanner.get_tasks(project_key=request.project_key)
        
        # Match requirements with implementations
        matches = scanner.match_requirements(
            requirements=request.requirements,
            implementations=inventory.get_script_info()
        )
        
        return {
            "matches": matches,
            "status": "success"
        }
    except Exception as e:
        logger.error(f"Error matching requirements: {e}")
        return {"error": str(e), "status": "error"}, 500


@app.post("/inventory/upload")
async def upload_suggestion(suggestion: Suggestion):
    """Upload a new suggestion"""
    try:
        db = db.connect_or_create()
        
        # Create suggestion in database
        suggestion_record = db.create_suggestion(suggestion_data={
            "title": suggestion.title,
            "description": suggestion.description,
            "task_id": suggestion.task_id,
            "severity": suggestion.severity,
            "category": suggestion.category,
            "file_path": suggestion.file_path,
            "created_at": datetime.now().isoformat()
        })
        
        db.commit()
        
        return {
            "status": "success",
            "id": suggestion_record.id
        }
    except Exception as e:
        logger.error(f"Error uploading suggestion: {e}")
        return {"error": str(e), "status": "error"}, 500


@app.get("/inventory/{suggestion_id}")
async def get_suggestion(suggestion_id: str):
    """Get suggestion details"""
    try:
        db = db.connect_or_create()
        suggestion = db.get_suggestion_by_id(suggestion_id)
        
        return {
            "suggestion": suggestion.to_dict(),
            "status": "success"
        }
    except Exception as e:
        logger.error(f"Error getting suggestion {suggestion_id}: {e}")
        return {"error": str(e), "status": "error"}, 500


if __name__ == "__main__":
    try:
        # Initialize database connection
        db = db.connect_or_create()
        
        # Mount suggestion API under /v1/suggestions
        mounted_api = SuggestionAPI(
            api_endpoint="/v1/suggestions",
            db=db
        )
        app.router.routes.extend(mounted_api.routes)
        
        # Log available routes
        logger.info("Available routes:")
        for route in app.routes:
            if hasattr(route, "path"):
                logger.info(f"  {route.methods} {route.path}")
        
        logger.info("Starting Jira/Gira Scanner API on http://0.0.0.0:8000")
        uvicorn.run(app, host="0.0.0.0", port=8000)
        
    except Exception as e:
        logger.error(f"Failed to start server: {e}")
        sys.exit(1)