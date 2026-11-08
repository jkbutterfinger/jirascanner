#!/usr/bin/env python3
"""
GIRA/Jira Task Scanner - Ein vollständiges Inventar-System für Ad-hoc Reports und Jira-Tasks
Version 1.0
Authored for German Reporting System
"""

import json
import os
import re
from datetime import datetime
from typing import List, Dict, Optional
from pathlib import Path
import argparse


class JiraScanner:
    """
    Scanner für Jira/Gira Tasks
    Extrahiert Aufgaben und generiert Inventar
    """
    
    def __init__(self, server: str, api_token: str, project_key: str, config_path: Optional[str] = None):
        """
        Initialisiere Scanner mit API-Zugangsdaten
        
        Args:
            server: Jira/Gira Server URL (https://...)
            api_token: API Authentication Token
            project_key: Project-Schlüssel (z.B. PROJ)
            config_path: Pfad zu optionaler Konfigurationsdatei
        """
        self.server = server.strip('/')
        self.api_token = api_token
        self.project_key = project_key.upper()
        self.base_url = f'{self.server}/rest/api/3'
        self.config_path = config_path
        self.inventory_db_path = os.path.join(Path(__file__).parent, 'inventory')
        
    def create_inventory_structure(self):
        """Erstellt Basis-Ordnerstruktur für Inventarisierung"""
        self.inventory_db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'inventory')
        os.makedirs(self.inventory_db_path, exist_ok=True)
        
        # Sicherstellen, dass das Verzeichnis existiert
        os.makedirs(self.inventory_db_path, exist_ok=True)
    
    def fetch_all_tasks(self, jql_query: str = None) -> Dict[str, any]:
        """
        Hole alle Jira Tasks
        
        Args:
            jql_query: JQL-Abfrage für erweiterte Suche
        
        Returns:
            Dictionary mit allen Tasks
        """
        return {}
    
    def scan_project(self) -> List[Dict]:
        """Scan alle Tasks im Projekt"""
        if not self.config_path:
            print("⚠️  Keine Konfiguration gefunden. Bitte konfigurieren Sie den Scanner.")
            return []
        
        with open(self.config_path, 'r') as f:
            config = json.load(f)
            
        tasks = []
        page_limit = config.get('max_results', 1000)
        
        # JQL Abfrage
        jql = f'project={self.project_key}' + f' ORDER BY created DESC'
        
        # Aufgabe wäre hier: Jira API Aufruf für echte Implementierung
        print(f"Scanned Tasks: {len(tasks)}")
        return tasks


class ScriptInventoryScanner:
    """
    Scanner für Python-Scripts und andere Reporting-Dateien
    """
    
    def __init__(self, root_path: str):
        self.root_path = Path(root_path)
        
    def find_report_scripts(self, extensions: List[str]) -> List[Path]:
        """Finde alle Reporting-Skripte"""
        found_files = []
        
        for ext in extensions:
            found_files.extend(
                self.root_path.rglob(f'*{ext}')
            )
        return found_files
    
    def extract_script_info(self, file_path: Path) -> Dict:
        """Extrahiere Informationen aus Script-Datei"""
        try:
            content = file_path.read_text(encoding='utf-8', errors='ignore')
            
            script_info = {
                'file_path': str(file_path),
                'file_name': file_path.name,
                'line_count': len(content.splitlines()),
                'size_bytes': file_path.stat().st_size,
                'last_modified': datetime.fromtimestamp(file_path.stat().st_mtime).isoformat(),
            }
            
            # Extrahiere Kommentare (Jira Referenzen)
            jira_refs = re.findall(r'[A-Z]+-\d+', content)
            script_info['jira_references'] = list(set(jira_refs))
            
            # Extrahiere Datenbank-Verbindungen
            db_patterns = [
                r'(?:mysql|postgres|sqlite)://[^ ]+',
                r'DATABASE_URL\s*=\s*["\']([^"\']+)["\']',
                r'(?:DB_|DATABASE_)[A-Z_]+)'
            ]
            
            db_connections = []
            for pattern in db_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    db_connections.extend(matches)
            
            script_info['database_connections'] = list(set(db_connections))
            
            # Extrahiere Importe
            script_info['imports'] = re.findall(r'^(import\s+\w+)', content, re.MULTILINE)
            
            # Extrahiere KPIs oder Metriken
            kpi_patterns = [
                r'[Kk]pi[_:]\s*([^\s]+)',
                r'[Mm]etric.*?([^\s:]+)',
                r'Dashboard.*?([^\s:]+)',
                r'Report.*?([^\s:]+)'
            ]
            
            kpi_matches = []
            for pattern in kpi_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                kpi_matches.extend(matches)
            
            script_info['kpi_detected'] = list(set([k.lower() for k in kpi_matches]))
            
            return script_info
            
        except Exception as e:
            print(f"⚠️  Fehler beim Parsen von {file_path}: {e}")
            return {
                'file_path': str(file_path),
                'error': str(e)
            }


class ConfluenceDocumentParser:
    """
    Parser für Confluence-Dokumente
    """
    
    def __init__(self, base_url: str, api_token: str):
        self.base_url = base_url
        self.api_token = api_token
        self.headers = {
            'Authorization': f'Basic {base64.b64encode(f"username:api_token".encode()).decode()}'
        }


class DatabaseManager:
    """Verwaltet Inventar-Datenbank"""
    
    def __init__(self, db_path: str = None):
        self.db_path = db_path or '/home/user1/work/jirascanner/inventory/inventory.json'
        
    def load_or_create(self) -> Dict:
        """Lade oder erstelle Inventar"""
        if os.path.exists(self.db_path):
            with open(self.db_path, 'r') as f:
                return json.load(f)
        return {'tasks': [], 'scripts': [], 'confluence_docs': [], 'gaps': []}
    
    def save(self):
        # Speichern würde hier implementiert werden


# Beispiel Hauptprogramm
def main():
    """Hauptausführung"""
    parser = argparse.ArgumentParser(description='Jira/Gira Aufgaben-Scanner')
    parser.add_argument('--server', required=True, help='Server URL')
    parser.add_argument('--token', required=True, help='API Token')
    parser.add_argument('--project', required=True, help='Project Key')
    parser.add_argument('--config', help='Pfad zu Konfigurationsdatei')
    parser.add_argument('--scripts', default='/home/user1/work/jirascanner/scripts', help='Pfad zu Scripts')
    parser.add_argument('--confluence', help='Confluence URLs')
    
    args = parser.parse_args()
    
    # Scanner initialisieren
    scanner = JiraScanner(
        server=args.server,
        api_token=args.token,
        project_key=args.project
    )
    
    script_scanner = ScriptInventoryScanner(args.scripts)
    
    # Inventar erstellen
    script_files = script_scanner.find_report_scripts(['.py'])
    for script in script_files:
        info = script_scanner.extract_script_info(script)
        print(f"Inventarisiert: {info['file_name']}")


if __name__ == '__main__':
    main()