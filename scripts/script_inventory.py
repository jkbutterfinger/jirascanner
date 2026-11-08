#!/usr/bin/env python3
"""
Script Inventory Scanner - Analysiert Reporting-Skripte
Verknüpft mit Jira-Tasks
"""

import os
import re
from pathlib import Path
from typing import Dict, List
from datetime import datetime
import json


class ScriptAnalyzer:
    """Analyse Reporting-Skripte"""
    
    def scan_directory(self, path: str) -> List[Dict]:
        """Scan gesamte Verzeichnisstruktur"""
        found_files = []
        scripts_path = Path(path)
        
        if not scripts_path.exists():
            print(f"⚠️  Pfad nicht gefunden: {path}")
            return found_files
        
        for file in scripts_path.rglob('*'):
            if file.is_file() and (file.suffix in ['.py', '.sh', '.sql', '.js', '.ts']):
                info = self.analyze_file(file)
                found_files.append(info)
                
        return found_files
    
    def analyze_file(self, file_path: Path) -> Dict:
        """Analysiere einzelne Datei"""
        try:
            content = file_path.read_text(encoding='utf-8', errors='replace')
            
            analysis = {
                'path': str(file_path),
                'name': file_path.name,
                'size': file_path.stat().st_size,
                'lines': len(content.strip().split('\n')),
                'timestamp': datetime.now().isoformat(),
                'jira_refs': [],
                'kpi_tags': [],
                'db_access': [],
                'api_calls': [],
                'external_services': [],
                'complexity': self.estimate_complexity(content),
            }
            
            # Jira Referenzen extrahieren
            jira_refs = re.findall(r'\b([A-Z]{2,4}-\d{4,6})\b', content)
            analysis['jira_refs'] = list(set(jira_refs))
            
            # KPI/Dashboard Tags
            kpi_patterns = [
                (r'[Kk]pi[_:]?\s*([^\s,:\n]+)', 'kpi'),
                (r'[Dd]ashboard[_:]?\s*([^\s,:\n]+)', 'dashboard'),
                (r'[Mm]etric[_:]?\s*([^\s,:\n]+)', 'metric'),
                (r'[Rr]eport[_:]?\s*([^\s,:\n]+)', 'report'),
            ]
            
            for pattern, category in kpi_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    analysis['kpi_tags'] += matches
            
            # Datenbankzugriff
            db_patterns = [
                r'DATABASE_URL.*?(?:postgres|mysql|sqlite)',
                r'mysql://[^"]+',
                r'postgres://[^"]+',
                r'sqlite://[^"]+',
                r'CONNECTION_STRING.*?(?:["]\\w+[")]\\w+[")]\\w+["])',
            ]
            for pattern in db_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                if matches:
                    analysis['db_access'] = list(set(matches))
            
            # API-Aufrufe/Eexterne Dienste
            api_patterns = [
                r'api\.requests\.(?:get|post|put|delete)',
                r'requests\.(?:get|post|put)',
                r'urllib?\.(?:request|context)',
                r'Requests\.(?:GET|POST|PUT|DELETE)',
            ]
            api_matches = []
            for pattern in api_patterns:
                matches = re.findall(pattern, content, re.IGNORECASE)
                api_matches.extend(matches)
            analysis['api_calls'] = list(set(api_matches))
            
            return analysis
            
        except Exception as e:
            return {
                'path': str(file_path),
                'error': str(e)
            }
    
    def estimate_complexity(self, code: str) -> int:
        """Schätze Code-Komplexität"""
        # Basis-Komplexität
        lines = len(code.split('\n'))
        
        # Funktionäre Bewertung
        imports = len(re.findall(r'import\s+\w+', code))
        classes = len(re.findall(r'\bclass\s+\w+', code))
        functions = len(re.findall(r'\bdef\s+\w+\s*\(', code))
        
        return min(5, max(1, lines // 50 + imports // 10 + classes + functions // 10))


class JiraMapper:
    """Mapping zwischen Jira-Tasks und Implementierungen"""
    
    def __init__(self):
        self.tasks = []
        self.definitions = []
        self.implementations = []
        
    def load_jira_tasks(self, tasks_data: List[Dict]):
        """Lade Jira-Tasks"""
        self.tasks = tasks_data
        
    def load_implementations(self, implementations: List[Dict]):
        """Lade Implementierungs-Daten"""
        self.implementations = implementations
        
    def map_requirements(self) -> List[Dict]:
        """Erstelle Mapping-Matrix"""
        results = []
        
        for task in self.tasks:
            matching_implementations = []
            missing_features = []
            partially_implemented = []
            
            # Prüfe Referenzen
            task_refs = set()
            for ref_pattern in [r'[A-Z]+-\d+', r'(?:PROJ|TASK|ISSUE)-\d+', r'#\d+', r'Issue:\s*[A-Z]+-\d+']:
                task_refs.update(re.findall(ref_pattern, task['description'] or ''))
            
            for impl in self.implementations:
                if task['key'] in str(impl.get('jira_refs', [])):
                    matching_implementations.append(impl)
            
            # Prüfe auf Lücken
            if task['status'] == 'in progress' and len(missing_features) > 0:
                partially_implemented.append({
                    'task': task['key'],
                    'missing': missing_features
                })
            
            results.append({
                'task': task['key'],
                'summary': task['summary'],
                'status': task['status'],
                'implementation_count': len(matching_implementations),
                'missing_features': len(missing_features),
                'fully_implemented': len(missing_features) == 0,
            })
        
        return results


class ReportGenerator:
    """Generate Reports und Analyse"""
    
    def generate_summary(self, inventory: List[Dict], jira_tasks: List[Dict]) -> Dict:
        """Generiere Gesamtübersicht"""
        return {
            'total_scripts': len(inventory),
            'total_tasks': len(jira_tasks),
            'scripts_with_jira_refs': sum(1 for s in inventory if s.get('jira_refs')),
            'date': datetime.now().isoformat()
        }
    
    def generate_gap_report(self, mapping: List[Dict]) -> str:
        """Generiere Gap Report"""
        gaps = [item for item in mapping if not item['fully_implemented']]
        
        report = "# Gap Report\n\n"
        report += f"## Total Gaps: {len(gaps)}\n\n"
        
        for gap in gaps:
            report += f"### {gap['task']} - {gap['summary']}\n"
            report += f"*Status*: {gap['status']}\n"
            report += f"*Implementation Count*: {gap['implementation_count']}\n"
            report += f"*Missing Features*: {gap['missing_features']}\n\n"
        
        return report


def main():
    """Hauptausführung"""
    # Scan Scripts
    script_paths = [
        '/home/user1/work/jirascanner/scripts/adhoc',
        '/home/user1/work/jirascanner/scripts/reports',
    ]
    
    all_scripts = []
    for path in script_paths:
        if os.path.exists(path):
            all_scripts.extend(ScriptAnalyzer().scan_directory(path))
    
    # Generiere Report
    summary = ReportGenerator().generate_summary(all_scripts, [])
    print(json.dumps(summary, indent=2))


if __name__ == '__main__':
    main()