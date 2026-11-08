#!/usr/bin/env python3
"""
Jira/Gira Scanner with Excel Export
====================================
Scans Jira tasks with CPU- prefix and exports to Excel

Required fields:
- Jira-Nummer
- Jira Auftraggeber
- Entwickler  
- Jahr/Monat/Tag (Completion Date)
- Zusammenfassung Jira Anforderung
- Filedatum
- File location
- Verwendete Technologien
"""

import os
import sys
import json
import yaml
import requests
import re
import base64
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any, Optional
from dataclasses import dataclass
from openpyxl import Workbook

# Configuration Path
CONFIG_PATH = Path(__file__).parent / "config.yaml"

@dataclass
class JiraIssue:
    """Jira Issue/Task Model"""
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
    labels: List[str] = None


class JiraClient:
    """Jira/Gira API Client"""
    
    def __init__(self, config_path: str = None):
        config = self._load_config(config_path or CONFIG_PATH)
        self.url = config.get("server", {}).get("url", "https://your-jira-server.com")
        self.token = config.get("api", {}).get("token", "")
        self.project_key = config.get("project", {}).get("key", "CPU")
    
    def _load_config(self, config_path: str) -> Dict:
        """Load configuration from YAML"""
        with open(config_path, 'r') as f:
            return yaml.safe_load(f)
    
    def search_tasks(self, prefix: str = "CPU-") -> List[JiraIssue]:
        """Search Jira tasks with CPU- prefix"""
        tasks = []
        jql = f'project={self.project_key} AND "summary" ~ "{prefix}"'
        
        try:
            url = f"{self.url}/rest/api/3/search"
            headers = {"Authorization": f"Bearer {self.token}"}
            params = {"jql": jql, "maxResults": 1000}
            
            response = requests.get(url, headers=headers, params=params)
            
            if response.status_code == 200:
                issues = response.json()
                for issue in issues:
                    fields = issue.get("fields", {})
                    
                    task = JiraIssue(
                        key=issue.get("key", ""),
                        issue_type=fields.get("issueType", {}).get("name", "Task"),
                        status=fields.get("status", {}).get("name", ""),
                        priority=self._get_priority(fields.get("priority", {}).get("name", "")),
                        summary=fields.get("summary", "") or fields.get("fields", {}).get("summary", ""),
                        description=fields.get("description", ""),
                        reporter=fields.get("reporter", {}).get("displayName", ""),
                        assignee=fields.get("assignee", {}).get("displayName", ""),
                        created=issue.get("created"),
                        resolved=fields.get("resolutionDate", ""),
                    )
                    tasks.append(task)
        except Exception as e:
            print(f"Error fetching Jira tasks: {e}")
        
        return tasks
    
    def _get_priority(self, priority_name: str) -> str:
        """Convert priority name"""
        priority_map = {
            "blocking": "Blocker",
            "critical": "Critical", 
            "high": "High",
            "medium": "Medium",
            "low": "Low",
            "lowest": "Low"
        }
        return priority_map.get(priority_name.lower(), "Medium")


class ScriptScanner:
    """Scan scripts and extract implementation details"""
    
    def __init__(self):
        self.config = self._load_config()
        self.scripts_path = self.config.get("scripts", {}).get("base_path", 
            "/home/user1/work/jirascanner/scripts/")
        self.extensions = self.config.get("scripts", {}).get("extensions", 
            [".py", ".sh", ".sql"])
    
    def _load_config(self):
        """Load config"""
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH, 'r') as f:
                return yaml.safe_load(f)
        return {}
    
    def find_scripts(self) -> List[Path]:
        """Find all scripts"""
        scripts = []
        base_path = Path(self.scripts_path)
        
        if base_path.exists():
            for ext in self.extensions:
                scripts.extend(base_path.rglob(f'*{ext}'))
        
        return scripts
    
    def extract_script_info(self, script_path: Path) -> Dict:
        """Extract implementation details from script"""
        info = {
            "path": str(script_path),
            "name": script_path.name,
            "size_bytes": script_path.stat().st_size if script_path.exists() else 0,
            "lines": 0,
            "modified_date": "",
            "jira_refs": [],
            "technologies": [],
            "kpi_reports": [],
        }
        
        if not script_path.exists():
            return info
        
        try:
            content = script_path.read_text(encoding='utf-8', errors='ignore')
            
            # Basic info
            info["lines"] = len(content.split('\n'))
            info["modified_date"] = datetime.fromtimestamp(
                script_path.stat().st_mtime
            ).strftime("%Y-%m-%d %H:%M:%S")
            
            # Jira references (CPU- prefix)
            info["jira_refs"] = list(set(re.findall(r'CPU-\d+', content)))
            
            # Technologies used (imports)
            info["technologies"] = []
            for match in re.findall(r'^import\s+(\w+(\.\w+)*)', content, re.MULTILINE):
                if match not in exc["technologies"]:
                    info["technologies"].append(match.split('.')[0])
            
            info["technologies"] = list(set(exc["technologies"]))
            
            # KPI/Dashboard mentions
            info["kpi_reports"] = list(set(re.findall(r'[kK]pi|Dashboard|Report|Metric', content, re.IGNORECASE)))
            
            # CPU requirements
            cpu_matches = re.findall(r'(?:#|//)\s*CPU\-?.*??:[\s\S]+?\n', content)
            info["cpu_requirements"] = cpu_matches
            
        except Exception as e:
            info["error"] = str(e)
        
        return info
    
    def get_task_info(self, task: JiraIssue) -> Optional[Dict]:
        """Get implementation for a Jira task"""
        if task.key and not task.key.startswith("CPU-"):
            return None
        
        # Search for matching script
        scripts = self.find_scripts()
        matching_scripts = []
        
        for script in scripts:
            try:
                script_info = self.extract_script_info(script)
                if task.key in script_info.get("jira_refs", []):
                    matching_scripts.append(script_info)
            except:
                continue
        
        if matching_scripts:
            return matching_scripts[0]
        return None


class ExcelExporter:
    """Export report to Excel"""
    
    def __init__(self):
        self.output_dir = "/home/user1/work/jirascanner/output/"
        os.makedirs(self.output_dir, exist_ok=True)
    
    def export(self, data: List[Dict]) -> str:
        """Export to Excel with proper formatting"""
        if not data:
            print("⚠️  No data to export")
            return ""
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Jira Tasks Report"
        
        # Headers row
        headers = [
            "Jira-Nummer",
            "Jira Auftraggeber", 
            "Entwickler",
            "Jahr/Monat/Tag Abgeschlossen",
            "Zusammenfassung Anforderung",
            "Filedatum",
            "File Location",
            "Verwendete Technologien",
            "Status",
            "Priorität"
        ]
        
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = type('obj', (object,), {'bold': True})()
            cell.fill = type('obj', (object,), {'fill_type': 'solid', 
                'start_color': '4472C4'})()
            cell.value = header
        
        # Empty line separator
        ws.cell(row=2, column=1, value="")
        
        # Data rows
        for row_idx, row_data in enumerate(data, 3):
            ws.cell(row=row_idx, column=1, value=row_data.get("Jira-Nummer", ""), border=type('obj', (object,), {'border': True, 'style': 'thin'})())
            ws.cell(row=row_idx, column=2, value=row_data.get("Jira Auftraggeber", ""))
            ws.cell(row=row_idx, column=3, value=row_data.get("Entwickler", ""))
            ws.cell(row=row_idx, column=4, value=row_data.get("Jahr/Monat/Tag Abgeschlossen", ""))
            ws.cell(row=row_idx, column=5, value=row_data.get("Zusammenfassung Anforderung", ""))
            ws.cell(row=row_idx, column=6, value=row_data.get("Filedatum", ""))
            ws.cell(row=row_idx, column=7, value=row_data.get("File Location", ""))
            ws.cell(row=row_idx, column=8, value=row_data.get("Verwendete Technologien", ""))
            ws.cell(row=row_idx, column=9, value=row_data.get("Status", ""))
            ws.cell(row=row_idx, column=10, value=row_data.get("Priorität", ""))
        
        # Apply borders
        ws.border = type('obj', (object,), {
            'top': {'style': 'thin'},
            'right': {'style': 'thin'}, 
            'bottom': {'style': 'thin'},
            'left': {'style': 'thin'}
        })()
        
        # Save
        filename = self.output_dir + f"jira_report_{datetime.now().strftime('%Y-%m-%d').replace('-', '_')}.xlsx"
        wb.save(filename)
        print(f"✅ Excel report saved: {filename}")
        return filename


def main():
    """Main execution"""
    print("=" * 70)
    print("🔍 Jira/Gira Scanner - CPU- Task Report Generator")
    print("Excel Export with complete task information")
    print("=" * 70)
    
    # Initialize
    jira_client = JiraClient()
    script_scanner = ScriptScanner()
    exporter = ExcelExporter()
    
    print("\n📋 Searching Jira tasks with prefix: CPU-")
    tasks = jira_client.search_tasks(prefix="CPU-")
    
    if not tasks:
        print("⚠️  No CPU- tasks found. Creating empty template...")
        template_data = [{"Jira-Nummer": "CPU-", "Zusammenfassung": "No data"}]
    else:
        print(f"✅ Found {len(tasks)} CPU- tasks\n")
        
        print("🔍 Scanning scripts for implementation details...")
        report = []
        
        for task in tasks:
            impl = script_scanner.get_task_info(task)
            
            if impl:
                report.append({
                    "Jira-Nummer": task.key,
                    "Jira Auftraggeber": task.reporter or "",
                    "Entwickler": task.assignee or "",
                    "Jahr/Monat/Tag Abgeschlossen": 
                        task.resolved or "Pending",
                    "Zusammenfassung Anforderung": task.summary,
                    "Filedatum": task.created or "",
                    "File Location": impl.get("path", ""),
                    "Verwendete Technologien": 
                        " | ".join(impl.get("technologies", [])[:3]) or "",
                    "Status": task.status or "Open",
                    "Priorität": task.priority
                })
            else:
                report.append({
                    "Jira-Nummer": task.key,
                    "Jira Auftraggeber": task.reporter or "",
                    "Entwickler": task.assignee or "",
                    "Jahr/Monat/Tag Abgeschlossen": "Pending",
                    "Zusammenfassung Anforderung": task.summary or "",
                    "Filedatum": task.created or "",
                    "File Location": "Not found",
                    "Verwendete Technologien": "",
                    "Status": task.status or "Missing",
                    "Priorität": task.priority
                })
        
        print(f"✅ Generated report for {len(report)} tasks")
    
    print("\n💾 Exporting to Excel...")
    filename = exporter.export(report if report else [{"Jira-Nummer": "N/A"}])
    
    print("=" * 70)
    print(f"✓ Report exported: {filename}")
    print("=" * 70)


if __name__ == "__main__":
    main()
