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
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

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
            tech_imports = re.findall(r'^(?:import\s+from\s+(\w+)|from\s+(\w+)', content, re.MULTILINE)
            for match in tech_imports:
                if match[0] and match[0] not in info["technologies"]:
                    info["technologies"].append(match[0])
                elif match[1] and match[1] not in info["technologies"]:
                    info["technologies"].append(match[1])
            
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
    """Export report to Excel with horizontal layout (all columns side-by-side)"""
    
    def __init__(self):
        self.output_dir = "/home/user1/work/jirascanner/output/"
        os.makedirs(self.output_dir, exist_ok=True)
    
    def export(self, data: List[Dict]) -> str:
        """
        Export to Excel with ALL COLUMNS SIDE-BY-SIDE (horizontal layout)
        Each row represents one Jira task, all 10 columns next to each other
        """
        if not data:
            print("⚠️  No data to export")
            return ""
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Jira-Task-Report-Excel"
        
        # Define all column headers (horizontal layout - all in one row)
        horizontal_headers = [
            {"value": "Jira-Nummer", "width": 20},
            {"value": "Jira-Auftraggeber", "width": 25},
            {"value": "Entwickler", "width": 25},
            {"value": "Jahr/Monat/Tag-Abgeschlossen", "width": 25},
            {"value": "Zusammenfassung-Anforderung", "width": 60},
            {"value": "File-Datum", "width": 20},
            {"value": "File-Location", "width": 50},
            {"value": "Verwendete-Technologien", "width": 60},
            {"value": "Status", "width": 15},
            {"value": "Priorität", "width": 15},
        ]
        
        # Set header cell styles
        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
            color="000000"
        )
        
        # Set column widths for horizontal layout
        column_widths = [20, 25, 25, 25, 60, 20, 50, 60, 15, 15]
        
        # Write headers (horizontal - all side-by-side in row 1)
        for col_idx, header_info in enumerate(horizontal_headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header_info["value"])
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = header_border
            ws.column_dimensions[chr(64 + col_idx)].width = column_widths[col_idx - 1]
        
        # Empty separator row
        ws.cell(row=2, column=1, value="")
        ws.cell(row=2, column=2, value="")
        ws.cell(row=2, column=3, value="")
        ws.cell(row=2, column=4, value="")
        ws.cell(row=2, column=5, value="")
        ws.cell(row=2, column=6, value="")
        ws.cell(row=2, column=7, value="")
        ws.cell(row=2, column=8, value="")
        ws.cell(row=2, column=9, value="")
        ws.cell(row=2, column=10, value="")
        
        # Define data row cell styles
        data_font = Font(size=10)
        data_alignment = Alignment(horizontal="left", vertical="center")
        data_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin'),
            color="000000"
        )
        normal_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        
        # Write data rows (all columns side-by-side for each task)
        for row_idx, row_data in enumerate(data, 3):
            # Column 1: Jira-Nummer
            ws.cell(row=row_idx, column=1, value=row_data.get("Jira-Nummer", ""))
            
            # Column 2: Jira-Auftraggeber
            ws.cell(row=row_idx, column=2, value=row_data.get("Jira-Auftraggeber", ""))
            
            # Column 3: Entwickler
            ws.cell(row=row_idx, column=3, value=row_data.get("Entwickler", ""))
            
            # Column 4: Jahr/Monat/Tag-Abgeschlossen
            ws.cell(row=row_idx, column=4, value=row_data.get("Jahr/Monat/Tag-Abgeschlossen", ""))
            
            # Column 5: Zusammenfassung-Anforderung
            ws.cell(row=row_idx, column=5, value=row_data.get("Zusammenfassung-Anforderung", ""))
            
            # Column 6: File-Datum
            ws.cell(row=row_idx, column=6, value=row_data.get("File-Datum", ""))
            
            # Column 7: File-Location
            ws.cell(row=row_idx, column=7, value=row_data.get("File-Location", ""))
            
            # Column 8: Verwendete-Technologien
            ws.cell(row=row_idx, column=8, value=row_data.get("Verwendete-Technologien", ""))
            
            # Column 9: Status
            ws.cell(row=row_idx, column=9, value=row_data.get("Status", ""))
            
            # Column 10: Priorität
            ws.cell(row=row_idx, column=10, value=row_data.get("Priorität", ""))
            
            # Apply border to all cells in this row
            for col in range(1, 11):
                cell = ws.cell(row=row_idx, column=col)
                cell.font = data_font
                cell.fill = normal_fill
                cell.alignment = data_alignment
                cell.border = data_border
        
        # Add total row at bottom
        ws.cell(row=len(data) + 3, column=1, value="GESAMT")
        ws.cell(row=len(data) + 3, column=10, value="=-")
        
        # Save workbook
        if "{YYYY}_{MM}_{DD}" in str(self.output_dir):
            output_filename = datetime.now().strftime("jira_report_%Y_%m_%d.xlsx")
        else:
            output_filename = datetime.now().strftime("jira_report_%Y_%m_%d.xlsx")
        
        output_path = os.path.join(self.output_dir, output_filename)
        wb.save(output_path)
        print(f"✅ Excel Report with ALL COLUMNS SIDE-BY-SIDE saved: {output_path}")
        return output_path
    
    def export_simple(self, data: List[Dict]) -> str:
        """Simple Excel export for debugging"""
        print("Simple export layout (all columns side-by-side):")
        print("-" * 180)
        for row in data:
            line = "| " + row.get("Jira-Nummer", "") + " | "
            line += row.get("Jira-Auftraggeber", "") + " | "
            line += row.get("Entwickler", "") + " | "
            line += row.get("Jahr/Monat/Tag-Abgeschlossen", "") + " | "
            line += row.get("Zusammenfassung-Anforderung", "")[:50] + " | "
            line += row.get("File-Datum", "") + " | "
            line += row.get("File-Location", "")[:40] + " | "
            line += row.get("Verwendete-Technologien", "")[:40] + " | "
            line += row.get("Status", "") + " | "
            line += row.get("Priorität", "") + " |"
            print(line)
        print("-" * 180)
        return ""


def main():
    """Main execution"""
    print("=" * 80)
    print("🔍 JIRA/GIRA SCANNER - CPU- Task Report Generator")
    print("Excel Export: ALLE KOLUMNEN NEBENEINANDER (Side-by-Side Layout)")
    print("=" * 80)
    
    # Initialize
    jira_client = JiraClient()
    script_scanner = ScriptScanner()
    exporter = ExcelExporter()
    
    print("\n📋 Searching Jira tasks with prefix: CPU-")
    tasks = jira_client.search_tasks(prefix="CPU-")
    
    if not tasks:
        print("⚠️  No CPU- tasks found. Creating empty template...")
        template_data = [{
            "Jira-Nummer": "CPU-",
            "Jira-Auftraggeber": "",
            "Entwickler": "",
            "Jahr/Monat/Tag Abgeschlossen": "Pending",
            "Zusammenfassung Anforderung": "No data",
            "File-Datum": "",
            "File-Location": "",
            "Verwendete-Technologien": "",
            "Status": "Open",
            "Priorität": "Medium"
        }]
    else:
        print(f"✅ Found {len(tasks)} CPU- tasks\n")
        
        print("🔍 Scanning scripts for implementation details...")
        report = []
        
        for task in tasks:
            impl = script_scanner.get_task_info(task)
            
            if impl:
                report.append({
                    "Jira-Nummer": task.key,
                    "Jira-Auftraggeber": task.reporter or "",
                    "Entwickler": task.assignee or "",
                    "Jahr/Monat/Tag-Abgeschlossen": 
                        task.resolved or "Pending",
                    "Zusammenfassung-Anforderung": task.summary or "",
                    "File-Datum": task.created or "",
                    "File-Location": impl.get("path", ""),
                    "Verwendete-Technologien": 
                        " | ".join(impl.get("technologies", [])[:3]) or "",
                    "Status": task.status or "Open",
                    "Priorität": task.priority
                })
            else:
                report.append({
                    "Jira-Nummer": task.key,
                    "Jira-Auftraggeber": task.reporter or "",
                    "Entwickler": task.assignee or "",
                    "Jahr/Monat/Tag-Abgeschlossen": "Pending",
                    "Zusammenfassung-Anforderung": task.summary or "",
                    "File-Datum": task.created or "",
                    "File-Location": "Not found",
                    "Verwendete-Technologien": "",
                    "Status": task.status or "Missing",
                    "Priorität": task.priority
                })
        
        print(f"✅ Generated report for {len(report)} tasks")
    
    print("\n💾 Exporting to Excel (ALL COLUMNS SIDE-BY-SIDE)...")
    filename = exporter.export(report if report else [{"Jira-Nummer": "N/A"}])
    
    print("=" * 80)
    print(f"✓ Report Excel saved: {filename}")
    print("=" * 80)


if __name__ == "__main__":
    main()
