# Jira/Gira Scanner

## Features
- Scan Jira/Gira tasks with CPU- prefix
- Generate comprehensive Excel reports with implementation details
- Track requirements and mappings
- Central configuration via YAML file

## Quick Start

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure
Update `config.yaml` with your Jira/Gira credentials.

### 3. Generate Report
```bash
python3 main.py
```

## Excel Report Columns

The generated Excel file includes:
- **Jira-Nummer**: CPU-XXXXXXXX (Task ID)
- **Jira Auftraggeber**: Requester/Demanding party
- **Entwickler**: Assigned developer
- **Jahr/Monat/Tag Abgeschlossen**: Completion date (YYYY-MM-DD HH:MM:SS)
- **Zusammenfassung Anforderung**: Complete task summary
- **Filedatum**: Date when task/file was created
- **File Location**: Script file path in repository
- **Verwendete Technologien**: Technologies/imports used
- **Status**: Current task status
- **Priorität**: Priority level (High/Medium/Low)

## Configuration

See `config.yaml` for all available options:

### Server Configuration
```yaml
server:
  url: "https://your-jira-server.com"
```

### API Authentication
```yaml
api:
  token: "your_api_token"
```

### Scripts Path
```yaml
scripts:
  base_path: "/path/to/scripts/"
  extensions: [".py", ".sh", ".sql", ".js"]
```

### Excels Export
```yaml
excel:
  output_path: "/home/user1/work/jirascanner/output/"
  datetime_format: "%Y-%m-%d %H:%M:%S"
```

## Jira Task Prefix

The scanner filters for tasks with the **CPU-** prefix:
- Matching: `CPU-1001`, `CPU-2024-001`
- Searching: Task summary/description contains "CPU-"

## Usage Examples

### Scan Current Directory
```bash
python main.py
```

### Scan Custom Path
```bash
python main.py --scripts /custom/path --output /custom/output
```

## Output

```
======================================================================
🔍 Jira/Gira Scanner - CPU- Task Report Generator
======================================================================

📋 Searching Jira tasks with prefix: CPU-
✅ Found 15 CPU- tasks

🔍 Scanning scripts for implementation details...
✅ Generated report for 15 tasks

💾 Exporting to Excel...
✅ Excel report saved: /home/user1/work/jirascanner/output/jira_report_2024_01_15.xlsx
======================================================================
✓ Report exported: /home/user1/work/jirascanner/output/jira_report_2024_01_15.xlsx
======================================================================
```

## Dependencies

```
pyyaml>=6.0
requests>=2.28.0
openpyxl>=3.1.0
lxml>=4.7.0
pandas>=1.5.0
```

## License

MIT License