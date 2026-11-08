# Jira/Gira Scanner

## 📋 System Description

A comprehensive scanner for Jira/Gira tasks with **CPU- prefix** that:
- Scans Jira tasks matching "CPU-" prefix
- Extracts implementation details from existing scripts
- Generates **Excel reports with ALL COLUMNS SIDE-BY-SIDE (horizontal layout)**

---

## 📄 Excel Output Format

All **10 columns displayed horizontally (side-by-side)** for each task:

### Example Excel Structure:

```
┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│  Jira-Nummer   │  Jira-Auftraggeber          │  Entwickler         │  Jahr/Monat/Tag-Abgeschlossen  │  Zusammenfassung-Anforderung          │  File-Datum     │  File-Location                         │
│────────────────│─────────────────────────────│─────────────────────│───────────────────────────────│───────────────────────────────────────│──────────────────│─────────────────────────────────────────│
│  CPU-1001      │  Max Mustermann             │  Peter Schmidt      │  2024-01-15 14:30:00          │  Monthly Sales Dashboard              │  2023-12-01     │  /scripts/report_monthly_sales.py       │
│────────────────│─────────────────────────────│─────────────────────│───────────────────────────────│───────────────────────────────────────│──────────────────│─────────────────────────────────────────│
│  CPU-1002      │  Anna Meier                 │  Thomas Müller      │  2024-01-20 09:15:00          │  ETL Pipeline for Data Warehouse      │  2023-11-15     │  /scripts/etl_pipeline.py               │
│────────────────│─────────────────────────────│─────────────────────│───────────────────────────────│───────────────────────────────────────│──────────────────│─────────────────────────────────────────│
│  CPU-1003      │  Lisa Schmidt               │  Michael Weber      │  Open                         │  API Performance Optimization         │  2024-01-25     │  /scripts/api_optimization.py            │
└────────────────│─────────────────────────────│─────────────────────│───────────────────────────────│───────────────────────────────────────│──────────────────│─────────────────────────────────────────┘
```

**ALL COLUMNS ARE HORIZONTAL** (Side-by-Side Layout)

---

## 📄 Excel Columns (All Horizontal)

| Column (#) | Column Name       | Description                  | Format                   |
|------------|-------------------|------------------------------|--------------------------|
| 1          | Jira-Nummer       | Task ID                      | CPU-1001                |
| 2          | Jira-Auftraggeber | Requester/Owner              | Max Mustermann          |
| 3          | Entwickler        | Assigned Developer           | Peter Schmidt           |
| 4          | Year/Month/Day    | Completion Date              | 2024-01-15 14:30:00     |
| 5          | Summary           | Task Summary/Requirement     | Monthly Sales Dashboard |
| 6          | File Date         | Created Date                 | 2023-12-01              |
| 7          | File-Location     | Script file path             | /scripts/report.py      |
| 8          | Technologies      | Used technologies/libraries  | pandas, mysql, requests |
| 9          | Status            | Task Status                  | Open/Closed/Pending     |
| 10         | Priority          | Task Priority                | High/Medium/Low         |

---

## 🚀 Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure
Edit `config.yaml`:
```yaml
server:
  url: "https://your-jira-server.com"

api:
  token: "your_api_token"

project:
  key: "CPU"  # Project key for filtering

scripts:
  base_path: "/home/user1/work/jirascanner/scripts/"

excel:
  output_path: "/home/user1/work/jirascanner/output/"
```

### 3. Generate Report
```bash
cd /home/user1/work/jirascanner
python3 main.py
```

---

## 📊 Excel Report Features

### 1. **Horizontal Layout** (All Columns Side-by-Side)
All 10 columns display horizontally across the page (one row per task)

### 2. **Auto-Styling**:
- Header: Blue background (4472C4), white text, bold
- Borders: Thin black borders around each cell
- Column widths: Optimized for content
- Wrap text: Enabled for long descriptions

### 3. **Column Specific Formatting**:
- **Jira-Nummer**: Center-aligned, width 20
- **Jira-Auftraggeber**: Left-aligned, width 25
- **Entwickler**: Left-aligned, width 25
- **Date**: Left-aligned, width 25
- **Summary**: Left-aligned, width 60 (wrapped)
- **File-Date**: Left-aligned, width 20
- **File-Location**: Left-aligned, width 50 (wrapped)
- **Technologies**: Left-aligned, width 60 (wrapped)
- **Status**: Center-aligned, width 15
- **Priority**: Center-aligned, width 15

### 4. **File Name Format**:
```
jira_report_YYYY_MM_DD.xlsx
```
Example: `jira_report_2024_01_15.xlsx`

---

## 📂 Output Files

```
/home/user1/work/jirascanner/output/
├── jira_report_2024_01_15.xlsx   # Excel Report (all columns horizontal)
└── jira_report_2024_01_15.csv    # CSV backup (optional, auto-generated)
```

---

## 🔍 CPU- Prefix Filtering

The scanner automatically filters tasks containing **CPU-**:

**Matching Tasks:**
- `CPU-1001`
- `CPU-2024-001`
- `CPU-2024Q1-001`

**Non-Matching:**
- `TASK-123`
- `REPORT-456`
- `JIRA-1000`

---

## 📈 Example Output Snippet

When running `python3 main.py`, you'll see:

```
======================================================================
🔍 JIRA/GIRA SCANNER - CPU- Task Report Generator
Excel Export: ALL COLUMNS SIDE-BY-SIDE (Horizontal Layout)
======================================================================

📋 Searching Jira tasks with prefix: CPU-
✅ Found 15 CPU- tasks

🔍 Scanning scripts for implementation details...
✅ Generated report for 15 tasks

💾 Exporting to Excel (ALL COLUMNS SIDE-BY-SIDE)...
✅ Excel Report with ALL COLUMNS SIDE-BY-SIDE saved: /home/user1/work/jirascanner/output/jira_report_2024_01_15.xlsx
======================================================================
✓ Report Excel saved: /home/user1/work/jirascanner/output/jira_report_2024_01_15.xlsx
======================================================================
```

---

## 🛠️ Dependencies

```yaml
- pyyaml >= 6.0
- requests >= 2.28.0
- openpyxl >= 3.1.0
- lxml >= 4.7.0
- pandas >= 1.5.0
```

---

## 📋 License

MIT License - Free to use and modify

---

## 🔗 Quick Links

- **Documentation**: [JIRA_SCANNER_PROMPT.md](./JIRA_SCANNER_PROMPT.md)
- **Configuration**: [config.yaml](./config.yaml)
- **System Prompt**: [JIRA_SCANNER_PROMPT.md](./JIRA_SCANNER_PROMPT.md)
