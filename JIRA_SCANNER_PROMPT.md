================================================================================
JIRA/GIRA SCANNER - SYSTEM PROMPT / NUTZUNGSANWEISUNG
================================================================================

SYSTEM BEZEICHNUNG: Jira/Gira Scanner with Excel Export
VERSION: 1.0.0
AUTHOR: Jira Scanner Development Team

================================================================================
ZWECK DES SYSTEMS
================================================================================

Dieses System analysiert Jira/Gira Tasks mit dem Prefix "CPU-" und generiert 
eine Excel-Datei mit vollständigen Informationen über:
- Task-Zuweisungen und Status
- Implementation-Nachweise
- Verwendete Technologien
- Missing Requirements

================================================================================
PRIMÄRE FUNKTIONEN
================================================================================

1. SCANNEN von Jira Tasks mit CPU- Prefix
2. AUTOMATISCHE VERKNÜPFUNG mit existierenden Skripten
3. GENERIERUNG von Excel-Bericht (XLSX Format)
4. TRACKING von Requirements und Implementierungen
5. GAP-ANALYSSE zwischen Tasks und Code

================================================================================
SYSTEM SETUP
================================================================================

1. INSTALLATION VON ABHÄNGIGKEITEN:
   
   cd /home/user1/work/jirascanner
   pip install -r requirements.txt

   Erforderliche Pakete:
   - pyyaml >= 6.0        # YAML-Konfiguration
   - requests >= 2.28.0   # API-Kommunikation
   - openpyxl >= 3.1.0    # Excel-Export
   - lxml >= 4.7.0        # XML Parsing
   - pandas >= 1.5.0      # Data Processing

2. KONFIGURATIONBearbeiten:

   a) Öffne config.yaml
   
   b) Fülle folgende Abschnitte:
   
    Server Configuration:
      url: "https://dein-jira-server.com"
      
    API Authentication:
      token: "dein_api_token"
      
    Project Configuration:
      key: "CPU"         # Jira Projekt-Schlüssel
      
    Scripts Path:
      base_path: "/home/user1/work/jirascanner/scripts/"
      extensions: [".py", ".sh", ".sql", ".js"]
      
    Excel Output:
      output_path: "/home/user1/work/jirascanner/output/"

   c) Speichere config.yaml

3. START DES SYSTEMS:

   python3 main.py
   
   oder
   
   bash run.sh

================================================================================
JIRA TASK FILTERING
================================================================================

SYSTEMATISCHES FILTERING:

Prefix: "CPU-"       (Erzwingt für alle Tasks)
Project Key: CPU     (Standard project)

MATCHING-MUSTER:
✓ CPU-1001
✓ CPU-2024-001
✓ CPU-2024Q1-001

NON MATCHING BEISPIELE:
✗ TASK-123
✗ REPORT-456
✗ JIRA-1000

================================================================================
EXCEL BEREICHS-STRUKTUR
================================================================================

Generierter Excel-Report enthält folgende Spalten:

1. Jira-Nummer          | CPU-XXXXXXXX (Task ID)
2. Jira Auftraggeber    | Wer verlangt die Task
3. Entwickler           | Zugeteilter Entwicklers
4. Jahr/Monat/Tag Abgeschlossen | YYYY-MM-DD HH:MM:SS
5. Zusammenfassung Anforderung | Taskbeschreibung
6. Filedatum            | Task-Erstellungsdatum
7. File Location        | Absoluter Pfad Skript
8. Verwendete Technologien   | Technologien/Importe
9. Status               | Aktuell Status (Open/Closed)
10. Priorität           | High/Medium/Low

================================================================================
EXAMPEL-OUTPUT
================================================================================

Excel Output Beispiel:

        Jira-Nummer     | Jira Auftraggeber       | Entwickler
CPU-1001   | MAX Mustermann        | HELG Schmidt      | 2024-01-15 09:30:00
CPU-1002   | ANA Meier             | TOM Mueller      | 2024-01-15 10:45:00

        | 
        | Zusammenfassung Anforderung
CPU-1001   | Monthly Sales Dashboard - KPI Dashboard mit regionaler Breakdown
CPU-1002   | Data Warehouse Migration - ETL Pipeline Setup

        | 
        | File Location (Script Paths)
CPU-1001   | /home/user1/work/jirascanner/scripts/adhoc/report_monthly_sales.py
CPU-1002   | /home/user1/work/jirascanner/scripts/reports/etl_migration.py

        | 
        | Verwendete Technologien
CPU-1001   | pandas, mysql, requests
CPU-1002   | sqlalchemy, postgresql, pandas

================================================================================
TECHNISCHE DETAILS
================================================================================

API INTEGRATION:
- REST API Integration (Jira API v3)
- Authentication via Bearer Token
- Projekt-basierte Filterung

SCRIPT ANALYSYE:
- Automatische Parse von Python/Shell Scripts
- Identifikation von Datenbankzugriffen
- Technologie-Erkennung aus Imports
- Jira-Referenz-Extraktion

EXCEL FORMATTING:
- Professionelles Layout mit Header-Formatting
- Automatische Sortierung nach Priority
- Datumsformatierung (YYYY-MM-DD HH:MM:SS)
- Farbcodierung für Status

================================================================================
ERROR HANDLING
================================================================================

Fehlerbehandlung ist integriert:

1. Jira API Down → Warnung + leere Reportgeneration
2. Mangelhafte Konfiguration → Standardwerte
3. Missing Scripts → "Not found" in Excel
4. Parse Fehler → Fehler Logging + Fortsetzung

Error Logging Location: /home/user1/work/jirascanner/logs/

================================================================================
SICHERHEITSHINWEISE
================================================================================

1. NIE API-Tokens im Code hardcodieren!

2. Verwende .gitignore für sensitive Daten:
   - Tokens
   - Passwords
   - API Keys
   
3. Nur authenticated Zugriff für Production-Server

4. API-Token regelmäßig rotieren

================================================================================
TROUBLESHOOTING
================================================================================

PROBLEM: Keine Tasks gefunden
LÖSUNG: 
- Prüfe config.yaml Server URL
- Prüfe API Token Gültigkeit  
- Prüfe Project Key
- Prüfe CPU- Prefix in Tasks

PROBLEM: Excel Export fehlschlägt
LÖSUNG:
- Prüfe pip install requirements.txt
- Prüfe Schreibe-Berechtigungen im Output-Verzeichnis
- Prüfe Python Version (>= 3.8)

PROBLEM: Scripts nicht analysiert
LÖSUNG:
- Prüfe Scripts-Pfad in config.yaml
- Prüfe Dateierweiterungen
- Prüfe Schreibrechte

================================================================================
NEXT STEPS / FUTURE ENHANCEMENTS
================================================================================

1. Integration mit Confluence-Dokumenten
2. CI/CD Pipeline Integration
3. Automatische Jira Updaten
4. Slack/Discord Notification
5. Advanced Gap Analysis mit AI
6. Multi-Project Support
7. Custom Template Export

================================================================================
CONTACT / SUPPORT
================================================================================

Fragen/Hilfe? 

GitHub Repo: https://github.com/user/jira-scanner
Issue Tracker: https://github.com/user/jira-scanner/issues
Documentation: https://github.com/user/jirascanner#readme

================================================================================
END OF SYSTEM PROMPT
================================================================================