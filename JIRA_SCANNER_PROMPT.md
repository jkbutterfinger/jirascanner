================================================================================
JIRA/GIRA SCANNER - SYSTEM- UND NUTZUNGSDOKUMENTATION
================================================================================

SYSTEMBEZEICHNUNG : Jira/Gira Scanner with Excel/CSV Export
VERSION           : 1.1.0
PYTHON            : >= 3.9
LIZENZ            : MIT
HAUPTEINSTIEG     : main.py

Dieses Dokument beschreibt den tatsächlichen Stand der Implementierung.
Es ist mit main.py, config.py, config.yaml und der API synchron zu halten.

================================================================================
1. ZWECK
================================================================================

Der Scanner analysiert Jira-Tasks mit einem konfigurierbaren Key-Prefix
(Standard: "CPU-"), verknüpft sie mit lokalen Skripten und erzeugt daraus
einen Excel-Bericht. Optional wird zusätzlich eine CSV-Datei geschrieben.

Er liefert:
- Task-Zuweisungen (Auftraggeber, Entwickler)
- Status und Priorität
- Abschlussdatum
- Zusammenfassung der Anforderung
- gefundene Script-Dateien inkl. Änderungsdatum
- erkannte Technologien
- eine einfache Gap-/Matching-Sicht über die FastAPI-API

================================================================================
2. KOMPONENTEN
================================================================================

main.py
  Kanonische Implementierung:
  - JiraClient       : Jira REST API v3, Pagination, Retry, ADF-Parsing
  - ScriptScanner    : einmaliger Script-Scan + Index nach Jira-Referenz
  - ExcelExporter    : XLSX- und CSV-Export
  - build_report()   : Zusammenführung Task <-> Script
  - main()           : CLI-Einstiegspunkt

config.py
  Gemeinsame Config-Schicht:
  - lädt config.yaml
  - ergänzt Defaults
  - wendet Umgebungsvariablen an
  - richtet Logging ein (Console + rotierende Logdatei)

scanner/main_scanner.py
  Kompatibilitäts-Wrapper. Delegiert vollständig an main.py.
  Historische Aufrufe wie "python -m scanner.main_scanner" funktionieren
  weiterhin, es gibt aber keine zweite Implementierung mehr.

scripts/script_inventory.py
  ScriptAnalyzer, JiraMapper, ReportGenerator sowie ein eigenes CLI
  für das Script-Inventar.

api/suggestion_api.py
  FastAPI-Service unter dem Prefix /api inkl. Auslieferung des Dashboards.

dashboard/dashboard.html
  Weboberfläche für Tasks, Script-Inventar und Suggestions.

tests/test_smoke.py
  Smoke-Tests für Parsing, Matching, Export und API.

config.yaml        Konfiguration (Token leer lassen, siehe Sicherheit)
run.sh             Startskript mit Vorabprüfungen
requirements.txt   Laufzeit-Abhängigkeiten
requirements-dev.txt  Test-/Lint-Abhängigkeiten

================================================================================
3. INSTALLATION
================================================================================

  cd jirascanner

  python3 -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt

Laufzeit-Abhängigkeiten:
  - PyYAML >= 6.0          YAML-Konfiguration
  - requests >= 2.28.0     Jira-HTTP-Requests
  - openpyxl >= 3.1.0      Excel-Export
  - fastapi >= 0.104.0     API
  - uvicorn[standard]      API-Server
  - pydantic >= 2.5.0      API-Modelle
  - python-dotenv >= 1.0   optionale .env-Unterstützung

Für Tests/Linting zusätzlich:
  pip install -r requirements-dev.txt

================================================================================
4. KONFIGURATION
================================================================================

Alle Module laden dieselbe Konfiguration über config.py. Fehlende Werte
werden durch Defaults ergänzt; Umgebungsvariablen haben Vorrang.

4.1 Wichtige Config-Blöcke (config.yaml)

  server:
    url: "https://your-jira-server.com"
    protocol: "https"
    port: 443

  api:
    token: ""                  # besser über JIRA_API_TOKEN setzen
    auth_type: "token"
    cors_origins: [...]        # Origins für die API

  project:
    key: "CPU"                 # Task-Keys beginnen mit "CPU-"
    jql: ""                    # optional: eigene JQL statt project.key

  scripts:
    base_path: "scripts"
    extensions: [".py", ".sh", ".sql", ".js", ".ts"]
    include_subdirectories: true
    exclude: ["__init__.py", "script_inventory.py"]

  filters:
    prefix: "CPU-"             # clientseitiger Key-Filter

  performance:
    max_results: 1000          # Obergrenze über alle Seiten
    timeout: 30                # Sekunden pro Request
    retry_attempts: 3

  excel:
    output_path: "output"
    output_filename: "jira_report_{YYYY}_{MM}_{DD}.xlsx"
    include_csv: true
    datetime_format: "%Y-%m-%d %H:%M:%S"
    sheet_name: "Jira-Task-Report"

  logging:
    level: "INFO"
    output_file: "logs/jira_scanner.log"
    max_size_mb: 10
    backup_count: 3

4.2 Umgebungsvariablen (haben Vorrang)

  JIRA_URL             -> server.url
  JIRA_API_TOKEN       -> api.token
  JIRA_PROJECT_KEY     -> project.key
  JIRA_SCRIPT_PATH     -> scripts.base_path
  JIRA_OUTPUT_PATH     -> excel.output_path
  JIRA_LOG_LEVEL       -> logging.level

  Optional für die API:
  API_HOST (Default 127.0.0.1), API_PORT (Default 8000)
  JIRA_SCANNER_DB (Pfad zur suggestions.json)

4.3 Verhalten ohne gültige Konfiguration

  Fehlt URL oder Token (oder steht noch "your-jira-server"), wird eine
  Warnung geloggt und ein leerer Bericht mit Hinweiszeile erzeugt.
  Das Tool bricht nicht ab.

================================================================================
5. NUTZUNG
================================================================================

5.1 Bericht erzeugen

  python3 main.py
  ./run.sh
  python3 main.py --dry-run          # nur Konsole, keine Datei
  python3 main.py --prefix CPU- -v   # DEBUG-Logging
  python3 main.py --config other.yaml

  run.sh prüft Abhängigkeiten, legt output/, logs/, inventory/ und
  jira_data/ an und gibt den Exit-Code von main.py weiter.

5.2 Script-Inventar

  python3 scripts/script_inventory.py
  python3 scripts/script_inventory.py --path /pfad/zu/scripts --json

5.3 API + Dashboard

  python3 -m api.suggestion_api
  # oder
  uvicorn api.suggestion_api:app --reload

  Dashboard : http://127.0.0.1:8000/
  OpenAPI   : http://127.0.0.1:8000/docs
  Health    : http://127.0.0.1:8000/api/health

5.4 Tests

  pytest

================================================================================
6. ABLAUF / TECHNISCHE DETAILS
================================================================================

6.1 Jira-Abfrage (REST API v3)

  1) Bevorzugt: /rest/api/3/search/jql mit nextPageToken-Pagination.
  2) Fallback:  /rest/api/3/search mit startAt-Pagination für ältere Server.
  3) Felder: summary, status, priority, reporter, assignee, created,
     resolutiondate, issuetype, description, labels.
  4) Beschreibungen im Atlassian Document Format (ADF) werden rekursiv
     in Text umgewandelt.
  5) Timeout und Retry sind über performance.timeout bzw.
     performance.retry_attempts konfigurierbar.
  6) HTTP 401/403 erzeugt eine klare Fehlermeldung; 5xx wird retryt.
  7) Pagination endet bei performance.max_results.

6.2 Prefix-Filter

  Die JQL lädt das Projekt (project.key bzw. project.jql). Anschließend
  wird jeder Task-Key clientseitig gegen filters.prefix geprüft.
  Beispiel: Prefix "CPU-" akzeptiert "CPU-1001", nicht "OPS-1001".

6.3 Script-Scan und Matching

  1) Alle Dateien unter scripts.base_path mit den konfigurierten
     Endungen werden gefunden (optional rekursiv).
  2) Dateien in scripts.exclude werden übersprungen.
  3) Jede Datei wird genau EINMAL gelesen.
  4) Jira-Referenzen werden per Regex [A-Z][A-Z0-9]+-\d+ erkannt.
  5) Der Index lautet: Jira-Key -> Liste von Script-Infos.
  6) Für jeden Task werden alle passenden Scripts zusammengeführt:
     - File-Location: alle Pfade (mehrzeilig)
     - File-Datum: Änderungsdatum des ersten Treffers (YYYY-MM-DD)
     - Verwendete-Technologien: erkannte Imports ohne Standardbibliothek
  7) Kein Treffer -> File-Location = "Not found".

  Dadurch entfällt das frühere O(Tasks × Scripts)-Verhalten.

6.4 Technologie-Erkennung

  Erkannt werden:
  - Python: "import x", "import x.y", "from x.y import z"
  - JavaScript: require("...")
  Standardbibliothek und Builtins werden herausgefiltert.
  Ausgabe: sortierte, eindeutige Liste (max. 8 im Excel-Bericht).

================================================================================
7. EXCEL-/CSV-AUSGABE
================================================================================

7.1 Spalten (alle nebeneinander, eine Zeile pro Task)

  +----------------+---------------------+----------------+-------------------------------+
  | Jira-Nummer    | Jira-Auftraggeber   | Entwickler     | Jahr/Monat/Tag-Abgeschlossen  |
  +----------------+---------------------+----------------+-------------------------------+
  | Zusammenfassung-Anforderung | File-Datum | File-Location | Verwendete-Technologien |
  +-----------------------------+------------+---------------+-------------------------+
  | Status | Priorität |
  +--------+-----------+

  Reihenfolge und Breite:

   1 Jira-Nummer                    18
   2 Jira-Auftraggeber              24
   3 Entwickler                     24
   4 Jahr/Monat/Tag-Abgeschlossen   26
   5 Zusammenfassung-Anforderung    55 (Umbruch)
   6 File-Datum                     16
   7 File-Location                  48 (Umbruch)
   8 Verwendete-Technologien        45 (Umbruch)
   9 Status                         14
  10 Priorität                      12

7.2 Beispiel (side-by-side)

  Jira-Nummer | Jira-Auftraggeber | Entwickler  | Jahr/Monat/Tag-Abgeschlossen | Zusammenfassung-Anforderung | File-Datum | File-Location                               | Verwendete-Technologien | Status      | Priorität
  CPU-1001    | Max Mustermann    | Anna Meier  | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-10 | /mount/drive/folder/report_monthly_sales.py | pandas, requests        | Done        | High
  CPU-1002    | Lisa Schmidt      | Tom Mueller | Pending                      | ETL Pipeline                | 2024-01-12 | /mount/drive/folder/etl.py                  | sqlalchemy              | In Progress | Medium

7.3 Formatierung

  - Kopfzeile: blau (4472C4), weiß, fett, zentriert, mit Rahmen
  - Datenzellen: dünne Rahmen, vertikal zentriert
  - Status-Farben: offen/rot, in Arbeit/orange, erledigt/grün
  - Umbruch für Summary, File-Location und Technologien
  - Freeze-Panes auf A2
  - AutoFilter über den Datenbereich
  - Summenzeile "GESAMT" mit Anzahl der Tasks
  - Datumsformat aus excel.datetime_format (Standard: YYYY-MM-DD HH:MM:SS)

7.4 Dateien

  Excel : <output_path>/jira_report_YYYY_MM_DD.xlsx
  CSV   : <output_path>/jira_report_YYYY_MM_DD.csv  (wenn include_csv=true)
  CSV-Format: UTF-8 mit BOM, Semikolon als Trennzeichen

  Der Dateiname kommt aus excel.output_filename; unterstützte Platzhalter:
  {YYYY}, {MM}, {DD}.

================================================================================
8. API-ENDPUNKTE
================================================================================

Alle fachlichen Endpunkte liegen unter /api:

  GET  /api/health
       Health-Check inkl. Version und Anzahl Suggestions.

  GET  /api/tasks?refresh=false
       Tasks aus jira_data/tasks.json. Mit refresh=true werden Tasks live
       über den JiraClient geladen.

  POST /api/tasks
       Task in inventory/suggestions.json speichern.

  GET  /api/scripts?path=...
       Script-Inventar scannen (Default: scripts.base_path).

  POST /api/scripts/analyze?path=...
       Einzelne Datei analysieren.

  GET  /api/suggestions
       Suggestions auflisten.

  POST /api/suggestions
       Suggestion anlegen (title, description, task_id erforderlich).

  POST /api/suggestions/batch
       Mehrere Suggestions anlegen.

  GET  /api/gaps?task_keys=CPU-1001
       Einfache Gap-Analyse: existiert ein Script mit dieser Referenz?

  GET  /api/match?task_key=CPU-1001&script_path=...
       Task gegen Scripts matchen (match_score 1.0 oder 0.0).

  GET  /
       Liefert dashboard/dashboard.html aus.

Antwortformate sind JSON. Fehler werden als HTTPException mit Statuscode
und Detailmeldung zurückgegeben.

================================================================================
9. FEHLERBEHANDLUNG UND LOGGING
================================================================================

  Jira nicht konfiguriert       -> Warnung + leerer Bericht
  Jira nicht erreichbar / 5xx   -> Retries, danach Warnung + leerer Bericht
  HTTP 401/403                  -> Fehlermeldung, leerer Bericht
  Scripts-Pfad fehlt            -> Warnung, File-Location = "Not found"
  Einzelnes Script unlesbar     -> Warnung, restliche Scripts laufen weiter
  Excel nicht schreibbar        -> Exit-Code 1 + Logeintrag
  Ungültige config.yaml         -> ValueError mit Pfadangabe

  Logging:
  - Console (Level aus logging.level, mit -v auf DEBUG)
  - Datei: logs/jira_scanner.log, rotierend
  - Format: Zeit | Level | Logger | Nachricht

================================================================================
10. SICHERHEIT
================================================================================

  1) Keine Tokens im Code hardcodieren.
  2) Empfohlen: JIRA_API_TOKEN als Umgebungsvariable oder in .env
     (bereits in .gitignore).
  3) config.yaml wird mit leerem Token eingecheckt. Echte Tokens nicht
     dort eintragen; alternativ config.local.yaml nutzen (ignoriert).
  4) .gitignore schützt .env, inventory/, jira_data/, confluence_data/,
     output/ und Logs.
  5) API bindet standardmäßig nur an 127.0.0.1. Für externen Zugriff
     API_HOST bewusst setzen und CORS-Origins einschränken.
  6) Tokens regelmäßig rotieren.

================================================================================
11. TROUBLESHOOTING
================================================================================

  PROBLEM: "Jira ist nicht konfiguriert"
  LÖSUNG:
  - server.url in config.yaml prüfen (kein "your-jira-server")
  - api.token bzw. JIRA_API_TOKEN setzen
  - project.key prüfen

  PROBLEM: Keine Tasks gefunden
  LÖSUNG:
  - JQL/project.key prüfen
  - filters.prefix prüfen (Standard "CPU-")
  - API-Berechtigungen des Tokens prüfen
  - Mit -v laufen lassen und logs/jira_scanner.log ansehen

  PROBLEM: Excel-Export schlägt fehl
  LÖSUNG:
  - pip install -r requirements.txt
  - Schreibrechte auf excel.output_path prüfen
  - Python-Version >= 3.9 prüfen

  PROBLEM: Scripts werden nicht gefunden
  LÖSUNG:
  - scripts.base_path prüfen
  - scripts.extensions prüfen
  - scripts.exclude prüfen
  - Jira-Referenz im Script prüfen (z. B. "# CPU-1001")

  PROBLEM: API startet nicht
  LÖSUNG:
  - pip install -r requirements.txt
  - Port 8000 belegt? API_PORT setzen
  - Importfehler: python -m api.suggestion_api direkt starten

================================================================================
12. TESTS UND QUALITÄT
================================================================================

  pytest
  ruff check .

  Die Tests decken ab:
  - ADF- und Datums-Parsing
  - Jira-Response-Verarbeitung inkl. Prefix-Filter
  - Verhalten ohne gültige Jira-Konfiguration
  - Script-Index und Report-Zusammenführung
  - Excel-/CSV-Export inkl. Header und Summenzeile
  - ScriptAnalyzer
  - API-Endpunkte (Health, Suggestions, Gaps, Dashboard)

================================================================================
13. BEKANNTE GRENZEN / ROADMAP
================================================================================

  - Confluence-Integration ist nur als Config-Block vorhanden.
  - Gap-Analyse ist eine Referenzprüfung, kein semantisches Matching.
  - Kein Rückschreiben nach Jira.
  - Kein Multi-Project-Support in einem Lauf.
  - Dashboard liest Tasks aus dem Cache; Live-Refresh über ?refresh=true.
  - Weitere Ideen: CI/CD-Integration, Slack/Teams-Benachrichtigung,
    Custom-Templates, erweiterte KI-Gap-Analyse.

================================================================================
ENDE DER DOKUMENTATION
================================================================================
