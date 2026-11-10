================================================================================
JIRA/GIRA SCANNER - SYSTEM- UND NUTZUNGSDOKUMENTATION
================================================================================

SYSTEMBEZEICHNUNG : Jira/Gira Scanner for SQL files with per-task Excel export
VERSION           : 1.3.0
PYTHON            : >= 3.9 (CI: 3.10 / 3.11 / 3.12)
LIZENZ            : MIT
HAUPTEINSTIEG     : main.py
STANDARD-PREFIX   : JRA-
REPOSITORY        : https://github.com/jkbutterfinger/jirascanner
BRANCH            : main
STAND             : 2026-11-10
DOKUMENTATION     : JIRA_SCANNER_PROMPT.md (dieses Dokument), README.md

SETUP             : ./setup.sh [--dev] [--test] [--recreate] [--venv DIR]
START             : ./run.sh [--mode scan|enrich|all] [--batch-size N] [--batch-delay S]
VENV              : .venv (wird von setup.sh angelegt, von run.sh automatisch genutzt)
KONFIGURATION     : config.yaml (Defaults und Umgebungsvariablen: config.py)
TESTS             : ./setup.sh --dev --test  bzw.  pytest && ruff check .
SCAN-INDEX        : inventory/scan_index.json (storage.scan_index)
EXCEL-BERICHT     : output/jira_report_<YYYY>_<MM>_<DD>.xlsx (Gesamtbericht)
                    optional pro Task: <base_path>/.../<Jira-Task>/<Jira-Task>.xlsx
CLI-MODI          : --mode scan | enrich | all (Default: all)
BATCH-LIMIT       : max. 100 Jira-Keys pro Request (hart begrenzt)
BATCH-PAUSE       : jira.batch_delay_seconds (Default 1.0 s), CLI --batch-delay

Dieses Dokument beschreibt den tatsächlichen Stand der Implementierung.
Es ist mit main.py, config.py, config.yaml, setup.sh, run.sh und der API
synchron zu halten.

================================================================================
1. ZWECK
================================================================================

Der Scanner sucht ausschließlich SQL-Dateien in einer definierten
Verzeichnisstruktur. Die Jira-Referenz wird aus dem Verzeichnisnamen
abgeleitet - nicht aus dem Dateiinhalt. Standardmäßig wird ein einziger
Gesamtbericht über alle gefundenen SQL-Dateien unter excel.output_path
erzeugt. Optional (excel.aggregate_report=false) entsteht für jeden
Jira-Task ein eigener Excel-Bericht direkt im Task-Verzeichnis.

Der Ablauf ist in zwei Aufrufmodi geteilt:

  1) --mode scan    Verzeichnisstruktur einlesen und Ausgabe soweit
                    möglich erstellen (ohne Jira-Zugriff).
  2) --mode enrich  Jira-Details im Nachgang anreichern. Die Keys werden
                    in Blöcken von maximal 100 abgefragt, mit Pause
                    zwischen den Blöcken.

  --mode all        Beide Phasen nacheinander (Default).

Erwartete Struktur:

  /mount/<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<name>.sql

Beispiel:

  /mount/2024/IT/ACME/JRA-1001/monthly_sales.sql
  /mount/2024/IT/ACME/JRA-1001/etl_refresh.sql
  /mount/2024/IT/ACME/JRA-1002/export.sql

Der Task-Key ist der Name des Elternverzeichnisses der SQL-Datei.
Der Gesamtbericht wird unter excel.output_path abgelegt:

  output/jira_report_<YYYY>_<MM>_<DD>.xlsx

Mit excel.aggregate_report=false wird pro Task ein Bericht im
Task-Verzeichnis abgelegt:

  /mount/2024/IT/ACME/JRA-1001/JRA-1001.xlsx

SQL-Dateien außerhalb eines Task-Verzeichnisses werden ebenfalls
analysiert und erscheinen ohne Jira-Nummer im Gesamtbericht.

Er liefert:
- Task-Zuweisungen (Auftraggeber, Entwickler)
- Status und Priorität
- Abschlussdatum
- Zusammenfassung der Anforderung
- gefundene SQL-Dateien inkl. Änderungsdatum
- erkannte Technologien (SQL und Dialekt-Hinweise)
- eine einfache Gap-/Matching-Sicht über die FastAPI-API

================================================================================
2. KOMPONENTEN
================================================================================

main.py
  Kanonische Implementierung:
  - JiraClient          : Jira REST API v3, Pagination, Retry, ADF-Parsing,
                          Batch-Abfrage nach Keys (max. 100 pro Request)
  - ScriptScanner       : SQL-Scan + Index nach Task-Verzeichnis
  - ExcelExporter       : Gesamtbericht oder ein XLSX-Bericht pro Task
  - build_scan_index()  : Inventur der Verzeichnisstruktur (JSON)
  - run_scan_phase()    : Phase 1 ohne Jira
  - run_enrich_phase()  : Phase 2 mit Jira-Anreicherung
  - build_report()      : Zusammenführung Task <-> SQL-Dateien
  - main()              : CLI-Einstiegspunkt (--mode scan|enrich|all)

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
  für das SQL-Inventar.

api/suggestion_api.py
  FastAPI-Service unter dem Prefix /api inkl. Auslieferung des Dashboards.

dashboard/dashboard.html
  Weboberfläche für Tasks, SQL-Inventar und Suggestions.

tests/test_smoke.py
  Smoke-Tests für Parsing, Task-Key-Ableitung, Matching, Export und API.

config.yaml        Konfiguration (Token leer lassen, siehe Sicherheit)
setup.sh           Python-/venv-Setup: findet Python >= 3.9, legt .venv an,
                   installiert requirements*.txt und prüft die Kernimporte
run.sh             Startskript mit Vorabprüfungen; nutzt .venv automatisch
requirements.txt   Laufzeit-Abhängigkeiten
requirements-dev.txt  Test-/Lint-Abhängigkeiten

================================================================================
3. INSTALLATION
================================================================================

  cd jirascanner

  # Empfohlen: Setup-Skript (idempotent)
  ./setup.sh                  # Laufzeit-Abhängigkeiten
  ./setup.sh --dev            # inkl. pytest, ruff, httpx
  ./setup.sh --recreate       # .venv neu aufbauen
  ./setup.sh --venv DIR       # andere virtuelle Umgebung
  PYTHON=python3.12 ./setup.sh

  # Manuell (entspricht dem Setup-Skript)
  python3 -m venv .venv
  source .venv/bin/activate
  python -m pip install --upgrade pip setuptools wheel
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
    key: "JRA"                 # Task-Keys beginnen mit "JRA-"
    jql: ""                    # optional: eigene JQL statt project.key

  scripts:
    base_path: "/mount"        # Root der Verzeichnisstruktur
    extensions: [".sql"]       # ausschließlich SQL
    include_subdirectories: true
    exclude: []

  filters:
    prefix: "JRA-"             # clientseitiger Key-Filter

  performance:
    max_results: 1000          # Obergrenze über alle Seiten
    timeout: 30                # Sekunden pro Request
    retry_attempts: 3

  storage:
    scan_index: "inventory/scan_index.json"   # Ergebnis von --mode scan

  jira:
    batch_size: 100            # 1-100 Keys pro Request (wird begrenzt)
    batch_delay_seconds: 1.0   # Pause zwischen zwei Blöcken

  excel:
    output_path: "output"                        # Ablage des Gesamtberichts
    aggregate_report: true                       # ein Bericht über alle Dateien
    aggregate_filename: "jira_report_{YYYY}_{MM}_{DD}.xlsx"
    output_filename: "{JIRA}.xlsx"               # pro Task (nur bei false)
    fallback_filename: "jira_report_{YYYY}_{MM}_{DD}.xlsx"
    include_csv: false
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

  Fehlt URL oder Token (oder steht noch "your-jira-server"), läuft
  --mode scan trotzdem vollständig durch und erzeugt die Berichte aus
  der Verzeichnisstruktur. --mode enrich wird mit einer Warnung
  übersprungen; die vorhandenen Scan-Berichte bleiben erhalten.
  Das Tool bricht nicht ab.

4.4 Dateiname des Excel-Berichts

  Gesamtbericht (Standard):

    excel.aggregate_report = true
    excel.aggregate_filename = "jira_report_{YYYY}_{MM}_{DD}.xlsx"

    Platzhalter: {YYYY}, {MM}, {DD}
    Ablage: excel.output_path (Standard: output/)

  Pro Task (optional):

    excel.aggregate_report = false
    excel.output_filename = "{JIRA}.xlsx"

    {JIRA}        Jira-Task-Key, z. B. JRA-1001
    {SQL}         Dateiname (ohne .sql) der ersten SQL-Datei
    {YEAR}        Jahr aus der Verzeichnisstruktur
    {DEPARTMENT}  Abteilung aus der Verzeichnisstruktur
    {CUSTOMER}    Auftraggeber aus der Verzeichnisstruktur
    {YYYY}        aktuelles Jahr
    {MM}          aktueller Monat
    {DD}          aktueller Tag

  Standard im Task-Modus: "{JIRA}.xlsx"
  -> /mount/2024/IT/ACME/JRA-1001/JRA-1001.xlsx

================================================================================
5. NUTZUNG
================================================================================

5.0 Python-Umgebung einrichten (setup.sh)

  ./setup.sh                  # Python >= 3.9 finden, .venv anlegen,
                              # Laufzeit-Abhängigkeiten installieren
  ./setup.sh --dev            # zusätzlich pytest, ruff, httpx
  ./setup.sh --dev --test     # danach pytest und ruff ausführen
  ./setup.sh --recreate       # .venv löschen und neu aufbauen
  ./setup.sh --venv DIR       # andere virtuelle Umgebung verwenden
  PYTHON=python3.12 ./setup.sh

  Das Skript ist idempotent, prüft die Python-Version, aktualisiert
  pip/setuptools/wheel und verifiziert die Kernimporte. run.sh erkennt
  .venv automatisch; ein manuelles "source" ist nicht erforderlich.

5.1 Berichte erzeugen (zweiphasig)

  Phase 1 - Verzeichnisstruktur einlesen (ohne Jira):
    python3 main.py --mode scan
    ./run.sh --mode scan

    - erzeugt den Gesamtbericht unter excel.output_path
    - analysiert alle SQL-Dateien, auch außerhalb von Task-Verzeichnissen
    - Jira-Spalten bleiben leer, Status = "Pending"
    - schreibt inventory/scan_index.json (storage.scan_index)

  Phase 2 - Jira-Details anreichern:
    python3 main.py --mode enrich
    python3 main.py --mode enrich --batch-size 50 --batch-delay 2

    - liest den Scan-Index (fehlt er, wird Phase 1 automatisch ausgeführt)
    - fragt nur die gefundenen Task-Keys ab
    - maximal 100 Keys pro Request
    - Pause zwischen den Blöcken (jira.batch_delay_seconds)
    - Keys ohne Jira-Treffer: Scan-Daten bleiben, Status = "Missing"

  Beides nacheinander (Default):
    python3 main.py --mode all

  Weitere Optionen:
    python3 main.py --dry-run          # nur Konsole, keine Datei
    python3 main.py --prefix JRA- -v   # DEBUG-Logging
    python3 main.py --config other.yaml

  run.sh prüft Abhängigkeiten, nutzt .venv automatisch, legt output/,
  logs/, inventory/ und jira_data/ an und gibt den Exit-Code von main.py
  weiter.

5.2 SQL-Inventar

  python3 scripts/script_inventory.py
  python3 scripts/script_inventory.py --path /mount --json

5.3 API + Dashboard

  python3 -m api.suggestion_api
  # oder
  uvicorn api.suggestion_api:app --reload

  Dashboard : http://127.0.0.1:8000/
  OpenAPI   : http://127.0.0.1:8000/docs
  Health    : http://127.0.0.1:8000/api/health

5.4 Tests

  ./setup.sh --dev --test     # richtet .venv ein, führt pytest + ruff aus
  # oder manuell:
  pytest
  ruff check .

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

6.2 Prefix-Filter und Key-Abfrage

  filters.prefix (Standard "JRA-") wird clientseitig geprüft. In Phase 1
  werden Task-Verzeichnisse außerhalb dieses Prefix ignoriert; --prefix
  überschreibt den konfigurierten Wert. Dateien ohne Task-Key bleiben im
  Gesamtbericht. Beim projektweiten Suchpfad (z. B. API
  /api/tasks?refresh=true) wird ebenfalls gefiltert.

  Die CLI-Anreicherung fragt gezielt die aus der Verzeichnisstruktur
  abgeleiteten Keys ab:
    key in ("JRA-1001", "JRA-1002", ...)
  Dadurch werden keine fremden Projekt-Tasks geladen.

6.3 SQL-Scan und Task-Zuordnung

  1) Es werden ausschließlich Dateien mit den Endungen aus
     scripts.extensions gefunden. Standard: [".sql"].
  2) Dateien in scripts.exclude werden übersprungen.
  3) Jede SQL-Datei wird genau EINMAL gelesen.
  4) Der Task-Key wird aus dem Elternverzeichnis abgeleitet:
       .../<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<name>.sql
     Der Verzeichnisname muss dem Muster [A-Z][A-Z0-9]+-\d+ entsprechen.
  5) Der Index lautet: Jira-Key -> Liste von SQL-Infos.
  6) Für jeden Task werden alle SQL-Dateien in seinem Verzeichnis
     zusammengeführt:
       - File-Location: Pfad der jeweiligen SQL-Datei
       - File-Datum: Änderungsdatum (YYYY-MM-DD)
       - Verwendete-Technologien: SQL plus erkannte Dialekt-Hinweise
  7) Liegt eine SQL-Datei nicht in einem Task-Verzeichnis, wird sie ohne
     Task-Zuordnung (leere Jira-Nummer) in den Gesamtbericht aufgenommen.
  8) Kein Treffer -> File-Location = "Not found".
  9) Phase 1 (--mode scan) schreibt die Inventur als JSON:
       storage.scan_index (Standard: inventory/scan_index.json)
     Sie enthält pro Task: Verzeichnis, Pfadkontext und SQL-Dateien sowie
     eine flache Liste "files" mit allen analysierten SQL-Dateien.
 10) Die Excel-Berichte aus Phase 1 enthalten bereits File-Datum,
     File-Location, Technologien sowie Jahr und Abteilung (lokale
     Spalten 11/12); die Jira-Spalten bleiben leer (Status = "Pending").

6.4 Technologie-Erkennung

  Für SQL-Dateien:
  - Basis: "SQL"
  - Zusätzlich erkannt: "PL/pgSQL", "T-SQL", "MySQL", "Oracle"
    anhand typischer Schlüsselwörter.
  Ausgabe: sortierte, eindeutige Liste (max. 8 im Excel-Bericht).

6.5 Jira-Anreicherung in Blöcken

  1) run_enrich_phase() liest die Task-Keys aus dem Scan-Index.
  2) JiraClient.fetch_tasks_by_keys() dedupliziert die Keys.
  3) Abfrage in Blöcken von jira.batch_size; harter Maximalwert 100.
     CLI: --batch-size (1-100).
  4) Zwischen den Blöcken wird jira.batch_delay_seconds pausiert.
     CLI: --batch-delay.
  5) JQL pro Block: key in ("JRA-1001", "JRA-1002", ...).
  6) Fehlgeschlagene Blöcke werden geloggt und übersprungen; die
     übrigen Blöcke laufen weiter.
  7) Nicht gefundene Keys bleiben mit Status "Missing" erhalten.
  8) Anschließend wird der Gesamtbericht (oder im Task-Modus die
     Task-Berichte) neu geschrieben.

================================================================================
7. EXCEL-AUSGABE
================================================================================

7.1 Ablage

  Standardmäßig entsteht ein Gesamtbericht über alle SQL-Dateien:

    <excel.output_path>/jira_report_<YYYY>_<MM>_<DD>.xlsx
    (Standard: output/jira_report_<YYYY>_<MM>_<DD>.xlsx)

  Er enthält eine Zeile pro SQL-Datei. Dateien in Task-Verzeichnissen
  erhalten eine Jira-Nummer; Dateien außerhalb bleiben ohne Task-Kontext.

  Mit excel.aggregate_report=false entsteht pro Jira-Task genau eine
  Excel-Datei im Task-Verzeichnis:

    /mount/<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<Jira-Task>.xlsx

  Enthält ein Task-Verzeichnis mehrere SQL-Dateien, stehen sie als
  mehrere Zeilen im selben Excel-Bericht.

  In Phase 1 (--mode scan) wird der Bericht bereits mit File-Datum,
  File-Location, Technologien sowie Jahr und Abteilung (aus der
  Verzeichnisstruktur) geschrieben; die Jira-Spalten bleiben leer
  (Status = "Pending"). Phase 2 (--mode enrich) schreibt denselben
  Bericht mit den Jira-Details neu.

  Nur wenn keine SQL-Dateien gefunden werden, wird ein leerer
  Fallback-Bericht unter excel.output_path geschrieben:
    <output_path>/jira_report_YYYY_MM_DD.xlsx

7.2 Spalten (alle nebeneinander, eine Zeile pro SQL-Datei)

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
  11 Jahr                            8 (lokal, aus Verzeichnisstruktur)
  12 Abteilung                      20 (lokal, aus Verzeichnisstruktur)

7.3 Beispiel (side-by-side)

  Datei: output/jira_report_2026_11_10.xlsx

  Jira-Nummer | Jira-Auftraggeber | Entwickler  | Jahr/Monat/Tag-Abgeschlossen | Zusammenfassung-Anforderung | File-Datum | File-Location                                  | Verwendete-Technologien | Status      | Priorität | Jahr | Abteilung
  JRA-1001    | Max Mustermann    | Anna Meier  | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-10 | /mount/2024/IT/ACME/JRA-1001/monthly_sales.sql | SQL                     | Done        | High      | 2024 | IT
  JRA-1001    | Max Mustermann    | Anna Meier  | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-12 | /mount/2024/IT/ACME/JRA-1001/etl_refresh.sql   | SQL                     | Done        | High      | 2024 | IT
  JRA-1002    | Lisa Schmidt      | Tom Mueller | Pending                      | ETL Pipeline                | 2024-01-15 | /mount/2024/IT/ACME/JRA-1002/export.sql        | SQL                     | In Progress | Medium    | 2024 | IT
              |                   |             |                              |                             | 2024-01-16 | /mount/oracle_sql/list_tables.sql              | SQL                     | Pending     |           |      |

7.4 Formatierung

  - Kopfzeile: blau (4472C4), weiß, fett, zentriert, mit Rahmen
  - Datenzellen: dünne Rahmen, vertikal zentriert
  - Status-Farben: offen/rot, in Arbeit/orange, erledigt/grün
  - Umbruch für Summary, File-Location und Technologien
  - Freeze-Panes auf A2
  - AutoFilter über den Datenbereich
  - Summenzeile "GESAMT" mit Anzahl der SQL-Dateien
  - Datumsformat aus excel.datetime_format (Standard: YYYY-MM-DD HH:MM:SS)

7.5 CSV (optional)

  Wenn excel.include_csv=true, wird zusätzlich eine CSV neben dem
  Excel-Bericht geschrieben:
    Gesamtbericht: <aggregate_filename>.csv
    Task-Modus:    <Jira-Task>.csv
  Format: UTF-8 mit BOM, Semikolon als Trennzeichen.

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
       SQL-Inventar scannen (Default: scripts.base_path).

  POST /api/scripts/analyze?path=...
       Einzelne SQL-Datei analysieren.

  GET  /api/suggestions
       Suggestions auflisten.

  POST /api/suggestions
       Suggestion anlegen (title, description, task_id erforderlich).

  POST /api/suggestions/batch
       Mehrere Suggestions anlegen.

  GET  /api/gaps?task_keys=JRA-1001
       Einfache Gap-Analyse: existiert eine SQL-Datei mit dieser Referenz?

  GET  /api/match?task_key=JRA-1001&script_path=...
       Task gegen SQL-Dateien matchen (match_score 1.0 oder 0.0).

  GET  /
       Liefert dashboard/dashboard.html aus.

Antwortformate sind JSON. Fehler werden als HTTPException mit Statuscode
und Detailmeldung zurückgegeben.

================================================================================
9. FEHLERBEHANDLUNG UND LOGGING
================================================================================

  Jira nicht konfiguriert            -> Scan-Berichte bleiben, enrich wird
                                        übersprungen
  Jira nicht erreichbar / 5xx        -> Retries, danach wird der Batch
                                        übersprungen
  HTTP 401/403                       -> Fehlermeldung, Batch wird übersprungen
  Scan-Index fehlt bei --mode enrich -> Verzeichnisstruktur wird automatisch
                                        zuerst eingelesen
  Scan-Root fehlt                    -> Warnung, keine Berichte
  SQL außerhalb Task-Verzeichnis     -> wird ohne Task-Zuordnung in den
                                        Gesamtbericht aufgenommen
  Task ohne SQL-Datei                -> Warnung, kein Bericht für den Task
  Task-Key nicht in Jira gefunden    -> Scan-Daten bleiben, Status "Missing"
  Einzelne SQL-Datei unlesbar        -> Warnung, restliche Dateien laufen weiter
  Excel nicht schreibbar             -> Exit-Code 1 + Logeintrag
  Ungültige config.yaml              -> ValueError mit Pfadangabe

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
  - filters.prefix prüfen (Standard "JRA-")
  - API-Berechtigungen des Tokens prüfen
  - Mit -v laufen lassen und logs/jira_scanner.log ansehen

  PROBLEM: Keine SQL-Dateien gefunden
  LÖSUNG:
  - scripts.base_path prüfen (z. B. "/mount")
  - scripts.extensions prüfen (Standard [".sql"])
  - Verzeichnisstruktur prüfen:
      <Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<name>.sql
  - Task-Verzeichnisname muss dem Muster JRA-1234 entsprechen

  PROBLEM: Excel-Export schlägt fehl
  LÖSUNG:
  - ./setup.sh ausführen (legt .venv an und installiert die Abhängigkeiten)
  - Schreibrechte auf dem Task-Verzeichnis bzw. excel.output_path prüfen
  - Python-Version >= 3.9 prüfen

  PROBLEM: API startet nicht
  LÖSUNG:
  - ./setup.sh ausführen
  - Port 8000 belegt? API_PORT setzen
  - Importfehler: python -m api.suggestion_api direkt starten

================================================================================
12. TESTS UND QUALITÄT
================================================================================

  ./setup.sh --dev --test    # Setup + pytest + ruff in einem Schritt
  pytest
  ruff check .

  Die Tests decken ab:
  - ADF- und Datums-Parsing
  - Jira-Response-Verarbeitung inkl. Prefix-Filter
  - Verhalten ohne gültige Jira-Konfiguration
  - Task-Key-Ableitung aus der Verzeichnisstruktur
  - SQL-Index und Report-Zusammenführung
  - Scan-Index (JSON-Roundtrip)
  - Scan- und Enrich-Phase inkl. CLI-Modi
  - Batch-Abfrage: maximal 100 Keys pro Request und Pause dazwischen
  - Excel-Gesamtbericht inkl. Dateien außerhalb von Task-Verzeichnissen
  - optionaler Excel-Export pro Task inkl. Header und Summenzeile
  - Dateinamen-Platzhalter
  - ScriptAnalyzer (nur SQL)
  - API-Endpunkte (Health, Suggestions, Gaps, Dashboard)

================================================================================
13. BEKANNTE GRENZEN / ROADMAP
================================================================================

  - Es werden ausschließlich .sql-Dateien berücksichtigt.
  - Die Task-Zuordnung hängt strikt an der Verzeichnisstruktur; Dateien
    ohne Task-Verzeichnis erscheinen ohne Jira-Nummer im Gesamtbericht.
  - Confluence-Integration ist nur als Config-Block vorhanden.
  - Gap-Analyse ist eine Referenzprüfung, kein semantisches Matching.
  - Kein Rückschreiben nach Jira.
  - Kein Multi-Project-Support in einem Lauf.
  - Dashboard liest Tasks aus dem Cache; Live-Refresh über ?refresh=true.
  - Weitere Ideen: CI/CD-Integration, Slack/Teams-Benachrichtigung,
    Custom-Templates, erweiterte KI-Gap-Analyse.

================================================================================
14. METADATEN, DATEIEN UND VERSIONSHISTORIE
================================================================================

14.1 Projekt-Metadaten

  Projektname       : jira-scanner
  Paketversion      : 1.3.0 (pyproject.toml)
  API-Version       : 1.3.0 (api/suggestion_api.py)
  Repository        : https://github.com/jkbutterfinger/jirascanner
  Default-Branch    : main
  Lizenz            : MIT
  Python            : >= 3.9
  Standard-Prefix   : JRA-
  Konfiguration     : config.yaml
  Scan-Index        : inventory/scan_index.json

14.2 Dateien und Aufgaben

  main.py                       Kanonischer Scanner, CLI, Excel-Export,
                                zweiphasiger Ablauf (scan/enrich/all)
  config.py                     Defaults, YAML-Laden, Umgebungsvariablen, Logging
  config.yaml                   Projektkonfiguration (Token leer lassen)
  setup.sh                      Python-/venv-Setup und Abhängigkeitsprüfung
  run.sh                        Startskript, nutzt .venv automatisch
  requirements.txt              Laufzeit-Abhängigkeiten
  requirements-dev.txt          Test-/Lint-Abhängigkeiten
  scanner/main_scanner.py       Kompatibilitäts-Wrapper auf main.py
  scripts/script_inventory.py   SQL-Analyse und Inventar-CLI
  api/suggestion_api.py         FastAPI-Service und Dashboard-Auslieferung
  dashboard/dashboard.html      Weboberfläche
  tests/test_smoke.py           Smoke-Tests
  README.md                     Nutzerdokumentation
  JIRA_SCANNER_PROMPT.md        System- und Nutzungsdokumentation (dieses Dokument)

14.3 Versionshistorie

  1.3.0   Zweiphasiger Ablauf:
          - --mode scan   : Verzeichnisstruktur einlesen, Gesamtbericht
                            mit Status "Pending" und Scan-Index erzeugen
          - --mode enrich : Jira-Details in Blöcken von max. 100 Keys
                            nachladen, konfigurierbare Pause
          - --mode all    : beide Phasen nacheinander (Default)
          - --batch-size / --batch-delay sowie jira.batch_size /
            jira.batch_delay_seconds
          - Prefix-Filter für Task-Verzeichnisse (Default "JRA-")
          - Gesamtbericht über alle SQL-Dateien (excel.aggregate_report)
          - setup.sh für reproduzierbares Python-/venv-Setup
          - run.sh nutzt .venv automatisch

  1.2.0   Einphasiger Lauf: projektweite Jira-Suche, SQL-Scan,
          Excel-Bericht pro Task.

================================================================================
ENDE DER DOKUMENTATION
================================================================================
