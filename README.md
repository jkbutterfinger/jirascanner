# Jira/Gira Scanner

Scanner für Jira-Tasks mit konfigurierbarem Key-Prefix (Standard: `CPU-`).
Er verknüpft Tasks mit lokalen Skripten und erzeugt einen Excel-Bericht
(optional zusätzlich CSV). Dazu gibt es eine FastAPI-API und ein Dashboard.

**Version:** 1.1.0 · **Python:** ≥ 3.9

---

## Schnellstart

```bash
cd jirascanner

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Konfiguration anpassen (URL, Token, Projekt)
$EDITOR config.yaml

# Bericht erzeugen
python3 main.py
# oder mit Komfortprüfungen:
./run.sh
```

Alternativ Token per Umgebungsvariable (empfohlen, kein Secret in Git):

```bash
export JIRA_URL="https://jira.example.com"
export JIRA_API_TOKEN="..."
export JIRA_PROJECT_KEY="CPU"
python3 main.py
```

### CLI-Optionen

```bash
python3 main.py --help
python3 main.py --dry-run          # nur Konsole, keine Datei
python3 main.py --prefix CPU- -v   # DEBUG-Logging
python3 main.py --config other.yaml
```

---

## Excel-Bericht

Alle 10 Spalten stehen nebeneinander, eine Zeile pro Task:

| # | Spalte | Inhalt |
|---|--------|--------|
| 1 | Jira-Nummer | Task-Key, z. B. `CPU-1001` |
| 2 | Jira-Auftraggeber | Reporter (`displayName`) |
| 3 | Entwickler | Assignee (`displayName`) |
| 4 | Jahr/Monat/Tag-Abgeschlossen | `resolutiondate` oder `Pending` |
| 5 | Zusammenfassung-Anforderung | Task-Summary |
| 6 | File-Datum | Änderungsdatum des gefundenen Scripts (`YYYY-MM-DD`) |
| 7 | File-Location | Pfad(e) der gefundenen Scripts oder `Not found` |
| 8 | Verwendete-Technologien | Erkannte Imports (ohne Stdlib) |
| 9 | Status | Jira-Status |
| 10 | Priorität | Jira-Priorität (normalisiert) |

Beispiel:

```text
Jira-Nummer | Jira-Auftraggeber | Entwickler | Jahr/Monat/Tag-Abgeschlossen | Zusammenfassung-Anforderung | File-Datum | File-Location                              | Verwendete-Technologien | Status | Priorität
CPU-1001    | Max Mustermann    | Anna Meier | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-10 | /mount/drive/folder/report_monthly_sales.py | pandas, requests        | Done   | High
CPU-1002    | Lisa Schmidt      | Tom Mueller| Pending                      | ETL Pipeline                | 2024-01-12 | /mount/drive/folder/etl.py                  | sqlalchemy              | In Progress | Medium
```

Der Export enthält:

- Kopfzeile mit blauem Hintergrund, weißer Schrift und Rahmen
- Status-Farbcodierung (offen/in Arbeit/erledigt)
- Zeilenumbruch für lange Summary-/Pfad-/Technologie-Spalten
- Freeze-Panes und AutoFilter
- Summenzeile `GESAMT` mit Anzahl der Tasks
- Dateiname `jira_report_YYYY_MM_DD.xlsx` (konfigurierbar)
- optional `jira_report_YYYY_MM_DD.csv` (UTF-8, Semikolon)

Ausgabeort: `output/` (siehe `excel.output_path`).

---

## Funktionsweise

1. **Jira-Abfrage** über REST API v3.
   - Bevorzugt `/rest/api/3/search/jql` mit `nextPageToken`.
   - Fällt bei älteren Servern auf `/rest/api/3/search` mit `startAt` zurück.
   - Pagination bis `performance.max_results`, Timeout und Retries konfigurierbar.
   - ADF-Beschreibungen (API v3) werden in Text umgewandelt.
2. **Prefix-Filter:** Task-Keys werden clientseitig gegen `filters.prefix` geprüft.
3. **Script-Scan:** Alle Dateien unter `scripts.base_path` werden **einmal** gelesen
   und nach gefundenen Jira-Referenzen (`[A-Z][A-Z0-9]+-\d+`) indexiert.
4. **Matching:** Für jeden Task werden alle referenzierenden Scripts gesucht;
   Pfade, Dateidatum und Technologien werden zusammengeführt.
5. **Export:** XLSX (und optional CSV) nach `excel.output_path`.

Ohne gültige Jira-Konfiguration läuft das Tool trotzdem durch und erzeugt
einen leeren Bericht mit Hinweiszeile.

---

## Konfiguration

`config.yaml` wird über `config.py` geladen; fehlende Werte werden durch
Defaults ergänzt. Umgebungsvariablen haben Vorrang:

| Variable | Config-Pfad |
|----------|-------------|
| `JIRA_URL` | `server.url` |
| `JIRA_API_TOKEN` | `api.token` |
| `JIRA_PROJECT_KEY` | `project.key` |
| `JIRA_SCRIPT_PATH` | `scripts.base_path` |
| `JIRA_OUTPUT_PATH` | `excel.output_path` |
| `JIRA_LOG_LEVEL` | `logging.level` |

Wichtige Blöcke: `server`, `api`, `project`, `scripts`, `filters`,
`performance`, `excel`, `logging`.

> **Sicherheit:** `config.yaml` wird mit leerem Token eingecheckt. Trage
> echte Tokens nicht dort ein, sondern nutze `JIRA_API_TOKEN` oder eine
> lokale, ignorierte Datei (`config.local.yaml`, `.env`).

---

## API und Dashboard

```bash
python3 -m api.suggestion_api
# oder
uvicorn api.suggestion_api:app --reload
```

Standard: `http://127.0.0.1:8000` (änderbar über `API_HOST`/`API_PORT`).
Das Dashboard wird unter `/` ausgeliefert, die OpenAPI-Doku unter `/docs`.

| Methode | Endpoint | Beschreibung |
|---------|----------|--------------|
| GET | `/api/health` | Health-Check |
| GET | `/api/tasks` | Tasks aus Cache; `?refresh=true` lädt live aus Jira |
| POST | `/api/tasks` | Task in JSON-DB speichern |
| GET | `/api/scripts` | Script-Inventar scannen |
| POST | `/api/scripts/analyze?path=...` | Einzelnes Script analysieren |
| GET | `/api/suggestions` | Suggestions auflisten |
| POST | `/api/suggestions` | Suggestion anlegen |
| POST | `/api/suggestions/batch` | Mehrere Suggestions anlegen |
| GET | `/api/gaps` | Einfache Gap-Analyse (Script-Referenzen) |
| GET | `/api/match` | Task gegen Scripts matchen |

Die API speichert Suggestions in `inventory/suggestions.json`
(überschreibbar über `JIRA_SCANNER_DB`).

---

## Script-Inventar (CLI)

```bash
python3 scripts/script_inventory.py
python3 scripts/script_inventory.py --path /pfad/zu/scripts --json
```

Das Inventar erkennt Jira-Referenzen, Technologien, Datenbankzugriffe,
API-Aufrufe, KPI-/Dashboard-Begriffe und eine grobe Komplexität.

---

## Projektstruktur

```text
config.py                     gemeinsame Config-/Logging-Schicht
main.py                       kanonischer Scanner + Excel-/CSV-Export
scanner/main_scanner.py       Kompatibilitäts-Wrapper (delegiert an main.py)
scripts/script_inventory.py   Script-Analyse und Inventar-CLI
api/suggestion_api.py         FastAPI-Service + Dashboard-Auslieferung
dashboard/dashboard.html      Weboberfläche
tests/test_smoke.py           Smoke-Tests
config.yaml                   Konfiguration
run.sh                        Startskript mit Vorabprüfungen
requirements.txt              Laufzeit-Abhängigkeiten
requirements-dev.txt          Test-/Lint-Abhängigkeiten
```

---

## Tests

```bash
source .venv/bin/activate
pip install -r requirements-dev.txt
pytest
```

Die Smoke-Tests prüfen u. a. ADF-/Datums-Parsing, Jira-Response-Verarbeitung,
Prefix-Filter, Script-Matching, Excel-/CSV-Export und die API-Endpunkte.

---

## Fehlerbehandlung

| Situation | Verhalten |
|-----------|-----------|
| Jira nicht konfiguriert | Warnung + leerer Bericht |
| Jira nicht erreichbar / 5xx | Retries, danach Warnung + leerer Bericht |
| 401/403 | Klare Fehlermeldung im Log, leerer Bericht |
| Scripts-Pfad fehlt | Warnung, `File-Location = Not found` |
| Einzelnes Script unlesbar | Warnung, restliche Scripts werden verarbeitet |
| Excel nicht schreibbar | Fehlercode 1 + Logeintrag |

Logs: `logs/jira_scanner.log` (rotierend).

---

## Bekannte Grenzen / nächste Schritte

- Confluence-Integration ist nur als Konfigurationsblock vorhanden.
- Gap-Analyse ist eine einfache Referenzprüfung, kein semantisches Matching.
- Kein automatisches Schreiben zurück nach Jira.
- Kein Multi-Project-Support in einem Lauf.
- Dashboard zeigt Tasks aus dem Cache; Live-Refresh über `?refresh=true`.

---

## Lizenz

MIT
