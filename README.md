# Jira/Gira Scanner

Der Scanner durchsucht ein Verzeichnis ausschließlich nach **SQL-Dateien**.
Die Jira-Referenz wird aus der Verzeichnisstruktur abgeleitet – nicht aus dem
Dateiinhalt. Für jeden Jira-Task (Auftrag) wird ein eigener Excel-Bericht
direkt im Task-Verzeichnis abgelegt.

**Standard-Prefix:** `JRA-` · **Version:** 1.2.0 · **Python:** ≥ 3.9

---

## Verzeichnisstruktur

Erwartet wird folgende Ablage:

```text
/mount/<Jahr>/<Abteilung>/<Auftraggeber>/<Jira-Task>/<name>.sql
```

Beispiel:

```text
/mount/2024/IT/ACME/JRA-1001/monthly_sales.sql
/mount/2024/IT/ACME/JRA-1001/etl_refresh.sql
/mount/2024/IT/ACME/JRA-1002/export.sql
```

Der Task-Key ist der Name des Elternverzeichnisses der SQL-Datei
(`JRA-1001`, `JRA-1002`, ...). Der Bericht wird dort abgelegt:

```text
/mount/2024/IT/ACME/JRA-1001/JRA-1001.xlsx
```

- **Nur `.sql`** wird gescannt; andere Dateien werden ignoriert.
- Ein Task-Verzeichnis kann mehrere SQL-Dateien enthalten. Sie erscheinen als
  mehrere Zeilen im selben Excel-Bericht.
- Liegt eine SQL-Datei nicht in einem `<Jira-Task>`-Verzeichnis, wird sie mit
  einer Warnung ignoriert.

Der Scan-Root wird über `scripts.base_path` gesetzt (z. B. `/mount`).

---

## Schnellstart

```bash
cd jirascanner

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Konfiguration anpassen (Jira-URL, Token, Scan-Root)
$EDITOR config.yaml

# Berichte erzeugen
python3 main.py
# oder mit Komfortprüfungen:
./run.sh
```

Alternativ Token per Umgebungsvariable (empfohlen, kein Secret in Git):

```bash
export JIRA_URL="https://jira.example.com"
export JIRA_API_TOKEN="..."
export JIRA_PROJECT_KEY="JRA"
export JIRA_SCRIPT_PATH="/mount"
python3 main.py
```

### CLI-Optionen

```bash
python3 main.py --help
python3 main.py --dry-run          # nur Konsole, keine Datei
python3 main.py --prefix JRA- -v   # DEBUG-Logging
python3 main.py --config other.yaml
```

---

## Excel-Bericht

Pro Jira-Task entsteht **eine** Excel-Datei im Task-Verzeichnis. Sie enthält
alle 10 Spalten nebeneinander und eine Zeile pro gefundener SQL-Datei:

| # | Spalte | Inhalt |
|---|--------|--------|
| 1 | Jira-Nummer | Task-Key, z. B. `JRA-1001` |
| 2 | Jira-Auftraggeber | Reporter (`displayName`) |
| 3 | Entwickler | Assignee (`displayName`) |
| 4 | Jahr/Monat/Tag-Abgeschlossen | `resolutiondate` oder `Pending` |
| 5 | Zusammenfassung-Anforderung | Task-Summary |
| 6 | File-Datum | Änderungsdatum der SQL-Datei (`YYYY-MM-DD`) |
| 7 | File-Location | Pfad der SQL-Datei |
| 8 | Verwendete-Technologien | `SQL` plus erkannte Dialekt-Hinweise |
| 9 | Status | Jira-Status |
| 10 | Priorität | Jira-Priorität (normalisiert) |

Beispiel für `/mount/2024/IT/ACME/JRA-1001/JRA-1001.xlsx`:

```text
Jira-Nummer | Jira-Auftraggeber | Entwickler | Jahr/Monat/Tag-Abgeschlossen | Zusammenfassung-Anforderung | File-Datum | File-Location                                  | Verwendete-Technologien | Status | Priorität
JRA-1001    | Max Mustermann    | Anna Meier | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-10 | /mount/2024/IT/ACME/JRA-1001/monthly_sales.sql | SQL                     | Done   | High
JRA-1001    | Max Mustermann    | Anna Meier | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-12 | /mount/2024/IT/ACME/JRA-1001/etl_refresh.sql   | SQL                     | Done   | High
```

Der Export enthält:

- Kopfzeile mit blauem Hintergrund, weißer Schrift und Rahmen
- Status-Farbcodierung (offen/in Arbeit/erledigt)
- Zeilenumbruch für lange Summary-/Pfad-/Technologie-Spalten
- Freeze-Panes und AutoFilter
- Summenzeile `GESAMT` mit Anzahl der SQL-Dateien
- Dateiname standardmäßig `<Jira-Task>.xlsx`

### Dateiname konfigurieren

`excel.output_filename` unterstützt Platzhalter:

| Platzhalter | Bedeutung |
|-------------|-----------|
| `{JIRA}` | Jira-Task-Key, z. B. `JRA-1001` |
| `{SQL}` | Dateiname (ohne `.sql`) der ersten gefundenen SQL-Datei |
| `{YEAR}` | Jahr aus der Verzeichnisstruktur |
| `{DEPARTMENT}` | Abteilung aus der Verzeichnisstruktur |
| `{CUSTOMER}` | Auftraggeber aus der Verzeichnisstruktur |
| `{YYYY}`, `{MM}`, `{DD}` | aktuelles Datum |

Beispiel: `"{YEAR}_{CUSTOMER}_{JIRA}.xlsx"` →
`2024_ACME_JRA-1001.xlsx`.

Nur wenn **keine** Tasks/SQL-Dateien gefunden werden, wird ein leerer
Fallback-Bericht unter `excel.output_path` mit `excel.fallback_filename`
geschrieben.

---

## Funktionsweise

1. **Jira-Abfrage** über REST API v3.
   - Bevorzugt `/rest/api/3/search/jql` mit `nextPageToken`.
   - Fällt bei älteren Servern auf `/rest/api/3/search` mit `startAt` zurück.
   - Pagination bis `performance.max_results`, Timeout und Retries konfigurierbar.
   - ADF-Beschreibungen (API v3) werden in Text umgewandelt.
2. **Prefix-Filter:** Task-Keys werden clientseitig gegen `filters.prefix`
   geprüft (Standard `JRA-`).
3. **SQL-Scan:** Alle `.sql`-Dateien unter `scripts.base_path` werden genau
   einmal gefunden und über den Namen ihres Elternverzeichnisses indexiert.
4. **Matching:** Für jeden Jira-Task werden die SQL-Dateien aus seinem
   Task-Verzeichnis gesucht.
5. **Export:** Ein Excel-Bericht pro Task direkt im Task-Verzeichnis.

Ohne gültige Jira-Konfiguration läuft das Tool trotzdem durch und erzeugt
einen leeren Fallback-Bericht mit Hinweiszeile.

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

```yaml
project:
  key: "JRA"
filters:
  prefix: "JRA-"
scripts:
  base_path: "/mount"      # Root der Verzeichnisstruktur
  extensions: [".sql"]
excel:
  output_filename: "{JIRA}.xlsx"
```

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
| GET | `/api/scripts` | SQL-Inventar scannen |
| POST | `/api/scripts/analyze?path=...` | Einzelne SQL-Datei analysieren |
| GET | `/api/suggestions` | Suggestions auflisten |
| POST | `/api/suggestions` | Suggestion anlegen |
| POST | `/api/suggestions/batch` | Mehrere Suggestions anlegen |
| GET | `/api/gaps` | Einfache Gap-Analyse (SQL-Referenzen) |
| GET | `/api/match` | Task gegen SQL-Dateien matchen |

Die API speichert Suggestions in `inventory/suggestions.json`
(überschreibbar über `JIRA_SCANNER_DB`).

---

## SQL-Inventar (CLI)

```bash
python3 scripts/script_inventory.py
python3 scripts/script_inventory.py --path /mount --json
```

Das Inventar erkennt den Task-Key aus dem Verzeichnis, SQL-Technologien,
Datenbankzugriffe, API-Aufrufe, KPI-/Dashboard-Begriffe und eine grobe
Komplexität.

---

## Projektstruktur

```text
config.py                     gemeinsame Config-/Logging-Schicht
main.py                       kanonischer Scanner + Excel-Export pro Task
scanner/main_scanner.py       Kompatibilitäts-Wrapper (delegiert an main.py)
scripts/script_inventory.py   SQL-Analyse und Inventar-CLI
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
ruff check .
```

Die Smoke-Tests prüfen u. a. ADF-/Datums-Parsing, Jira-Response-Verarbeitung,
Prefix-Filter, Task-Key-Ableitung aus der Verzeichnisstruktur, SQL-Matching,
den Excel-Export pro Task und die API-Endpunkte.

---

## Fehlerbehandlung

| Situation | Verhalten |
|-----------|-----------|
| Jira nicht konfiguriert | Warnung + leerer Fallback-Bericht |
| Jira nicht erreichbar / 5xx | Retries, danach Warnung + leerer Fallback-Bericht |
| 401/403 | Klare Fehlermeldung im Log, leerer Fallback-Bericht |
| Scan-Root fehlt | Warnung, keine Berichte |
| SQL-Datei außerhalb eines Task-Verzeichnisses | Warnung, Datei wird ignoriert |
| Task ohne SQL-Datei | Warnung, kein Bericht für diesen Task |
| Einzelne SQL-Datei unlesbar | Warnung, restliche Dateien werden verarbeitet |
| Excel nicht schreibbar | Fehlercode 1 + Logeintrag |

Logs: `logs/jira_scanner.log` (rotierend).

---

## Bekannte Grenzen / nächste Schritte

- Es werden ausschließlich `.sql`-Dateien berücksichtigt.
- Die Task-Zuordnung hängt strikt an der Verzeichnisstruktur.
- Confluence-Integration ist nur als Konfigurationsblock vorhanden.
- Gap-Analyse ist eine einfache Referenzprüfung, kein semantisches Matching.
- Kein automatisches Schreiben zurück nach Jira.
- Kein Multi-Project-Support in einem Lauf.
- Dashboard zeigt Tasks aus dem Cache; Live-Refresh über `?refresh=true`.

---

## Lizenz

MIT
