# Jira/Gira Scanner

Der Scanner durchsucht ein Verzeichnis ausschließlich nach **SQL-Dateien**.
Die Jira-Referenz wird aus der Verzeichnisstruktur abgeleitet – nicht aus dem
Dateiinhalt. Für jeden Jira-Task (Auftrag) wird ein eigener Excel-Bericht
direkt im Task-Verzeichnis abgelegt.

**Standard-Prefix:** `JRA-` · **Version:** 1.3.0 · **Python:** ≥ 3.9

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

# Python-Umgebung einrichten (findet Python >= 3.9, legt .venv an,
# installiert alle Abhängigkeiten)
./setup.sh

# optional inkl. Test-/Lint-Werkzeugen:
./setup.sh --dev

# Konfiguration anpassen (Jira-URL, Token, Scan-Root)
$EDITOR config.yaml

# Phase 1: Verzeichnisstruktur einlesen (ohne Jira)
python3 main.py --mode scan

# Phase 2: Jira-Details nachladen (max. 100 Keys pro Request)
python3 main.py --mode enrich

# Beides nacheinander (Default)
python3 main.py --mode all

# oder mit Komfortprüfungen (nutzt .venv automatisch):
./run.sh --mode scan
```

Alternativ Token per Umgebungsvariable (empfohlen, kein Secret in Git):

```bash
export JIRA_URL="https://jira.example.com"
export JIRA_API_TOKEN="..."
export JIRA_PROJECT_KEY="JRA"
export JIRA_SCRIPT_PATH="/mount"
python3 main.py
```

### Setup-Skript

`setup.sh` richtet die Python-Umgebung ein und ist idempotent:

```bash
./setup.sh                  # Python >= 3.9 finden, .venv anlegen, Laufzeit-Pakete
./setup.sh --dev            # zusätzlich pytest, ruff, httpx
./setup.sh --dev --test     # danach pytest und ruff ausführen
./setup.sh --recreate       # .venv löschen und komplett neu aufbauen
./setup.sh --venv /tmp/venv # andere virtuelle Umgebung verwenden
PYTHON=python3.12 ./setup.sh
```

Das Skript prüft die Python-Version, legt bei Bedarf `.venv` an, aktualisiert
`pip`/`setuptools`/`wheel`, installiert `requirements.txt` bzw.
`requirements-dev.txt` und verifiziert die Kernimporte. `run.sh` erkennt die
`.venv` danach automatisch – ein manuelles `source` ist nicht nötig.

### CLI-Optionen

```bash
python3 main.py --help
python3 main.py --mode scan                      # nur Verzeichnisstruktur
python3 main.py --mode enrich                    # Jira-Details nachladen
python3 main.py --mode all                       # beides (Default)
python3 main.py --mode enrich --batch-size 50 --batch-delay 2
python3 main.py --mode all --dry-run             # nur Konsole, keine Datei
python3 main.py --prefix JRA- -v                 # DEBUG-Logging
python3 main.py --config other.yaml
```

---

## Aufrufmodi

Der Scanner ist in zwei Phasen aufgeteilt, die einzeln oder zusammen laufen.

### 1. Verzeichnisstruktur einlesen (ohne Jira)

```bash
python3 main.py --mode scan
```

- Liest ausschließlich die Verzeichnisstruktur und die `.sql`-Dateien.
- Erzeugt pro Task bereits `<Jira-Task>.xlsx` mit File-Datum, Pfad,
  erkannten Technologien sowie Jahr und Abteilung aus der Verzeichnisstruktur.
- Jira-Spalten bleiben leer, Status = `Pending`.
- Schreibt die Inventur nach `storage.scan_index` (Standard:
  `inventory/scan_index.json`).
- Berücksichtigt nur Task-Keys mit `filters.prefix` (Standard `JRA-`);
  `--prefix` überschreibt den Wert.
- Benötigt keine Jira-Konfiguration und keinen Jira-Zugriff.

### 2. Jira-Details anreichern

```bash
python3 main.py --mode enrich
```

- Liest den Scan-Index. Fehlt er, wird die Verzeichnisstruktur automatisch
  zuerst eingelesen.
- Fragt nur die tatsächlich gefundenen Task-Keys ab (kein projektweiter Scan).
- Holt die Details in Blöcken von **maximal 100 Keys** pro Request.
- Pausiert zwischen den Blöcken (`jira.batch_delay_seconds`, Standard 1 s),
  damit der Jira-Server nicht überlastet wird.
- Schreibt die angereicherten Berichte in die jeweiligen Task-Verzeichnisse.
- Keys ohne Jira-Treffer behalten die Scan-Daten und erhalten Status
  `Missing`.

Parameter:

```bash
python3 main.py --mode enrich --batch-size 50 --batch-delay 2
```

| Option | Bedeutung | Default |
|--------|-----------|---------|
| `--batch-size` | Keys pro Jira-Request (wird auf 1–100 begrenzt) | `jira.batch_size` (100) |
| `--batch-delay` | Pause zwischen Blöcken in Sekunden | `jira.batch_delay_seconds` (1.0) |

### Beides nacheinander

```bash
python3 main.py --mode all      # Default
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
| 11 | Jahr | Jahr aus der Verzeichnisstruktur (lokal) |
| 12 | Abteilung | Abteilung aus der Verzeichnisstruktur (lokal) |

In Phase 1 (`--mode scan`) bleiben die Jira-Spalten 2–5, 9 und 10 leer;
Status ist `Pending`. Die lokalen Spalten 1, 6–8, 11 und 12 werden bereits
im Scan-Modus gefüllt. Phase 2 (`--mode enrich`) ergänzt die Jira-Details.

Beispiel für `/mount/2024/IT/ACME/JRA-1001/JRA-1001.xlsx`:

```text
Jira-Nummer | Jira-Auftraggeber | Entwickler | Jahr/Monat/Tag-Abgeschlossen | Zusammenfassung-Anforderung | File-Datum | File-Location                                  | Verwendete-Technologien | Status | Priorität | Jahr | Abteilung
JRA-1001    | Max Mustermann    | Anna Meier | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-10 | /mount/2024/IT/ACME/JRA-1001/monthly_sales.sql | SQL                     | Done   | High      | 2024 | IT
JRA-1001    | Max Mustermann    | Anna Meier | 2024-02-01 10:00:00          | Monthly Sales Dashboard     | 2024-01-12 | /mount/2024/IT/ACME/JRA-1001/etl_refresh.sql   | SQL                     | Done   | High      | 2024 | IT
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

1. **Phase 1 – Verzeichnis-Scan:** Alle `.sql`-Dateien unter
   `scripts.base_path` werden genau einmal gefunden und über den Namen ihres
   Elternverzeichnisses indexiert. Die Inventur wird als JSON gespeichert.
2. **Phase 2 – Jira-Anreicherung:** Nur die im Scan gefundenen Task-Keys
   werden in Blöcken von maximal 100 über die Jira-REST-API v3 abgefragt.
   Zwischen den Blöcken wird pausiert.
3. **Jira-Abfrage:** Bevorzugt `/rest/api/3/search/jql` mit `nextPageToken`.
   Fällt bei älteren Servern auf `/rest/api/3/search` mit `startAt` zurück.
   Pagination bis `performance.max_results`, Timeout und Retries sind
   konfigurierbar. ADF-Beschreibungen (API v3) werden in Text umgewandelt.
4. **Prefix-Filter:** Task-Keys werden clientseitig gegen `filters.prefix`
   geprüft (Standard `JRA-`).
5. **Export:** Ein Excel-Bericht pro Task direkt im Task-Verzeichnis.

Ohne gültige Jira-Konfiguration läuft `--mode scan` vollständig durch.
`--mode enrich` wird übersprungen; die vorhandenen Scan-Berichte bleiben
erhalten.

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
`performance`, `jira`, `storage`, `excel`, `logging`.

```yaml
project:
  key: "JRA"
filters:
  prefix: "JRA-"
scripts:
  base_path: "/mount"      # Root der Verzeichnisstruktur
  extensions: [".sql"]
storage:
  scan_index: "inventory/scan_index.json"
jira:
  batch_size: 100             # max. 100 Keys pro Request
  batch_delay_seconds: 1.0    # Pause zwischen Blöcken
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
setup.sh                      Python-/venv-Setup (Interpreter finden, .venv anlegen)
run.sh                        Startskript mit Vorabprüfungen (nutzt .venv automatisch)
requirements.txt              Laufzeit-Abhängigkeiten
requirements-dev.txt          Test-/Lint-Abhängigkeiten
```

---

## Tests

```bash
./setup.sh --dev --test    # richtet .venv ein, installiert Dev-Pakete,
                           # führt pytest und ruff aus

# oder manuell:
source .venv/bin/activate
pytest
ruff check .
```

Die Smoke-Tests prüfen u. a. ADF-/Datums-Parsing, Jira-Response-Verarbeitung,
Prefix-Filter, Task-Key-Ableitung aus der Verzeichnisstruktur, SQL-Matching,
den Excel-Export pro Task, die Scan-/Enrich-Phasen, die Batch-Grenze von
maximal 100 Keys und die API-Endpunkte.

---

## Fehlerbehandlung

| Situation | Verhalten |
|-----------|-----------|
| Jira nicht konfiguriert | Scan-Berichte bleiben erhalten, `enrich` wird übersprungen |
| Jira nicht erreichbar / 5xx | Retries, danach wird der Batch übersprungen |
| 401/403 | Klare Fehlermeldung im Log, Batch wird übersprungen |
| Scan-Index fehlt bei `enrich` | Verzeichnisstruktur wird automatisch eingelesen |
| Scan-Root fehlt | Warnung, keine Berichte |
| SQL-Datei außerhalb eines Task-Verzeichnisses | Warnung, Datei wird ignoriert |
| Task ohne SQL-Datei | Warnung, kein Bericht für diesen Task |
| Task-Key nicht in Jira gefunden | Scan-Daten bleiben, Status `Missing` |
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
