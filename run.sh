#!/bin/bash
# Jira/Gira Scanner – Startskript
# ---------------------------------------------------------------------------
# Prüft Voraussetzungen, legt Ausgabeordner an und startet main.py.
# Alle Argumente werden an main.py weitergegeben, z. B.:
#   ./run.sh --mode scan
#   ./run.sh --mode enrich --batch-size 50 --batch-delay 2
#   ./run.sh --mode all --dry-run
# ---------------------------------------------------------------------------

set -u

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

# Virtuelle Umgebung automatisch bevorzugen (kein "source" nötig).
if [[ -z "${PYTHON:-}" ]]; then
    if [[ -x "$SCRIPT_DIR/.venv/bin/python" ]]; then
        PYTHON="$SCRIPT_DIR/.venv/bin/python"
    elif [[ -x "$SCRIPT_DIR/venv/bin/python" ]]; then
        PYTHON="$SCRIPT_DIR/venv/bin/python"
    else
        PYTHON="python3"
    fi
fi

echo "=============================================="
echo "🔍 JIRA/GIRA SCANNER"
echo "SQL-Scanner: Phase 1 Verzeichnis, Phase 2 Jira-Anreicherung"
echo "=============================================="
echo ""

if ! command -v "$PYTHON" >/dev/null 2>&1; then
    echo "❌ Fehler: $PYTHON wurde nicht gefunden."
    exit 1
fi

# Abhängigkeiten prüfen
if ! "$PYTHON" -c "import yaml, requests, openpyxl" >/dev/null 2>&1; then
    echo "❌ Fehlende Abhängigkeiten."
    echo ""
    echo "Bitte zuerst die Umgebung einrichten:"
    echo "   ./setup.sh"
    echo "   # oder inkl. Test-Werkzeugen:"
    echo "   ./setup.sh --dev"
    echo ""
    exit 1
fi

# Konfiguration anlegen, falls sie fehlt
if [[ ! -f "config.yaml" ]]; then
    echo "⚠️  config.yaml fehlt – Defaults aus config.py werden verwendet."
    echo "   Für eine echte Konfiguration: config.yaml anlegen (siehe README)."
    echo ""
fi

# Ausgabe- und Logordner
mkdir -p output logs inventory jira_data

echo "Starte Scanner..."
echo ""

"$PYTHON" main.py "$@"
EXIT_CODE=$?

echo ""
if [[ $EXIT_CODE -eq 0 ]]; then
    echo "✅ Bericht erfolgreich erzeugt."
    echo ""
    echo "Ausgabedateien:"
    shopt -s nullglob
    reports=(output/*.xlsx output/*.csv)
    if [[ ${#reports[@]} -gt 0 ]]; then
        ls -lh "${reports[@]}"
    else
        echo "  (keine Dateien unter output/)"
    fi
    shopt -u nullglob
    echo "Gesamtbericht liegt unter output/ (excel.output_path)."
    echo "Scan-Index: inventory/scan_index.json"
else
    echo "❌ Berichtserzeugung fehlgeschlagen (Exit-Code $EXIT_CODE)."
    echo "Details siehe logs/jira_scanner.log"
fi

echo "=============================================="
exit $EXIT_CODE
