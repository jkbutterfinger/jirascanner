#!/usr/bin/env bash
# Jira/Gira Scanner – Setup-Skript
# ---------------------------------------------------------------------------
# Richtet die Python-Umgebung des Projekts ein:
#   1. passenden Python-Interpreter (>= 3.9) finden
#   2. virtuelle Umgebung .venv anlegen (oder erneuern)
#   3. pip/setuptools/wheel aktualisieren
#   4. Abhängigkeiten aus requirements.txt bzw. requirements-dev.txt
#      installieren
#   5. Installation prüfen
#
# Beispiele:
#   ./setup.sh                 # Laufzeit-Abhängigkeiten
#   ./setup.sh --dev           # inkl. pytest, ruff, httpx
#   ./setup.sh --dev --test    # zusätzlich Tests und Lint ausführen
#   ./setup.sh --recreate      # .venv komplett neu aufbauen
#   ./setup.sh --venv /tmp/venv
# ---------------------------------------------------------------------------

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="${VENV_DIR:-.venv}"
MODE="runtime"
RECREATE=0
SKIP_INSTALL=0
RUN_CHECKS=0
EDITABLE=0

usage() {
    cat <<EOF
Verwendung: $(basename "$0") [Optionen]

Optionen:
  --dev            Entwicklungs-Abhängigkeiten installieren (requirements-dev.txt)
  --recreate       Vorhandene virtuelle Umgebung löschen und neu anlegen
  --venv DIR       Anderes Verzeichnis für die virtuelle Umgebung (Default: .venv)
  --no-install     Umgebung anlegen, aber keine Pakete installieren
  --editable       Projekt zusätzlich editierbar installieren (pip install -e .)
  --test           Nach dem Setup pytest und ruff ausführen
  -h, --help       Diese Hilfe anzeigen

Umgebungsvariablen:
  PYTHON           Bevorzugter Interpreter, z. B. PYTHON=python3.12 ./setup.sh
  VENV_DIR         Verzeichnis der virtuellen Umgebung (Default: .venv)

Beispiele:
  ./setup.sh
  ./setup.sh --dev --test
  PYTHON=python3.12 ./setup.sh --recreate
EOF
}

log()  { printf '%s\n' "$*"; }
info() { printf 'ℹ️  %s\n' "$*"; }
ok()   { printf '✅ %s\n' "$*"; }
warn() { printf '⚠️  %s\n' "$*"; }
fail() { printf '❌ %s\n' "$*" >&2; exit 1; }

while [[ $# -gt 0 ]]; do
    case "$1" in
        --dev)        MODE="dev" ;;
        --recreate)   RECREATE=1 ;;
        --venv)       [[ $# -ge 2 ]] || fail "--venv benötigt ein Verzeichnis"; VENV_DIR="$2"; shift ;;
        --no-install) SKIP_INSTALL=1 ;;
        --editable)   EDITABLE=1 ;;
        --test)       RUN_CHECKS=1 ;;
        -h|--help)    usage; exit 0 ;;
        *)            fail "Unbekannte Option: $1 (siehe --help)" ;;
    esac
    shift
done

# --- Python-Interpreter finden ---------------------------------------------
find_python() {
    local candidates=()
    if [[ -n "${PYTHON:-}" ]]; then
        candidates+=("$PYTHON")
    fi
    candidates+=(python3.13 python3.12 python3.11 python3.10 python3.9 python3)

    local candidate
    for candidate in "${candidates[@]}"; do
        if ! command -v "$candidate" >/dev/null 2>&1; then
            continue
        fi
        if "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 9) else 1)' >/dev/null 2>&1; then
            printf '%s\n' "$candidate"
            return 0
        fi
    done
    return 1
}

if ! BASE_PYTHON="$(find_python)"; then
    fail "Kein Python >= 3.9 gefunden. Bitte installieren, z. B.: sudo apt install python3 python3-venv"
fi

BASE_VERSION="$("$BASE_PYTHON" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])')"
info "Python-Interpreter: $BASE_PYTHON ($BASE_VERSION)"

# --- Virtuelle Umgebung anlegen --------------------------------------------
if [[ "$RECREATE" -eq 1 && -d "$VENV_DIR" ]]; then
    info "Entferne vorhandene Umgebung: $VENV_DIR"
    rm -rf "$VENV_DIR"
fi

if [[ -d "$VENV_DIR" && ! -x "$VENV_DIR/bin/python" && ! -x "$VENV_DIR/Scripts/python.exe" ]]; then
    warn "Verzeichnis $VENV_DIR existiert, ist aber keine gültige venv – wird neu angelegt."
    rm -rf "$VENV_DIR"
fi

if [[ ! -d "$VENV_DIR" ]]; then
    info "Lege virtuelle Umgebung an: $VENV_DIR"
    if ! "$BASE_PYTHON" -m venv "$VENV_DIR"; then
        fail "venv konnte nicht erstellt werden. Fehlt das Paket python3-venv? (sudo apt install python3-venv)"
    fi
else
    info "Vorhandene virtuelle Umgebung wird verwendet: $VENV_DIR"
fi

if [[ -x "$VENV_DIR/bin/python" ]]; then
    VENV_PY="$VENV_DIR/bin/python"
elif [[ -x "$VENV_DIR/Scripts/python.exe" ]]; then
    VENV_PY="$VENV_DIR/Scripts/python.exe"
else
    fail "Kein Python in $VENV_DIR gefunden."
fi

VENV_VERSION="$("$VENV_PY" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])')"
ok "Virtuelle Umgebung aktiv: $VENV_DIR (Python $VENV_VERSION)"

# --- Pakete installieren ---------------------------------------------------
if [[ "$SKIP_INSTALL" -eq 0 ]]; then
    info "Aktualisiere pip, setuptools und wheel"
    "$VENV_PY" -m pip install --upgrade pip setuptools wheel

    REQUIREMENTS="requirements.txt"
    if [[ "$MODE" == "dev" ]]; then
        REQUIREMENTS="requirements-dev.txt"
    fi

    if [[ -f "$REQUIREMENTS" ]]; then
        info "Installiere Abhängigkeiten aus $REQUIREMENTS"
        "$VENV_PY" -m pip install -r "$REQUIREMENTS"
    else
        warn "$REQUIREMENTS nicht gefunden – installiere aus pyproject.toml"
        if [[ "$MODE" == "dev" ]]; then
            "$VENV_PY" -m pip install -e ".[dev]"
            EDITABLE=0
        else
            "$VENV_PY" -m pip install .
        fi
    fi

    if [[ "$EDITABLE" -eq 1 ]]; then
        info "Installiere Projekt editierbar (pip install -e .)"
        "$VENV_PY" -m pip install -e .
    fi
fi

# --- Installation prüfen ---------------------------------------------------
if [[ "$SKIP_INSTALL" -eq 0 ]]; then
    info "Prüfe Kernabhängigkeiten"
    "$VENV_PY" - <<'PY'
import importlib
import sys

modules = ["yaml", "requests", "openpyxl", "fastapi", "uvicorn", "pydantic"]
missing = []
for name in modules:
    try:
        importlib.import_module(name)
    except Exception as exc:  # noqa: BLE001 - Diagnose
        missing.append(f"{name} ({exc})")

if missing:
    print("Fehlende Module: " + ", ".join(missing), file=sys.stderr)
    raise SystemExit(1)
print("Alle Kernabhängigkeiten importierbar.")
PY
    ok "Abhängigkeiten geprüft"
fi

if [[ "$RUN_CHECKS" -eq 1 ]]; then
    info "Führe Tests aus"
    "$VENV_PY" -m pytest
    info "Führe Lint aus"
    "$VENV_PY" -m ruff check .
    ok "Tests und Lint erfolgreich"
fi

# --- Abschluss -------------------------------------------------------------
log ""
ok "Setup abgeschlossen"
log ""
log "Aktivieren:"
log "  source \"$VENV_DIR/bin/activate\""
log ""
log "Starten:"
log "  $VENV_PY main.py --mode scan"
log "  ./run.sh --mode scan"
