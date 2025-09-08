#!/usr/bin/env zsh
# Prepare local environment for the contact orchestrator
# - Creates Python venv and installs requirements
# - Checks PowerShell 7 and installs ExchangeOnlineManagement module

set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${(%):-%N}")" && pwd)"
cd "$SCRIPT_DIR"

VENV_DIR="$SCRIPT_DIR/.venv"

echo "[BOOTSTRAP] Vérification Python3…"
if ! command -v python3 >/dev/null 2>&1; then
  echo "[ERROR] python3 introuvable. Installez Xcode CLT ou Python.org." >&2
  exit 1
fi

if [[ ! -d "$VENV_DIR" ]]; then
  echo "[BOOTSTRAP] Création venv: $VENV_DIR"
  python3 -m venv "$VENV_DIR"
fi

echo "[BOOTSTRAP] Installation dépendances Python…"
"$VENV_DIR/bin/python3" -m pip install --upgrade pip
if [[ -f "$SCRIPT_DIR/csv_fr_csv_list/requirements.txt" ]]; then
  "$VENV_DIR/bin/python3" -m pip install -r "$SCRIPT_DIR/csv_fr_csv_list/requirements.txt"
fi

echo "[BOOTSTRAP] Vérification PowerShell 7 (pwsh)…"
if ! command -v pwsh >/dev/null 2>&1; then
  echo "[WARN] PowerShell 7 (pwsh) introuvable. Requis pour Exchange."
  echo "       Installation via Homebrew: brew install --cask powershell"
  echo "       Ou: https://learn.microsoft.com/powershell/scripting/install/installing-powershell"
else
  echo "[BOOTSTRAP] Installation module ExchangeOnlineManagement (scope utilisateur)…"
  pwsh -NoLogo -NoProfile -Command "try { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop; Write-Host '[OK] Module installé' } catch { Write-Host '[INFO] Skip: ' + \
    (\"$($_.Exception.Message)\") }"
fi

echo "[BOOTSTRAP] Terminé."

