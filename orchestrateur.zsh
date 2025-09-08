#!/usr/bin/env zsh
# Portable Mac orchestrator for contact workflows
# - Menu via AppleScript
# - Uses local Python venv for pandas
# - Optionally calls PowerShell 7 + Exchange Online module

set -euo pipefail

SCRIPT_DIR="$(cd -- "$(dirname -- "${(%):-%N}")" && pwd)"
cd "$SCRIPT_DIR"

CONFIG_FILE="$SCRIPT_DIR/config.env"
[[ -f "$CONFIG_FILE" ]] && source "$CONFIG_FILE" || true

# Defaults (can be overridden in config.env)
OUTPUT_DIR="${OUTPUT_DIR:-$HOME/Desktop}"
VAULT_DIR="${VAULT_DIR:-}"
VAULT_NAME="${VAULT_NAME:-}"
VAULT_SUBPATH="${VAULT_SUBPATH:-}"
DATA_GOUV_URL="${DATA_GOUV_URL:-}"
SMTP_DOMAIN="${SMTP_DOMAIN:-}"
EXCHANGE_ENABLE_REMOVAL="${EXCHANGE_ENABLE_REMOVAL:-false}"
EXCHANGE_HARD_DELETE="${EXCHANGE_HARD_DELETE:-false}"
SUBCHOICE="${SUBCHOICE:-}"
EXCHANGE_SHOW_TERMINAL="${EXCHANGE_SHOW_TERMINAL:-true}"

VENV_DIR="$SCRIPT_DIR/.venv"
PYBIN="$VENV_DIR/bin/python3"

die() { echo "[ERROR] $*" >&2; exit 1; }

require_cmd() { command -v "$1" >/dev/null 2>&1 || die "Commande introuvable: $1"; }

notify() {
  local msg="$1"
  osascript -e 'display notification '"$msg"' with title "Contacts Tool"' >/dev/null 2>&1 || true
}

# (legacy helper removed)

# Deterministic menus that return exactly the chosen label (no prefix)
ask_main_choice() {
  osascript \
    -e 'set theButtons to {"Microsoft List", "data.gouv.fr", "Annuler"}' \
    -e 'set theAns to button returned of (display dialog "Que souhaitez-vous faire ?" with title "Contacts Tool" buttons theButtons default button "Microsoft List")' \
    -e 'theAns' 2>/dev/null || true
}

ask_mslist_choice() {
  sleep 0.2
  osascript <<'APPLESCRIPT'
set theChoices to {"VCF (Apple Contacts)", "Obsidian Vault", "Exchange"}
set theSel to choose from list theChoices with title "Contacts Tool" with prompt "Traitement Microsoft List" OK button name "Valider" cancel button name "Annuler" without multiple selections allowed
if theSel is false then
  return ""
else
  return (item 1 of theSel)
end if
APPLESCRIPT
}

tty_mslist_choice() {
  echo "Sélection (terminal) :" >&2
  local options=("VCF (Apple Contacts)" "Obsidian Vault" "Exchange" "Annuler")
  select opt in "${options[@]}"; do
    case "$REPLY" in
      1|2|3|4) echo "$opt"; return 0 ;;
      *) echo "Choix invalide" >&2 ;;
    esac
  done
}

# Combined AppleScript: choose file + show submenu in one GUI session
ask_mslist_flow() {
  osascript <<'APPLESCRIPT'
try
  set theFile to POSIX path of (choose file with prompt "Sélectionnez le fichier CSV (Microsoft List)")
on error errMsg number errNum
  return "ERR:" & errNum & ":" & errMsg
end try
set theChoices to {"VCF (Apple Contacts)", "Obsidian Vault", "Exchange"}
set theSel to choose from list theChoices with title "Contacts Tool" with prompt "Traitement Microsoft List" OK button name "Valider" cancel button name "Annuler" without multiple selections allowed
if theSel is false then
  return "ERR:-128:Cancelled"
else
  return "OK:" & theFile & linefeed & (item 1 of theSel)
end if
APPLESCRIPT
}

choose_file_csv() {
  # Use a permissive chooser (no type filter) for better compatibility across macOS versions
  # Returns empty on cancel or error
  osascript -e 'try
    POSIX path of (choose file with prompt "Sélectionnez le fichier CSV (Microsoft List)")
  end try' 2>/dev/null || true
}

choose_folder() {
  osascript -e 'POSIX path of (choose folder with prompt "Sélectionnez le dossier" )'
}

run_bootstrap_if_needed() {
  if [[ ! -x "$PYBIN" ]]; then
    echo "[INFO] Environnement Python manquant: exécution du bootstrap…"
    zsh "$SCRIPT_DIR/bootstrap.zsh" || die "Bootstrap échoué."
  fi
}

ensure_pwsh_hint() {
  if ! command -v pwsh >/dev/null 2>&1; then
    echo "[WARN] PowerShell 7 (pwsh) introuvable. Requis pour l'import Exchange."
    echo "      Installation via Homebrew: brew install --cask powershell"
    echo "      Ou: https://learn.microsoft.com/powershell/scripting/install/installing-powershell"
    return 1
  fi
  return 0
}

reveal_in_finder() {
  local p="$1"
  [[ -n "$p" ]] || return 0
  osascript <<APPLESCRIPT >/dev/null 2>&1 || true
tell application "Finder"
  try
    set theItem to POSIX file "$p" as alias
    reveal theItem
    activate
  end try
end tell
APPLESCRIPT
}

open_terminal_with_command() {
  local cmd="$1"
  local tmp
  tmp=$(mktemp -t exch_import.XXXXXX)
  printf '%s
' "$cmd" > "$tmp"
  chmod +x "$tmp"
  osascript <<APPLESCRIPT >/dev/null 2>&1 || true
 tell application "Terminal"
  activate
  do script "zsh " & quoted form of POSIX path of "$tmp"
 end tell
APPLESCRIPT
}


# --- Preflight installation assistant ---
apple_confirm_install() {
  # Returns 0 if user confirms (Installer), 1 otherwise
  osascript <<'APPLESCRIPT' >/dev/null 2>&1
set theButtons to {"Installer", "Plus tard"}
set theBtn to button returned of (display dialog "Installer les dépendances manquantes ?" with title "Contacts Tool" buttons theButtons default button "Installer" with icon note)
if theBtn is equal to "Installer" then
  return
else
  error number -128
end if
APPLESCRIPT
}

has_python_deps() {
  [[ -x "$PYBIN" ]] || return 1
  "$PYBIN" - <<'PY'
try:
    import pandas  # noqa: F401
except Exception:
    raise SystemExit(2)
raise SystemExit(0)
PY
}

has_pwsh() { command -v pwsh >/dev/null 2>&1; }
has_exchange_module() {
  pwsh -NoLogo -NoProfile -Command 'if (Get-Module -ListAvailable ExchangeOnlineManagement) { "ok" }' 2>/dev/null | grep -q '^ok$'
}

install_pwsh_via_brew() {
  if command -v brew >/dev/null 2>&1; then
    echo "[INSTALL] Homebrew: powershell"
    brew install --cask powershell || return 1
    return 0
  else
    echo "[WARN] Homebrew introuvable; impossible d'installer pwsh automatiquement."
    osascript -e "display dialog \"Homebrew n'est pas installé. Installez PowerShell 7 depuis la page Microsoft.\" with title \"Contacts Tool\" buttons {\"Ouvrir la page\", \"OK\"} default button \"Ouvrir la page\"" >/dev/null 2>&1 && \
      open "https://learn.microsoft.com/powershell/scripting/install/installing-powershell"
    return 1
  fi
}

install_exchange_module() {
  echo "[INSTALL] ExchangeOnlineManagement (scope utilisateur)"
  pwsh -NoLogo -NoProfile -Command 'try { Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop; "ok" } catch { Write-Host ("ERR: " + $_.Exception.Message); exit 1 }' || return 1
  return 0
}

preflight_setup() {
  local need_py=1 need_pwsh=1 need_exo=1
  if has_python_deps >/dev/null 2>&1; then need_py=0; fi
  if has_pwsh; then need_pwsh=0; fi
  if [[ $need_pwsh -eq 0 ]] && has_exchange_module; then need_exo=0; fi

  if [[ $need_py -eq 0 && $need_pwsh -eq 0 && $need_exo -eq 0 ]]; then
    return 0
  fi

  echo "[SETUP] Dépendances manquantes:" \
       $([[ $need_py -eq 1 ]] && echo " python(pandas)") \
       $([[ $need_pwsh -eq 1 ]] && echo " pwsh") \
       $([[ $need_exo -eq 1 ]] && echo " ExchangeOnlineManagement")

  if apple_confirm_install; then
    # Install pwsh first if needed
    if [[ $need_pwsh -eq 1 ]]; then
      install_pwsh_via_brew || true
    fi
    # Always run bootstrap (creates venv + installs pandas; installs Exchange module if pwsh exists)
    zsh "$SCRIPT_DIR/bootstrap.zsh" || true
    # If Exchange module still missing, try explicit install
    if has_pwsh && ! has_exchange_module; then
      install_exchange_module || true
    fi
  fi
}

# --- Helpers for flows ---
do_csv_to_vcf() {
  local in_csv="$1"
  run_bootstrap_if_needed
  "$PYBIN" "$SCRIPT_DIR/csv_outlook/csv_contact_batch.py" "$in_csv"
  local out_guess="${in_csv:h}/${in_csv:t:r}_contacts.vcf"
  if [[ -f "$out_guess" ]]; then
    mkdir -p "$OUTPUT_DIR"
    local final="$OUTPUT_DIR/${out_guess:t}"
    mv -f "$out_guess" "$final"
    notify "VCF généré: ${final:t}"
    echo "[OK] $final"
  else
    echo "[WARN] Fichier VCF de sortie introuvable (attendu: $out_guess)"
  fi
}

do_csv_to_vault() {
  local in_csv="$1"
  run_bootstrap_if_needed
  local cfg_json="${VAULT_MAPPING_JSON:-$SCRIPT_DIR/csv_vault/mapping.example.json}"

  local args=("$SCRIPT_DIR/csv_vault/csv_to_obsidian_contacts.py" --csv "$in_csv" --config "$cfg_json")
  if [[ -n "$VAULT_NAME" ]]; then
    args+=(--vault-name "$VAULT_NAME")
    [[ -n "$VAULT_SUBPATH" ]] && args+=(--out-subpath "$VAULT_SUBPATH")
  elif [[ -n "$VAULT_DIR" ]]; then
    if [[ -n "$VAULT_SUBPATH" ]]; then
      args+=(--vault "$VAULT_DIR" --out-subpath "$VAULT_SUBPATH")
    else
      args+=(--out "$VAULT_DIR")
    fi
  else
    echo "[INFO] VAULT non défini. Choisissez un dossier cible (Vault)."
    local chosen
    chosen="$(choose_folder)" || die "Sélection annulée."
    args+=(--out "$chosen")
  fi

  "$PYBIN" "${args[@]}"
  notify "Obsidian: génération terminée."
}

do_csv_to_exchange() {
  local in_csv="$1"
  ensure_pwsh_hint || return 1

  local ps1="$SCRIPT_DIR/csv_exchange/sync_gal_by_id.ps1"
  [[ -f "$ps1" ]] || die "Script Exchange introuvable: $ps1"

  local flags=(-File "$ps1" -CsvPath "$in_csv")
  if [[ "${EXCHANGE_ENABLE_REMOVAL}" == "true" ]]; then flags+=(-EnableRemoval); fi
  if [[ "${EXCHANGE_HARD_DELETE}" == "true" ]]; then flags+=(-HardDelete); fi
  if [[ -n "${SMTP_DOMAIN}" ]]; then flags+=(-SmtpDomain "$SMTP_DOMAIN"); fi
  flags+=(-Apply)

  echo "[INFO] Lancement import Exchange (pwsh)…"
  local cmd q
  # Build a shell-safe command line for Terminal using %q quoting
  q=$(printf %q "$SCRIPT_DIR")
  cmd="cd $q; pwsh"
  for a in "${flags[@]}"; do
    q=$(printf %q "$a")
    cmd+=" $q"
  done
  cmd+="; echo; echo 'Import Exchange terminé.'; echo; read -r '?Appuyez sur Entrée pour fermer…'"

  if [[ "$EXCHANGE_SHOW_TERMINAL" == "true" ]]; then
    open_terminal_with_command "$cmd"
  else
    pwsh "${flags[@]}"
  fi
  notify "Import Exchange terminé."
}

do_datagouv_to_exchange() {
  [[ -n "$DATA_GOUV_URL" ]] || die "DATA_GOUV_URL non défini dans config.env"
  # Prépare les répertoires de travail
  run_bootstrap_if_needed
  local in_dir="$SCRIPT_DIR/csv_fr_csv_list/input"
  local out_dir="$SCRIPT_DIR/csv_fr_csv_list/output"
  mkdir -p "$in_dir" "$out_dir"
  local ts="$(date +%Y%m%d_%H%M%S)"
  local base="data_gouv_${ts}.csv"
  local staged="$in_dir/$base"

  echo "[INFO] Téléchargement: $DATA_GOUV_URL"
  curl -L --fail "$DATA_GOUV_URL" -o "$staged" || die "Téléchargement échoué."

  echo "[INFO] Formatage data.gouv…"
  "$PYBIN" "$SCRIPT_DIR/csv_fr_csv_list/csv_fr_csv_list.py"

  # Chercher d'abord le fichier transformé correspondant au nom téléchargé
  local transformed expected
  expected="$out_dir/${base%.*}__transformed.csv"
  if [[ -f "$expected" ]]; then
    transformed="$expected"
  else
    # Fallback: dernier __transformed.csv (au cas où le script renomme différemment)
    transformed="$(ls -t "$out_dir"/*__transformed.csv 2>/dev/null | head -n1 || true)"
  fi
  [[ -n "$transformed" ]] || die "Aucun CSV transformé trouvé dans $out_dir"
  echo "[OK] Transformé: $transformed"

  # Déposer aussi la version transformée sur le bureau pour clarté
  mkdir -p "$OUTPUT_DIR"
  local final_out="$OUTPUT_DIR/${base%.*}__transformed.csv"
  if cp -f "$transformed" "$final_out"; then
    echo "[INFO] Copie du CSV transformé: $final_out"
    notify "CSV transformé copié: ${final_out:t}"
    reveal_in_finder "$final_out"
  else
    echo "[WARN] Échec copie transformé vers $final_out" >&2
  fi

  do_csv_to_exchange "$transformed"
}

# --- Main menu ---

# Allow Quick Action file argument
INPUT_FROM_ARGS="${1:-}"
if [[ -n "$INPUT_FROM_ARGS" && ! -f "$INPUT_FROM_ARGS" ]]; then
  echo "[WARN] Argument fourni mais introuvable: $INPUT_FROM_ARGS" >&2
  INPUT_FROM_ARGS=""
fi

# Offer to install missing deps up front
preflight_setup || true

main_choice=$(ask_main_choice)
case "$main_choice" in
  Microsoft\ List )
    # If SUBCHOICE provided (from Shortcuts/Automator), skip GUI submenu
    if [[ -n "$SUBCHOICE" ]]; then
      in_csv="$INPUT_FROM_ARGS"
      if [[ -z "$in_csv" ]]; then
        in_csv="$(choose_file_csv)"
      fi
      [[ -n "$in_csv" ]] || die "Sélection annulée."
      [[ -f "$in_csv" ]] || die "Fichier introuvable."
      case "$SUBCHOICE" in
        VCF*|vcf*|Apple* ) do_csv_to_vcf "$in_csv" ;;
        Obsidian*|vault*|VAULT* ) do_csv_to_vault "$in_csv" ;;
        Exchange*|exchange*|EXCHANGE* ) do_csv_to_exchange "$in_csv" ;;
        * ) die "SUBCHOICE inconnu: $SUBCHOICE" ;;
      esac
      exit 0
    fi

    if [[ -z "$INPUT_FROM_ARGS" ]]; then
      echo "[INFO] Sélection fichier + sous-menu (GUI)…"
      flow_res=$(ask_mslist_flow || true)
      if [[ -z "$flow_res" ]]; then
        die "Sélection annulée."
      elif [[ "$flow_res" == OK:* ]]; then
        tmp="${flow_res#OK:}"
        in_csv="${tmp%%$'\n'*}"
        sub_choice="${tmp#*$'\n'}"
      elif [[ "$flow_res" == ERR:* ]]; then
        err_payload="${flow_res#ERR:}"
        err_code="${err_payload%%:*}"
        err_msg="${err_payload#*:}"
        if [[ "$err_code" == "-128" ]]; then
          die "Sélection annulée."
        else
          echo "[ERROR] AppleScript ($err_code): $err_msg" >&2
          die "Erreur AppleScript ($err_code)"
        fi
      else
        # Backward compatibility: two-line output
        in_csv="${flow_res%%$'\n'*}"
        sub_choice="${flow_res#*$'\n'}"
      fi
    else
      in_csv="$INPUT_FROM_ARGS"
      [[ -f "$in_csv" ]] || die "Fichier introuvable."
      echo "[INFO] Ouverture du sous-menu…"
      sub_choice=$(ask_mslist_choice || true)
      [[ -n "$sub_choice" ]] || sub_choice=$(tty_mslist_choice)
    fi
    case "$sub_choice" in
      VCF* ) do_csv_to_vcf "$in_csv" ;;
      Obsidian* ) do_csv_to_vault "$in_csv" ;;
      Exchange ) do_csv_to_exchange "$in_csv" ;;
      * ) exit 0 ;;
    esac
    ;;
  data.gouv.fr )
    do_datagouv_to_exchange ;;
  * )
    exit 0 ;;
esac
