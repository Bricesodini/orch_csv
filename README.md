# orch_csv — Orchestrateur de flux CSV (macOS)

Outils et scripts pour orchestrer des conversions et synchronisations de contacts à partir de CSV : vCard pour Apple Contacts, notes Obsidian, et import Exchange Online. Un menu macOS (AppleScript) pilote les étapes courantes.

## Vue d’ensemble
- `orchestrateur.zsh` ouvre un menu avec deux entrées principales :
  - « Microsoft List » : traite un CSV exporté depuis Microsoft Lists.
    - Sorties possibles : `VCF (Apple Contacts)`, `Obsidian Vault`, `Exchange`.
  - « data.gouv.fr » : télécharge un CSV depuis `DATA_GOUV_URL`, le normalise, puis lance l’import Exchange.
- `bootstrap.zsh` prépare l’environnement (Python venv + dépendances, PowerShell/ExchangeOnlineManagement si présent).

Sous-dossiers :
- `csv_outlook/` : conversion CSV → vCard (`.vcf`) et CSV Outlook.
- `csv_vault/` : génération de notes Obsidian par contact (frontmatter YAML configurable).
- `csv_fr_csv_list/` : normalisation des CSV officiels (Annuaire Éducation) vers un format « Microsoft Lists friendly ».
- `csv_exchange/` : script PowerShell pour synchroniser la GAL Exchange Online à partir d’un CSV.

## Prérequis
- macOS (pour les boîtes de dialogue AppleScript et l’ouverture de Terminal/Finder).
- Python 3 (installé par défaut sur macOS récents) ; le script crée un venv local `./.venv`.
- `pandas` et dépendances : installées automatiquement via `bootstrap.zsh` (depuis `csv_fr_csv_list/requirements.txt`).
- Pour Exchange (optionnel) :
  - PowerShell 7 (`pwsh`) : `brew install --cask powershell`.
  - Module `ExchangeOnlineManagement` : installé dans le scope utilisateur par `bootstrap.zsh` ou via `Install-Module`.

## Démarrage rapide
1) Facultatif : exécuter l’installation locale
```
./bootstrap.zsh
```
2) Lancer l’orchestrateur
```
./orchestrateur.zsh
```
3) Choisir un flux :
- « Microsoft List » puis :
  - `VCF (Apple Contacts)` → produit `..._contacts.vcf` et le copie dans `OUTPUT_DIR`.
  - `Obsidian Vault` → crée/mettre à jour des notes Markdown dans un coffre Obsidian.
  - `Exchange` → appelle PowerShell pour créer/mettre à jour des Mail Contacts dans Exchange Online.
- « data.gouv.fr » → nécessite `DATA_GOUV_URL` dans `config.env` (voir ci‑dessous).

## Configuration (`config.env`)
Créer un fichier `config.env` à la racine si besoin ; variables supportées (toutes optionnelles) :
- `OUTPUT_DIR` : dossier de sortie par défaut pour les artefacts (défaut : `~/Desktop`).
- `VAULT_DIR` : chemin absolu du coffre Obsidian cible (alternative à `VAULT_NAME`).
- `VAULT_NAME` : nom du coffre Obsidian dans iCloud Drive (`~/Library/Mobile Documents/iCloud~md~obsidian/Documents/<vault>`).
- `VAULT_SUBPATH` : sous-dossier relatif à l’intérieur du coffre (ex. `Contacts`).
- `VAULT_MAPPING_JSON` : chemin vers le mapping JSON pour Obsidian (défaut : `csv_vault/mapping.example.json`).
- `DATA_GOUV_URL` : URL directe d’un CSV à télécharger (utilisé par « data.gouv.fr »).
- `SMTP_DOMAIN` : domaine SMTP pour les groupes dynamiques Exchange (optionnel).
- `EXCHANGE_ENABLE_REMOVAL` : `true|false` — masquer/supprimer les contacts absents du CSV courant (scope : ListName).
- `EXCHANGE_HARD_DELETE` : `true|false` — si `true`, supprime au lieu de masquer.
- `EXCHANGE_SHOW_TERMINAL` : `true|false` — ouvrir Terminal et afficher la commande pwsh.
- `SUBCHOICE` : pour automatiser le sous-choix sans GUI (`VCF`, `Obsidian`, `Exchange`).

Exemple minimal :
```
OUTPUT_DIR="$HOME/Desktop"
VAULT_NAME="Brice knowledge"
VAULT_SUBPATH="Contacts"
DATA_GOUV_URL="https://www.data.gouv.fr/…/fichier.csv"
SMTP_DOMAIN="exemple.org"
EXCHANGE_ENABLE_REMOVAL=false
EXCHANGE_HARD_DELETE=false
```

## Flux « Microsoft List »
Après sélection du CSV :

1) Générer VCF (Apple Contacts)
- Script : `csv_outlook/csv_contact_batch.py`
- Détection d’encodage/délimiteur, mapping de colonnes usuelles (Prénom, Nom, Organisation, Mails, Tels…),
  catégories vCard à partir de `Zone_Com`/`Département`, UID stable.
- Sortie : `nom_source_contacts.vcf` (copié dans `OUTPUT_DIR`).

2) Générer des notes Obsidian
- Script : `csv_vault/csv_to_obsidian_contacts.py`
- Mapping flexible via JSON : `csv_vault/mapping.example.json` (colonnes alias, champs liste, règles de merge, normalisation tel FR, champs requis).
- Options utiles :
  - `--vault-name <Nom>` + `--out-subpath <SousDossier>` (iCloud Obsidian) ou `--out <Chemin>`.
  - `--dry-run` pour prévisualiser.
  - `--id-prefix`/`--id-fallback` pour composer des identifiants lisibles.

3) Importer vers Exchange Online
- Script : `csv_exchange/sync_gal_by_id.ps1` (appelé via `pwsh`).
- Principe :
  - Clé d’unicité : `CustomAttribute3 = "<ListName>:<ID>"` (ListName dérivé du nom de fichier CSV).
  - Attributs marquant la source : `CustomAttribute1='Source:Lists'`, `CustomAttribute2='List:<ListName>'`.
  - Mise à jour/création des Mail Contacts ; visibilité GAL garantie ; notes enrichies avec champs « extras ».
  - Gestion des conflits d’adresse : réutilisation d’un `MailContact` existant si adéquat, sinon skip.
- Paramètres clés : `-CsvPath <fichier> [-Apply] [-EnableRemoval] [-HardDelete] [-SmtpDomain <domaine>]`.

## Flux « data.gouv.fr » → Exchange
1) Téléchargement du CSV depuis `DATA_GOUV_URL` vers `csv_fr_csv_list/input/`.
2) Normalisation : `csv_fr_csv_list/csv_fr_csv_list.py`
   - Nettoyage CP et téléphones FR, enrichissement « Niveau », tags `Zone_Com` (proximité/périphérique), filtrage géographique, etc.
   - Export `…__transformed.csv` (UTF‑8 BOM, séparateur `,`, guillemets sur toutes les cellules) dans `csv_fr_csv_list/output/`.
3) Copie du CSV transformé dans `OUTPUT_DIR` et import Exchange (cf. ci‑dessus).

## Détails utiles
- Environnement Python :
  - Créé/enrichi automatiquement au premier lancement ou via `bootstrap.zsh`.
  - `pandas` est requis pour les transformations `csv_fr_csv_list`.
- PowerShell/Exchange :
  - `orchestrateur.zsh` propose d’installer ce qui manque (Homebrew pour `pwsh`, module Exchange).
  - L’import s’exécute en `Terminal` par défaut pour visibilité, configurable via `EXCHANGE_SHOW_TERMINAL`.
- Automatisation :
  - `orchestrateur.zsh` accepte un chemin CSV en argument (ex. depuis Raccourci/Automator) et peut sauter le sous-menu via `SUBCHOICE`.

## Structure du dépôt
```
bootstrap.zsh
orchestrateur.zsh
config.env (optionnel)
csv_outlook/
  ├── csv_contact_batch.py
  └── csv_outlook_batch.py
csv_vault/
  ├── csv_to_obsidian_contacts.py
  ├── mapping.example.json
  └── README_CSV_TO_OBSIDIAN.md
csv_fr_csv_list/
  ├── csv_fr_csv_list.py
  └── requirements.txt
csv_exchange/
  └── sync_gal_by_id.ps1
```

## Dépannage
- `pandas introuvable` : exécuter `./bootstrap.zsh` ou `python3 -m pip install -r csv_fr_csv_list/requirements.txt`.
- `pwsh introuvable` : `brew install --cask powershell` (ou installer depuis Microsoft Docs).
- `ExchangeOnlineManagement` manquant : `Install-Module ExchangeOnlineManagement -Scope CurrentUser` dans `pwsh`.
- Encodage CSV : pour Obsidian, régler `csv_encoding` dans le mapping si besoin (`utf-8-sig`, `latin-1`, …).
- Aucune action lors de l’import Exchange : vérifier que le CSV contient des e‑mails valides et, pour le diff par ID, une colonne `ID`.

—
Mainteneur : Brice Sodini — Utilisation personnelle/équipe ; adaptez les mappings selon vos sources.

