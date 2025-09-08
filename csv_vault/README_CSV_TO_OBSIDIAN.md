# CSV → Obsidian Contacts (flexible mapping)

Ce kit convertit un CSV (export Microsoft Lists ou autre) en **notes de contacts** Obsidian, en s’adaptant à des **en-têtes variables** grâce à un fichier de **mapping**.

## Fichiers
- `csv_to_obsidian_contacts.py` — script principal
- `mapping.example.json` — configuration de mapping (à copier/adapter)
- Ce README

## Installation
1. Installe Python 3 (macOS en a généralement un).
2. Place tes fichiers CSV et le script au même endroit ou n'importe où.

## Utilisation (exemples)
```
python csv_to_obsidian_contacts.py   --csv "/chemin/contacts.csv"   --out "/chemin/Obsidian Vault/Contacts"   --config "/chemin/mapping.json"   --project ""   --dry-run
```
- `--dry-run` : affiche ce qui serait créé/mis à jour **sans écrire**.
- Quand OK, relance **sans** `--dry-run` pour générer les `.md`.

### Portabilité macOS (iCloud Obsidian)
Sur Mac, si ton coffre Obsidian est synchronisé via iCloud, tu peux éviter les chemins absolus différents d’une machine à l’autre en utilisant `--vault-name`:
```
python csv_to_obsidian_contacts.py \
  --csv "~/Desktop/contacts.csv" \
  --vault-name "Brice knowledge" \
  --out-subpath "Contacts" \
  --config "mapping.json" \
  --dry-run
```
- `--vault-name` vise `~/Library/Mobile Documents/iCloud~md~obsidian/Documents/<vault>` automatiquement.
- `--out` devient optionnel si tu fournis `--vault-name` (ou `--vault` + `--out-subpath`).
- `--config` par défaut cherche `mapping.json` dans le dossier courant.
- Le script vérifie l’existence du coffre iCloud et affiche une erreur claire s’il est introuvable.

### ID composé avec `--id-prefix`
- Si tu fournis `--id-prefix`, le script compose un identifiant lisible (ex: `SCO-001`).
- Si la colonne `id_mslist` (ou l’ID logique choisi) est **numérique**, on obtient `PREFIX-<id>` avec zero‑padding.
- Si l’ID est vide ou non numérique, la stratégie se règle via `--id-fallback`:
  - `seq` (défaut): **compteur séquentiel par exécution** → `PREFIX-001`, `PREFIX-002`, ...
  - `hash`: hash stable (SHA‑1, 8 hex) des données de la ligne → `PREFIX-abcdef12`
  - `raw`: garde la valeur brute si présente → `PREFIX-<valeur>` (sinon rien)
  - `skip`: ignore la ligne si l’ID n’est pas numérique
  - Exemple: `--id-prefix "SCO" --id-pad 3 --id-fallback hash`

### Champs requis (skip des lignes vides)
- Tu peux définir les champs considérés comme « essentiels » dans `mapping.json`:
```
{
  "required_fields": ["Nom", "Prénom", "organisation"]
}
```
- Si aucun de ces champs n’est rempli pour une ligne, elle est marquée `[SKIPPED]` (et le détail est affiché en dry‑run).

### Dry‑run et logging
- Dry‑run affiche maintenant un e‑mail pertinent (détection `email`/`Mail_Pro`/`Mail_Perso`).
- Active des logs détaillés avec `--log-level DEBUG` pour voir les décisions de merge et la composition des IDs.

## Configuration (mapping.json)
- **aliases** : liste les **noms possibles** d’en-têtes CSV pour chaque **clé logique** (`nom`, `email`, etc.).
- **projects** : profils d’override. Tu peux lancer avec `--project projetX` pour appliquer des alias spécifiques.
- **extras** : `"all"` pour inclure **toutes** les autres colonnes CSV comme champs YAML additionnels (normalisées). Ou liste de colonnes à inclure.
- **transforms.normalize_phone_fr** : normalise les numéros FR en **E.164** (`+33...`).

## Mise à jour vs création
- Le script **met à jour** une note existante si elle a le **même `id_mslist`** OU la **même adresse `email`**.
- Sinon, il **crée** un nouveau fichier `Contacts/<nom>.md`.
- Certains champs perso (`notes`, `custom_tags`) sont **préservés** (configurable via `preserve_keys`).

## YAML produit (exemple)
```yaml
---
id_mslist: 123
nom: Jean Dupont
email: [email protected]
telephone: +33612345678
organisation: ACME
groupes: [Clients, VIP]
type: contact
source_updated: 2025-09-01T12:00:00Z
---
```

## Astuces
- Mets un **ID** dans ton CSV pour fiabiliser les mises à jour (colonne `ID`, `Id`, etc.).
- Les champs de liste (comme `groupes`) sont séparés sur `; , | • ·`.
- Pour forcer un mapping, ajoute l'en-tête voulu dans `aliases` (ou dans un **profil projet**).
- Les champs supplémentaires du CSV sont **ajoutés** en YAML (noms normalisés).

## Intégration Obsidian
- Place les notes générées dans ton dossier `Contacts/` de ton coffre.
- Active **Bases** (Obsidian 1.9) et crée une vue filtrée sur `type: contact`.
- Tes templates d’email peuvent référencer les contacts via `[[Nom du fichier]]` et récupérer `email`.
- Délimiteur CSV: si `csv_delimiter` vaut `"auto"` (recommandé), le script détecte automatiquement `, ; \t |`.
