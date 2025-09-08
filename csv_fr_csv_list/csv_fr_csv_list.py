#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CSV officiel Annuaire Éducation (FR) -> CSV "Microsoft Lists"-friendly
---------------------------------------------------------------------
- Lit tous les .csv du dossier ./input
- Normalise:
  * Code_postal -> 5 chiffres
  * Telephone -> 10 chiffres FR (0XXXXXXXXX) si possible
  * Nom_etablissement: ajoute " de {Commune}" seulement si
        Type_etablissement == "Ecole" ET
        Nom_etablissement ∈ {"ECOLE PUBLIQUE", "ECOLE ELEMENTAIRE PUBLIQUE"}
- Tagging Zone_Com (2 passes):
  1) Marque proximité si l'un: PIAL contient "0071156U" OU circonscription contient "Ardèche Nord" OU ZAP/bassin contient "Ardèche Verte" OU CP dans liste fournie -> "proximité"
  2) (sur la source) départements {Ardèche, Drôme, Haute-Loire, Loire, Rhône, Isère, Ain}
     ET pas PIAL 0071156U ET Type_etablissement != "Ecole" ET non déjà taggé -> "Secteur périphérique"
- Fusion "Niveau": Type_etablissement + drapeaux (ULIS, SEGPA, voies, sections, etc.) en "texte" séparé par ";"
- Mapping & export minimal:
    ID, Organisation, Niveau, OrgaType, Adresse_1, Code_postal, Commune,
    Tel_Fixe, Web, Mail_Pro, Département, Zone_Com
"""

import os
import re
import sys
import glob
import csv
import json
try:
    import pandas as pd
except Exception as e:
    print("[ERROR] Le module 'pandas' est introuvable.\n"
          "Installez les dépendances avec: 'python3 -m pip install -r requirements.txt'\n"
          "ou au minimum: 'python3 -m pip install pandas'\n"
          f"Détail: {e}", file=sys.stderr)
    sys.exit(1)

# --- Chemins (adapter si besoin)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# --- Constantes
SEPARATOR = ";"  # CSV FR / Microsoft Lists (par défaut historique)
# Options de format d'export pour l'import demandé
EXPORT_SEPARATOR = ","  # Format import: virgule
EXPORT_QUOTE_ALL = True   # Toujours entourer de guillemets
NIVEAU_AS_JSON_LIST = False  # Exporter 'Niveau' en texte: "Maternelles, Primaires"
ENCODING_READS = ["utf-8-sig", "utf-8", "cp1252"]
PIAL_CODE = "0071156U"
DEPS_PERIPH = {"Ardèche", "Drôme", "Haute-Loire", "Loire", "Rhône", "Isère", "Ain"}
ENABLE_GEO_FILTER = True  # Filtrer uniquement ces départements ciblés

# CP inclus en proximité (tous établissements, 1ère passe)
POSTAL_PROXIMITY_RAW = {
    "42220","26140","42410","38550","38150","26240","7300","42220","42660",
    "69420","42520","38550","38150","38150","26140","26140","26140","26240",
    "42410","42520","42520","42220","42220"
}

# --- Utilitaires
def read_csv_smart(path: str) -> pd.DataFrame:
    """Lecture robuste du CSV avec séparateur ';' et différents encodages possibles."""
    last_err = None
    for enc in ENCODING_READS:
        try:
            return pd.read_csv(path, sep=SEPARATOR, dtype=str, encoding=enc, keep_default_na=False).fillna("")
        except Exception as e:
            last_err = e
    raise last_err

def _strip_accents_lower(s: str) -> str:
    """Retourne une version normalisée: sans accents, en minuscules, espaces trim."""
    import unicodedata
    if s is None:
        return ""
    s = str(s).strip()
    if not s:
        return ""
    # Décomposition Unicode puis filtrage accents
    s_norm = unicodedata.normalize("NFKD", s)
    s_no_acc = "".join(ch for ch in s_norm if not unicodedata.combining(ch))
    return s_no_acc.lower()

def zfill_cp(cp: str) -> str:
    s = (cp or "").strip()
    if not s:
        return ""
    # garder seulement chiffres (au cas où il y aurait des espaces)
    s = re.sub(r"\D", "", s)
    return s.zfill(5)[:5] if s else ""

def normalize_phone_fr(s: str) -> str:
    """Retourne 10 chiffres FR ou '' si non conforme.
       Règles:
       - enlève tout sauf chiffres
       - si commence par 33 et longueur >=11 -> drop '33' puis s'assure du 0 en tête
       - si commence par 0 et longueur == 10 -> ok
    """
    if not s:
        return ""
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if digits.startswith("33") and len(digits) >= 11:
        digits = digits[2:]
        if not digits.startswith("0"):
            digits = "0" + digits
    if digits.startswith("0") and len(digits) == 10:
        return digits
    return ""  # sinon on préfère laisser vide

def normalize_web(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^(?:https?://)", u, flags=re.I):
        u = "https://" + u
    return u

def normalize_mail(m: str) -> str:
    m = (m or "").strip()
    if m and "@" in m and "." in m.split("@")[-1]:
        return m
    return ""

def enrich_nom_etablissement(nom: str, type_etab: str, commune: str) -> str:
    """Ajoute ' - {commune}' seulement pour cas EXACTS:
       - type_etab == 'Ecole'
       - nom dans une liste exacte (après normalisation, accents/casse ignorés):
         'Ecole publique', 'Ecole élémentaire publique',
         'Ecole primaire', 'Ecole primaire privée', 'Ecole élémentaire',
         'Ecole primaire Bourg', 'Ecole élémentaire Bourg'
    """
    if not nom or not type_etab:
        return nom

    # Normalisations
    from typing import Set
    name_key = re.sub(r"\s+", " ", _strip_accents_lower(nom))
    commune_str = (commune or "").strip()
    commune_key = re.sub(r"\s+", " ", _strip_accents_lower(commune_str))

    # Cas exacts autorisés (sans accents, en minuscules)
    allowed: Set[str] = {
        "ecole publique",
        "ecole elementaire publique",
        "ecole primaire",
        "ecole primaire privee",
        "ecole elementaire",
        "ecole primaire bourg",
        "ecole elementaire bourg",
    }

    if type_etab.strip().lower() == "ecole" and name_key in allowed:
        # Anti-doublon: ne pas ajouter si la commune est déjà présente (approx.)
        if commune_key and commune_key not in name_key:
            return f"{nom} de {commune_str}"
    return nom

def normalize_org_casing(s: str) -> str:
    """Normalise la casse/structure des noms d'organisation.
    Règles pragmatiques:
    - Trim + collapse espaces
    - Title-case souple sauf mots-outils (de, du, des, la, le, les, et, en, sur, sous, aux, au, a, pour, par, chez, dans)
    - Force certains adjectifs génériques en minuscule: publique, privee, elementaire, primaire, bourg
    - Conserve séparateurs (espaces, tirets, slash)
    """
    if not s:
        return s
    s = re.sub(r"\s+", " ", str(s).strip())
    small_words = {
        "de","du","des","la","le","les","l","d","et","en","sur","sous","aux","au","a","pour","par","chez","dans",
        "publique","privee","elementaire","primaire","bourg"
    }

    parts = re.split(r"([\s\-/]+)", s)
    out = []
    is_word = True
    word_index = 0
    for part in parts:
        if not part:
            continue
        if re.fullmatch(r"[\s\-/]+", part):
            out.append(part)
            is_word = True
            continue
        token = part.lower()
        if word_index > 0 and token in small_words:
            out.append(token)
        else:
            # Capitalise première lettre seulement
            out.append(token[:1].upper() + token[1:])
        word_index += 1
    return "".join(out)

def build_niveau(row: pd.Series) -> str:
    """Construit le champ 'Niveau' en concaténant des libellés distincts, séparés par ';'."""
    parts = []
    # Base : Type_etablissement
    te = row.get("Type_etablissement", "").strip()
    if te:
        parts.append(te)

    # Ajouts selon colonnes drapeaux == '1' ou non vide
    flags_map = {
        "Ecole_maternelle": "Maternelle",
        "Ecole_elementaire": "Élémentaire",
        "Voie_generale": "Voie générale",
        "Voie_technologique": "Voie technologique",
        "Voie_professionnelle": "Voie professionnelle",
        "ULIS": "ULIS",
        "Apprentissage": "Apprentissage",
        "Segpa": "SEGPA",
        "Section_arts": "Section arts",
        "Section_cinema": "Section cinéma",
        "Section_theatre": "Section théâtre",
        "Section_sport": "Section sport",
        "Section_internationale": "Section internationale",
        "Section_europeenne": "Section européenne",
        "Lycee_Agricole": "Lycée agricole",
        "Lycee_des_metiers": "Lycée des métiers",
        "Post_BAC": "Post-bac",
        "GRETA": "GRETA",
    }
    for col, label in flags_map.items():
        v = str(row.get(col, "")).strip()
        if v and v != "0":
            parts.append(label)

    # Appartenance Éducation Prioritaire (valeur telle quelle si non vide)
    aep = row.get("Appartenance_Education_Prioritaire", "").strip()
    if aep:
        parts.append(aep)

    # Dédup + ordre stable
    seen = set()
    deduped = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            deduped.append(p)
    return ";".join(deduped)

def build_niveau_canonical(row: pd.Series) -> str:
    """Construit une version canonique pour l'export JSON ("Maternelles", "Primaires", "Collège", "Lycée").
    - Ecole_maternelle -> "Maternelles"
    - Ecole_elementaire -> "Primaires"
    - Type_etablissement contient collège -> "Collège"
    - Type_etablissement contient lycée -> "Lycée"
    Si type = Ecole et aucun drapeau trouvé, fallback sur "Primaires".
    """
    parts = []
    type_raw = str(row.get("Type_etablissement", ""))
    type_norm = _strip_accents_lower(type_raw)

    def flag_truthy(val: str) -> bool:
        v = str(val or "").strip()
        return bool(v) and v != "0"

    if "ecole" in type_norm:
        if flag_truthy(row.get("Ecole_maternelle", "")):
            parts.append("Maternelles")
        if flag_truthy(row.get("Ecole_elementaire", "")):
            parts.append("Primaires")
        if not parts:  # fallback
            parts.append("Primaires")
    else:
        if "college" in type_norm:
            parts.append("Collège")
        if "lycee" in type_norm or "lycee" in _strip_accents_lower(type_raw.replace("é", "e")):
            # Le test précédent couvre déjà les accents normalisés
            parts.append("Lycée")

    # Dédup
    seen = set()
    canon = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            canon.append(p)
    return ";".join(canon)

def compute_zone_com(df: pd.DataFrame) -> pd.Series:
    """Calcule Zone_Com en deux passes, sans écraser le tag PIAL."""
    zone = pd.Series([""] * len(df), index=df.index)

    # Pré-normalisations robustes (accents/casse/espaces)
    dep_norm = df.get("Libelle_departement", "").map(_strip_accents_lower)
    type_norm = df.get("Type_etablissement", "").map(_strip_accents_lower)
    pial_series = df.get("PIAL", "").astype(str)
    circo_norm = df.get("nom_circonscription", "").map(_strip_accents_lower)
    zap_norm = df.get("libelle_zone_animation_pedagogique", "").map(_strip_accents_lower)
    bassin_norm = df.get("libelle_bassin_formation", "").map(_strip_accents_lower)
    cp_norm = df.get("Code_postal", "").astype(str).map(zfill_cp)

    # Set normalisé pour comparaison départements
    deps_periph_norm = {_strip_accents_lower(x) for x in DEPS_PERIPH}

    # Pass 1 : Proximité par marqueurs source (au-delà des écoles)
    # - PIAL contient le code du PIAL Nord Ardèche
    # - OU nom_circonscription contient "ardèche nord"
    # - OU libelle_zone_animation_pedagogique contient "ardèche verte"
    # - OU CP dans la liste fournie
    mask_pial = pial_series.str.contains(PIAL_CODE, na=False)
    mask_circo = circo_norm.str_contains("ardeche nord", na=False) if hasattr(circo_norm, 'str_contains') else circo_norm.str.contains("ardeche nord", na=False)
    mask_zap = zap_norm.str_contains("ardeche verte", na=False) if hasattr(zap_norm, 'str_contains') else zap_norm.str.contains("ardeche verte", na=False)
    mask_bassin = bassin_norm.str_contains("ardeche verte", na=False) if hasattr(bassin_norm, 'str_contains') else bassin_norm.str.contains("ardeche verte", na=False)
    postal_prox_norm = {zfill_cp(x) for x in POSTAL_PROXIMITY_RAW}
    mask_cp = cp_norm.isin(postal_prox_norm)
    mask_proximite = mask_pial | mask_circo | mask_zap | mask_bassin | mask_cp
    zone.loc[mask_proximite] = "proximité"

    # Pass 2 : périphérique (sur la source, sans écraser PIAL)
    # Ne taguer en périphérique que les lignes non encore affectées
    unassigned = zone.eq("")
    mask_periph = (
        unassigned &
        dep_norm.isin(deps_periph_norm) &
        ~pial_series.str.contains(PIAL_CODE, na=False) &
        (type_norm != "ecole")
    )
    zone.loc[mask_periph] = "Secteur périphérique"

    return zone

def transform_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Sécurité: s'assurer que les colonnes existent
    needed = [
        "Nom_etablissement","Type_etablissement","Statut_public_prive","Adresse_1",
        "Code_postal","Nom_commune","Telephone","Web","Mail","Libelle_departement","PIAL",
        "Ecole_maternelle","Ecole_elementaire","Voie_generale","Voie_technologique",
        "Voie_professionnelle","ULIS","Apprentissage","Segpa","Section_arts","Section_cinema",
        "Section_theatre","Section_sport","Section_internationale","Section_europeenne",
        "Lycee_Agricole","Lycee_des_metiers","Post_BAC","Appartenance_Education_Prioritaire","GRETA"
    ]
    for col in needed:
        if col not in df.columns:
            df[col] = ""

    # Normalisations
    df["Code_postal"] = df["Code_postal"].map(zfill_cp)
    df["Tel_Fixe"] = df["Telephone"].map(normalize_phone_fr)
    df["Web_norm"] = df["Web"].map(normalize_web)
    df["Mail_norm"] = df["Mail"].map(normalize_mail)

    # Enrichissement nom d'établissement (règle précise)
    df["Organisation"] = [
        enrich_nom_etablissement(nom, te, com)
        for nom, te, com in zip(df["Nom_etablissement"], df["Type_etablissement"], df["Nom_commune"])
    ]

    # Normalisations additionnelles avant export
    # - Code postal strict 5 chiffres déjà appliqué (df["Code_postal"]) via zfill_cp
    # - Téléphone normalisé à 10 chiffres (df["Tel_Fixe"]) via normalize_phone_fr
    # - Casse/structure 'Organisation'
    df["Organisation"] = df["Organisation"].map(normalize_org_casing)

    # Zone_Com (2 passes)
    df["Zone_Com"] = compute_zone_com(df)

    # Niveau (fusion)
    df["Niveau_out"] = df.apply(build_niveau, axis=1)
    # Niveau canonique pour export JSON
    df["Niveau_canon"] = df.apply(build_niveau_canonical, axis=1)

    # Colonnes finales (mapping)
    out = pd.DataFrame({
        "Organisation": df["Organisation"],
        # Utiliser la version canonique pour l'export JSON
        "Niveau": df["Niveau_canon"],
        "OrgaType": df["Statut_public_prive"],
        "Adresse_1": df["Adresse_1"],
        "Code_postal": df["Code_postal"],
        "Commune": df["Nom_commune"],
        "Tel_Fixe": df["Tel_Fixe"],
        "Web": df["Web_norm"],
        "Mail_Pro": df["Mail_norm"],
        "Département": df["Libelle_departement"],
        "Zone_Com": df["Zone_Com"],
    })

    # Optionnel: garde aussi PIAL brut pour vérifs
    # out["PIAL"] = df["PIAL"]

    return out

def process_file(path_in: str, out_dir: str) -> str:
    df = read_csv_smart(path_in)

    # Optionnel: filtrage géographique par département
    if ENABLE_GEO_FILTER:
        # s'assurer de la colonne
        if "Libelle_departement" not in df.columns:
            df["Libelle_departement"] = ""
        dep_norm = df["Libelle_departement"].map(_strip_accents_lower)
        deps_periph_norm = {_strip_accents_lower(x) for x in DEPS_PERIPH}
        before = len(df)
        df = df[dep_norm.isin(deps_periph_norm)].copy()
        after = len(df)
        print(f"[INFO] Filtre départements actif: {after}/{before} lignes conservées")

    # Transform
    out_df = transform_dataframe(df)

    # Supprimer les lignes sans email (Mail_Pro vide)
    try:
        before_mail = len(out_df)
        out_df = out_df[out_df["Mail_Pro"].astype(str).str.strip() != ""].copy()
        after_mail = len(out_df)
        if after_mail != before_mail:
            print(f"[INFO] Lignes sans email supprimées: {before_mail - after_mail}")
    except Exception:
        pass

    # Petit récap des tags Zone_Com pour contrôle
    try:
        vc = out_df["Zone_Com"].value_counts(dropna=False)
        info = ", ".join([f"{idx if idx else 'vide'}={cnt}" for idx, cnt in vc.items()])
        print(f"[INFO] Zone_Com: {info}")
    except Exception:
        pass

    # Supprimer les lignes sans Zone_Com (vide)
    try:
        before_rows = len(out_df)
        out_df = out_df[out_df["Zone_Com"].astype(str).str.strip() != ""].copy()
        after_rows = len(out_df)
        if after_rows != before_rows:
            print(f"[INFO] Lignes sans Zone_Com supprimées: {before_rows - after_rows}")
    except Exception:
        pass

    # Ajout de la colonne ID (SCOLFR-<n>) en tête, après filtrages
    try:
        out_df = out_df.reset_index(drop=True)
        out_df.insert(0, "ID", [f"SCOLFR-{i}" for i in range(1, len(out_df) + 1)])
    except Exception as e:
        print(f"[WARN] Impossible d'ajouter la colonne ID: {e}")

    # Nom de sortie
    base = os.path.splitext(os.path.basename(path_in))[0]
    path_out = os.path.join(out_dir, f"{base}__transformed.csv")

    # Mise en forme finale des valeurs avant export
    if NIVEAU_AS_JSON_LIST:
        out_df["Niveau"] = [
            json.dumps([p for p in str(v).split(";") if p], ensure_ascii=False) if str(v).strip() else json.dumps([], ensure_ascii=False)
            for v in out_df["Niveau"]
        ]
    else:
        out_df["Niveau"] = [
            ", ".join([p for p in str(v).split(";") if p]) if str(v).strip() else ""
            for v in out_df["Niveau"]
        ]

    # Export
    # UTF-8 avec BOM pour une compatibilité Excel accrue (accents)
    quoting_mode = csv.QUOTE_ALL if EXPORT_QUOTE_ALL else csv.QUOTE_MINIMAL
    out_df.to_csv(path_out, index=False, sep=EXPORT_SEPARATOR, encoding="utf-8-sig", quoting=quoting_mode)

    return path_out

def main():
    os.makedirs(INPUT_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    files = sorted(glob.glob(os.path.join(INPUT_DIR, "*.csv")))
    if not files:
        print(f"[INFO] Aucun CSV trouvé dans {INPUT_DIR}")
        return 0

    print(f"[INFO] Fichiers détectés ({len(files)}):")
    for f in files:
        print("  -", os.path.basename(f))

    ok, ko = 0, 0
    for f in files:
        try:
            outp = process_file(f, OUTPUT_DIR)
            print(f"[OK]  {os.path.basename(f)} -> {os.path.basename(outp)}")
            ok += 1
        except Exception as e:
            print(f"[ERR] {os.path.basename(f)} : {e}", file=sys.stderr)
            ko += 1

    print(f"\n[SUMMARY] OK={ok} | ERR={ko}")
    return 0 if ko == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
