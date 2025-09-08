#!/usr/bin/env python3
import csv
import sys
from pathlib import Path
import unicodedata
import uuid
from datetime import datetime, timezone

# We now output vCard (.vcf) instead of CSV for Apple Contacts.

# Normalize strings (for robust header match if needed)
def norm(s: str) -> str:
    """Light normalization keeping accents (for filtering/labels).
    - Trim spaces
    - Normalize to NFC (preserve diacritics)
    """
    if s is None:
        return ""
    return unicodedata.normalize("NFC", str(s)).strip()

# Fixed mapping for Brice's Microsoft Lists export (Contacts Scolaires)
# Source header -> Apple Contacts column (or special handling)
MAP_DIRECT = {
    "Prénom": "First name",
    "Nom": "Last name",
    "Organisation": "Company",
    "Fonction": "Job Title",
    "Tel_Mobile": "Mobile Phone",
    "Tel_Fixe": "Work Phone",
    "Commune": "Work City",
    "Département": "Work State",
    "Adresse_1": "Work Street",
    "Code_postal": "Work ZIP",
}

# Default country for work address if any work-address component exists
DEFAULT_WORK_COUNTRY = "France"

# Fields that go to Notes verbatim (if present)
# Do not include ID in notes per request
FIELDS_TO_NOTES = ["OrgaType", "Niveau", "Classes", "Projets", "Zone_Com"]

def detect_delimiter(sample: str) -> str:
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",",";","|","\t"])
        return dialect.delimiter
    except Exception:
        # heuristic fallback
        counts = {d: sample.count(d) for d in [",",";","|","\t"]}
        return max(counts, key=counts.get)

def read_csv_text(p: Path):
    for enc in ("utf-8-sig","utf-8","cp1252"):
        try:
            return p.read_text(encoding=enc), enc
        except UnicodeDecodeError:
            continue
    return p.read_bytes().decode("utf-8", errors="replace"), "utf-8 (errors=replace)"

def build_contacts_row(src: dict) -> dict:
    # We keep a normalized mapping dict for convenience, then emit a vCard from it.
    out = {
        "First name": "",
        "Last name": "",
        "Company": "",
        "Job Title": "",
        "Department": "",
        "Work Email": "",
        "Home Email": "",
        "Work Phone": "",
        "Mobile Phone": "",
        "Home Phone": "",
        "Work Street": "",
        "Work City": "",
        "Work State": "",
        "Work ZIP": "",
        "Work Country": "",
        "Home Street": "",
        "Home City": "",
        "Home State": "",
        "Home ZIP": "",
        "Home Country": "",
        "Note": "",
    }

    # Direct mappings
    for src_key, dst_key in MAP_DIRECT.items():
        if src_key in src:
            val = norm(src.get(src_key, ""))
            if val:
                out[dst_key] = val

    # Emails: prefer mapping each to its label
    work_email = norm(src.get("Mail_Pro", "")) if "Mail_Pro" in src else ""
    home_email = norm(src.get("Mail_Perso", "")) if "Mail_Perso" in src else ""
    if work_email:
        out["Work Email"] = work_email
    if home_email:
        out["Home Email"] = home_email
    # If only one email exists and the other is empty, leave as-is; Contacts import lets you map.

    # No explicit Full Name column in Apple CSV; Contacts derives it.

    # Default country for Work address when we have partial address info
    if not out.get("Work Country") and any(out.get(k) for k in ["Work Street", "Work City", "Work State", "Work ZIP"]):
        out["Work Country"] = DEFAULT_WORK_COUNTRY

    # Department field on the contact: prefer Départements, fallback to Zone_Com
    dep_val = norm(src.get("Département", "")) if "Département" in src else ""
    zc = norm(src.get("Zone_Com", "")) if "Zone_Com" in src else ""
    if dep_val:
        out["Department"] = dep_val
    elif zc:
        out["Department"] = zc

    # Notes aggregation
    notes_parts = []
    for key in FIELDS_TO_NOTES:
        if key in src:
            val = norm(src.get(key, ""))
            if val:
                notes_parts.append(f"{key}: {val}")
    if notes_parts:
        out["Note"] = "\n".join(notes_parts)

    return out

def vescape(s: str) -> str:
    if s is None:
        return ""
    # vCard escaping per RFC: backslash, comma, semicolon, and newline
    return (
        str(s)
        .replace("\\", "\\\\")
        .replace(";", "\\;")
        .replace(",", "\\,")
        .replace("\n", "\\n")
    )

def fold(line: str, limit: int = 75) -> str:
    # Simple line folding: split long lines with CRLF + space
    if len(line) <= limit:
        return line
    parts = []
    s = line
    while len(s) > limit:
        parts.append(s[:limit])
        s = s[limit:]
    parts.append(s)
    return "\r\n ".join(parts)

def stable_uid(src: dict, mapped: dict) -> str:
    # Prefer a stable source ID; else primary email; else name+company
    key = ""
    if src.get("ID"):
        key = f"mslist:{norm(src.get('ID'))}"
    elif mapped.get("Work Email") or mapped.get("Home Email"):
        key = f"email:{mapped.get('Work Email') or mapped.get('Home Email')}"
    else:
        key = f"name:{mapped.get('First name','')} {mapped.get('Last name','')}|{mapped.get('Company','')}"
    u = uuid.uuid5(uuid.NAMESPACE_URL, f"csv-contacts/{key}")
    return str(u)

def make_vcard(src: dict) -> str:
    m = build_contacts_row(src)
    fn = (m.get("First name", "") + " " + m.get("Last name", "")).strip()
    if not fn:
        fn = m.get("Company", "") or ""
    rev = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    uid = stable_uid(src, m)

    lines = [
        "BEGIN:VCARD",
        "VERSION:3.0",
        f"N:{vescape(m.get('Last name',''))};{vescape(m.get('First name',''))};;;",
        f"FN:{vescape(fn)}",
    ]

    if m.get("Company") or m.get("Department"):
        lines.append(f"ORG:{vescape(m.get('Company',''))};{vescape(m.get('Department',''))}")
    if m.get("Job Title"):
        lines.append(f"TITLE:{vescape(m.get('Job Title'))}")

    if m.get("Work Email"):
        lines.append(f"EMAIL;TYPE=INTERNET,WORK,PREF:{vescape(m.get('Work Email'))}")
    if m.get("Home Email"):
        lines.append(f"EMAIL;TYPE=INTERNET,HOME:{vescape(m.get('Home Email'))}")

    if m.get("Work Phone"):
        lines.append(f"TEL;TYPE=VOICE,WORK:{vescape(m.get('Work Phone'))}")
    if m.get("Mobile Phone"):
        lines.append(f"TEL;TYPE=VOICE,CELL,PREF:{vescape(m.get('Mobile Phone'))}")
    if m.get("Home Phone"):
        lines.append(f"TEL;TYPE=VOICE,HOME:{vescape(m.get('Home Phone'))}")

    if any(m.get(k) for k in ["Work Street","Work City","Work State","Work ZIP","Work Country"]):
        lines.append(
            "ADR;TYPE=WORK:;;" +
            f"{vescape(m.get('Work Street',''))};{vescape(m.get('Work City',''))};{vescape(m.get('Work State',''))};{vescape(m.get('Work ZIP',''))};{vescape(m.get('Work Country',''))}"
        )
    if any(m.get(k) for k in ["Home Street","Home City","Home State","Home ZIP","Home Country"]):
        lines.append(
            "ADR;TYPE=HOME:;;" +
            f"{vescape(m.get('Home Street',''))};{vescape(m.get('Home City',''))};{vescape(m.get('Home State',''))};{vescape(m.get('Home ZIP',''))};{vescape(m.get('Home Country',''))}"
        )

    if m.get("Note"):
        # Replace newlines already escaped with \n in vescape
        lines.append(f"NOTE:{vescape(m.get('Note'))}")

    # Categories from Zone_Com and Département for easier filtering
    categories = []
    zc_src = norm(src.get("Zone_Com", "")) if "Zone_Com" in src else ""
    dep_src = norm(src.get("Département", "")) if "Département" in src else ""
    for val in (zc_src, dep_src):
        if val and val not in categories:
            categories.append(val)
    if categories:
        lines.append(f"CATEGORIES:{vescape(','.join(categories))}")

    # Stable identifiers to allow updates on re-import
    lines.append(f"UID:{uid}")
    lines.append(f"X-ABUID:{uid}:ABPerson")
    lines.append(f"REV:{rev}")
    lines.append("END:VCARD")

    # Fold long lines
    return "\r\n".join(fold(l) for l in lines) + "\r\n"

def convert_file(in_path: Path, out_path: Path):
    text, enc = read_csv_text(in_path)
    delim = detect_delimiter(text[:5000])
    reader = csv.DictReader(text.splitlines(), delimiter=delim)
    cards = [make_vcard(row) for row in reader]

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        f.writelines(cards)
    print(f"✔ {in_path} → {out_path} | {len(cards)} contacts | enc={enc} sep={delim}")

def main():
    args = sys.argv[1:]

    # If paths are provided, process those files/folders and write outputs
    # next to the source file(s).
    if args:
        for raw in args:
            p = Path(raw).expanduser().resolve()
            if not p.exists():
                print(f"⚠ Chemin introuvable: {p}")
                continue

            if p.is_file() and p.suffix.lower() == ".csv":
                out_p = p.with_name(f"{p.stem}_contacts.vcf")
                convert_file(p, out_p)
            elif p.is_dir():
                csv_files = sorted(p.glob("*.csv"))
                if not csv_files:
                    print(f"(aucun .csv dans {p})")
                for f in csv_files:
                    out_p = f.with_name(f"{f.stem}_contacts.vcf")
                    convert_file(f, out_p)
            else:
                print(f"(ignoré: {p} — pas un .csv)")
        return

    # Fallback: legacy behaviour using ./input → ./output
    base = Path(__file__).resolve().parent
    input_dir = base / "input"
    output_dir = base / "output"
    output_dir.mkdir(exist_ok=True)

    csv_files = list(input_dir.glob("*.csv"))
    if not csv_files:
        print("Aucun CSV en argument et aucun dans input/. Ajoute un fichier ou passe un chemin au script.")
        return

    for p in csv_files:
        out_p = output_dir / f"{p.stem}_contacts.vcf"
        convert_file(p, out_p)

if __name__ == "__main__":
    main()
