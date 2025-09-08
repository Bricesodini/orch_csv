#!/usr/bin/env python3
import csv
import sys
from pathlib import Path
import unicodedata

# Outlook target columns
OUTLOOK_COLUMNS = [
    "First Name",
    "Last Name",
    "Full Name",
    "E-mail Address",
    "Company",
    "Job Title",
    "Department",
    "Business Phone",
    "Mobile Phone",
    "Home Phone",
    "Business Street",
    "Business City",
    "Business State",
    "Business Postal Code",
    "Business Country/Region",
    "Home Street",
    "Home City",
    "Home State",
    "Home Postal Code",
    "Home Country/Region",
    "Notes",
]

# Normalize strings (for robust header match if needed)
def norm(s: str) -> str:
    if s is None:
        return ""
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s.strip()

# Fixed mapping for Brice's Microsoft Lists export (Contacts Scolaires)
# Source header -> Outlook column (or special handling)
MAP_DIRECT = {
    "Prénom": "First Name",
    "Nom": "Last Name",
    "Organisation": "Company",
    "Fonction": "Job Title",
    "Tel_Mobile": "Mobile Phone",
    "Tel_Fixe": "Business Phone",
    "Commune": "Business City",
    "Département": "Business State",
    # "Department" (Outlook) intentionally left blank because "Département" is mapped to Business State.
}

# Fields that go to Notes verbatim (if present)
FIELDS_TO_NOTES = ["ID", "OrgaType", "Niveau", "Classes", "Projets"]

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

def build_outlook_row(src: dict) -> dict:
    out = {c: "" for c in OUTLOOK_COLUMNS}

    # Direct mappings
    for src_key, dst_key in MAP_DIRECT.items():
        if src_key in src:
            val = norm(src.get(src_key, ""))
            if val:
                out[dst_key] = val

    # Email preference: Mail_Pro > Mail_Perso
    email = ""
    if "Mail_Pro" in src and norm(src.get("Mail_Pro","")):
        email = norm(src.get("Mail_Pro",""))
    elif "Mail_Perso" in src and norm(src.get("Mail_Perso","")):
        email = norm(src.get("Mail_Perso",""))
    if email:
        # Outlook expects the header name with hyphen: "E-mail Address"
        out["E-mail Address"] = email

    # Full Name if missing
    if not out["Full Name"]:
        fn, ln = out["First Name"].strip(), out["Last Name"].strip()
        if fn or ln:
            out["Full Name"] = (fn + " " + ln).strip()

    # Notes aggregation
    notes_parts = []
    for key in FIELDS_TO_NOTES:
        if key in src:
            val = norm(src.get(key, ""))
            if val:
                notes_parts.append(f"{key}: {val}")
    if notes_parts:
        out["Notes"] = "\n".join(notes_parts)

    return out

def convert_file(in_path: Path, out_path: Path):
    text, enc = read_csv_text(in_path)
    delim = detect_delimiter(text[:5000])
    reader = csv.DictReader(text.splitlines(), delimiter=delim)
    rows = [build_outlook_row(row) for row in reader]

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=OUTLOOK_COLUMNS, delimiter=",")
        writer.writeheader()
        writer.writerows(rows)
    print(f"✔ {in_path} → {out_path} | {len(rows)} lignes | enc={enc} sep={delim}")

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
                out_p = p.with_name(f"{p.stem}_outlook.csv")
                convert_file(p, out_p)
            elif p.is_dir():
                csv_files = sorted(p.glob("*.csv"))
                if not csv_files:
                    print(f"(aucun .csv dans {p})")
                for f in csv_files:
                    out_p = f.with_name(f"{f.stem}_outlook.csv")
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
        out_p = output_dir / f"{p.stem}_outlook.csv"
        convert_file(p, out_p)

if __name__ == "__main__":
    main()
