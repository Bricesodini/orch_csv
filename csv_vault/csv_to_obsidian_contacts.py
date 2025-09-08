#!/usr/bin/env python3
"""
CSV -> Obsidian Contacts (flexible mapping)
-------------------------------------------
Reads a CSV exported from Microsoft Lists (or any source) and creates/updates
one Markdown note per contact in an Obsidian vault folder.

Key features:
- Robust, *configurable* mapping: rename/alias columns without changing code.
- Auto-detects headers: accepts varying CSV column names between projects.
- Merges updates: preserves your custom fields in existing notes.
- Handles "groups" list fields, phone normalization, and project-specific extras.
- Dry-run mode to preview without writing files.

Usage:
  python csv_to_obsidian_contacts.py \
      --csv /Users/bricesodini/Desktop/Scolaires.csv \
      --out "/Users/bricesodini/Library/Mobile Documents/iCloud~md~obsidian/Documents/Brice knowledge" \
      --config mapping.json \
      [--project PROJECT_NAME] \
      [--id-key id_mslist] \
      [--dry-run]

See README_CSV_TO_OBSIDIAN.md for details and examples.
"""
import argparse
import csv
import re
import sys
import unicodedata
from pathlib import Path
from datetime import datetime
import json
import logging
from functools import lru_cache
import hashlib

# ----------------------- Logging -----------------------
logger = logging.getLogger("csv_to_obsidian")

def setup_logger(level: str = "INFO"):
    lvl = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(level=lvl, format='[%(levelname)s] %(message)s')
    logger.setLevel(lvl)

# ----------------------- YAML helpers & parsing -----------------------
SAFE_RE = re.compile(r'^[A-Za-z0-9_-]+$')
EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")

# Defaults for merge behavior
DEFAULT_PRESERVE_KEYS = ['notes', 'custom_tags']
DEFAULT_OVERWRITE_KEYS = {
    'Nom','Prénom','Tel_Mobile','Tel_Fixe','Mail_Pro','Mail_Perso',
    'organisation','groupes','type','id_mslist','source_updated'
}

def parse_maybe_list(val):
    """If value looks like a JSON list (e.g. ' ["A", "B"] '), return list, else None."""
    if isinstance(val, list):
        return val
    if not isinstance(val, str):
        return None
    s = val.strip()
    if s.startswith('[') and s.endswith(']'):
        try:
            parsed = json.loads(s)
            if isinstance(parsed, list):
                return parsed
        except Exception:
            return None
    return None

def quote_scalar(v):
    """Return a YAML-safe scalar; quote strings with spaces/accents or digit-only values."""
    if isinstance(v, bool):
        return 'true' if v else 'false'
    if v is None or v == '':
        return '""'
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v)
    # Quote strings that are all digits or begin with '+' or '0' to preserve phone numbers/leading zeros
    if s.isdigit() or s.startswith('+') or (len(s) > 1 and s[0] == '0'):
        return f'"{s}"'
    if SAFE_RE.match(s):
        return s
    s2 = s.replace('"', '\\"')
    return f'"{s2}"'

# Normalize a field name for matching (ignore accents/case/punct)
@lru_cache(maxsize=4096)
def _norm_key_for_match(name: str) -> str:
    s = unicodedata.normalize('NFKD', str(name)).encode('ascii','ignore').decode('ascii').lower()
    return re.sub(r'[^a-z0-9]', '', s)

# Simple singularizer to treat 'emails' == 'email', 'mails' == 'mail', etc.
def _singularize_norm(s: str) -> str:
    return s[:-1] if s.endswith('s') else s

# Build a value from multiple possible field names inside the already-built contact dict
# (accent/case insensitive). Returns first non-empty string.
def get_contact_value(contact: dict, aliases):
    if not aliases:
        return ''
    # map normalized key -> original key present in contact
    norm_map = {}
    for k in contact.keys():
        normalized_key = _norm_key_for_match(k)
        norm_map[normalized_key] = k
    for a in aliases:
        alias_norm = _norm_key_for_match(a)
        if alias_norm in norm_map:
            v = contact.get(norm_map[alias_norm])
            if isinstance(v, str) and v.strip():
                return v.strip()
    return ''

# Sanitize a human-readable filename (keep accents/case/spaces, remove forbidden chars)
def sanitize_filename(title: str) -> str:
    s = str(title or '').strip()
    # forbid characters on macOS/Windows: / \ : * ? " < > |
    s = re.sub(r'[\\/:*?"<>|]+', ' ', s)
    # collapse spaces
    s = re.sub(r'\s+', ' ', s).strip()
    # avoid empty
    return s if s else 'contact'

def is_phone_field(name: str) -> bool:
    normalized_name = _norm_key_for_match(name)
    return any(k in normalized_name for k in ['phone','tel','telephone','mobile','portable','fixe'])

def is_bool_field(name: str, cfg) -> bool:
    normalized_name = _norm_key_for_match(name)
    # from config explicit list
    for b in (cfg.get('boolean_fields') or []):
        if _norm_key_for_match(b) == normalized_name:
            return True
    # heuristic
    return (
        normalized_name.startswith('is') or normalized_name.startswith('has') or
        'actif' in normalized_name or 'active' in normalized_name or 'enabled' in normalized_name or 'inscrit' in normalized_name
    )

def coerce_bool_if_needed(name: str, val, cfg):
    if isinstance(val, bool):
        return val
    if not isinstance(val, str):
        return val
    if not is_bool_field(name, cfg):
        return val
    s = val.strip().lower()
    if s in ('true','vrai','yes','oui','1'):
        return True
    if s in ('false','faux','no','non','0'):
        return False
    return val

# ----------------------- Utilities -----------------------

def slugify(value: str, allow_unicode=False):
    value = str(value)
    if allow_unicode:
        value = unicodedata.normalize('NFKC', value)
    else:
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
    value = re.sub(r'[^\w\s-]', '', value).strip().lower()
    return re.sub(r'[-\s]+', '-', value)

def split_list(val, seps=None):
    if val is None:
        return []
    if isinstance(val, list):
        return [v for v in (x.strip() for x in val) if v]
    s = str(val)
    if seps is None:
        seps = [';', ',', '|', '•', '·']
    for sp in seps:
        s = s.replace(sp, ',')
    parts = [p.strip() for p in s.split(',')]
    return [p for p in parts if p]

def read_frontmatter(md_text: str):
    if not md_text.startswith('---'):
        return {}, md_text
    parts = md_text.split('---', 2)
    if len(parts) < 3:
        return {}, md_text
    _, yaml_block, body = parts

    data = {}
    lines = yaml_block.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        if not line.strip():
            i += 1
            continue
        # Handle "key: value" or "key:" (for block lists)
        if ':' in line and not line.lstrip().startswith('- '):
            key, val = line.split(':', 1)
            key = key.strip()
            val = val.strip()
            # Block list: key: on its own, followed by indented "- item" lines
            if val == '':
                j = i + 1
                items = []
                while j < len(lines):
                    nxt = lines[j]
                    if nxt.startswith('  - '):
                        items.append(nxt[4:].strip().strip('"').strip("'"))
                        j += 1
                    elif nxt.strip().startswith('- '):  # tolerate missing indentation
                        items.append(nxt.strip()[2:].strip().strip('"').strip("'"))
                        j += 1
                    else:
                        break
                if items:
                    data[key] = items
                    i = j
                    continue
                else:
                    data[key] = ''
            else:
                # Inline list [a, b]
                if val.startswith('[') and val.endswith(']'):
                    inner = val[1:-1].strip()
                    if inner:
                        data[key] = [x.strip().strip('"').strip("'") for x in inner.split(',')]
                    else:
                        data[key] = []
                else:
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        val = val[1:-1]
                    data[key] = val
        # Lines starting with "- " without a current key are ignored
        i += 1
    return data, body

def dump_frontmatter(data: dict):
    # preserve insertion order to allow custom key ordering
    lines = ['---']
    for k in data.keys():
        v = data[k]
        if isinstance(v, list):
            if not v:
                lines.append(f"{k}: []")
            else:
                lines.append(f"{k}:")
                for item in v:
                    lines.append(f"  - {quote_scalar(item)}")
        else:
            lines.append(f"{k}: {quote_scalar(v)}")
    lines.append('---')
    return '\n'.join(lines) + '\n'

def load_config(path: Path):
    try:
        with path.open('r', encoding='utf-8') as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON in config '{path}': line {e.lineno} column {e.colno} — {e.msg}")
        sys.exit(2)
    except OSError as e:
        logger.error(f"Cannot read config '{path}': {e}")
        sys.exit(2)

def validate_config(cfg: dict):
    problems = False
    if not isinstance(cfg, dict):
        logger.error("Config must be a JSON object at top-level")
        sys.exit(2)
    for key in ('aliases', 'projects', 'transforms', 'list_fields'):
        if key in cfg and not isinstance(cfg[key], dict):
            logger.warning(f"Config '{key}' should be an object; got {type(cfg[key]).__name__}")
            problems = True
    if 'preserve_keys' in cfg and not isinstance(cfg['preserve_keys'], list):
        logger.warning("Config 'preserve_keys' should be a list")
        problems = True
    if 'overwrite_keys' in cfg and not isinstance(cfg['overwrite_keys'], list):
        logger.warning("Config 'overwrite_keys' should be a list")
        problems = True
    if problems:
        logger.info("Proceeding despite non-fatal config warnings.")

def is_valid_email(s: str) -> bool:
    if not s:
        return False
    return EMAIL_RE.match(s.strip()) is not None

def normalize_phone(fr_phone):
    if not fr_phone:
        return ''
    s = re.sub(r'[^\d+]', '', str(fr_phone))
    if s.startswith('00'):
        s = '+' + s[2:]
    if s.startswith('+33') and len(s) in (12, 13):
        s = '+33' + re.sub(r'^0', '', s[3:])
    elif s.startswith('0') and len(s) >= 10:
        s = '+33' + s[1:]
    return s

# ----------------------- ID prefix helpers -----------------------
def make_id_prefix(s: str) -> str:
    """
    Build a safe, readable prefix from user input (e.g., 'Ville d’Annonay' -> 'VILLE_ANNONAY').
    Keeps A-Z, 0-9 and underscores. Converts spaces/dashes to underscores.
    """
    if not s:
        return ''
    # reuse slugify to ascii, then upper-case, replace '-' by '_'
    base = slugify(s, allow_unicode=False).upper().replace('-', '_')
    # remove anything not alnum or underscore
    return re.sub(r'[^A-Z0-9_]', '', base)

def format_composed_id(prefix: str, raw_id: str, pad: int = 3) -> str:
    """
    Compose '<PREFIX>-NNN' from a prefix and a raw id (ideally numeric).
    If raw_id is not numeric, fallback to '<PREFIX>-<raw_id>'.
    """
    pfx = make_id_prefix(prefix)
    if not raw_id:
        return pfx if pfx else ''
    s = str(raw_id).strip()
    try:
        n = int(s)
        return f"{pfx}-{n:0{pad}d}" if pfx else f"{n:0{pad}d}"
    except ValueError:
        # keep as is if not an integer
        return f"{pfx}-{s}" if pfx else s

# ----------------------- Core mapping -----------------------

def select_first_present(row, candidates):
    if not candidates:
        return ''
    # Build a normalized map of row headers -> original header
    norm_to_orig = {}
    for k in row.keys():
        normalized_key = _norm_key_for_match(k)
        norm_to_orig[normalized_key] = k
        # also map singular form
        norm_to_orig[_singularize_norm(normalized_key)] = k
    for c in candidates:
        # 1) exact match first
        if c in row and str(row[c]).strip():
            return row[c]
        # 2) normalized (ignore case/accents, allow trailing 's')
        candidate_norm = _norm_key_for_match(c)
        cand_keys = [candidate_norm, _singularize_norm(candidate_norm)]
        for candidate_key in cand_keys:
            if candidate_key in norm_to_orig:
                v = row[norm_to_orig[candidate_key]]
                if str(v).strip():
                    return v
    return ''

def build_contact(row, cfg, project=None):
    aliases = cfg.get('aliases', {})
    extras_policy = cfg.get('extras', {'include': 'all'})
    list_fields = cfg.get('list_fields', {'groupes': True})
    transforms = cfg.get('transforms', {})
    project_overrides = (cfg.get('projects') or {}).get(project or '', {})

    lf_keys = set((cfg.get('list_fields') or {}).keys())
    def is_list_field(name: str) -> bool:
        normalized_name = _norm_key_for_match(name)
        for k in lf_keys:
            if _norm_key_for_match(k) == normalized_name:
                return True
        return False

    def candidates(field_key):
        if 'aliases' in project_overrides and field_key in project_overrides['aliases']:
            return project_overrides['aliases'][field_key]
        # Prefer explicit aliases; fallback to the logical key itself so exact header matches (e.g. 'Prénom') work without config
        return aliases.get(field_key, [field_key])

    contact = {}
    for logical in ['id_mslist', 'Nom', 'Prénom', 'Tel_Mobile', 'Tel_Fixe', 'Mail_Pro', 'Mail_Perso', 'organisation', 'groupes', 'type']:
        # Use aliases only; if none configured, leave empty
        val = select_first_present(row, candidates(logical)) if candidates(logical) else ''
        if logical in (cfg.get('list_fields') or {}):
            parsed = parse_maybe_list(val)
            if parsed is not None:
                val = [str(x).strip() for x in parsed if str(x).strip()]
            else:
                val = split_list(val)
        elif logical == 'type' and not val:
            val = 'contact'
        contact[logical] = val

    if extras_policy == 'all':
        # Mark primary logical fields as used so they aren't re-added as extras
        used = set()
        for x in ['id_mslist','Nom','Prénom','Tel_Mobile','Tel_Fixe','Mail_Pro','Mail_Perso','organisation','groupes','type']:
            used.update(candidates(x))

        # Build a normalized set of existing keys in `contact` to avoid duplicates
        existing_norm = set(_norm_key_for_match(k) for k in contact.keys())

        for k, v in row.items():
            # Skip if this header is one of the mapped primaries via alias list
            if k in used:
                continue
            # Skip if a normalized equivalent already exists in contact (e.g., 'Organisation' vs 'organisation')
            if _norm_key_for_match(k) in existing_norm:
                continue

            # Use the original CSV header name as the key to keep things dynamic
            parsed_list = parse_maybe_list(v)
            if parsed_list is not None:
                contact[k] = [str(x).strip() for x in parsed_list if str(x).strip()]
            elif is_list_field(k):
                contact[k] = split_list(v)
            else:
                # keep raw value, with boolean coercion if declared in config
                contact[k] = coerce_bool_if_needed(k, v, cfg)
    elif isinstance(extras_policy, list):
        for name in extras_policy:
            if name in row:
                key_norm = _norm_key_for_match(name)
                if key_norm in ( _norm_key_for_match(k) for k in contact.keys() ):
                    # Already mapped; skip to avoid duplication
                    continue
                v = row[name]
                parsed_list = parse_maybe_list(v)
                if parsed_list is not None:
                    contact[name] = [str(x).strip() for x in parsed_list if str(x).strip()]
                elif is_list_field(name):
                    contact[name] = split_list(v)
                else:
                    contact[name] = coerce_bool_if_needed(name, v, cfg)

    return contact

def merge_frontmatter(existing: dict, newdata: dict, preserve_keys=None, overwrite_keys=None):
    preserve_keys_set = set(preserve_keys or [])
    overwrite_keys_set = set(overwrite_keys or [])
    out = dict(existing)
    overwritten, preserved, filled = [], [], []
    def is_empty(val):
        if val is None:
            return True
        if isinstance(val, str):
            return val.strip() == ''
        if isinstance(val, list):
            return len(val) == 0
        return False
    for k, v in newdata.items():
        if k in preserve_keys_set and k in existing:
            preserved.append(k)
            continue
        if k in overwrite_keys_set or k not in existing or is_empty(existing.get(k)):
            prev = existing.get(k)
            out[k] = v
            if k in overwrite_keys_set and prev != v:
                overwritten.append(k)
            elif prev is None:
                filled.append(k)
    if logger.isEnabledFor(logging.DEBUG):
        if overwritten:
            logger.debug(f"merge overwrite: {sorted(overwritten)}")
        if preserved:
            logger.debug(f"merge preserved: {sorted(preserved)}")
        if filled:
            logger.debug(f"merge filled new keys: {sorted(filled)}")
    return out

# ----------------------- Refactored helpers -----------------------
def resolve_output_directory(args) -> Path:
    if args.vault_name:
        icloud_obsidian = Path.home() / 'Library' / 'Mobile Documents' / 'iCloud~md~obsidian' / 'Documents'
        vault_dir = icloud_obsidian / args.vault_name
        if not vault_dir.exists():
            logger.error(f"Vault '{args.vault_name}' introuvable dans iCloud: {vault_dir}")
            sys.exit(2)
        out_dir = vault_dir / args.out_subpath if args.out_subpath else vault_dir
        return out_dir
    if args.vault and args.out_subpath:
        return Path(args.vault).expanduser() / args.out_subpath
    if args.out:
        return Path(args.out).expanduser()
    logger.error("must provide either --out, or --vault + --out-subpath, or --vault-name (+ optional --out-subpath)")
    sys.exit(2)

def read_csv_rows(csv_path: Path, cfg: dict):
    try:
        with open(csv_path, 'r', encoding=cfg.get('csv_encoding','utf-8-sig'), newline='') as f:
            delim = cfg.get('csv_delimiter', 'auto')
            if delim == 'auto':
                sample = f.read(8192)
                try:
                    sniffed = csv.Sniffer().sniff(sample, delimiters=[',',';','\t','|'])
                    delimiter = sniffed.delimiter
                except Exception:
                    delimiter = ','
                f.seek(0)
            else:
                delimiter = delim
            reader = csv.DictReader(f, delimiter=delimiter)
            return list(reader)
    except UnicodeDecodeError as e:
        logger.error(f"CSV encoding error while reading {csv_path}: {e}. Hint: set 'csv_encoding' in config (e.g., 'utf-8-sig', 'latin-1').")
        sys.exit(2)
    except csv.Error as e:
        logger.error(f"CSV format error in {csv_path}: {e}")
        sys.exit(2)

def build_existing_indexes(out_dir: Path):
    existing_by_id = {}
    existing_by_email = {}
    for md in out_dir.glob('*.md'):
        try:
            txt = md.read_text(encoding='utf-8')
        except OSError as e:
            logger.warning(f"Cannot read note '{md}': {e}")
            continue
        fm, _ = read_frontmatter(txt)
        if not fm:
            continue
        if 'id_mslist' in fm and fm['id_mslist']:
            existing_by_id[str(fm['id_mslist'])] = md
        # Index email-like fields: any key containing 'email' or 'mail'
        email_candidates = []
        for key, val in fm.items():
            key_norm = _norm_key_for_match(key)
            if 'email' in key_norm or 'mail' in key_norm:
                if isinstance(val, list):
                    for item in val:
                        if isinstance(item, str):
                            email_candidates.append(item.strip().lower())
                else:
                    email_candidates.append(str(val).strip().lower())
        for e in email_candidates:
            if e and is_valid_email(e):
                existing_by_email[e] = md
    return existing_by_id, existing_by_email

def choose_primary_email_lower(contact: dict) -> str:
    for ek in ('email', 'Mail_Pro', 'Mail_Perso', 'mail_pro', 'mail_perso'):
        v = contact.get(ek)
        if v and str(v).strip():
            candidate = str(v).strip().lower()
            if is_valid_email(candidate):
                return candidate
    return ''

def build_frontmatter(contact: dict, cfg: dict) -> dict:
    new_fm = {}
    new_fm['Nom'] = contact.get('Nom','')
    new_fm['Prénom'] = contact.get('Prénom','')
    new_fm['Tel_Mobile'] = contact.get('Tel_Mobile','')
    new_fm['Tel_Fixe'] = contact.get('Tel_Fixe','')
    new_fm['Mail_Pro'] = contact.get('Mail_Pro','')
    new_fm['Mail_Perso'] = contact.get('Mail_Perso','')
    new_fm['organisation'] = contact.get('organisation','')
    # Then the remaining core metadata
    new_fm['id_mslist'] = contact.get('id_mslist','')
    new_fm['groupes'] = contact.get('groupes',[])
    new_fm['type'] = contact.get('type','contact')
    new_fm['source_updated'] = datetime.utcnow().isoformat(timespec='seconds') + 'Z'
    # Optional transforms
    if (cfg.get('transforms') or {}).get('normalize_phone_fr'):
        if new_fm.get('Tel_Mobile'):
            new_fm['Tel_Mobile'] = normalize_phone(new_fm['Tel_Mobile'])
        if new_fm.get('Tel_Fixe'):
            new_fm['Tel_Fixe'] = normalize_phone(new_fm['Tel_Fixe'])
    # Add remaining keys
    for k, v in contact.items():
        if k == 'src_id_mslist':
            continue
        if k in ('email', 'telephone') and not (v and str(v).strip()):
            continue
        if k not in new_fm:
            new_fm[k] = v
    return new_fm

def process_contact_row(row, cfg, args, existing_by_id, existing_by_email, out_dir: Path):
    contact = build_contact(row, cfg, project=args.project)
    # Skip rows with no essential data (configurable)
    required = cfg.get('required_fields', ['Nom','Prénom','Tel_Mobile','Tel_Fixe','Mail_Pro','Mail_Perso','organisation'])
    if not any(str(contact.get(k) or '').strip() for k in required):
        return 'SKIPPED', None, None, contact
    # Compose custom id if requested
    src_ms_id = str(contact.get('id_mslist') or '').strip()
    if args.id_prefix:
        composed = None
        if src_ms_id.isdigit():
            composed = format_composed_id(args.id_prefix, src_ms_id, args.id_pad)
            fb = 'numeric'
        else:
            if args.id_fallback == 'seq':
                composed = format_composed_id(args.id_prefix, str(process_contact_row.seq_counter), args.id_pad)
                process_contact_row.seq_counter += 1
                fb = 'seq'
            elif args.id_fallback == 'hash':
                basis = json.dumps(row, sort_keys=True, ensure_ascii=False)
                h8 = hashlib.sha1(basis.encode('utf-8')).hexdigest()[:8]
                composed = format_composed_id(args.id_prefix, h8, args.id_pad)
                fb = 'hash'
            elif args.id_fallback == 'raw':
                if src_ms_id:
                    composed = format_composed_id(args.id_prefix, src_ms_id, args.id_pad)
                    fb = 'raw'
                else:
                    fb = 'raw-empty'
            elif args.id_fallback == 'skip':
                if args.dry_run:
                    logger.debug("[id] skipping row due to --id-fallback=skip and non-numeric/empty source id")
                return 'SKIPPED', None, None, contact
        if composed:
            contact['id_mslist'] = composed
            if args.dry_run:
                logger.debug(f"[id] composed id_mslist={composed} (strategy={fb}, src='{src_ms_id or '∅'}')")
    # Title and filename
    last_name = (contact.get('Nom') or '').strip()
    first_name = (contact.get('Prénom') or '').strip()
    etab = get_contact_value(contact, ['Etablissement','Établissement','etablissement','organisation','Organisation'])
    if last_name:
        title_human = last_name.upper()
        if first_name:
            title_human += f" {first_name}"
    else:
        title_human = first_name or contact.get('organisation') or contact.get('email') or 'contact'
    if etab:
        title_human += f" - {etab}"
    filename_base = sanitize_filename(title_human)
    filename = f"{filename_base}.md"
    # Match existing
    unique_id = str(contact.get(args.id_key) or '').strip()
    email_lc = choose_primary_email_lower(contact)
    if unique_id and unique_id in existing_by_id:
        target = existing_by_id[unique_id]
    elif email_lc and email_lc in existing_by_email:
        target = existing_by_email[email_lc]
    else:
        target = out_dir / filename
        i = 2
        while target.exists():
            target = out_dir / f"{filename_base} - {i}.md"
            i += 1
    # FM
    new_fm = build_frontmatter(contact, cfg)
    body_default = "\n"
    if target.exists():
        txt = target.read_text(encoding='utf-8')
        old_fm, old_body = read_frontmatter(txt)
        cfg_overwrite = set(cfg.get('overwrite_keys', []))
        authoritative = DEFAULT_OVERWRITE_KEYS | cfg_overwrite
        merged = merge_frontmatter(
            existing=old_fm,
            newdata=new_fm,
            preserve_keys=cfg.get('preserve_keys', DEFAULT_PRESERVE_KEYS),
            overwrite_keys=authoritative
        )
        md_text = dump_frontmatter(merged) + (old_body if old_body else body_default)
        return 'UPDATED', target, md_text, contact
    else:
        md_text = dump_frontmatter(new_fm) + body_default
        return 'CREATED', target, md_text, contact

# Initialize a per-run sequential counter for composed IDs
process_contact_row.seq_counter = 1

def main():
    ap = argparse.ArgumentParser(description="CSV -> Obsidian contacts generator (flexible mapping)")
    ap.add_argument('--csv', required=True, help='Path to input CSV')
    ap.add_argument('--out', default='', help='Output folder (Obsidian Contacts directory). Optional if --vault or --vault-name is provided')
    ap.add_argument('--config', default='mapping.json', help='Mapping JSON file (default: mapping.json in current folder)')
    ap.add_argument('--project', default='', help='Project profile key from config (optional)')
    ap.add_argument('--id-key', default='id_mslist', help='Logical key used as unique id (default: id_mslist)')
    ap.add_argument('--id-prefix', default='', help='Optional prefix to compose a custom id (e.g., "Scolaires" -> SCO-001). Will be sanitized to A_Z/0-9/_ and uppercased.')
    ap.add_argument('--id-pad', type=int, default=3, help='Zero-padding width for numeric IDs when composing with --id-prefix (default: 3)')
    ap.add_argument('--id-fallback', default='seq', choices=['seq','hash','raw','skip'], help='When --id-prefix is set and source ID is non-numeric: seq (default), hash, raw, or skip the row')
    ap.add_argument('--dry-run', action='store_true', help='Preview actions without writing files')
    ap.add_argument('--vault', default='', help='Path to Obsidian vault root (optional)')
    ap.add_argument('--out-subpath', default='', help='Relative subpath inside the vault for contacts (e.g., \"Projets/MJC/Contacts\")')
    ap.add_argument('--vault-name', default='', help='Name of the Obsidian vault stored in iCloud Drive (~/Library/Mobile Documents/iCloud~md~obsidian/Documents/<vault>). Useful across Macs')
    ap.add_argument('--log-level', default='INFO', choices=['DEBUG','INFO','WARNING','ERROR'], help='Logging verbosity (default: INFO)')
    args = ap.parse_args()

    setup_logger(args.log_level)

    # Expand user on common paths
    csv_path = Path(args.csv).expanduser()
    config_path = Path(args.config).expanduser()

    # Resolve output directory with portable options
    out_dir: Path = resolve_output_directory(args)

    if not csv_path.exists():
        logger.error(f"CSV not found: {csv_path}")
        sys.exit(2)
    if not config_path.exists():
        logger.error(f"config not found: {config_path}. Pass --config or place mapping.json next to the script.")
        sys.exit(2)

    cfg = load_config(config_path)
    validate_config(cfg)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = read_csv_rows(csv_path, cfg)

    existing_by_id, existing_by_email = build_existing_indexes(out_dir)

    created, updated, skipped = 0, 0, 0
    for row in rows:
        action, target, md_text, contact = process_contact_row(
            row=row,
            cfg=cfg,
            args=args,
            existing_by_id=existing_by_id,
            existing_by_email=existing_by_email,
            out_dir=out_dir,
        )
        if action == 'SKIPPED':
            skipped += 1
            if args.dry_run:
                req = cfg.get('required_fields', ['Nom','Prénom','Tel_Mobile','Tel_Fixe','Mail_Pro','Mail_Perso','organisation'])
                print("[SKIP] Empty row: no essential fields present →", {k: contact.get(k, '') for k in req})
            continue
        if action == 'CREATED':
            created += 1
        elif action == 'UPDATED':
            updated += 1
        if args.dry_run:
            primary_email = choose_primary_email_lower(contact) or contact.get('Mail_Pro','') or contact.get('Mail_Perso','')
            final_id = contact.get('id_mslist','')
            print(f"[{action}] {final_id} | {target.name}  —  {primary_email}")
        else:
            try:
                target.write_text(md_text, encoding='utf-8')
            except OSError as e:
                logger.error(f"Failed to write '{target}': {e}")

    summary = f"Done. Created: {created}, Updated: {updated}, Skipped: {skipped}"
    if args.dry_run:
        print(summary)
    else:
        logger.info(summary)

if __name__ == '__main__':
    main()
