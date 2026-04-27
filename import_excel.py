"""
import_excel.py — Convertit un fichier Excel mensuel → data.js

Usage :
    python import_excel.py lead-mars-2026.xlsx          # auto-détection depuis le nom
    python import_excel.py <fichier.xlsx> <YYYY-MM> "<Mois Année>"

Dépendances :
    pip install pandas openpyxl
"""

import sys
import json
import re
import os
import pandas as pd

FILE_RE = re.compile(r'lead-(\w+)-(\d{4})\.xlsx', re.IGNORECASE)
MONTH_NAMES = {
    'janv':'Janvier','fev':'Février','mars':'Mars','avr':'Avril',
    'mai':'Mai','juin':'Juin','juil':'Juillet','aout':'Août',
    'sept':'Septembre','oct':'Octobre','nov':'Novembre','dec':'Décembre',
}
MONTH_NUMS = {
    'janv':'01','fev':'02','mars':'03','avr':'04','mai':'05','juin':'06',
    'juil':'07','aout':'08','sept':'09','oct':'10','nov':'11','dec':'12',
}

def parse_filename(filename):
    m = FILE_RE.match(os.path.basename(filename))
    if not m:
        return None, None
    mois = m.group(1).lower()
    annee = m.group(2)
    num  = MONTH_NUMS.get(mois)
    name = MONTH_NAMES.get(mois)
    if not num:
        return None, None
    return f"{annee}-{num}", f"{name} {annee}"

# ─── CONFIG ─────────────────────────────────────────────────────────────────
# Adaptez ces valeurs selon la structure de votre Excel

SHEET_NAME  = 0          # 0 = première feuille, ou nom ex: "Stats"
HEADER_ROW  = 0          # Ligne d'entête (0-indexé, après skip_rows)
SKIP_ROWS   = None       # Nombre de lignes à sauter avant l'entête (None = 0)

# Plage de données à lire (None = tout)
USE_COLS    = "A:R"      # Colonnes Excel à lire, ex: "A:R"
MAX_ROWS    = 19         # Nb max de lignes de données (hors entête) — exclut les totaux en bas

# Correspondance colonnes Excel → champs JSON
# Clé = nom de colonne dans l'Excel (insensible à la casse)
# Valeur = nom du champ JSON
COL_MAP = {
    "sources":                "source",
    "budget":                 "budget",
    "nb cv":                  "nb_cv",
    "cpa cv":                 "cpa_cv",
    "candidat uniq":          "candidat_uniq",
    "new candidats validés":  "new_cand",
    "tx new cand validés":    "tx_new_cand",
    "cpnc":                   "cpnc",
    "nb intérimaires":        "nb_int",
    "tx mise à l'emploi":     "tx_emploi",
    "cme":                    "cme",
    "nint":                   "n_int",
    "% nint":                 "pct_n_int",
    "ca total hrfa":          "ca",
    "roi brut ca":            "roi_brut",
    "marge total":            "marge",
    "% marge":                "pct_marge",
    "roi réel marge":         "roi_reel",
}

# Sources gratuites (budget = 0 ou vide)
# Les autres seront automatiquement "payant" si budget > 0
GRATUIT_SOURCES = {
    "site carrière", "cvthèque", "candidature spontanée",
    "gmb", "google my business", "google jobs", "google jobs apply",
    "apec", "jooble", "jooble.org", "france travail", "talent", "talent.com",
    "linkedin limited", "monster organic", "truckfly",
}

# Lignes résumé à ignorer (contiennent ces mots)
SKIP_ROWS_CONTAINING = ["résumé", "resume", "total", "calcul"]

# ─── FONCTIONS ───────────────────────────────────────────────────────────────

def clean_val(v):
    """Convertit une cellule Excel en valeur Python propre."""
    if v is None:
        return None
    s = str(v).strip()
    if s in ('', '-', '—', '#DIV/0!', 'nan', 'NaN', 'N/A', '#N/A'):
        return None
    # Enlever symboles €, espaces insécables, virgules → points
    s = s.replace('\xa0', '').replace(' ', '').replace('€', '').replace(',', '.')
    # Pourcentage : "1,0%" → 0.01
    if s.endswith('%'):
        try:
            return round(float(s[:-1]) / 100, 6)
        except ValueError:
            return None
    try:
        return round(float(s), 2)
    except ValueError:
        return v  # retourne la chaîne originale (ex: nom de source)


def normalize_col(name):
    """Normalise un nom de colonne pour la correspondance."""
    return str(name).strip().lower().replace('\xa0', ' ')


def detect_type(source_name, budget):
    """Détermine si la source est gratuite ou payante."""
    if source_name.lower().strip() in GRATUIT_SOURCES:
        return "gratuit"
    if budget and budget > 0:
        return "payant"
    return "gratuit"


def import_month(excel_path, month_key, month_label):
    """Lit l'Excel et retourne un dict prêt pour data.js."""
    print(f"Lecture de : {excel_path}")
    df = pd.read_excel(
        excel_path,
        sheet_name=SHEET_NAME,
        header=HEADER_ROW,
        skiprows=SKIP_ROWS,
        usecols=USE_COLS,
        nrows=MAX_ROWS,
        dtype=str,          # tout en string, on parse nous-mêmes
    )

    # Normaliser les noms de colonnes
    col_norm = {normalize_col(c): c for c in df.columns}
    col_map_resolved = {}
    for excel_col_norm, json_key in COL_MAP.items():
        matched = col_norm.get(excel_col_norm)
        if matched:
            col_map_resolved[matched] = json_key
        else:
            # Recherche partielle
            for ec_norm, ec_orig in col_norm.items():
                if excel_col_norm in ec_norm or ec_norm in excel_col_norm:
                    col_map_resolved[ec_orig] = json_key
                    break

    if not col_map_resolved:
        print("ERREUR : Aucune colonne reconnue. Vérifiez COL_MAP dans le script.")
        print("Colonnes trouvées :", list(df.columns))
        sys.exit(1)

    print(f"Colonnes mappées : {len(col_map_resolved)} / {len(COL_MAP)}")

    rows = []
    for _, row in df.iterrows():
        # Récupérer le nom de la source
        src_col = next((c for c, j in col_map_resolved.items() if j == 'source'), None)
        if not src_col:
            continue
        source_name = str(row.get(src_col, '')).strip()

        # Ignorer lignes vides ou résumés
        if not source_name or source_name.lower() in ('nan', '', 'none'):
            continue
        if any(kw in source_name.lower() for kw in SKIP_ROWS_CONTAINING):
            continue

        # Construire l'objet
        obj = {"source": source_name}
        for excel_col, json_key in col_map_resolved.items():
            if json_key == 'source':
                continue
            raw = row.get(excel_col)
            obj[json_key] = clean_val(raw)

        # S'assurer que tous les champs existent
        for field in ['budget','nb_cv','cpa_cv','candidat_uniq','new_cand','tx_new_cand',
                      'cpnc','nb_int','tx_emploi','cme','n_int','pct_n_int',
                      'ca','roi_brut','marge','pct_marge','roi_reel']:
            obj.setdefault(field, None)

        # Budget = 0 si None pour les calculs
        budget = obj.get('budget') or 0
        obj['budget'] = budget if budget else 0

        obj['type'] = detect_type(source_name, budget)
        rows.append(obj)

    print(f"{len(rows)} sources importées.")
    return {"label": month_label, "rows": rows}


def extract_month_block(content, month_key):
    """Trouve le bloc { ... } d'un mois dans data.js. Retourne (start, end) ou None."""
    marker = f'"{month_key}":'
    pos = content.find(marker)
    if pos == -1:
        return None
    # Avancer jusqu'à l'accolade ouvrante
    brace_start = content.index('{', pos)
    depth = 0
    for i in range(brace_start, len(content)):
        if content[i] == '{':
            depth += 1
        elif content[i] == '}':
            depth -= 1
            if depth == 0:
                return (pos, i + 1)
    return None


def update_data_js(month_key, month_data, data_js_path="data.js"):
    """Insère ou remplace le mois dans data.js, en préservant les champs manuels (ats, ...)."""
    if not os.path.exists(data_js_path):
        print(f"ERREUR : {data_js_path} introuvable.")
        sys.exit(1)

    with open(data_js_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # ── Préserver les champs saisis manuellement ──────────────────
    MANUAL_FIELDS = ['ats']
    block = extract_month_block(content, month_key)
    if block:
        existing_json_str = content[block[0]:block[1]]
        # Extraire uniquement le dict {} du mois (sans la clé "YYYY-MM":)
        brace_idx = existing_json_str.index('{')
        try:
            existing = json.loads(existing_json_str[brace_idx:])
            for field in MANUAL_FIELDS:
                if field in existing and existing[field] is not None:
                    month_data[field] = existing[field]
                    print(f"  Champ '{field}' préservé : {existing[field]}")
        except Exception:
            pass

    # ── Sérialiser le nouveau mois ────────────────────────────────
    json_str = json.dumps(month_data, ensure_ascii=False, indent=2)
    indented  = '\n'.join('  ' + line for line in json_str.split('\n'))
    new_entry = f'  "{month_key}": {indented}'

    # ── Remplacer ou insérer ──────────────────────────────────────
    if block:
        # Remplacer le bloc existant (en incluant la clé "YYYY-MM":)
        # Vérifier s'il y a une virgule après le bloc
        end = block[1]
        suffix = content[end:end+2].strip()
        if suffix.startswith(','):
            end = content.index(',', end) + 1
        new_content = content[:block[0]] + new_entry + content[end:]
        print(f"Mois {month_key} mis à jour.")
    else:
        # Insérer avant le dernier }
        last_brace = content.rfind('};')
        if last_brace == -1:
            print("ERREUR : Structure de data.js non reconnue.")
            sys.exit(1)
        # Ajouter une virgule si le fichier contient déjà des mois
        needs_comma = content[:last_brace].rstrip().endswith('}')
        sep = ',\n' if needs_comma else '\n'
        new_content = content[:last_brace] + sep + new_entry + '\n' + content[last_brace:]
        print(f"Mois {month_key} ajouté.")

    with open(data_js_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f"data.js mis à jour avec le mois {month_key} ({month_data['label']}).")


# ─── MAIN ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    excel_path = sys.argv[1]

    if not os.path.exists(excel_path):
        print(f"ERREUR : Fichier introuvable : {excel_path}")
        sys.exit(1)

    # Auto-détection depuis le nom de fichier ou arguments explicites
    if len(sys.argv) >= 4:
        month_key   = sys.argv[2]
        month_label = sys.argv[3]
    else:
        month_key, month_label = parse_filename(excel_path)
        if not month_key:
            print(f"ERREUR : Impossible de détecter le mois depuis '{excel_path}'.")
            print("  Nommez le fichier : lead-mars-2026.xlsx")
            print("  Ou passez les arguments : python3 import_excel.py fichier.xlsx 2026-03 'Mars 2026'")
            sys.exit(1)
        print(f"Mois détecté : {month_key} ({month_label})")

    if not re.match(r'^\d{4}-\d{2}$', month_key):
        print(f"ERREUR : Format de mois invalide '{month_key}'. Attendu : YYYY-MM")
        sys.exit(1)

    month_data = import_month(excel_path, month_key, month_label)
    update_data_js(month_key, month_data)
    print("\nTerminé ! Rechargez le dashboard dans votre navigateur.")
