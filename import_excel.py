"""
import_excel.py — Convertit un fichier Excel mensuel → data.js

Usage :
    python import_excel.py <fichier.xlsx> <YYYY-MM> "<Mois Année>"

Exemples :
    python import_excel.py stats_mars_2026.xlsx 2026-03 "Mars 2026"
    python import_excel.py stats_avril_2026.xlsx 2026-04 "Avril 2026"

Dépendances :
    pip install pandas openpyxl
"""

import sys
import json
import re
import os
import pandas as pd

# ─── CONFIG ─────────────────────────────────────────────────────────────────
# Adaptez ces valeurs selon la structure de votre Excel

SHEET_NAME  = 0          # 0 = première feuille, ou nom ex: "Stats"
HEADER_ROW  = 0          # Ligne d'entête (0-indexé, après skip_rows)
SKIP_ROWS   = None       # Nombre de lignes à sauter avant l'entête (None = 0)

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
SKIP_ROWS_CONTAINING = ["résumé", "resume", "total"]

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
        return float(s)
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


def update_data_js(month_key, month_data, data_js_path="data.js"):
    """Insère ou remplace le mois dans data.js."""
    if not os.path.exists(data_js_path):
        print(f"ERREUR : {data_js_path} introuvable. Exécutez depuis le dossier du dashboard.")
        sys.exit(1)

    with open(data_js_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Sérialiser le nouveau mois (JSON indenté)
    json_str = json.dumps(month_data, ensure_ascii=False, indent=2)
    # Indenter pour s'intégrer dans l'objet JS
    indented = '\n'.join('  ' + line for line in json_str.split('\n'))
    new_entry = f'  "{month_key}": {indented}'

    # Si le mois existe déjà → remplacer
    pattern = rf'"({re.escape(month_key)})":\s*\{{[^}}]*(?:\{{[^}}]*\}}[^}}]*)?\}}'
    # Approche simple : chercher le bloc entre "YYYY-MM": { ... }
    # On reconstruit plutôt le fichier entièrement

    # Extraire l'objet STATS_DATA existant et parser les clés/mois
    # On cherche où insérer (avant le commentaire de fin ou avant le })
    # Méthode robuste : on remplace le commentaire sentinelle

    SENTINEL = "  // Ajoutez les mois suivants ici via import_excel.py"

    if f'"{month_key}":' in content:
        print(f"Mois {month_key} déjà présent → mise à jour.")
        # Stratégie simple : ré-écrire le fichier en remplaçant le bloc du mois
        # Pour rester simple, on demande à l'utilisateur de vérifier
        # On utilise une regex multi-ligne conservative
        print("Mise à jour automatique limitée. Supprimez le mois existant manuellement si besoin,")
        print("ou laissez le script ajouter le mois — les deux coexisteront temporairement.")

    if SENTINEL in content:
        insert = f"{new_entry},\n{SENTINEL}"
        new_content = content.replace(SENTINEL, insert)
    else:
        # Fallback : insérer avant la dernière ligne de l'objet
        # Trouver le dernier }; et insérer avant
        last_brace = content.rfind('};')
        if last_brace == -1:
            print("ERREUR : Structure de data.js non reconnue.")
            sys.exit(1)
        new_content = content[:last_brace] + f"{new_entry}\n" + content[last_brace:]

    with open(data_js_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f"data.js mis à jour avec le mois {month_key} ({month_data['label']}).")


# ─── MAIN ────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(0)

    excel_path  = sys.argv[1]
    month_key   = sys.argv[2]   # ex: "2026-03"
    month_label = sys.argv[3]   # ex: "Mars 2026"

    if not os.path.exists(excel_path):
        print(f"ERREUR : Fichier introuvable : {excel_path}")
        sys.exit(1)

    if not re.match(r'^\d{4}-\d{2}$', month_key):
        print(f"ERREUR : Format de mois invalide '{month_key}'. Attendu : YYYY-MM")
        sys.exit(1)

    month_data = import_month(excel_path, month_key, month_label)
    update_data_js(month_key, month_data)
    print("\nTerminé ! Rechargez dashboard.html dans votre navigateur.")
