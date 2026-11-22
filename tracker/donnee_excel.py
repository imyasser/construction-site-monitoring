import sys
import pandas as pd
from tracker.config import COLONNES


def lire_planning_excel(fichier_excel):
    excel = pd.ExcelFile(fichier_excel)
    df = None

    for feuille in excel.sheet_names:
        temp_excel = pd.read_excel(fichier_excel, sheet_name=feuille, header=0)
        # Nettoyage des espaces dans les noms de colonnes
        temp_excel.columns = [str(c).strip() for c in temp_excel.columns]

        if COLONNES["lot_ou_tache"] in temp_excel.columns and COLONNES["avancement"] in temp_excel.columns:
            df = temp_excel
            break

    if df is None:
        print(f"Erreur : aucune feuille ne contient les colonnes requises "
              f"'{COLONNES['lot_ou_tache']}' et '{COLONNES['avancement']}'.")
        sys.exit(1)

    taches = []

    # CORRECTION BUG : on vérifie que obs_col existe et est dans le dataframe
    # avant d'essayer d'y accéder — évite un TypeError si la clé est absente de config
    obs_col = COLONNES.get("observations")
    obs_disponible = (obs_col is not None) and (obs_col in df.columns)

    for _, row in df.iterrows():
        val = row[COLONNES["lot_ou_tache"]]

        # Ligne vide -> ignorée
        if pd.isna(val) or str(val).strip() in ("", "nan"):
            continue

        lot_valeur = str(val).strip()

        # Ligne titre "LOT 1 — ..." -> ignorée
        if lot_valeur.upper().startswith("LOT"):
            continue

        taches.append({
            "code_tache":  lot_valeur,
            "avancement":  row.get(COLONNES["avancement"]),
            "observations": row.get(obs_col) if obs_disponible else None,
        })

    return taches