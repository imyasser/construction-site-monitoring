import json
import datetime as dt
import warnings
from pathlib import Path
import pandas as pd
import mysql.connector
from tracker import config

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = ROOT / "output"
DASHBOARD_DIR = ROOT / "dashboard"

SQL_CREATE_VIEW = """
CREATE OR REPLACE VIEW v_dernier_avancement AS
SELECT a.tache_id,
       a.avancement_pct,
       a.statut,
       a.observations,
       cr.date_cr
FROM avancement_hebdo a
JOIN comptes_rendus cr ON cr.cr_id = a.cr_id
JOIN (
    SELECT tache_id, MAX(cr.date_cr) AS max_date
    FROM avancement_hebdo a2
    JOIN comptes_rendus cr ON cr.cr_id = a2.cr_id
    GROUP BY tache_id
) latest ON latest.tache_id = a.tache_id AND latest.max_date = cr.date_cr;
"""

Q_LOTS = """
SELECT l.numero_lot, l.designation AS lot, COUNT(t.tache_id) AS nb_taches,
       ROUND(SUM(va.avancement_pct * COALESCE(t.duree_jours, 1)) /
             SUM(COALESCE(t.duree_jours, 1)), 0) AS avancement_pct
FROM lots l
JOIN taches t ON t.lot_id = l.lot_id
JOIN v_dernier_avancement va ON va.tache_id = t.tache_id
GROUP BY l.lot_id, l.numero_lot, l.designation
ORDER BY CAST(l.numero_lot AS UNSIGNED);
"""

# date_fin_prevue is computed on the fly; the contractor comes from lots -> entreprises.
Q_RETARDS = """
SELECT
    l.numero_lot,
    t.code_tache,
    t.designation,
    e.nom AS entreprise,
    DATE_ADD(t.date_debut_prevue, INTERVAL t.duree_jours DAY) AS date_fin_prevue,
    DATEDIFF(CURDATE(), DATE_ADD(t.date_debut_prevue, INTERVAL t.duree_jours DAY)) AS jours_de_retard,
    va.avancement_pct,
    va.observations
FROM taches t
JOIN lots l ON l.lot_id = t.lot_id
JOIN entreprises e ON e.entreprise_id = l.entreprise_id
JOIN v_dernier_avancement va ON va.tache_id = t.tache_id
WHERE t.date_debut_prevue IS NOT NULL
  AND t.duree_jours IS NOT NULL
  AND DATE_ADD(t.date_debut_prevue, INTERVAL t.duree_jours DAY) < CURDATE()
  AND va.avancement_pct < 100
ORDER BY jours_de_retard DESC;
"""

Q_CHARGE = """
SELECT e.nom AS entreprise, COUNT(t.tache_id) AS nb_taches,
       SUM(COALESCE(t.duree_jours, 0)) AS jours_planifies,
       ROUND(AVG(va.avancement_pct), 0) AS avancement_moyen_pct
FROM entreprises e
JOIN lots l ON l.entreprise_id = e.entreprise_id
JOIN taches t ON t.lot_id = l.lot_id
JOIN v_dernier_avancement va ON va.tache_id = t.tache_id
GROUP BY e.entreprise_id, e.nom
ORDER BY jours_planifies DESC;
"""

Q_GLOBAL = """
SELECT
    ROUND(SUM(va.avancement_pct * COALESCE(t.duree_jours, 1)) /
          SUM(COALESCE(t.duree_jours, 1)), 0) AS avancement_global_pct,
    DATEDIFF(c.date_livraison_prevue, CURDATE()) AS jours_restants
FROM taches t
JOIN v_dernier_avancement va ON va.tache_id = t.tache_id
JOIN lots l ON l.lot_id = t.lot_id
JOIN chantiers c ON c.chantier_id = l.chantier_id
GROUP BY c.chantier_id, c.date_livraison_prevue;
"""


def main():
    print("Connexion à la base de données...")
    cnx = mysql.connector.connect(**config.DB_CONFIG)
    cur = cnx.cursor()
    cur.execute(SQL_CREATE_VIEW)
    cnx.commit()
    cur.close()

    df_lots    = pd.read_sql(Q_LOTS,    cnx)
    df_retards = pd.read_sql(Q_RETARDS, cnx)
    df_charge  = pd.read_sql(Q_CHARGE,  cnx)
    df_global  = pd.read_sql(Q_GLOBAL,  cnx)
    cnx.close()

    # Export Excel
    OUTPUT_DIR.mkdir(exist_ok=True)
    fichier_excel = OUTPUT_DIR / "export_suivi_mhamid.xlsx"
    with pd.ExcelWriter(fichier_excel, engine="openpyxl") as writer:
        df_lots.to_excel(writer,    sheet_name="Avancement par lot",    index=False)
        df_retards.to_excel(writer, sheet_name="Retards en cours",      index=False)
        df_charge.to_excel(writer,  sheet_name="Charge par entreprise", index=False)
    print(f"Export Excel : {fichier_excel}")

    # Nettoyage NaN pour JSON valide
    df_lots    = df_lots.fillna("")
    df_retards = df_retards.fillna("")
    df_charge  = df_charge.fillna("")
    df_global  = df_global.fillna("")

    data = {
        "date_export": dt.date.today().isoformat(),
        "chantier":    config.CHANTIER_NOM,
        "global":      df_global.to_dict(orient="records")[0] if len(df_global) > 0 else {},
        "lots":        df_lots.to_dict(orient="records"),
        "retards":     df_retards.to_dict(orient="records"),
        "charge":      df_charge.to_dict(orient="records"),
    }

    fichier_json = DASHBOARD_DIR / "dashboard_data.json"
    with open(fichier_json, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"Export JSON généré : {fichier_json}")


if __name__ == "__main__":
    main()