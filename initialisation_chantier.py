import os
import sys
from datetime import date, datetime, timedelta

import pandas as pd
import openpyxl
import mysql.connector

import config    # DB_CONFIG, COLONNES, PLANNING_ANNEE

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION DIRECTE (Remplacement d'argparse)
# ─────────────────────────────────────────────────────────────────────────────
FICHIER_EXCEL = "planning_chantier.xlsx"  # Remplacez par le nom de votre fichier
NOM_CHANTIER_OVERRIDE = None              # Laissez à None pour lecture automatique ou mettez un nom en dur

ANNEE = getattr(config, "PLANNING_ANNEE", 2026)    # année à appliquer aux dates JJ/MM


# ─────────────────────────────────────────────────────────────────────────────
# UTILITAIRES
# ─────────────────────────────────────────────────────────────────────────────

def _parse_date(txt: str | None) -> date | None:
    """
    Convertit 'JJ/MM' ou 'JJ/MM/AAAA' en date Python.
    Retourne None si le texte est vide, '-', ou non parseable.
    
    Exemples:
        "18/06"    → 2026-06-18 (avec ANNEE=2026)
        "18/06/2026" → 2026-06-18
        "18-06-2026" → 2026-06-18
    """
    if not txt or str(txt).strip() in ("", "-", "nan"):
        return None
    
    txt = str(txt).strip()
    
    # Remplacer les séparateurs '-' par '/' pour uniformiser
    txt = txt.replace("-", "/")
    
    parts = txt.split("/")
    
    if len(parts) == 3:
        # Format JJ/MM/AAAA ou JJ/MM/AA
        try:
            jour = int(parts[0])
            mois = int(parts[1])
            annee = int(parts[2])
            # Si année sur 2 chiffres (ex: 26 → 2026)
            if annee < 100:
                annee = 2000 + annee if annee >= 24 else 2000 + annee
            return date(annee, mois, jour)
        except ValueError:
            return None
    
    elif len(parts) == 2:
        # Format JJ/MM → on utilise ANNEE (2026 par défaut)
        try:
            jour = int(parts[0])
            mois = int(parts[1])
            return date(ANNEE, mois, jour)
        except ValueError:
            return None
    
    return None


def _parse_duree(txt) -> int | None:
    """'12j', '12', 12  →  12 ;  '-', None, '' → None"""
    if txt is None or str(txt).strip() in ("", "-", "nan"):
        return None
    txt = str(txt).lower().replace("j", "").strip()
    return int(txt) if txt.isdigit() else None


def _pct(val) -> int | None:
    """
    0.75 → 75,  1 → 100,  0 → 0,  None → None.
    Accepte aussi 75 ou '75%'.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, str):
        val = val.replace("%", "").strip()
        try:
            val = float(val)
        except ValueError:
            return None
    if isinstance(val, float) and val <= 1.0:
        return round(val * 100)
    return int(val)


def _statut_depuis_pct(pct: int | None) -> str:
    """Déduit le statut ENUM avancement_hebdo depuis le % initial."""
    if pct is None or pct == 0:
        return "a_demarrer"
    if pct >= 100:
        return "termine"
    return "en_cours"


def _clean(val) -> str | None:
    """Renvoie la valeur nettoyée, ou None si vide/nan."""
    if val is None:
        return None
    s = str(val).strip()
    return None if s in ("", "nan", "None", "—") else s


def _calculer_date_fin(date_debut: date | None, duree_jours: int | None) -> date | None:
    """
    Calcule la date de fin prévue à partir de la date de début et de la durée.
    
    Exemple:
        date_debut = 18/06/2026, duree_jours = 12 → 30/06/2026
    """
    if date_debut is None or duree_jours is None:
        return None
    return date_debut + timedelta(days=duree_jours)


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE TABLE 1 — TITRE / MÉTADONNÉES DU CHANTIER
# ─────────────────────────────────────────────────────────────────────────────

def lire_metadonnees(wb: openpyxl.Workbook, nom_override: str | None) -> dict:
    meta = {
        "nom": nom_override or "Chantier sans nom",
        "date_debut": None,
        "date_livraison_prevue": None,
    }

    if "Table 1" not in wb.sheetnames:
        print("  [!] Table 1 absente — métadonnées non extraites.")
        return meta

    ws = wb["Table 1"]
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            if not cell:
                continue
            txt = str(cell)

            if not nom_override and "OCIC" in txt and "—" in txt:
                parts = txt.split("|")
                if len(parts) >= 2:
                    meta["nom"] = parts[1].strip()

            if "LIVRAISON" in txt.upper():
                parts_liv = txt.split(":")
                if len(parts_liv) >= 2:
                    date_str = parts_liv[1].strip().split()[0]
                    meta["date_livraison_prevue"] = _parse_date(date_str)

    print(f"  Chantier     : {meta['nom']}")
    print(f"  Livraison    : {meta['date_livraison_prevue'].strftime('%d-%m-%Y') if meta['date_livraison_prevue'] else 'Non définie'}")
    return meta


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE TABLE 7 — ENTREPRISES
# ─────────────────────────────────────────────────────────────────────────────

def lire_entreprises(wb: openpyxl.Workbook) -> dict[str, dict]:
    entreprises: dict[str, dict] = {}

    if "Table 7" not in wb.sheetnames:
        print("  [!] Table 7 absente — entreprises non extraites.")
        return entreprises

    ws = wb["Table 7"]
    rows = list(ws.iter_rows(values_only=True))

    header_idx = None
    for i, row in enumerate(rows):
        if row and str(row[0]).strip().upper() == "LOT":
            header_idx = i
            break

    if header_idx is None:
        print("  [!] En-tête non trouvée dans Table 7.")
        return entreprises

    for row in rows[header_idx + 1:]:
        if not row or not row[0]:
            continue
        lot_raw  = _clean(row[0])
        societe  = _clean(row[1])
        contact  = _clean(row[2])
        concerne = _clean(row[3])

        if not lot_raw or not societe or societe in ("—", "-"):
            continue

        if "LOT" in lot_raw.upper():
            key = lot_raw.upper().replace("LOT", "").strip()
            role = "Entreprise exécutante"
        else:
            key = lot_raw
            role = "MOD" if lot_raw.upper() == "MOD" else "MO"

        entreprises[key] = {
            "nom": societe,
            "contact": contact,
            "role": role,
            "lots_concernes": concerne,
        }

    print(f"  Entreprises  : {len(entreprises)} lues depuis Table 7")
    return entreprises


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE TABLE 2 — PLANNING GANTT (lots + tâches)
# ─────────────────────────────────────────────────────────────────────────────

def lire_planning(wb: openpyxl.Workbook) -> tuple[list[dict], date | None]:
    if "Table 2" not in wb.sheetnames:
        raise ValueError("Table 2 (planning Gantt) introuvable dans le classeur.")

    ws = wb["Table 2"]
    rows = list(ws.iter_rows(values_only=True))

    col_code  = config.COLONNES.get("lot_ou_tache", "N° LOT")
    col_desig = config.COLONNES.get("designation",  "DÉSIGNATION DES TRAVAUX")
    col_resp  = config.COLONNES.get("responsable",  "RESPONSABLE")
    col_debut = config.COLONNES.get("debut",        "DÉBUT")
    col_duree = config.COLONNES.get("duree",        "DURÉE (j)")
    col_avanc = config.COLONNES.get("avancement",   "AVANC. %")
    col_obs   = config.COLONNES.get("observations", "OBSERVATIONS")

    header_idx = None
    col_map: dict[str, int] = {}
    for i, row in enumerate(rows):
        row_str = [str(c).strip() if c else "" for c in row]
        if col_code in row_str:
            header_idx = i
            col_map = {v: j for j, v in enumerate(row_str) if v}
            break

    if header_idx is None:
        raise ValueError(f"Colonne '{col_code}' introuvable dans Table 2.")

    def get(row, col_name):
        idx = col_map.get(col_name)
        return row[idx] if idx is not None and idx < len(row) else None

    date_debut_chantier = None
    lots: list[dict] = []
    lot_actuel = None

    for row in rows[header_idx + 1:]:
        if not row or all(c is None for c in row):
            continue

        code_raw = get(row, col_code)
        if code_raw is None:
            continue

        code_str = str(code_raw).strip()
        if not code_str or code_str in ("nan", "None"):
            continue

        if code_str.upper().startswith("LOT"):
            txt = code_str.replace("—", "|").replace(":", "|")
            parts = [p.strip() for p in txt.split("|")]
            num_lot = parts[0].upper().replace("LOT", "").strip()

            if len(parts) > 1 and parts[1]:
                designation = parts[1]
            else:
                designation = _clean(get(row, col_desig)) or f"LOT {num_lot}"

            lot_actuel = {
                "numero_lot": num_lot,
                "designation": designation,
                "taches": [],
            }
            lots.append(lot_actuel)
            continue

        if lot_actuel is None:
            continue

        # Nettoyer le code tâche (ex: 1.1000000000000001 → 1.1)
        try:
            code_tache = str(float(code_str)).rstrip('0').rstrip('.')
        except ValueError:
            code_tache = code_str

        designation   = _clean(get(row, col_desig)) or code_tache
        responsable   = _clean(get(row, col_resp))
        date_debut    = _parse_date(get(row, col_debut))
        duree_jours   = _parse_duree(get(row, col_duree))
        avanc_pct     = _pct(get(row, col_avanc))
        observations  = _clean(get(row, col_obs))

        if date_debut and (date_debut_chantier is None or date_debut < date_debut_chantier):
            date_debut_chantier = date_debut

        # ==============================================================
        # CALCUL DE LA DATE DE FIN PREVUE
        # ==============================================================
        date_fin_prevue = _calculer_date_fin(date_debut, duree_jours)

        # Formatage pour l'affichage
        date_debut_str = date_debut.strftime('%d-%m-%Y') if date_debut else None
        date_fin_str = date_fin_prevue.strftime('%d-%m-%Y') if date_fin_prevue else None

        lot_actuel["taches"].append({
            "code_tache":         code_tache,
            "designation":        designation,
            "responsable":        responsable,
            "date_debut_prevue":  date_debut,
            "date_debut_str":     date_debut_str,
            "duree_jours":        duree_jours,
            "date_fin_prevue":    date_fin_prevue,  # Date calculée
            "date_fin_str":       date_fin_str,     # Format français
            "avancement_pct_init": avanc_pct if avanc_pct is not None else 0,
            "observations_init":  observations,
        })

    nb_taches = sum(len(l["taches"]) for l in lots)
    print(f"  Lots         : {len(lots)}")
    print(f"  Tâches       : {nb_taches}")
    print(f"  Date début ↦  {date_debut_chantier.strftime('%d-%m-%Y') if date_debut_chantier else 'Non définie'}")
    return lots, date_debut_chantier


# ─────────────────────────────────────────────────────────────────────────────
# LECTURE TABLE 6 — DATE DU CR INITIAL
# ─────────────────────────────────────────────────────────────────────────────

def lire_date_cr(wb: openpyxl.Workbook) -> date | None:
    for sheet_name in ("Table 4", "Table 6", "Table 1"):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if not cell:
                    continue
                cell_str = str(cell)
                if "MÀJ" in cell_str.upper() or "MAJ" in cell_str.upper():
                    parts = cell_str.split(":")
                    if len(parts) >= 2:
                        date_str = parts[1].strip().split()[0]
                        d = _parse_date(date_str)
                        if d:
                            print(f"  Date CR init : {d.strftime('%d-%m-%Y')} (depuis {sheet_name})")
                            return d
    print("  [!] Date CR introuvable — date du jour utilisée.")
    return date.today()


# ─────────────────────────────────────────────────────────────────────────────
# INSERTION EN BASE
# ─────────────────────────────────────────────────────────────────────────────

def inserer_chantier(nom_chantier: str, fichier_excel: str):
    if not os.path.exists(fichier_excel):
        print(f"[ERREUR] Fichier introuvable : '{fichier_excel}'")
        return False

    print(f"\n{'='*60}")
    print(f"  Initialisation chantier depuis : {fichier_excel}")
    print(f"{'='*60}")

    wb = openpyxl.load_workbook(fichier_excel, read_only=True, data_only=True)

    meta          = lire_metadonnees(wb, nom_chantier)
    entreprises   = lire_entreprises(wb)
    lots, date_debut_chantier = lire_planning(wb)
    date_cr       = lire_date_cr(wb)

    if not lots:
        print("[ERREUR] Aucun lot extrait — opération annulée.")
        return False

    if meta["date_debut"] is None:
        meta["date_debut"] = date_debut_chantier

    try:
        conn = mysql.connector.connect(**config.DB_CONFIG)
        cur  = conn.cursor()
    except Exception as e:
        print(f"[ERREUR] Connexion MySQL impossible : {e}")
        return False

    try:
        # ── 1. CHANTIER ───────────────────────────────────────────
        cur.execute(
            """
            INSERT INTO chantiers (nom, date_debut, date_livraison_prevue, statut)
            VALUES (%s, %s, %s, 'en_cours')
            ON DUPLICATE KEY UPDATE
                date_debut            = COALESCE(VALUES(date_debut), date_debut),
                date_livraison_prevue = COALESCE(VALUES(date_livraison_prevue), date_livraison_prevue)
            """,
            (meta["nom"], meta["date_debut"], meta["date_livraison_prevue"])
        )
        conn.commit()

        cur.execute("SELECT chantier_id FROM chantiers WHERE nom = %s", (meta["nom"],))
        chantier_id = cur.fetchone()[0]
        print(f"\n  ✔ Chantier ID={chantier_id} : '{meta['nom']}'")

        # ── 2. ENTREPRISES ──────────────────────────────────────
        lot_to_entreprise_id: dict[str, int] = {}

        for lot_num, ent in entreprises.items():
            cur.execute(
                """
                INSERT INTO entreprises (nom, contact_terrain, role)
                VALUES (%s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    contact_terrain = COALESCE(VALUES(contact_terrain), contact_terrain),
                    role            = COALESCE(VALUES(role), role)
                """,
                (ent["nom"], ent["contact"], ent["role"])
            )
            conn.commit()
            cur.execute("SELECT entreprise_id FROM entreprises WHERE nom = %s", (ent["nom"],))
            ent_id = cur.fetchone()[0]
            lot_to_entreprise_id[lot_num] = ent_id

        print(f"  ✔ {len(lot_to_entreprise_id)} entreprises insérées/mises à jour")

        # ── 3. LOTS ─────────────────────────────────────────────
        lot_id_map: dict[str, int] = {}

        for lot in lots:
            num  = str(lot["numero_lot"])
            ent_id = lot_to_entreprise_id.get(num)

            cur.execute(
                """
                INSERT INTO lots (chantier_id, numero_lot, designation, entreprise_id)
                VALUES (%s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    designation   = VALUES(designation),
                    entreprise_id = COALESCE(VALUES(entreprise_id), entreprise_id)
                """,
                (chantier_id, num, lot["designation"], ent_id)
            )
            conn.commit()

            cur.execute(
                "SELECT lot_id FROM lots WHERE chantier_id = %s AND numero_lot = %s",
                (chantier_id, num)
            )
            lot_id = cur.fetchone()[0]
            lot_id_map[num] = lot_id

        print(f"  ✔ {len(lot_id_map)} lots insérés")

        # ── 4. COMPTE-RENDU INITIAL ────────────────────────────
        fichier_source = os.path.basename(fichier_excel)
        reference      = f"Réf. : {date_cr.strftime('%d/%m/%Y')}"

        cur.execute(
            """
            INSERT INTO comptes_rendus (chantier_id, date_cr, reference, fichier_source)
            VALUES (%s, %s, %s, %s)
            ON DUPLICATE KEY UPDATE
                reference      = VALUES(reference),
                fichier_source = VALUES(fichier_source)
            """,
            (chantier_id, date_cr, reference, fichier_source)
        )
        conn.commit()

        cur.execute(
            "SELECT cr_id FROM comptes_rendus WHERE chantier_id = %s AND date_cr = %s",
            (chantier_id, date_cr)
        )
        cr_id = cur.fetchone()[0]
        print(f"  ✔ Compte-rendu initial ID={cr_id} ({date_cr.strftime('%d-%m-%Y')})")

        # ── 5. TÂCHES + AVANCEMENT_HEBDO ────────────────────────
        total_taches   = 0
        total_avanc    = 0
        taches_sans_date = []

        for lot in lots:
            num    = str(lot["numero_lot"])
            lot_id = lot_id_map[num]

            for t in lot["taches"]:
                ent_id = lot_to_entreprise_id.get(num)

                # Surcharge si le responsable contient le nom d'une autre société
                for ent_key, ent in entreprises.items():
                    nom_soc = ent["nom"].lower()
                    resp    = (t["responsable"] or "").lower()
                    if nom_soc in resp and ent_key != num:
                        new_id = lot_to_entreprise_id.get(ent_key)
                        if new_id:
                            ent_id = new_id
                        break

                # ==============================================================
                # INSERTION DE LA TÂCHE
                # On insère date_fin_prevue calculée (car elle est GENERATED dans 
                # la base, mais on la calcule côté Python pour être sûr)
                # ==============================================================
                cur.execute(
                    """
                    INSERT INTO taches (
                        lot_id, code_tache, designation, responsable,
                        entreprise_id, date_debut_prevue, duree_jours,
                        date_fin_prevue,
                        avancement_pct_init, observations_init
                    )
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                        designation          = VALUES(designation),
                        responsable          = COALESCE(VALUES(responsable), responsable),
                        entreprise_id        = COALESCE(VALUES(entreprise_id), entreprise_id),
                        date_debut_prevue    = COALESCE(VALUES(date_debut_prevue), date_debut_prevue),
                        duree_jours          = COALESCE(VALUES(duree_jours), duree_jours),
                        date_fin_prevue      = COALESCE(VALUES(date_fin_prevue), date_fin_prevue),
                        avancement_pct_init  = VALUES(avancement_pct_init),
                        observations_init    = COALESCE(VALUES(observations_init), observations_init)
                    """,
                    (
                        lot_id,
                        t["code_tache"],
                        t["designation"],
                        t["responsable"],
                        ent_id,
                        t["date_debut_prevue"],
                        t["duree_jours"],
                        t["date_fin_prevue"],  # Date calculée
                        t["avancement_pct_init"],
                        t["observations_init"],
                    )
                )
                total_taches += 1

                # ==============================================================
                # RÉCUPÉRATION DE L'ID DE LA TÂCHE
                # ==============================================================
                cur.execute(
                    "SELECT tache_id FROM taches WHERE lot_id = %s AND code_tache = %s",
                    (lot_id, t["code_tache"])
                )
                result = cur.fetchone()
                if result is None:
                    raise Exception(f"Tâche non trouvée après insertion: {t['code_tache']}")
                tache_id = result[0]

                # Log des dates
                date_debut_str = t["date_debut_prevue"].strftime('%d-%m-%Y') if t["date_debut_prevue"] else "Non définie"
                date_fin_str = t["date_fin_prevue"].strftime('%d-%m-%Y') if t["date_fin_prevue"] else "Non définie"
                
                print(f"    Tâche {t['code_tache']}: début={date_debut_str}, fin={date_fin_str}, durée={t['duree_jours']}j")

                if t["date_debut_prevue"] is None:
                    taches_sans_date.append(f"    {t['code_tache']} — {t['designation']}")

                # ==============================================================
                # INSERTION DANS AVANCEMENT_HEBDO
                # ==============================================================
                pct    = t["avancement_pct_init"]
                statut = _statut_depuis_pct(pct)
                obs    = t["observations_init"]

                cur.execute(
                    """
                    INSERT INTO avancement_hebdo (tache_id, cr_id, avancement_pct, statut, observations)
                    VALUES (%s, %s, %s, %s, %s)
                    ON DUPLICATE KEY UPDATE
                        avancement_pct = VALUES(avancement_pct),
                        statut         = VALUES(statut),
                        observations   = COALESCE(VALUES(observations), observations)
                    """,
                    (tache_id, cr_id, pct, statut, obs)
                )
                total_avanc += 1

        conn.commit()
        cur.close()
        conn.close()

        print(f"\n{'='*60}")
        print(f"  ✅ Initialisation terminée avec succès")
        print(f"{'='*60}")
        print(f"  Chantier  : {meta['nom']} (ID={chantier_id})")
        print(f"  Lots      : {len(lots)}")
        print(f"  Tâches    : {total_taches}  (avancement_hebdo : {total_avanc} lignes)")
        print(f"  CR init   : {date_cr.strftime('%d-%m-%Y')}")

        if taches_sans_date:
            print(f"\n  ⚠ {len(taches_sans_date)} tâche(s) sans date de début (à compléter) :")
            for t in taches_sans_date:
                print(t)
        else:
            print("  ✔ Toutes les tâches ont une date de début.")

        return True

    except Exception as e:
        conn.rollback()
        print(f"\n[ERREUR] Insertion annulée : {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    ok = inserer_chantier(nom_chantier=NOM_CHANTIER_OVERRIDE, fichier_excel=FICHIER_EXCEL)
    sys.exit(0 if ok else 1)