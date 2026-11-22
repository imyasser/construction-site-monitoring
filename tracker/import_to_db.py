import sys
import datetime

import mysql.connector
import pandas as pd

from tracker import config

def se_connecter_a_la_base():
  
    return mysql.connector.connect(**config.DB_CONFIG)

def recuperer_id_chantier(cur):
    query = "SELECT chantier_id FROM chantiers WHERE nom = %s"
    cur.execute(query,(config.CHANTIER_NOM, ))
    resultat = cur.fetchone()

    if not resultat:
        print(f"Le chantier '{config.CHANTIER_NOM}' n'existe pas dans la base de données.")
        sys.exit(1)

    return resultat[0]

def creer_recuperer_compte_rendue(cur,chantier_id, date_cr, reference, fichier_source):
    query  = """
    INSERT INTO comptes_rendus (chantier_id, date_cr, reference, fichier_source)
    VALUES(%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE fichier_source = VALUES(fichier_source)
    """

    cur.execute(query,(chantier_id, date_cr, reference or date_cr, fichier_source))

    cur.execute("SELECT cr_id FROM comptes_rendus WHERE chantier_id = %s AND date_cr = %s",(chantier_id,date_cr))

    return cur.fetchone()[0]

def transformer_en_pourcentage(val):
    try:
        temp = str(val).replace("%","").strip()
        valeur = float(temp)
        if 0 < valeur <=1:
            pourcentage = valeur * 100
        else:
            pourcentage = valeur
        return round(pourcentage)
    except(ValueError,TypeError):
        return 0


def statut_avancement(pourcentage):
    if pourcentage >=100:
        return "termine"
    elif pourcentage ==0:
        return "a_demarrer"
    else:
        return "en_cours"
  

def detecter_retard(cur,tache_id,cr_id,date_debut,duree,pourcentage,observation,date_cr):
    if date_debut is None or duree is None:
        return False
    date_debut_ini = datetime.date.fromisoformat(str(date_debut))
    date_fin_prevue = date_debut_ini + datetime.timedelta(days=int(duree))

    date_donnée = datetime.date.fromisoformat(str(date_cr))

    if pourcentage < 100 and date_donnée > date_fin_prevue:
        nbr_jrs_retard = (date_donnée - date_fin_prevue).days
        if observation is None:
            cause = "Retard non documenté"
        else:
            cause = observation

        sql = """
        INSERT INTO retards (tache_id,cr_id,jours_retard,cause,statut)
        VALUES(%s,%s,%s,%s,'ouvert')
        ON DUPLICATE KEY UPDATE jours_retard = VALUES(jours_retard), cause = VALUES(cause)
        """
        cur.execute(sql,(tache_id,cr_id,nbr_jrs_retard,cause))
        return True
    else:
        return False

def enregister_tache(cur,code_tache,id_cr,avancement_brut,observations,date_cr):
    if pd.isna(avancement_brut):
        return "vide"

    pourcentage = transformer_en_pourcentage(avancement_brut)

    if pd.isna(observations):
        observations = None
    else:
        observations = str(observations)

    #chercher la tache dans la table taches

    cur.execute("SELECT tache_id,date_debut_prevue,duree_jours FROM taches WHERE code_tache = %s",(code_tache,))

    resultat = cur.fetchone()
    if resultat is None:
        return "Inconnue"

    id_tache = resultat[0]
    date_debut_prevue = resultat[1]
    duree = resultat[2]

    statut = statut_avancement(pourcentage)

    sql = """INSERT INTO avancement_hebdo (tache_id,cr_id,avancement_pct,statut,observations)
    VALUES (%s,%s,%s,%s,%s)
    ON DUPLICATE KEY UPDATE avancement_pct = VALUES(avancement_pct), statut = VALUES(statut), observations = VALUES(observations)
    """

    cur.execute(sql,(id_tache,id_cr,pourcentage,statut,observations))

    check_retard = detecter_retard(cur,id_tache,id_cr,date_debut_prevue,duree,pourcentage,observations,date_cr)

    if check_retard:
        return "Retard détecté"
    else:
        return "Aucun retard"