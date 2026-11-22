import donnee_excel
import import_to_db
import initialisation_chantier
import export_hebdo


def demande_fichier():
    fichier = input("Entrer le nom du fichier Excel : ")
    return fichier.strip()


def demande_date():
    date = input("Date du relevé hebdo (AAAA-MM-JJ) : ")
    return date.strip()


def import_hebd():
    print("---- Import hebdomadaire du planning ----")
    fichier     = demande_fichier()
    date_releve = demande_date()

    print("\nEn cours de lecture...")
    taches_liste = donnee_excel.lire_planning_excel(fichier)
    print(f"{len(taches_liste)} tâches trouvées.\n")

    print("Connexion à la base de données...")
    connexion = import_to_db.se_connecter_a_la_base()
    cur       = connexion.cursor()

    id_chantier = import_to_db.recuperer_id_chantier(cur)

    # CORRECTION BUG 2 : creer_recuperer_compte_rendue attend 4 arguments
    # (cur, chantier_id, date_cr, fichier_source) — on avait passé 5 par erreur
    id_cr = import_to_db.creer_recuperer_compte_rendue(
        cur, id_chantier, date_releve,date_releve, fichier
    )

    print("Enregistrement des tâches...\n")

    aucun_retard_nbr   = 0
    retard_detecte_nbr = 0
    ignorees_nbr       = 0
    taches_non_trouvees = []

    for tache in taches_liste:
        resultat = import_to_db.enregister_tache(
            cur,
            tache["code_tache"],
            id_cr,
            tache["avancement"],
            tache["observations"],
            date_releve,
        )

        if resultat == "Aucun retard":
            aucun_retard_nbr += 1
        elif resultat == "Retard détecté":
            retard_detecte_nbr += 1
        elif resultat == "vide":
            ignorees_nbr += 1
        elif resultat == "Inconnue":
            ignorees_nbr += 1
            taches_non_trouvees.append(tache["code_tache"])

    connexion.commit()

    print("--- RÉSULTAT ---")
    print(f"CR du {date_releve} :")
    print(f"- {aucun_retard_nbr} tâches à jour")
    print(f"- {retard_detecte_nbr} tâches en retard")
    print(f"- {ignorees_nbr} tâches ignorées / non trouvées")

    if taches_non_trouvees:
        print(f"\nTâches non trouvées en base : {', '.join(taches_non_trouvees)}")

    cur.close()
    connexion.close()


def main():
    while True:
        print("\n------  Système de gestion des chantiers  ------")
        print("1. Importer un compte-rendu hebdomadaire")
        print("2. Ajouter un nouveau chantier (depuis Excel)")
        print("3. Exporter (Excel & JSON Dashboard)")
        print("4. Quitter")

        # CORRECTION BUG 4 : le menu va de 1 à 4, pas 1 à 5
        choix = input("Choisissez entre (1-4) : ").strip()

        if choix == "1":
            import_hebd()
        elif choix == "2":
            nm_chantier = input("Entrer le nom du chantier : ").strip()
            fichier     = demande_fichier()
            initialisation_chantier.inserer_chantier(nm_chantier, fichier)
        elif choix == "3":
            export_hebdo.main()
        elif choix == "4":
            print("Terminé.")
            break
        else:
            print("Choix invalide, entrer un nombre entre 1 et 4.")


if __name__ == "__main__":
    main()