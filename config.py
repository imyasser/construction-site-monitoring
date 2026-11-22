DB_CONFIG = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "root",
    "password": "",
    "database": "suivi_chantier_mhamid"
}

CHANTIER_NOM = "OCIC - LBV Sud Carrefour Mhamid - Marrakech"

# Noms des colonnes EXACTEMENT comme ils apparaissent dans le fichier Excel
# (avec accents et caractères spéciaux) — toute différence = colonne non trouvée
COLONNES = {
    "lot_ou_tache": "N° LOT",
    "designation":  "DÉSIGNATION DES TRAVAUX",   # accent sur É
    "responsable":  "RESPONSABLE",
    "debut":        "DÉBUT",                      # accent sur É, clé sans accent
    "duree":        "DURÉE (j)",                  # accent sur É, clé sans accent
    "avancement":   "AVANC. %",
    "observations": "OBSERVATIONS",
}