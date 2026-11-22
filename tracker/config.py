import os

# Credentials are read from environment variables (see .env.example).
# Defaults are for a local development database only.
DB_CONFIG = {
    "host":     os.getenv("DB_HOST", "127.0.0.1"),
    "port":     int(os.getenv("DB_PORT", "3307")),
    "user":     os.getenv("DB_USER", "root"),
    "password": os.getenv("DB_PASSWORD", ""),
    "database": os.getenv("DB_NAME", "suivi_chantier_mhamid"),
}

CHANTIER_NOM = os.getenv("CHANTIER_NOM", "OCIC - LBV Sud Carrefour Mhamid - Masiira")

# Year applied to planning dates written as DD/MM (no year)
PLANNING_ANNEE = int(os.getenv("PLANNING_ANNEE", "2026"))

# Column names EXACTLY as they appear in the Excel file
# (accents and special characters included) — any difference = column not found
COLONNES = {
    "lot_ou_tache": "N° LOT",
    "designation":  "DÉSIGNATION DES TRAVAUX",
    "responsable":  "RESPONSABLE",
    "debut":        "DÉBUT",
    "duree":        "DURÉE (j)",
    "avancement":   "AVANC. %",
    "observations": "OBSERVATIONS",
}
