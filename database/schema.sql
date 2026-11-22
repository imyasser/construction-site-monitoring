-- Schema reconstructed from the queries used in tracker/*.py.
-- Compare with your existing database before running it on live data.

CREATE DATABASE IF NOT EXISTS suivi_chantier_mhamid
  CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE suivi_chantier_mhamid;

CREATE TABLE IF NOT EXISTS chantiers (
  chantier_id           INT AUTO_INCREMENT PRIMARY KEY,
  nom                   VARCHAR(255) NOT NULL,
  date_debut            DATE NULL,
  date_livraison_prevue DATE NULL,
  statut                VARCHAR(30) NOT NULL DEFAULT 'en_cours',
  UNIQUE KEY uq_chantier_nom (nom)
);

CREATE TABLE IF NOT EXISTS entreprises (
  entreprise_id   INT AUTO_INCREMENT PRIMARY KEY,
  nom             VARCHAR(255) NOT NULL,
  contact_terrain VARCHAR(255) NULL,
  role            VARCHAR(100) NULL,
  UNIQUE KEY uq_entreprise_nom (nom)
);

CREATE TABLE IF NOT EXISTS lots (
  lot_id        INT AUTO_INCREMENT PRIMARY KEY,
  chantier_id   INT NOT NULL,
  numero_lot    VARCHAR(10) NOT NULL,
  designation   VARCHAR(255) NOT NULL,
  entreprise_id INT NULL,
  UNIQUE KEY uq_lot (chantier_id, numero_lot),
  FOREIGN KEY (chantier_id)   REFERENCES chantiers(chantier_id),
  FOREIGN KEY (entreprise_id) REFERENCES entreprises(entreprise_id)
);

CREATE TABLE IF NOT EXISTS taches (
  tache_id            INT AUTO_INCREMENT PRIMARY KEY,
  lot_id              INT NOT NULL,
  code_tache          VARCHAR(20) NOT NULL,
  designation         VARCHAR(255) NOT NULL,
  responsable         VARCHAR(255) NULL,
  entreprise_id       INT NULL,
  date_debut_prevue   DATE NULL,
  duree_jours         INT NULL,
  date_fin_prevue     DATE NULL,
  avancement_pct_init INT NOT NULL DEFAULT 0,
  observations_init   TEXT NULL,
  UNIQUE KEY uq_tache (lot_id, code_tache),
  FOREIGN KEY (lot_id)        REFERENCES lots(lot_id),
  FOREIGN KEY (entreprise_id) REFERENCES entreprises(entreprise_id)
);

CREATE TABLE IF NOT EXISTS comptes_rendus (
  cr_id          INT AUTO_INCREMENT PRIMARY KEY,
  chantier_id    INT NOT NULL,
  date_cr        DATE NOT NULL,
  reference      VARCHAR(100) NULL,
  fichier_source VARCHAR(255) NULL,
  UNIQUE KEY uq_cr (chantier_id, date_cr),
  FOREIGN KEY (chantier_id) REFERENCES chantiers(chantier_id)
);

CREATE TABLE IF NOT EXISTS avancement_hebdo (
  avancement_id  INT AUTO_INCREMENT PRIMARY KEY,
  tache_id       INT NOT NULL,
  cr_id          INT NOT NULL,
  avancement_pct INT NOT NULL,
  statut         ENUM('a_demarrer','en_cours','termine') NOT NULL,
  observations   TEXT NULL,
  UNIQUE KEY uq_avancement (tache_id, cr_id),
  FOREIGN KEY (tache_id) REFERENCES taches(tache_id),
  FOREIGN KEY (cr_id)    REFERENCES comptes_rendus(cr_id)
);

CREATE TABLE IF NOT EXISTS retards (
  retard_id    INT AUTO_INCREMENT PRIMARY KEY,
  tache_id     INT NOT NULL,
  cr_id        INT NOT NULL,
  jours_retard INT NOT NULL,
  cause        TEXT NULL,
  statut       VARCHAR(20) NOT NULL DEFAULT 'ouvert',
  UNIQUE KEY uq_retard (tache_id, cr_id),
  FOREIGN KEY (tache_id) REFERENCES taches(tache_id),
  FOREIGN KEY (cr_id)    REFERENCES comptes_rendus(cr_id)
);
