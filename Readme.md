# 🏗️ Construction Site Tracking & Management System (OCIC - Mhamid)

A complete **Python**, **MySQL**, and **HTML/JS** solution designed to automate construction progress tracking from Excel files (Gantt schedules and weekly progress reports), detect task delays, and visualize key performance indicators (KPIs) through an interactive web dashboard.

## 📋 Table of Contents

* [Key Features](#-key-features)
* [Project Architecture](#-project-architecture)
* [Prerequisites & Installation](#-prerequisites--installation)
* [Configuration](#-configuration)
* [Usage](#-usage)
* [Viewing the Dashboard](#-viewing-the-dashboard)

## ✨ Key Features

* **Automated Site Initialization**: Extracts work packages (*lots*), tasks, planned start dates, durations, and contractors directly from an initial Excel schedule (`Table 1`, `Table 2`, `Table 7`).
* **Weekly Progress Import**: Parses weekly Excel status reports, updating task completion percentages and field observations in the database.
* **Automatic Delay Detection**: Compares planned end dates (`start_date + duration_days`) against the report date to automatically flag overdue tasks and calculate delay durations.
* **Multi-Format Exports**:
  * Generates a summary Excel workbook (`export_suivi_mhamid.xlsx`) detailing progress by work package, ongoing delays, and contractor workload.
  * Generates a JSON data feed (`dashboard_data.json`) to power the front-end dashboard.
* **Interactive Web Dashboard**: Dynamic interface built with Chart.js displaying global KPIs, progress per work package, contractor workload distribution, and a detailed delay log.

## 📂 Project Architecture

```
├── main.py                      # Main entry point (interactive CLI menu)
├── config.py                    # MySQL connection settings and Excel column mappings
├── initialisation_chantier.py   # Initial site import script (Gantt schedule & contractors)
├── donnee_excel.py              # Module for reading and cleaning weekly Excel reports
├── import_to_db.py              # Handles MySQL inserts/updates and delay detection logic
├── export_hebdo.py              # Creates SQL views, Excel reports, and JSON dashboard data
├── dashboard.html               # Web dashboard user interface
├── js.js                        # Front-end logic and Chart.js rendering
└── dashboard_data.json          # Exported JSON data consumed by the dashboard
```

## 🛠️ Prerequisites & Installation

### 1. System Requirements

* **Python 3.10+**
* **MySQL / MariaDB Server** (configured by default on port `3307`)
* A modern web browser (and a local HTTP server such as VS Code *Live Server* or Python's built-in server to avoid CORS issues)

### 2. Install Python Dependencies

Clone the repository and install the required packages:

```bash
git clone https://github.com/your-username/construction-site-tracker.git
cd construction-site-tracker
pip install pandas openpyxl mysql-connector-python
```

## ⚙️ Configuration

1. **Database Setup**: Ensure your `suivi_chantier_mhamid` database is created along with the required tables (`chantiers`, `entreprises`, `lots`, `taches`, `comptes_rendus`, `avancement_hebdo`, `retards`).
2. **Application Settings (`config.py`)**: Update the MySQL credentials and site name to match your environment:

```python
DB_CONFIG = {
    "host": "127.0.0.1",
    "port": 3307,
    "user": "root",
    "password": "your_password",
    "database": "suivi_chantier_mhamid"
}

CHANTIER_NOM = "OCIC - LBV Sud Carrefour Mhamid - Marrakech"
```

## 🚀 Usage

Run the main application from your terminal:

```bash
python main.py
```

The interactive CLI menu provides 4 options:

1. **Import a weekly progress report**: Enter the weekly Excel filename and the report date (`YYYY-MM-DD`) to update task progress and detect delays.
2. **Add a new construction site (from Excel)**: Initializes a complete site structure from an initial Excel schedule.
3. **Export (Excel & JSON Dashboard)**: Generates/updates `export_suivi_mhamid.xlsx` and `dashboard_data.json`.
4. **Quit**: Exits the application.

## 📊 Viewing the Dashboard

Because `dashboard.html` uses `fetch()` to load `dashboard_data.json`, browsers will block local `file://` requests due to CORS policies. Start a lightweight local HTTP server in the project directory:

```bash
python -m http.server 8000
```

Then open your browser and navigate to: `http://localhost:8000/dashboard.html`.