# 📊 Excel Automation & Reporting Tool

A production-grade Python application for automated data cleaning, validation, analytics, and Excel report generation — built for freelancing, portfolio, and resume use.

[![Python](https://img.shields.io/badge/Python-3.9%2B-blue)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.35%2B-red)](https://streamlit.io)
[![Pandas](https://img.shields.io/badge/Pandas-2.0%2B-green)](https://pandas.pydata.org)
[![License](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

---

## Overview

This tool accepts CSV and Excel files, runs a configurable cleaning pipeline, validates data against user-defined rules, generates analytics, and exports a formatted multi-sheet Excel workbook — all through a clean Streamlit dashboard.

**Practical use cases:**
- Cleaning raw CRM exports before import
- Validating vendor-supplied data files
- Generating monthly sales summary reports
- Automating repetitive Excel work for clients

---

## Features

| Category | Capability |
|---|---|
| **Upload** | CSV, XLSX, XLS · multiple sample datasets |
| **Cleaning** | Deduplication · whitespace trim · column name standardization · auto type conversion · 4 missing-value strategies |
| **Validation** | Required column check · email format regex · numeric field validation · per-row issue annotation |
| **Analytics** | Numeric describe · histograms · category bar charts · correlation heatmap · time-series line chart |
| **Export** | Cleaned CSV · invalid records CSV · formatted 4-sheet Excel workbook with embedded chart |
| **Logging** | Structured logs written to `exports/logs/YYYY-MM-DD.log` |

---

## Architecture

```
excel-automation-tool/
│
├── main.py                       # Streamlit dashboard (entry point)
│
├── automation/
│   ├── cleaner.py                # DataCleaner — composable 6-step cleaning pipeline
│   └── validator.py              # DataValidator — email, numeric & column rules
│
├── reports/
│   └── excel_exporter.py         # ExcelExporter — formatted 4-sheet workbook builder
│
├── utils/
│   ├── helpers.py                # Analytics computation, formatting, name mapping
│   └── logger.py                 # Centralized daily-file logging setup
│
├── data/
│   ├── sample_sales.xlsx         # 235-row sales dataset with injected dirty data
│   ├── sample_sales.csv          # CSV version of the same dataset
│   ├── sample_customers.xlsx     # 168-row customer dataset with injected dirty data
│   ├── sample_customers.csv      # CSV version of the same dataset
│   └── generate_sample_data.py   # Bootstrap script — re-run to regenerate all files
│
├── docs/
│   └── screenshots/              # App screenshots used in this README
│
├── exports/
│   └── logs/                     # Auto-created daily log files (gitignored)
│
├── requirements.txt
├── .gitignore
└── README.md
```

---

## Local Setup

### 1. Clone the repository

```bash
git clone https://github.com/your-username/excel-automation-tool.git
cd excel-automation-tool
```

### 2. Create a virtual environment

```bash
python -m venv venv
source venv/bin/activate        # macOS / Linux
venv\Scripts\activate           # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

### 4. Generate sample data

```bash
python data/generate_sample_data.py
```

### 5. Launch the dashboard

```bash
streamlit run main.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser.

---

## Usage Walkthrough

1. **Select data source** — upload your own file or pick a sample from the sidebar.
2. **Configure cleaning** — choose which steps to apply (dedup, trim, type conversion, missing-value strategy).
3. **Set validation rules** — mark required columns, email fields, and numeric fields.
4. **Click Process Data** — results appear across five tabs.
5. **Download reports** — grab the cleaned CSV, invalid records CSV, or the full formatted Excel workbook.

---

## Dashboard Tabs

| Tab | Contents |
|---|---|
| **Overview** | KPIs · raw data preview · column profile table |
| **Cleaning** | Before/after metrics · cleaned data preview · processing log |
| **Validation** | Missing column alerts · missing-value bar chart · invalid row table |
| **Analytics** | Numeric distributions · category charts · correlation heatmap · time-series |
| **Export** | Download buttons · processing summary |

---

## Deployment — Streamlit Community Cloud

1. Push the repository to GitHub (ensure `requirements.txt` is at the root).
2. Visit [share.streamlit.io](https://share.streamlit.io) and connect your GitHub account.
3. Select the repository, set **Main file path** to `main.py`, and deploy.

> **Note:** The `data/` folder must contain the sample `.xlsx` and `.csv` files. Run `python data/generate_sample_data.py` locally, commit the generated files, and push before deploying.

---

## Tech Stack

| Tool | Purpose |
|---|---|
| **Python 3.9+** | Core language |
| **pandas** | Data manipulation and analysis |
| **openpyxl** | Excel workbook generation and styling |
| **Streamlit** | Interactive web dashboard |
| **Plotly** | Interactive charts |
| **NumPy** | Numeric operations |

---

## Screenshots

### Landing Page
![Landing page showing step cards and feature overview](docs/screenshots/01_landing_page.png)

### Sidebar — Excel File Loaded
![Sidebar with sample_sales.xlsx loaded and cleaning/validation options configured](docs/screenshots/02_sidebar_excel_file_loaded.png)

### Overview Tab — KPIs & Data Preview
![Overview tab showing KPI cards (235 rows, 12 columns, 13 duplicates, 58 missing values) and raw data preview](docs/screenshots/03_overview_kpi_and_data_preview.png)

### Cleaning Tab — Results
![Cleaning tab showing 235 → 220 rows after removing 15 rows and 13 duplicates](docs/screenshots/04_cleaning_results.png)

### Validation Tab — Summary
![Validation tab showing 7 invalid emails flagged in red and 99.2% data completeness in green](docs/screenshots/05_validation_summary.png)

### Analytics — Numeric Summary
![Analytics tab showing describe() statistics for quantity, unit price and total amount](docs/screenshots/06_analytics_numeric_summary.png)

### Analytics — Distribution Charts
![Three Plotly histograms showing distributions of quantity, unit price and total amount](docs/screenshots/07_analytics_plotly_histograms.png)

### Analytics — Correlation Heatmap & Category Chart
![Pearson correlation heatmap and top-product category bar chart](docs/screenshots/08_analytics_correlation_and_category.png)

### Export Tab — Download Panel
![Export tab with three download buttons: Cleaned CSV, Invalid Records, and Full Excel Report](docs/screenshots/09_export_download_panel.png)

---

## Live Demo

[Open the Excel Automation Tool](https://excel-automation-azarsyed.streamlit.app)

Hosted on Streamlit Community Cloud. Upload your own CSV or XLSX, or
use the packaged sample datasets, to see the full clean → validate →
analyse → export pipeline in one click.

---

## Extending the Tool

- **Email report sending** — add `smtplib` integration to `reports/` and trigger after export.
- **SQLite history** — log every processing run to a local database via `sqlite3`.
- **Configurable rules** — load validation config from a JSON file instead of the sidebar.
- **Scheduled runs** — wrap `main.py` logic in a script and schedule with `cron` or `APScheduler`.

---

## License

MIT — free to use in commercial projects and client work.
