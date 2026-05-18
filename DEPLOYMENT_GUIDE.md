# BIR Form 2307 Generator — Deployment & User Guide

---

## Table of Contents
1. [System Requirements](#1-system-requirements)
2. [Project Structure](#2-project-structure)
3. [Installation (Development)](#3-installation-development)
4. [Configuration](#4-configuration)
5. [Database Setup](#5-database-setup)
6. [Running the Application](#6-running-the-application)
7. [Building the EXE (Production)](#7-building-the-exe-production)
8. [Using the Application](#8-using-the-application)
9. [Cell Mapping Reference](#9-cell-mapping-reference)
10. [Troubleshooting](#10-troubleshooting)
11. [Architecture Overview](#11-architecture-overview)

---

## 1. System Requirements

| Requirement | Minimum |
|---|---|
| **OS** | Windows 10 / 11 (64-bit) |
| **Python** | 3.10 or higher |
| **Microsoft Excel** | 2016 or later (for PDF export) |
| **SQL Server** | 2014 or later |
| **ODBC Driver** | ODBC Driver 17 for SQL Server |
| **RAM** | 4 GB recommended |
| **Disk** | 200 MB (app + dependencies) |

---

## 2. Project Structure

```
bir2307_system/
│
├── main.py                     ← Entry point
│
├── config/
│   ├── app_config.json         ← DB, payor, defaults (EDIT THIS)
│   └── cell_mapping.json       ← Excel cell addresses (edit if template changes)
│
├── src/
│   ├── config_loader.py        ← Configuration manager
│   ├── logger_setup.py         ← Logging setup
│   ├── database.py             ← SQL Server connection & queries
│   ├── validators.py           ← Data validation & quarter logic
│   ├── excel_automation.py     ← Template filling (openpyxl)
│   ├── pdf_export.py           ← PDF generation (win32com)
│   └── batch_processor.py     ← Orchestration layer
│
├── ui/
│   ├── main_window.py          ← CustomTkinter main window
│   └── settings_dialog.py     ← Settings modal
│
├── assets/
│   └── BIR2307_template.xlsx  ← BIR 2307 blank template (DO NOT MODIFY)
│
├── logs/
│   └── bir2307.log            ← Auto-generated rotating log
│
├── OUTPUT/                     ← Generated files (auto-created)
│   └── Q1_2026/
│       ├── Juan Dela Cruz.xlsx
│       └── Juan Dela Cruz.pdf
│
├── requirements.txt
├── bir2307.spec                ← PyInstaller build config
└── DEPLOYMENT_GUIDE.md
```

---

## 3. Installation (Development)

### Step 1 — Install ODBC Driver
Download from Microsoft:
https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server

### Step 2 — Create Python virtual environment
```bash
cd bir2307_system
python -m venv venv
venv\Scripts\activate          # Windows
```

### Step 3 — Install dependencies
```bash
pip install -r requirements.txt
```

### Step 4 — Verify pywin32 post-install
```bash
python venv\Scripts\pywin32_postinstall.py -install
```

---

## 4. Configuration

### app_config.json (primary config)

Open `config/app_config.json` and fill in your values:

```json
{
  "database": {
    "driver": "ODBC Driver 17 for SQL Server",
    "server": "YOUR_SERVER_NAME_OR_IP",
    "database": "YOUR_DATABASE_NAME",
    "uid": "YOUR_SQL_USERNAME",
    "pwd": "YOUR_SQL_PASSWORD",
    "trusted_connection": false
  },
  "query": {
    "table": "WithholdingTax",
    "payee_name_col": "PAYEE_NAME",
    "month_col": "MONTH",
    "amount_col": "AMOUNT"
  },
  "payor": {
    "name": "ABC Corporation",
    "tin": "123-456-789-000",
    "address": "123 Business Ave, Makati City",
    "zip_code": "1200",
    "signatory": "Juan Dela Cruz / Chief Accountant / 000-123-456-000"
  },
  "defaults": {
    "atc_code": "WI010",
    "income_description": "Professional/Service Fees",
    "tax_rate": 0.05
  }
}
```

> **Tip:** You can also configure these settings from inside the app via **⚙ Settings**.

### Windows Authentication (Domain Users)
Set `"trusted_connection": true` and leave `uid`/`pwd` empty.

---

## 5. Database Setup

### Required Table Schema

```sql
CREATE TABLE WithholdingTax (
    ID          INT IDENTITY(1,1) PRIMARY KEY,
    PAYEE_NAME  NVARCHAR(255)  NOT NULL,
    MONTH       NVARCHAR(50)   NOT NULL,   -- e.g. 'January 2026'
    AMOUNT      DECIMAL(18, 2) NOT NULL,

    -- Optional columns (system uses defaults if absent):
    ATC_CODE    NVARCHAR(10),
    DESCRIPTION NVARCHAR(255),
    TAX_WITHHELD DECIMAL(18, 2)
);
```

### Sample Data

```sql
INSERT INTO WithholdingTax (PAYEE_NAME, MONTH, AMOUNT) VALUES
('Juan Dela Cruz',  'January 2026',  10000.00),
('Juan Dela Cruz',  'February 2026', 12000.00),
('Juan Dela Cruz',  'March 2026',    15000.00),
('Maria Santos',    'January 2026',   8000.00),
('Maria Santos',    'February 2026',  9000.00),
('Maria Santos',    'March 2026',    11000.00);
```

### Accepted MONTH Formats
The system parses all common formats:

| Format | Example |
|--------|---------|
| Full month + year | `January 2026` |
| Abbreviated month | `Jan 2026` |
| MM/YYYY | `01/2026` |
| YYYY-MM | `2026-01` |

---

## 6. Running the Application

```bash
# From the project root:
python main.py
```

---

## 7. Building the EXE (Production)

### Prerequisites
```bash
pip install pyinstaller
```

### Build
```bash
pyinstaller bir2307.spec --clean
```

Output: `dist\BIR2307Generator\BIR2307Generator.exe`

### Deployment to End-User PCs
1. Copy the entire `dist\BIR2307Generator\` folder to the target PC.
2. Ensure Microsoft Excel is installed on the target PC (for PDF export).
3. Install **ODBC Driver 17 for SQL Server** on the target PC.
4. Run `BIR2307Generator.exe`.
5. Configure DB settings on first launch via **⚙ Settings**.

> **Note:** No Python installation required on end-user PCs.

---

## 8. Using the Application

### Main Screen Walkthrough

```
┌─────────────────────────────────────────────────────────────┐
│ 🇵🇭 BIR Form 2307 Generator                    [↻] [⚙]    │
├──────────────────┬──────────────────────────────────────────┤
│ Payees           │ Output Folder: [C:\OUTPUT]  [Browse][Open]│
│ [🔍 Search]      │                                          │
│ ┌──────────────┐ │ [▶ Generate Selected] [▶▶ All] [⏹ Stop] │
│ │ Juan Dela C  │ │                                          │
│ │ Maria Santos │ │ Status: Processing 2 of 5…               │
│ │ Pedro Reyes  │ │ ████████░░░░░░░░  40%                    │
│ └──────────────┘ │                                          │
│ [Select All][Clr]│ Event Log:                               │
│ 3 payees         │ 10:01 [INFO] Juan Dela Cruz: Done ✓      │
└──────────────────┴──────────────────────────────────────────┘
```

### Step-by-Step Workflow

1. **Launch the app** — payees load automatically from the database.
2. **Select payees** — click one or Ctrl+click multiple in the list.
3. **Set output folder** — defaults to `OUTPUT/` next to the exe.
4. **Click Generate** — the system will:
   - Query the database for each payee's monthly records
   - Detect the quarter automatically
   - Fill the BIR 2307 Excel template
   - Export a PDF via Microsoft Excel
   - Save files as `OUTPUT/Q1_2026/Juan Dela Cruz.xlsx` and `.pdf`
5. **Monitor progress** — watch the progress bar and event log.
6. **Open output folder** — click "Open Folder" to view generated files.

### File Naming Convention
```
OUTPUT/
└── Q1_2026/
    ├── Juan Dela Cruz.xlsx
    ├── Juan Dela Cruz.pdf
    ├── Maria Santos.xlsx
    └── Maria Santos.pdf
```

---

## 9. Cell Mapping Reference

The file `config/cell_mapping.json` maps SQL data to BIR 2307 cells.
Update this only if you change the Excel template layout.

| Field | Default Cell | Notes |
|-------|-------------|-------|
| Period Date From | `J11` | MM/DD/YYYY format |
| Period Date To | `AB11` | MM/DD/YYYY format |
| Payee Name | `B17` | Row after label |
| Payee TIN | `N14` | After TIN label |
| Payor Name | `B29` | Row after label |
| Payor TIN | `N26` | After TIN label |
| Payor Address | `B32` | Row after label |
| ATC Code (row 1) | `L38` | First income row |
| 1st Month Amount | `O38` | Jan/Apr/Jul/Oct |
| 2nd Month Amount | `T38` | Feb/May/Aug/Nov |
| 3rd Month Amount | `Y38` | Mar/Jun/Sep/Dec |
| Total | `AD38` | Auto-calculated |
| Tax Withheld | `AI38` | Or computed from rate |

---

## 10. Troubleshooting

### "Template not found"
Place `BIR2307_template.xlsx` in the `assets/` folder next to `main.py`.

### "Connection failed"
- Verify server name/IP is correct.
- Check SQL Server Browser service is running.
- Ensure ODBC Driver 17 is installed.
- Test with: `sqlcmd -S SERVER_NAME -U USERNAME -P PASSWORD`

### PDF not generated
- Ensure Microsoft Excel is installed.
- Run the app as a user with Excel license.
- Check `logs/bir2307.log` for the exact COM error.
- PDF export is skipped gracefully if Excel is absent — the `.xlsx` is still saved.

### "Could not parse month"
Your `MONTH` column value is in an unrecognized format.
Supported: `January 2026`, `Jan 2026`, `01/2026`, `2026-01`.
Update your SQL data or add a computed column.

### Multiple quarters in one payee's records
The system uses the quarter of the **first chronological month**.
Records outside that quarter are skipped with a warning in the log.

---

## 11. Architecture Overview

```
┌─────────────┐     ┌──────────────────┐     ┌───────────────────┐
│  UI Layer   │────▶│  Batch Processor │────▶│  Database Manager │
│(CustomTkinter)     │  (Orchestrator)  │     │  (pyodbc)         │
└─────────────┘     └──────────────────┘     └───────────────────┘
                             │
                    ┌────────┴────────┐
                    ▼                 ▼
          ┌──────────────────┐  ┌──────────────────┐
          │ Excel Automation │  │   PDF Exporter   │
          │   (openpyxl)     │  │  (win32com/Excel)│
          └──────────────────┘  └──────────────────┘
                    │
          ┌──────────────────┐
          │   Validators     │
          │ (Quarter logic,  │
          │  data checking)  │
          └──────────────────┘
                    │
          ┌──────────────────┐
          │  Config Loader   │  ← app_config.json
          │  (Singleton)     │  ← cell_mapping.json
          └──────────────────┘
```

### Data Flow
```
SQL Server
    │  (pyodbc query)
    ▼
pandas DataFrame
    │  (validators.py)
    ▼
PayeeRecord (validated, quarter-detected)
    │  (excel_automation.py)
    ▼
Filled BIR2307 .xlsx → OUTPUT/Q1_2026/Name.xlsx
    │  (pdf_export.py / win32com)
    ▼
BIR2307 .pdf  → OUTPUT/Q1_2026/Name.pdf
```
