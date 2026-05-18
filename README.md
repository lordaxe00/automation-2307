# BIR Form 2307 Automation

A Python desktop application for generating BIR Form 2307 Excel files from database payee data.

## Overview

This repository contains a small automation tool that:
- connects to a SQL Server database
- retrieves payee and withholding data
- validates and formats it
- fills a BIR 2307 Excel template
- optionally exports PDF output when `pywin32` / Excel COM is available

## Files

- `main.py` — application entry point
- `requirement.txt` — required Python dependencies
- `config/app_config.json` — app configuration and defaults
- `config/cell_mapping.json` — Excel template cell mappings
- `assets/REFORMATTED-2307.xlsx` — Excel template used for output
- `src/` — application modules
- `ui/` — CustomTkinter user interface

## Setup

1. Create a Python virtual environment:

   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   ```

2. Install dependencies:

   ```powershell
   pip install -r requirement.txt
   ```

3. Edit `config/app_config.json` and set your database and payor settings.

## Running

Run the app from the repository root:

```powershell
python main.py
```

## Settings

Open `Settings` in the app to configure:
- database connection
- payor information
- period from / period to
- income description
- ATC code
- tax rate

The `Period & Income` settings now override default values for record creation.

## Notes

- `pywin32` is optional and only needed for Excel COM-based PDF export on Windows.
- The app uses `openpyxl` fallback when `win32com` is not available.
- Generated `.pyc`, `.venv`, `logs/`, and `OUTPUT/` are ignored by `.gitignore`.

## GitHub Upload

If you upload this folder manually to GitHub, include:
- `README.md`
- `.gitignore`
- `requirement.txt`
- `src/`, `ui/`, `config/`, `assets/`, and `main.py`

## License

Add a license file if you want to publish this repository publicly.
