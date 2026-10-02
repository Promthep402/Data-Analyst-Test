# FruitBlend24 Analysis

Python analysis project for the FruitBlend24 internship case.

## Requirements

- Python 3.10 or newer
- The workbook `FruitBlend24_Intern_Case_Data.xlsx`

## Run locally (recommended)

Open PowerShell in the project folder:

```powershell
cd "C:\path\to\Data-Analyst-Test"
```

If `uv` is installed, run:

```powershell
uv run --with pandas --with openpyxl python analyze.py
```

Then open the generated report:

```powershell
Start-Process .\FruitBlend24_Analysis_Report.html
```

## Run locally (standard Python)

Python 3.10 or newer is required. Create an environment and install the
dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python analyze.py
Start-Process .\FruitBlend24_Analysis_Report.html
```

If PowerShell blocks activation, run this once in the same terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

The script reads and validates the Excel workbook:

- `FruitBlend24_Analysis_Summary.json`: machine-readable analysis results included with the project
- `FruitBlend24_Analysis_Report.html`: browser-ready dashboard included with the project

The generated artifacts are already included as submission files. The Python
script preserves the dashboard layout and matching summary when rerun.

## Project files

- `analyze.py`: analysis and report generation script
- `requirements.txt`: Python dependencies
- `FruitBlend24_Intern_Case_Data.xlsx`: source workbook
- `FruitBlend24_Case_Brief.docx`: case brief
- `FruitBlend24_Analysis_Report.html`: self-contained browser report
- `FruitBlend24_Analysis_Summary.json`: machine-readable output

The HTML and JSON files in the repository are generated outputs and can be recreated by running the script.