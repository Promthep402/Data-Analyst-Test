# FruitBlend24 Analysis

Python analysis project for the FruitBlend24 internship case.

## Requirements

- Python 3.10 or newer
- The workbook `FruitBlend24_Intern_Case_Data.xlsx`

## Run locally

```powershell
pip install -r requirements.txt
python analyze.py
```

Using `uv` is also supported:

```powershell
uv run --with pandas --with openpyxl python analyze.py
```

## Create the presentation

```powershell
pip install -r requirements.txt
python create_presentation.py
```

This creates `FruitBlend24_Data_Analyst_Presentation.pptx`, an 8-slide Data Analyst case presentation with embedded charts.

The script reads the Excel workbook and generates:

- `FruitBlend24_Analysis_Summary.json`: machine-readable analysis results
- `FruitBlend24_Analysis_Report.html`: browser-ready analysis report

## Project files

- `analyze.py`: analysis and report generation script
- `create_presentation.py`: presentation and chart generation script
- `requirements.txt`: Python dependencies
- `FruitBlend24_Intern_Case_Data.xlsx`: source workbook
- `FruitBlend24_Case_Brief.docx`: case brief

The HTML and JSON files in the repository are generated outputs and can be recreated by running the script.