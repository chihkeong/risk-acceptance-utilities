# Risk Acceptance Utilities

Utilities for processing and converting security assessment results to Risk Assessment (RA) templates.

## Description

This repository contains tools to automate the conversion of security findings into standardized Risk Assessment documentation.

### pt2ra.py

Converts Penetration Testing (PT) results from Excel format to Risk Assessment templates.

**Features:**
- Read PT results from Excel files
- Extract key risk information (S/N, Risk Rating, Issue Title, Observation, Implications)
- Auto-fill RA templates with processed data
- Color-coded console output for easy review

### v2ra.py

Converts Vulnerability Assessment (VA) results from Excel format to Risk Assessment templates.

**Features:**
- Read VA results from Excel files (dkp_gcc.xlsx, dkp_onprem.xlsx)
- Extract CVE details (CVE ID, CVSS, CVSS String, CVSS Rating, RHA, plugin_text, Title, CIA)
- Auto-fill RA templates with processed vulnerability data
- Color-coded console output for easy review

## Prerequisites

- Python 3.6+
- openpyxl library (`pip install openpyxl`)
- colorama library (`pip install colorama`)

## Installation

```bash
pip install openpyxl colorama
```

## Usage

### PT to RA Conversion (pt2ra.py)

1. Place your PT results in `pt_results.xlsx` (sheet name: "Risk Register")
2. Ensure you have `RA_Blank_Template_Only.xlsx` in the same directory
3. Run the script:

```bash
python pt2ra.py
```

The script will generate `Filled_RA_Template.xlsx` with the converted data.

### VA to RA Conversion (v2ra.py)

1. Place your VA results in `dkp_gcc.xlsx` (sheet name: "Sheet")
2. Ensure you have `RA_Blank_Template_Only.xlsx` in the same directory
3. Run the script:

```bash
python v2ra.py
```

The script will generate `RA_gcc.xlsx` with the converted vulnerability data.

## Input File Format

### PT Results (pt2ra.py)
The PT results Excel file should have the following columns:
- S/N
- Overall Risk Rating
- Issue Title
- Observation
- Implications

### VA Results (v2ra.py)
The VA results Excel file should have the following columns:
- CVE
- CVSS
- CVSS String
- CVSS Rating
- RHA
- plugin_text
- Title
- CIA

## License

This project is for internal use.
