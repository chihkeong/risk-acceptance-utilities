# Risk Acceptance Utilities

Utilities for processing and converting Penetration Testing (PT) results to Risk Assessment (RA) templates.

## Description

This repository contains tools to automate the conversion of PT findings into standardized Risk Assessment documentation.

### pt2ra.py

Converts PT results from Excel format to Risk Assessment templates.

**Features:**
- Read PT results from Excel files
- Extract key risk information (S/N, Risk Rating, Issue Title, Observation, Implications)
- Auto-fill RA templates with processed data
- Color-coded console output for easy review

## Prerequisites

- Python 3.6+
- openpyxl library (`pip install openpyxl`)
- colorama library (`pip install colorama`)

## Usage

1. Place your PT results in `pt_results.xlsx` (sheet name: "Risk Register")
2. Ensure you have `RA_Blank_Template_Only.xlsx` in the same directory
3. Run the script:

```bash
python pt2ra.py
```

The script will generate `Filled_RA_Template.xlsx` with the converted data.
