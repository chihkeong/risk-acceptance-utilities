# Changelog

All notable changes to the Risk Acceptance Utilities project will be documented in this file.

## [1.0.0] - 2026-02-26

### Added
- Initial release of PT to RA conversion utility
- `pt2ra.py` - Main conversion script
- Process PT results from Excel files
- Auto-fill RA templates with extracted data
- Color-coded console output using colorama
- Support for standard PT result columns (S/N, Risk Rating, Issue Title, Observation, Implications)

### Features
- Read PT results from `pt_results.xlsx`
- Validate RA template structure before filling
- Generate `Filled_RA_Template.xlsx` output
