# EDAA-FAB Excel Comparison Tool

## Overview
This repository contains a Python script that compares data between two sheets (CDL and GITHUB) in an Excel file and generates a report of matching records.

## Requirements
- Python 3.x
- pandas
- openpyxl

## Installation
Install the required dependencies:
```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage
Run the script with default settings (reads `compare.xlsx`, outputs `matched_records.xlsx`):
```bash
python3 compare_sheets.py
```

### Custom Input/Output Files
Specify custom input and output file paths:
```bash
python3 compare_sheets.py input_file.xlsx output_file.xlsx
```

## How It Works

The script compares two sheets in the input Excel file:

1. **CDL Sheet** - Contains source system data with:
   - Column I (index 8): "Table Field Name"
   - Column K (index 10): "Biz Name"

2. **GITHUB Sheet** - Contains GitHub data with:
   - Column C (index 2): "cdm_column"
   - Column D (index 3): "pdm_column"

### Matching Logic
The script identifies matches using two criteria:
- **Match 1**: CDL Column I (Table Field Name) matches GITHUB Column C (cdm_column)
- **Match 2**: CDL Column K (Biz Name) matches GITHUB Column D (pdm_column)

Matches are case-insensitive and whitespace is trimmed.

### Output
The script generates an Excel file with:
- All columns from both CDL and GITHUB sheets for matched records
- A `Match_Type` column indicating:
  - `Both_Matches`: Both Column I→C and Column K→D match
  - `CDL_Column_I_matches_GITHUB_Column_C`: Only Column I→C matches
  - `CDL_Column_K_matches_GITHUB_Column_D`: Only Column K→D matches

## Example Output
```
================================================================================
CDL and GITHUB Sheet Comparison Tool
================================================================================
Input file: compare.xlsx
Output file: matched_records.xlsx

Reading compare.xlsx...
CDL sheet: 190 rows, 11 columns
GITHUB sheet: 179 rows, 4 columns

Found 138 matching records
Matched records saved to matched_records.xlsx

Match Type Summary:
Match_Type
CDL_Column_K_matches_GITHUB_Column_D    67
Both_Matches                            64
CDL_Column_I_matches_GITHUB_Column_C     7

Comparison completed successfully!
```

## Files
- `compare.xlsx` - Input Excel file with CDL and GITHUB sheets
- `compare_sheets.py` - Main comparison script
- `matched_records.xlsx` - Output file with matching records (generated after running the script)
