# EDAA-FAB Excel Comparison and Merge Tools

## Overview
This repository contains Python scripts for comparing and merging data between two sheets (CDL and GITHUB) in Excel files.

## Requirements
- Python 3.x
- pandas
- openpyxl

## Installation
Install the required dependencies:
```bash
pip install pandas openpyxl
```

## Scripts

### 1. merge_loc_compare.py - Merge Records Tool

Merges records between CDL and GITHUB tabs from `Loc_Compare.xlsx`.

#### Usage
```bash
# Basic usage (default input: Loc_Compare.xlsx, output: merged_output.xlsx)
python3 merge_loc_compare.py

# Custom input/output files
python3 merge_loc_compare.py input_file.xlsx output_file.xlsx
```

#### Matching Logic
- **Match Rule 1**: CDL Column I (Table Field Name) matches GITHUB Column D (cdm_column)
- **Match Rule 2**: CDL Column K (EDL Tables) matches GITHUB Column E (pdm_column)

#### Output
A single merged sheet containing:
1. All CDL rows (with matching GITHUB data appended when a match is found)
2. Unmatched GITHUB rows appended at the end

The output includes a `Match_Type` column indicating:
- `CDL_I_matches_GITHUB_D`: Column I matches Column D
- `CDL_K_matches_GITHUB_E`: Column K matches Column E
- `Both_Matches`: Both matching conditions met
- `No_Match`: CDL row with no matching GITHUB record
- `Unmatched_GITHUB`: GITHUB row with no matching CDL record

### 2. compare_sheets.py - Comparison Tool

Compares CDL and GITHUB sheets and creates separate files for matched and unmatched records.

#### Basic Usage
Run the script with default settings (reads `compare.xlsx`, outputs `matched_records.xlsx` and `unmatched_records.xlsx`):
```bash
python3 compare_sheets.py
```

#### Custom Input/Output Files
Specify custom input and output file paths:
```bash
python3 compare_sheets.py input_file.xlsx matched_output.xlsx unmatched_output.xlsx
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
The script generates two Excel files:

#### 1. Matched Records File (`matched_records.xlsx`)
Contains all matched records with:
- All columns from both CDL and GITHUB sheets for matched records
- A `Match_Type` column indicating:
  - `Both_Matches`: Both Column I→C and Column K→D match
  - `CDL_Column_I_matches_GITHUB_Column_C`: Only Column I→C matches
  - `CDL_Column_K_matches_GITHUB_Column_D`: Only Column K→D matches

#### 2. Unmatched Records File (`unmatched_records.xlsx`)
Contains two sheets:
- **Unmatched_CDL**: All CDL records that didn't match any GITHUB records
- **Unmatched_GITHUB**: All GITHUB records that didn't match any CDL records

## Example Output
```
================================================================================
CDL and GITHUB Sheet Comparison Tool
================================================================================
Input file: compare.xlsx
Matched records output: matched_records.xlsx
Unmatched records output: unmatched_records.xlsx

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

Found 54 unmatched CDL records
Found 45 unmatched GITHUB records
Unmatched CDL records saved to sheet 'Unmatched_CDL' in unmatched_records.xlsx
Unmatched GITHUB records saved to sheet 'Unmatched_GITHUB' in unmatched_records.xlsx

Comparison completed successfully!
```

## Files
- `Loc_Compare.xlsx` - Input Excel file for merge operation (CDL and GITHUB sheets)
- `merge_loc_compare.py` - Script to merge CDL and GITHUB tabs into a single output
- `merged_output.xlsx` - Output file with merged records (generated after running merge_loc_compare.py)
- `compare.xlsx` - Input Excel file for comparison operation (CDL and GITHUB sheets)
- `compare_sheets.py` - Script to compare and separate matched/unmatched records
- `matched_records.xlsx` - Output file with matching records (generated after running compare_sheets.py)
- `unmatched_records.xlsx` - Output file with unmatched records from both sheets (generated after running compare_sheets.py)
