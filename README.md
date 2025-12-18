# EDAA-FAB Excel Comparison and Merge Tools

## Overview
This repository contains Python scripts for comparing and merging data between CDL and GITHUB sheets in Excel files.

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

### 1. Compare Sheets Script (`compare_sheets.py`)

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

### 2. Merge Sheets Script (`merge_sheets.py`)

#### Basic Usage
Run the merge script with default settings (reads `Loc_Compare.xlsx`, outputs `Loc_Compare_Merged.xlsx`):
```bash
python3 merge_sheets.py
```

#### Custom Input/Output Files
Specify custom input and output file paths:
```bash
python3 merge_sheets.py input_file.xlsx merged_output.xlsx
```

### 3. Customer Compare Merge Script (`merge_cust_compare.py`)

#### Basic Usage
Run the merge script with default settings (reads `cust_compare.xlsx`, outputs `cust_compare_merged.xlsx`):
```bash
python3 merge_cust_compare.py
```

#### Custom Input/Output Files
Specify custom input and output file paths:
```bash
python3 merge_cust_compare.py input_file.xlsx merged_output.xlsx
```

## How It Works

### Compare Sheets Script (`compare_sheets.py`)

The script compares two sheets in the input Excel file:

1. **CDL Sheet** - Contains source system data with:
   - Column I (index 8): "Table Field Name"
   - Column K (index 10): "Biz Name"

2. **GITHUB Sheet** - Contains GitHub data with:
   - Column C (index 2): "cdm_column"
   - Column D (index 3): "pdm_column"

#### Matching Logic
The script identifies matches using two criteria:
- **Match 1**: CDL Column I (Table Field Name) matches GITHUB Column C (cdm_column)
- **Match 2**: CDL Column K (Biz Name) matches GITHUB Column D (pdm_column)

Matches are case-insensitive and whitespace is trimmed.

#### Output
The script generates two Excel files:

##### 1. Matched Records File (`matched_records.xlsx`)
Contains all matched records with:
- All columns from both CDL and GITHUB sheets for matched records
- A `Match_Type` column indicating:
  - `Both_Matches`: Both Column I→C and Column K→D match
  - `CDL_Column_I_matches_GITHUB_Column_C`: Only Column I→C matches
  - `CDL_Column_K_matches_GITHUB_Column_D`: Only Column K→D matches

##### 2. Unmatched Records File (`unmatched_records.xlsx`)
Contains two sheets:
- **Unmatched_CDL**: All CDL records that didn't match any GITHUB records
- **Unmatched_GITHUB**: All GITHUB records that didn't match any CDL records

### Merge Sheets Script (`merge_sheets.py`)

The merge script combines CDL and GITHUB sheets from `Loc_Compare.xlsx`:

1. **CDL Sheet** - Contains source system data with:
   - Column I (index 8): "Table Field Name"
   - Column M (index 12): "c"

2. **GITHUB Sheet** - Contains GitHub data with:
   - Column D (index 3): "cdm_column"
   - Column E (index 4): "pdm_column"

#### Merging Logic
For each row in CDL:
- If **CDL Column I** matches **GITHUB Column D**, OR
- If **CDL Column M** matches **GITHUB Column E**

Then the matching GITHUB record's fields are appended to the CDL record.

Matches are case-insensitive and whitespace is trimmed.

#### Output
The script generates a single Excel file with a **Merged** sheet containing:

1. **All CDL rows** (56 rows) - Each CDL row is preserved with:
   - All original CDL columns
   - GITHUB columns prefixed with `GITHUB_` (populated if matched, empty if no match)

2. **Unmatched GITHUB rows** (appended at the end) - GITHUB records that didn't match any CDL record:
   - Empty CDL columns
   - Populated GITHUB columns prefixed with `GITHUB_`

The output preserves all data from both sheets while clearly indicating which records matched.

### Customer Compare Merge Script (`merge_cust_compare.py`)

The merge script combines CDL and GITHUB sheets from `cust_compare.xlsx`:

1. **CDL Sheet** - Contains customer data with:
   - Column I (index 8): "Table Field Name"
   - Column K (index 10): "Biz Name"

2. **GITHUB Sheet** - Contains GitHub data with:
   - Column D (index 3): "cdm_column"
   - Column E (index 4): "pdm_column"

#### Merging Logic
For each row in CDL:
- If **CDL Column I** (Table Field Name) matches **GITHUB Column D** (cdm_column), OR
- If **CDL Column K** (Biz Name) matches **GITHUB Column E** (pdm_column)

Then the matching GITHUB record's fields are appended to the CDL record.

Matches are case-insensitive and whitespace is trimmed.

#### Output
The script generates a single Excel file with a **Merged** sheet containing:

1. **All CDL rows** (190 rows) - Each CDL row is preserved with:
   - All original CDL columns
   - GITHUB columns prefixed with `GITHUB_` (populated if matched, empty if no match)

2. **Unmatched GITHUB rows** (appended at the end) - GITHUB records that didn't match any CDL record:
   - Empty CDL columns
   - Populated GITHUB columns prefixed with `GITHUB_`

The output preserves all data from both sheets while clearly indicating which records matched.

## Example Output

### Compare Sheets Example
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

### Merge Sheets Example
```
================================================================================
CDL and GITHUB Sheet Merge Tool
================================================================================
Input file: Loc_Compare.xlsx
Output file: Loc_Compare_Merged.xlsx

Reading Loc_Compare.xlsx...
CDL sheet: 56 rows, 13 columns
GITHUB sheet: 35 rows, 6 columns

Processing CDL rows...
Processed 56 CDL records
Matched 19 GITHUB records with CDL records
Found 16 unmatched GITHUB records to append

Total merged records: 72

Merged sheet saved to Loc_Compare_Merged.xlsx

Merge completed successfully!
```

### Customer Compare Merge Example
```
================================================================================
CDL and GITHUB Sheet Merge Tool for cust_compare.xlsx
================================================================================
Input file: cust_compare.xlsx
Output file: cust_compare_merged.xlsx

Reading cust_compare.xlsx...
CDL sheet: 190 rows, 11 columns
GITHUB sheet: 219 rows, 6 columns

Processing CDL rows...
Processed 190 CDL records
Matched 127 GITHUB records with CDL records
Found 92 unmatched GITHUB records to append

Total merged records: 282

Merged sheet saved to cust_compare_merged.xlsx

Merge completed successfully!
```

## Files
- `compare.xlsx` - Input Excel file with CDL and GITHUB sheets (for compare_sheets.py)
- `Loc_Compare.xlsx` - Input Excel file with CDL and GITHUB sheets (for merge_sheets.py)
- `cust_compare.xlsx` - Input Excel file with CDL and GITHUB sheets (for merge_cust_compare.py)
- `compare_sheets.py` - Comparison script that creates separate matched/unmatched files
- `merge_sheets.py` - Merge script that creates a single merged output sheet
- `merge_cust_compare.py` - Merge script for cust_compare.xlsx
- `matched_records.xlsx` - Output from compare_sheets.py with matching records
- `unmatched_records.xlsx` - Output from compare_sheets.py with unmatched records
- `Loc_Compare_Merged.xlsx` - Output from merge_sheets.py with merged data
- `cust_compare_merged.xlsx` - Output from merge_cust_compare.py with merged data
