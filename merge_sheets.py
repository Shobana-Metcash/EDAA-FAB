#!/usr/bin/env python3
"""
Script to merge CDL and GITHUB tabs from Loc_Compare.xlsx

Matching Rules:
- CDL Column I (Table Field Name) matches GITHUB Column D (cdm_column), OR
- CDL Column M (c) matches GITHUB Column E (pdm_column)

Output:
- A separate "Merged" sheet with:
  1. All CDL rows enriched with matched GITHUB data
  2. Unmatched CDL rows (CDL data only)
  3. Unmatched GITHUB rows appended at the end
"""

import pandas as pd
import sys
from datetime import datetime

def merge_sheets(input_file='Loc_Compare.xlsx', output_file='Loc_Compare_Merged.xlsx'):
    """
    Merge CDL and GITHUB sheets based on matching criteria.
    
    Args:
        input_file: Path to the input Excel file (default: Loc_Compare.xlsx)
        output_file: Path to the output Excel file (default: Loc_Compare_Merged.xlsx)
    """
    try:
        # Read the sheets
        print(f"Reading {input_file}...")
        cdl_df = pd.read_excel(input_file, sheet_name='CDL')
        github_df = pd.read_excel(input_file, sheet_name='GITHUB')
        
        print(f"CDL sheet: {cdl_df.shape[0]} rows, {cdl_df.shape[1]} columns")
        print(f"GITHUB sheet: {github_df.shape[0]} rows, {github_df.shape[1]} columns")
        
        # Get the column names for matching
        cdl_col_i = 'Table Field Name'  # Column I (index 8)
        cdl_col_m = 'c'  # Column M (index 12)
        github_col_d = 'cdm_column'  # Column D (index 3)
        github_col_e = 'pdm_column'  # Column E (index 4)
        
        # Validate that required columns exist
        if cdl_col_i not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_i}'")
        if cdl_col_m not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_m}'")
        if github_col_d not in github_df.columns:
            raise ValueError(f"GITHUB sheet must contain column '{github_col_d}'")
        if github_col_e not in github_df.columns:
            raise ValueError(f"GITHUB sheet must contain column '{github_col_e}'")
        
        # Track which GITHUB records have been matched
        matched_github_indices = set()
        
        # List to store merged records
        merged_records = []
        
        # Process each CDL row
        print("\nProcessing CDL rows...")
        for cdl_idx, cdl_row in cdl_df.iterrows():
            cdl_value_i = cdl_row[cdl_col_i]
            cdl_value_m = cdl_row[cdl_col_m]
            
            # Look for a matching GITHUB record
            matched_github_row = None
            matched_github_idx = None
            
            for github_idx, github_row in github_df.iterrows():
                # Skip if already matched
                if github_idx in matched_github_indices:
                    continue
                
                github_value_d = github_row[github_col_d]
                github_value_e = github_row[github_col_e]
                
                # Check if CDL Column I matches GITHUB Column D
                match_found = False
                if not pd.isna(cdl_value_i) and not pd.isna(github_value_d):
                    if str(cdl_value_i).strip().upper() == str(github_value_d).strip().upper():
                        match_found = True
                
                # Check if CDL Column M matches GITHUB Column E
                if not match_found and not pd.isna(cdl_value_m) and not pd.isna(github_value_e):
                    if str(cdl_value_m).strip().upper() == str(github_value_e).strip().upper():
                        match_found = True
                
                if match_found:
                    matched_github_row = github_row
                    matched_github_idx = github_idx
                    matched_github_indices.add(github_idx)
                    break
            
            # Create merged record
            merged_record = cdl_row.to_dict()
            
            # If match found, append GITHUB columns
            if matched_github_row is not None:
                for col in github_df.columns:
                    # Prefix GITHUB columns to avoid conflicts
                    merged_record[f'GITHUB_{col}'] = matched_github_row[col]
            else:
                # No match found, add empty GITHUB columns
                for col in github_df.columns:
                    merged_record[f'GITHUB_{col}'] = None
            
            merged_records.append(merged_record)
        
        print(f"Processed {len(merged_records)} CDL records")
        print(f"Matched {len(matched_github_indices)} GITHUB records with CDL records")
        
        # Append unmatched GITHUB records
        unmatched_github_indices = set(github_df.index) - matched_github_indices
        print(f"Found {len(unmatched_github_indices)} unmatched GITHUB records to append")
        
        for github_idx in sorted(unmatched_github_indices):
            github_row = github_df.loc[github_idx]
            
            # Create a record with empty CDL columns
            merged_record = {}
            for col in cdl_df.columns:
                merged_record[col] = None
            
            # Add GITHUB columns
            for col in github_df.columns:
                merged_record[f'GITHUB_{col}'] = github_row[col]
            
            merged_records.append(merged_record)
        
        # Create DataFrame from merged records
        merged_df = pd.DataFrame(merged_records)
        
        print(f"\nTotal merged records: {len(merged_df)}")
        
        # Save to Excel file
        merged_df.to_excel(output_file, index=False, sheet_name='Merged')
        print(f"\nMerged sheet saved to {output_file}")

        return merged_df

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'Loc_Compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'Loc_Compare_Merged.xlsx'
    
    print("="*80)
    print("CDL and GITHUB Sheet Merge Tool")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()
    
    result = merge_sheets(input_file, output_file)
    
    if result is not None:
        print("\nMerge completed successfully!")
    else:
        print("\nMerge failed.")
        sys.exit(1)
