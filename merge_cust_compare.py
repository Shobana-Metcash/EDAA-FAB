#!/usr/bin/env python3
"""
Script to merge CDL and GITHUB tabs from cust_compare.xlsx

Matching Rules:
- CDL Column I (Table Field Name) matches GITHUB Column D (cdm_column), OR
- CDL Column K (Biz Name) matches GITHUB Column E (pdm_column)

Output:
- A separate "Merged" sheet with:
  1. All CDL rows enriched with matched GITHUB data
  2. Unmatched CDL rows (CDL data only)
  3. Unmatched GITHUB rows appended at the end
"""

import pandas as pd
import sys
from openpyxl import load_workbook

def merge_cust_compare(input_file='cust_compare.xlsx', output_file=None):
    """
    Merge CDL and GITHUB sheets based on matching criteria.
    
    Args:
        input_file: Path to the input Excel file (default: cust_compare.xlsx)
        output_file: Path to the output Excel file (default: None, which adds Merged sheet to input_file)
    """
    try:
        # Read the sheets
        print(f"Reading {input_file}...")
        cdl_df = pd.read_excel(input_file, sheet_name='CDL')
        github_df = pd.read_excel(input_file, sheet_name='GITHUB')
        
        print(f"CDL sheet: {cdl_df.shape[0]} rows, {cdl_df.shape[1]} columns")
        print(f"GITHUB sheet: {github_df.shape[0]} rows, {github_df.shape[1]} columns")
        
        # Get the column names for matching
        # CDL Column I (index 8) = Table Field Name
        # CDL Column K (index 10) = Biz Name
        # GITHUB Column D (index 3) = cdm_column
        # GITHUB Column E (index 4) = pdm_column
        cdl_col_i = 'Table Field Name'  # Column I (index 8)
        cdl_col_k = 'Biz Name'  # Column K (index 10)
        github_col_d = 'cdm_column'  # Column D (index 3)
        github_col_e = 'pdm_column'  # Column E (index 4)
        
        # Validate that required columns exist
        if cdl_col_i not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_i}'")
        if cdl_col_k not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_k}'")
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
        # Note: Using nested loop for matching. This is O(n*m) but acceptable for datasets
        # of this size (190 CDL x 219 GITHUB records). We break after finding the first match
        # to minimize iterations. For larger datasets, consider using pandas merge or dict lookup.
        for cdl_idx, cdl_row in cdl_df.iterrows():
            cdl_value_i = cdl_row[cdl_col_i]
            cdl_value_k = cdl_row[cdl_col_k]
            
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
                
                # Check if CDL Column K matches GITHUB Column E
                if not match_found and not pd.isna(cdl_value_k) and not pd.isna(github_value_e):
                    if str(cdl_value_k).strip().upper() == str(github_value_e).strip().upper():
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
        
        # Determine output file
        if output_file is None:
            output_file = input_file
            print(f"\nAdding 'Merged' sheet to {output_file}...")
        else:
            print(f"\nSaving to {output_file}...")
        
        # Save to Excel file with the Merged sheet
        # If output is the same as input, preserve existing sheets
        if output_file == input_file:
            # Use pandas ExcelWriter with if_sheet_exists='replace' to handle existing sheet
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Merged')
        else:
            # Create new file with only Merged sheet
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Merged')
        
        print(f"Merged sheet saved successfully!")

        return merged_df

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'cust_compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    print("="*80)
    print("CDL and GITHUB Sheet Merge Tool for cust_compare.xlsx")
    print("="*80)
    print(f"Input file: {input_file}")
    if output_file:
        print(f"Output file: {output_file}")
    else:
        print(f"Output: Adding 'Merged' sheet to {input_file}")
    print()
    
    result = merge_cust_compare(input_file, output_file)
    
    if result is not None:
        print("\nMerge completed successfully!")
    else:
        print("\nMerge failed.")
        sys.exit(1)
