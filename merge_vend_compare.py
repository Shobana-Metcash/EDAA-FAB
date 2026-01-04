#!/usr/bin/env python3
"""
Script to merge CDL and GITHUB tabs from vend_compare.xlsx

Matching Rules:
- CDL Column G (Table Field Name) matches GITHUB Column D (cdm_column), OR
- CDL Column K (Biz Name) matches GITHUB Column E (pdm_column)

Output:
- A sheet named "vend_compare_merged" with:
  1. All CDL rows enriched with matched GITHUB data
  2. Unmatched CDL rows (CDL data only)
  3. Unmatched GITHUB rows appended at the end
  4. A "Comments" column indicating match status
"""

import pandas as pd
import sys

def values_match(val1, val2):
    """
    Check if two values match (case-insensitive, whitespace-trimmed).
    
    Args:
        val1: First value to compare
        val2: Second value to compare
        
    Returns:
        bool: True if values match, False otherwise
    """
    if pd.isna(val1) or pd.isna(val2):
        return False
    return str(val1).strip().upper() == str(val2).strip().upper()

def merge_vend_compare(input_file='vend_compare.xlsx', output_file=None):
    """
    Merge CDL and GITHUB sheets based on matching criteria.
    
    Args:
        input_file: Path to the input Excel file (default: vend_compare.xlsx)
        output_file: Path to the output Excel file (default: vend_compare_merged.xlsx)
                    If None, creates vend_compare_merged.xlsx
    """
    try:
        # Set default output file
        if output_file is None:
            output_file = 'vend_compare_merged.xlsx'
        
        # Read the sheets
        print(f"Reading {input_file}...")
        cdl_df = pd.read_excel(input_file, sheet_name='CDL')
        github_df = pd.read_excel(input_file, sheet_name='GITHUB')
        
        print(f"CDL sheet: {cdl_df.shape[0]} rows, {cdl_df.shape[1]} columns")
        print(f"GITHUB sheet: {github_df.shape[0]} rows, {github_df.shape[1]} columns")
        
        # Get the column names for matching
        cdl_col_g = 'Table Field Name'  # Column G
        cdl_col_k = 'Biz Name'  # Column K
        github_col_d = 'cdm_column'  # Column D
        github_col_e = 'pdm_column'  # Column E
        
        # Validate that required columns exist
        if cdl_col_g not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_g}'")
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
        for cdl_idx, cdl_row in cdl_df.iterrows():
            cdl_value_g = cdl_row[cdl_col_g]
            cdl_value_k = cdl_row[cdl_col_k]
            
            # Look for a matching GITHUB record
            matched_github_row = None
            matched_github_idx = None
            match_type = None
            
            for github_idx, github_row in github_df.iterrows():
                # Skip if already matched (to ensure one-to-one matching)
                if github_idx in matched_github_indices:
                    continue
                
                github_value_d = github_row[github_col_d]
                github_value_e = github_row[github_col_e]
                
                # Check if CDL Column G matches GITHUB Column D
                if values_match(cdl_value_g, github_value_d):
                    matched_github_row = github_row
                    matched_github_idx = github_idx
                    match_type = 'CDL Column G matches GITHUB Column D'
                    break
                
                # Check if CDL Column K matches GITHUB Column E
                if values_match(cdl_value_k, github_value_e):
                    matched_github_row = github_row
                    matched_github_idx = github_idx
                    match_type = 'CDL Column K matches GITHUB Column E'
                    break
            
            # Create merged record
            merged_record = cdl_row.to_dict()
            
            # If match found, append GITHUB columns
            if matched_github_row is not None:
                for col in github_df.columns:
                    # Prefix GITHUB columns to avoid conflicts
                    merged_record[f'GITHUB_{col}'] = matched_github_row[col]
                
                matched_github_indices.add(matched_github_idx)
                
                # Set comment based on match type
                merged_record['Comments'] = match_type
            else:
                # No match found, add empty GITHUB columns
                for col in github_df.columns:
                    merged_record[f'GITHUB_{col}'] = None
                
                # Set comment for no match
                merged_record['Comments'] = 'No matching record in GITHUB'
            
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
            
            # Set comment for unmatched GITHUB record
            merged_record['Comments'] = 'No matching record in CDL'
            
            merged_records.append(merged_record)
        
        # Create DataFrame from merged records
        merged_df = pd.DataFrame(merged_records)
        
        print(f"\nTotal merged records: {len(merged_df)}")
        
        # Save to Excel file with sheet name "vend_compare_merged"
        print(f"\nSaving to {output_file}...")
        merged_df.to_excel(output_file, index=False, sheet_name='vend_compare_merged')
        print(f"Merged sheet saved successfully!")

        return merged_df

    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'vend_compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'vend_compare_merged.xlsx'
    
    print("="*80)
    print("CDL and GITHUB Sheet Merge Tool for vend_compare.xlsx")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()
    
    result = merge_vend_compare(input_file, output_file)
    
    if result is not None:
        print("\nMerge completed successfully!")
    else:
        print("\nMerge failed.")
        sys.exit(1)
