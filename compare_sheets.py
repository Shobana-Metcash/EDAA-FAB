#!/usr/bin/env python3
"""
Script to compare CDL and GITHUB tabs from compare.xlsx
Matches:
- CDL Column I (Table Field Name) with GITHUB Column C (cdm_column)
- CDL Column K (Biz Name) with GITHUB Column D (pdm_column)
Creates two output files:
1. matched_records.xlsx - Contains matched records as single rows
2. unmatched_records.xlsx - Contains unmatched records from both CDL and GITHUB in separate sheets
"""

import pandas as pd
import sys
from datetime import datetime

def compare_sheets(input_file='compare.xlsx', output_file='matched_records.xlsx', unmatched_file='unmatched_records.xlsx'):
    """
    Compare CDL and GITHUB sheets and create output with matched and unmatched records.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output Excel file for matched records
        unmatched_file: Path to the output Excel file for unmatched records
    """
    try:
        # Read the sheets
        print(f"Reading {input_file}...")
        cdl_df = pd.read_excel(input_file, sheet_name='CDL')
        github_df = pd.read_excel(input_file, sheet_name='GITHUB')
        
        print(f"CDL sheet: {cdl_df.shape[0]} rows, {cdl_df.shape[1]} columns")
        print(f"GITHUB sheet: {github_df.shape[0]} rows, {github_df.shape[1]} columns")
        
        # Get the column names for clarity
        cdl_col_i = 'Table Field Name'  # Column I (index 8)
        cdl_col_k = 'Biz Name'  # Column K (index 10)
        github_col_c = 'cdm_column'  # Column C (index 2)
        github_col_d = 'pdm_column'  # Column D (index 3)
        
        # Validate that required columns exist
        if cdl_col_i not in cdl_df.columns or cdl_col_k not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain columns '{cdl_col_i}' and '{cdl_col_k}'")
        if github_col_c not in github_df.columns or github_col_d not in github_df.columns:
            raise ValueError(f"GITHUB sheet must contain columns '{github_col_c}' and '{github_col_d}'")
        
        matched_records = []
        matched_cdl_indices = set()
        matched_github_indices = set()
        
        # Iterate through CDL rows
        for cdl_idx, cdl_row in cdl_df.iterrows():
            cdl_value_i = cdl_row[cdl_col_i]
            cdl_value_k = cdl_row[cdl_col_k]
            
            # Skip if both values are NaN
            if pd.isna(cdl_value_i) and pd.isna(cdl_value_k):
                continue
            
            # Check for matches in GITHUB sheet
            for github_idx, github_row in github_df.iterrows():
                github_value_c = github_row[github_col_c]
                github_value_d = github_row[github_col_d]
                
                # Check if Column I matches Column C
                match_type = None
                if not pd.isna(cdl_value_i) and not pd.isna(github_value_c):
                    if str(cdl_value_i).strip().upper() == str(github_value_c).strip().upper():
                        match_type = 'CDL_Column_I_matches_GITHUB_Column_C'
                
                # Check if Column K matches Column D
                if not pd.isna(cdl_value_k) and not pd.isna(github_value_d):
                    if str(cdl_value_k).strip().upper() == str(github_value_d).strip().upper():
                        if match_type:
                            match_type = 'Both_Matches'
                        else:
                            match_type = 'CDL_Column_K_matches_GITHUB_Column_D'
                
                # If there's a match, create a combined record
                if match_type:
                    matched_record = {
                        'Match_Type': match_type,
                        # CDL columns
                        'CDL_Source_System': cdl_row['Source System'],
                        'CDL_Table': cdl_row['CDL Table'],
                        'CDL_Column': cdl_row['CDL Column'],
                        'CDL_Source_Table': cdl_row['Source Table'],
                        'CDL_Source_Column': cdl_row['Source Column'],
                        'CDL_Join_Logic': cdl_row['Join Logic'],
                        'CDL_Transformation': cdl_row['Transformation'],
                        'CDL_Current_Outcome': cdl_row['Current Outcome'],
                        'CDL_Table_Field_Name': cdl_value_i,
                        'CDL_Definition': cdl_row['Definition'],
                        'CDL_Biz_Name': cdl_value_k,
                        # GITHUB columns
                        'GITHUB_d365_source': github_row['d365_source'],
                        'GITHUB_d365_table': github_row['d365_table'],
                        'GITHUB_cdm_column': github_value_c,
                        'GITHUB_pdm_column': github_value_d,
                    }
                    matched_records.append(matched_record)
                    matched_cdl_indices.add(cdl_idx)
                    matched_github_indices.add(github_idx)
        
        # Create DataFrame from matched records
        if matched_records:
            matched_df = pd.DataFrame(matched_records)
            print(f"\nFound {len(matched_records)} matching records")
            
            # Save to Excel file
            matched_df.to_excel(output_file, index=False, sheet_name='Matched_Records')
            print(f"Matched records saved to {output_file}")
            
            # Display summary
            print("\nMatch Type Summary:")
            print(matched_df['Match_Type'].value_counts())
        else:
            print("No matching records found.")
            matched_df = None
        
        # Find unmatched records
        unmatched_cdl_indices = set(cdl_df.index) - matched_cdl_indices
        unmatched_github_indices = set(github_df.index) - matched_github_indices
        
        print(f"\nFound {len(unmatched_cdl_indices)} unmatched CDL records")
        print(f"Found {len(unmatched_github_indices)} unmatched GITHUB records")
        
        # Create unmatched records file
        if unmatched_cdl_indices or unmatched_github_indices:
            with pd.ExcelWriter(unmatched_file, engine='openpyxl') as writer:
                if unmatched_cdl_indices:
                    unmatched_cdl_df = cdl_df.loc[list(unmatched_cdl_indices)]
                    unmatched_cdl_df.to_excel(writer, sheet_name='Unmatched_CDL', index=False)
                    print(f"Unmatched CDL records saved to sheet 'Unmatched_CDL' in {unmatched_file}")
                
                if unmatched_github_indices:
                    unmatched_github_df = github_df.loc[list(unmatched_github_indices)]
                    unmatched_github_df.to_excel(writer, sheet_name='Unmatched_GITHUB', index=False)
                    print(f"Unmatched GITHUB records saved to sheet 'Unmatched_GITHUB' in {unmatched_file}")
        
        return matched_df
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'matched_records.xlsx'
    unmatched_file = sys.argv[3] if len(sys.argv) > 3 else 'unmatched_records.xlsx'
    
    print("="*80)
    print("CDL and GITHUB Sheet Comparison Tool")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Matched records output: {output_file}")
    print(f"Unmatched records output: {unmatched_file}")
    print()
    
    result = compare_sheets(input_file, output_file, unmatched_file)
    
    print("\nComparison completed successfully!")
