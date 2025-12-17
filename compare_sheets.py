#!/usr/bin/env python3
"""
Script to compare CDL and GITHUB tabs from compare.xlsx
Matches:
- CDL Column I (Table Field Name) with GITHUB Column C (cdm_column)
- CDL Column K (Biz Name) with GITHUB Column D (pdm_column)
Creates a new file with matching records.
"""

import pandas as pd
import sys
from datetime import datetime

def compare_sheets(input_file='compare.xlsx', output_file='matched_records.xlsx'):
    """
    Compare CDL and GITHUB sheets and create output with matched records.
    
    Args:
        input_file: Path to the input Excel file
        output_file: Path to the output Excel file
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
        
        matched_records = []
        
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
            
            return matched_df
        else:
            print("No matching records found.")
            return None
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'matched_records.xlsx'
    
    print("="*80)
    print("CDL and GITHUB Sheet Comparison Tool")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()
    
    result = compare_sheets(input_file, output_file)
    
    if result is not None:
        print("\nComparison completed successfully!")
    else:
        print("\nComparison failed or no matches found.")
        sys.exit(1)
