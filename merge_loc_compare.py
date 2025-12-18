#!/usr/bin/env python3
"""
Script to merge records between CDL and GITHUB tabs from Loc_Compare.xlsx

Matching Rules:
- CDL Column I (Table Field Name) matches GITHUB Column D (cdm_column), OR
- CDL Column K (EDL Tables) matches GITHUB Column E (pdm_column)

Output:
- Single merged sheet with all CDL rows (enriched with matching GITHUB data)
- Unmatched GITHUB rows appended at the end
"""

import pandas as pd
import sys
from datetime import datetime


def merge_loc_compare(input_file='Loc_Compare.xlsx', output_file='merged_output.xlsx'):
    """
    Merge CDL and GITHUB sheets from Loc_Compare.xlsx
    
    Args:
        input_file: Path to the input Excel file (Loc_Compare.xlsx)
        output_file: Path to the output Excel file for merged records
    """
    try:
        # Read the sheets
        print(f"Reading {input_file}...")
        cdl_df = pd.read_excel(input_file, sheet_name='CDL')
        github_df = pd.read_excel(input_file, sheet_name='GITHUB')
        
        print(f"CDL sheet: {cdl_df.shape[0]} rows, {cdl_df.shape[1]} columns")
        print(f"GITHUB sheet: {github_df.shape[0]} rows, {github_df.shape[1]} columns")
        
        # Define expected column names for matching
        # CDL Column I (index 8) and Column K (index 10)
        # GITHUB Column D (index 3) and Column E (index 4)
        cdl_col_i_expected = 'Table Field Name'
        cdl_col_k_expected = 'EDL Tables'
        github_col_d_expected = 'cdm_column'
        github_col_e_expected = 'pdm_column'
        
        # Validate that expected columns exist
        if cdl_col_i_expected not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_i_expected}' (expected at index 8)")
        if cdl_col_k_expected not in cdl_df.columns:
            raise ValueError(f"CDL sheet must contain column '{cdl_col_k_expected}' (expected at index 10)")
        if github_col_d_expected not in github_df.columns:
            raise ValueError(f"GITHUB sheet must contain column '{github_col_d_expected}' (expected at index 3)")
        if github_col_e_expected not in github_df.columns:
            raise ValueError(f"GITHUB sheet must contain column '{github_col_e_expected}' (expected at index 4)")
        
        # Use validated column names
        cdl_col_i = cdl_col_i_expected
        cdl_col_k = cdl_col_k_expected
        github_col_d = github_col_d_expected
        github_col_e = github_col_e_expected
        
        print(f"\nMatching columns:")
        print(f"  CDL: '{cdl_col_i}' (Column I)")
        print(f"  CDL: '{cdl_col_k}' (Column K)")
        print(f"  GITHUB: '{github_col_d}' (Column D)")
        print(f"  GITHUB: '{github_col_e}' (Column E)")
        
        # Track which GITHUB rows have been matched
        matched_github_indices = set()
        
        # List to store merged records
        merged_records = []
        
        # Process each CDL row
        print("\nProcessing CDL rows...")
        for cdl_idx, cdl_row in cdl_df.iterrows():
            cdl_value_i = cdl_row[cdl_col_i]
            cdl_value_k = cdl_row[cdl_col_k]
            
            # Start with the CDL row data
            merged_row = cdl_row.to_dict()
            
            # Look for matching GITHUB record
            # Note: Takes first match only. If multiple GITHUB rows match the same CDL row,
            # only the first match will be used. This is consistent with the requirement
            # to append matching GITHUB data to each CDL row.
            match_found = False
            for github_idx, github_row in github_df.iterrows():
                github_value_d = github_row[github_col_d]
                github_value_e = github_row[github_col_e]
                
                # Check if CDL Column I matches GITHUB Column D
                match_on_i_d = False
                if not pd.isna(cdl_value_i) and not pd.isna(github_value_d):
                    if str(cdl_value_i).strip().upper() == str(github_value_d).strip().upper():
                        match_on_i_d = True
                
                # Check if CDL Column K matches GITHUB Column E
                match_on_k_e = False
                if not pd.isna(cdl_value_k) and not pd.isna(github_value_e):
                    if str(cdl_value_k).strip().upper() == str(github_value_e).strip().upper():
                        match_on_k_e = True
                
                # If either match is found, append GITHUB data
                if match_on_i_d or match_on_k_e:
                    match_found = True
                    matched_github_indices.add(github_idx)
                    
                    # Add GITHUB columns with prefix to avoid conflicts
                    for github_col in github_df.columns:
                        merged_row[f'GITHUB_{github_col}'] = github_row[github_col]
                    
                    # Add match type indicator
                    if match_on_i_d and match_on_k_e:
                        merged_row['Match_Type'] = 'Both_Matches'
                    elif match_on_i_d:
                        merged_row['Match_Type'] = 'CDL_I_matches_GITHUB_D'
                    else:
                        merged_row['Match_Type'] = 'CDL_K_matches_GITHUB_E'
                    
                    break  # Take first match only
            
            # Add the merged row (with or without GITHUB data)
            if not match_found:
                merged_row['Match_Type'] = 'No_Match'
            
            merged_records.append(merged_row)
        
        print(f"Processed {len(merged_records)} CDL rows")
        
        # Find unmatched GITHUB rows
        unmatched_github_indices = set(github_df.index) - matched_github_indices
        print(f"Found {len(matched_github_indices)} matched GITHUB rows")
        print(f"Found {len(unmatched_github_indices)} unmatched GITHUB rows")
        
        # Append unmatched GITHUB rows to the end
        if unmatched_github_indices:
            print("\nAppending unmatched GITHUB rows...")
            for github_idx in sorted(unmatched_github_indices):
                github_row = github_df.loc[github_idx]
                
                # Create a row with only GITHUB data (CDL columns will be NaN)
                unmatched_row = {}
                
                # Add empty CDL columns
                for cdl_col in cdl_df.columns:
                    unmatched_row[cdl_col] = None
                
                # Add GITHUB columns with prefix
                for github_col in github_df.columns:
                    unmatched_row[f'GITHUB_{github_col}'] = github_row[github_col]
                
                unmatched_row['Match_Type'] = 'Unmatched_GITHUB'
                
                merged_records.append(unmatched_row)
        
        # Create merged DataFrame
        merged_df = pd.DataFrame(merged_records)
        
        # Reorder columns: Match_Type first, then CDL columns, then GITHUB columns
        cols = ['Match_Type'] + list(cdl_df.columns) + [f'GITHUB_{col}' for col in github_df.columns]
        # Only include columns that exist in the merged_df
        cols = [col for col in cols if col in merged_df.columns]
        merged_df = merged_df[cols]
        
        print(f"\nMerged output: {merged_df.shape[0]} rows, {merged_df.shape[1]} columns")
        
        # Display match type summary
        print("\nMatch Type Summary:")
        print(merged_df['Match_Type'].value_counts())
        
        # Save to Excel file
        merged_df.to_excel(output_file, index=False, sheet_name='Merged_Data')
        print(f"\nMerged data saved to {output_file}")
        
        return merged_df
            
    except FileNotFoundError as e:
        print(f"Error: Input file not found - {e}")
        print(f"Please ensure '{input_file}' exists in the current directory.")
        return None
    except ValueError as e:
        print(f"Error: Invalid data or column structure - {e}")
        print("Please check that the Excel file has the expected column names.")
        return None
    except KeyError as e:
        print(f"Error: Missing expected column - {e}")
        print("Please verify the CDL and GITHUB sheets have the required columns.")
        return None
    except PermissionError as e:
        print(f"Error: Permission denied - {e}")
        print(f"Please ensure you have write permissions for '{output_file}'.")
        return None
    except Exception as e:
        print(f"Unexpected error: {e}")
        import traceback
        traceback.print_exc()
        return None


if __name__ == "__main__":
    # Allow custom input/output file paths as command-line arguments
    input_file = sys.argv[1] if len(sys.argv) > 1 else 'Loc_Compare.xlsx'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'merged_output.xlsx'
    
    print("="*80)
    print("CDL and GITHUB Sheet Merge Tool")
    print("="*80)
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print()
    
    result = merge_loc_compare(input_file, output_file)
    
    if result is not None:
        print("\nMerge completed successfully!")
    else:
        print("\nMerge failed.")
        sys.exit(1)
