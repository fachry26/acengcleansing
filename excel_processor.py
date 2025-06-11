import pandas as pd
import os
import time
import re # Import the re module for regular expressions
import openpyxl # Import openpyxl to handle workbooks directly
from openpyxl.utils.dataframe import dataframe_to_rows # To write DataFrame to existing sheet

def process_data_excel(input_filepath, cleaned_output_filepath, excluded_output_filepath, keywords_list=None, input_sheet_name='Sheet1', output_sheet_name='Processed Data'):

    # Use default keywords if none are provided or the list is empty
    if keywords_list is None or not keywords_list:
        keywords_list = []

    # Create a temporary file to work with, in case the original has hidden sheets
    temp_input_filepath = input_filepath + ".tmp"
    
    original_workbook = None 
    try:
        # Load the original workbook using openpyxl
        original_workbook = openpyxl.load_workbook(input_filepath)

        # Check if all sheets are hidden. If so, unhide the input_sheet_name temporarily.
        # This prevents the "At least one sheet must be visible" error.
        all_sheets_hidden = all(sheet.sheet_state == 'hidden' or sheet.sheet_state == 'veryHidden' for sheet in original_workbook._sheets)
        
        if all_sheets_hidden:
            if input_sheet_name not in original_workbook.sheetnames:
                raise ValueError(f"Input sheet '{input_sheet_name}' not found in the Excel file, and all sheets are hidden.")
            
            # Temporarily unhide the input sheet
            original_workbook[input_sheet_name].sheet_state = 'visible'
            print(f"Temporarily unhid sheet '{input_sheet_name}' as all sheets were hidden.")
        
        # Save the potentially unhidden workbook to a temporary file
        original_workbook.save(temp_input_filepath)

        # Read the Excel data from the specified input sheet into a pandas DataFrame
        # Use the temporary file to ensure pandas reads the updated visibility
        df = pd.read_excel(temp_input_filepath, sheet_name=input_sheet_name)

        # --- Debugging: Print all columns found in the Excel file ---
        print(f"Columns found in sheet '{input_sheet_name}' of {os.path.basename(input_filepath)}: {df.columns.tolist()}")

        # Try to find the 'KONTEN' column robustly (case-insensitive, strip spaces)
        konten_col_name = None
        for col in df.columns:
            normalized_col = str(col).strip().upper()
            if normalized_col == 'KONTEN':
                konten_col_name = col
                break

        if konten_col_name is None:
            # If 'KONTEN' column is still not found after robust search
            raise ValueError(
                f"'KONTEN' column not found in sheet '{input_sheet_name}' of {os.path.basename(input_filepath)}. "
                f"Available columns are: {df.columns.tolist()}"
            )

        # Convert the identified 'KONTEN' column to string type and fill NaN values
        df[konten_col_name] = df[konten_col_name].astype(str).fillna('')

        # Initialize an empty mask that is all False, to collect rows for exclusion
        combined_mask_exclude = pd.Series([False] * len(df), index=df.index)

        # 1. Create mask for keywords exclusion
        for keyword in keywords_list:
            # Combine individual keyword matches using logical OR
            combined_mask_exclude = combined_mask_exclude | \
                                    df[konten_col_name].str.contains(str(keyword).strip(), case=False, na=False)

        # 2. Create mask for foreign language character exclusion (allowing emojis and common symbols)
        # This regex targets specific Unicode ranges for common non-Latin scripts:
        # Chinese (Han): \u4E00-\u9FFF
        # Japanese (Hiragana/Katakana): \u3040-\u30FF (Hiragana), \u30A0-\u30FF (Katakana)
        # Korean (Hangul): \uAC00-\uD7AF
        # Devanagari (Hindi, etc.): \u0900-\u097F
        # Arabic: \u0600-\u06FF
        # Cyrillic: \u0400-\u04FF
        # This pattern will match any character within these specific ranges.
        foreign_character_pattern = r'[\u4E00-\u9FFF\uAC00-\uD7AF\u0900-\u097F\u0600-\u06FF\u0400-\u04FF]'
        mask_foreign_characters = df[konten_col_name].str.contains(foreign_character_pattern, regex=True, na=False)

        # Combine keyword exclusion mask AND foreign character exclusion mask using logical OR
        # A row is excluded if it matches a keyword OR contains foreign language characters.
        combined_mask_exclude = combined_mask_exclude | mask_foreign_characters

        # Create two DataFrames based on the combined exclusion mask:
        # 1. Cleaned data: rows where the combined_mask_exclude is False (do NOT contain any keywords or foreign language characters)
        df_cleaned = df[~combined_mask_exclude]
        # 2. Excluded data: rows where the combined_mask_exclude is True (DO contain any keywords or foreign language characters)
        df_excluded = df[combined_mask_exclude]

        # --- Save the DataFrames to output files, preserving other sheets ---
        for output_df, output_filepath, file_type in [(df_cleaned, cleaned_output_filepath, "Cleaned data"), (df_excluded, excluded_output_filepath, "Excluded items")]:
            # Always start with a fresh workbook for each output file
            # This simplifies the logic of copying sheets
            output_workbook = openpyxl.Workbook()
            
            # Remove the default 'Sheet' created by openpyxl
            if 'Sheet' in output_workbook.sheetnames:
                output_workbook.remove(output_workbook['Sheet'])

            # Copy all sheets from the original workbook EXCEPT the input sheet
            # and ensure they are visible in the output workbook
            for sheet_name in original_workbook.sheetnames:
                if sheet_name != input_sheet_name: # Do not copy the input sheet itself
                    original_sheet = original_workbook[sheet_name]
                    new_ws = output_workbook.create_sheet(title=sheet_name)
                    new_ws.sheet_state = 'visible' # Ensure copied sheets are visible

                    # Copy all cells, including formatting
                    for row in original_sheet.iter_rows():
                        for cell in row:
                            new_cell = new_ws[cell.coordinate]
                            new_cell.value = cell.value
                            if cell.has_style:
                                new_cell.font = cell.font.copy()
                                new_cell.fill = cell.fill.copy()
                                new_cell.border = cell.border.copy()
                                new_cell.alignment = cell.alignment.copy()
                                new_cell.number_format = cell.number_format
                                new_cell.protection = cell.protection.copy()

                    # Copy column dimensions
                    for col_dim in original_sheet.column_dimensions.values():
                        new_ws.column_dimensions[col_dim.index].width = col_dim.width
                    # Copy row dimensions
                    for row_dim in original_sheet.row_dimensions.values():
                        new_ws.row_dimensions[row_dim.index].height = row_dim.height
            
            # Create or replace the processed data sheet
            if output_sheet_name in output_workbook.sheetnames:
                output_workbook.remove(output_workbook[output_sheet_name])
            
            processed_ws = output_workbook.create_sheet(title=output_sheet_name, index=0) # Add at the beginning
            processed_ws.sheet_state = 'visible' # Ensure processed sheet is visible

            # Write the DataFrame to the new sheet
            for r_idx, row in enumerate(dataframe_to_rows(output_df, index=False, header=True), 1):
                processed_ws.append(row)
            
            # Auto-fit column widths for the processed sheet for better readability
            for column in processed_ws.columns:
                max_length = 0
                column_name = column[0].column_letter # Get the column name
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                processed_ws.column_dimensions[column_name].width = adjusted_width

            output_workbook.save(output_filepath)
            print(f"{file_type} saved to {output_filepath} in sheet '{output_sheet_name}' and other sheets preserved.")

    except FileNotFoundError:
        raise FileNotFoundError(f"Input file not found at {input_filepath}")
    except ValueError as e:
        # Catch specific ValueErrors (e.g., sheet not found)
        raise e
    except Exception as e:
        # Catch any other unexpected errors during pandas/openpyxl operations
        raise Exception(f"An error occurred during data processing: {e}")
    finally:
        # --- Delete the original uploaded input file and the temporary file immediately after processing ---
        if os.path.exists(input_filepath):
            try:
                time.sleep(0.1) # Small delay to ensure file handles are released
                os.remove(input_filepath)
                print(f"Deleted uploaded temporary file: {input_filepath}")
            except Exception as e:
                print(f"Error deleting input file {input_filepath}: {e}")
        if os.path.exists(temp_input_filepath):
            try:
                time.sleep(0.1) # Small delay
                os.remove(temp_input_filepath)
                print(f"Deleted temporary working file: {temp_input_filepath}")
            except Exception as e:
                print(f"Error deleting temporary file {temp_input_filepath}: {e}")