import os
import time
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection
from openpyxl.worksheet.hyperlink import Hyperlink

def copy_cell_properties(source_cell, target_cell):
    """Copies all relevant properties from a source cell to a target cell."""
    target_cell.value = source_cell.value
    if source_cell.font: target_cell.font = source_cell.font.copy()
    if source_cell.fill: target_cell.fill = source_cell.fill.copy()
    if source_cell.border: target_cell.border = source_cell.border.copy()
    if source_cell.alignment: target_cell.alignment = source_cell.alignment.copy()
    target_cell.number_format = source_cell.number_format
    if source_cell.protection: target_cell.protection = source_cell.protection.copy()
    if source_cell.hyperlink:
        target_cell.hyperlink = Hyperlink(ref=source_cell.hyperlink.ref,
                                          target=source_cell.hyperlink.target,
                                          tooltip=source_cell.hyperlink.tooltip,
                                          display=source_cell.hyperlink.display)

def process_data_excel(input_filepath, cleaned_output_filepath, excluded_output_filepath, keywords_list=None, input_sheet_name='Sheet1', output_sheet_name='Processed Data'):

    if keywords_list is None:
        keywords_list = []

    original_workbook = None # Initialize to None
    try:
        # Load the original workbook using openpyxl
        original_workbook = openpyxl.load_workbook(input_filepath)

        all_sheets_hidden = all(sheet.sheet_state == 'hidden' or sheet.sheet_state == 'veryHidden' for sheet in original_workbook._sheets)

        if all_sheets_hidden:
            if input_sheet_name not in original_workbook.sheetnames:
                raise ValueError(f"Input sheet '{input_sheet_name}' not found in the Excel file, and all sheets are hidden.")
            original_workbook[input_sheet_name].sheet_state = 'visible'
            print(f"Temporarily unhid sheet '{input_sheet_name}' as all sheets were hidden.")

        try:
            input_sheet = original_workbook[input_sheet_name]
        except KeyError:
            raise ValueError(f"Input sheet '{input_sheet_name}' not found in the Excel file.")

        # Identify 'KONTEN' and 'UUID' columns dynamically from the first row
        header_row_cells = [cell for cell in input_sheet[1]] # Get the actual cell objects for the header row
        konten_col_index = -1
        uuid_col_index = -1 # Index for the 'UUID' column

        for idx, cell in enumerate(header_row_cells):
            col_val = str(cell.value or '').strip().upper()
            if col_val == 'KONTEN':
                konten_col_index = idx
            elif col_val == 'UUID':
                uuid_col_index = idx # Found UUID column

        if konten_col_index == -1:
            raise ValueError(
                f"'KONTEN' column not found in sheet '{input_sheet_name}'. "
                f"Available columns (from first row) are: {[cell.value for cell in header_row_cells]}"
            )
        
        # Determine which column indices to actually copy (exclude UUID_col_index)
        columns_to_copy_indices = [idx for idx in range(len(header_row_cells)) if idx != uuid_col_index]

        # Create new workbooks
        cleaned_workbook = openpyxl.Workbook()
        excluded_workbook = openpyxl.Workbook()

        if 'Sheet' in cleaned_workbook.sheetnames: cleaned_workbook.remove(cleaned_workbook['Sheet'])
        if 'Sheet' in excluded_workbook.sheetnames: excluded_workbook.remove(excluded_workbook['Sheet'])

        cleaned_ws = cleaned_workbook.create_sheet(title=output_sheet_name, index=0)
        excluded_ws = excluded_workbook.create_sheet(title=output_sheet_name, index=0)
        cleaned_ws.sheet_state = 'visible'
        excluded_ws.sheet_state = 'visible'

        # --- Copy HEADER ROW to both CLEANED and EXCLUDED sheets (SKIPPING UUID) ---
        current_row_in_cleaned = 1 # Start at row 1 for headers
        current_row_in_excluded = 1 # Start at row 1 for headers

        target_col_for_header_cleaned = 1
        target_col_for_header_excluded = 1

        for idx in columns_to_copy_indices:
            original_cell = header_row_cells[idx]
            
            # Copy to cleaned sheet
            new_cleaned_cell = cleaned_ws.cell(row=current_row_in_cleaned, column=target_col_for_header_cleaned)
            copy_cell_properties(original_cell, new_cleaned_cell)
            target_col_for_header_cleaned += 1

            # Copy to excluded sheet
            new_excluded_cell = excluded_ws.cell(row=current_row_in_excluded, column=target_col_for_header_excluded)
            copy_cell_properties(original_cell, new_excluded_cell)
            target_col_for_header_excluded += 1
        
        # Copy header row dimension
        original_header_row_dim = input_sheet.row_dimensions[1]
        cleaned_ws.row_dimensions[current_row_in_cleaned].height = original_header_row_dim.height
        excluded_ws.row_dimensions[current_row_in_excluded].height = original_header_row_dim.height


        # Foreign character pattern
        foreign_character_pattern = r'[\u4E00-\u9FFF\uAC00-\uD7AF\u0900-\u097F\u0600-\u06FF\u0400-\u04FF]'

        # Iterate through DATA ROWS of the input sheet (starting from row 2, skipping the header)
        for r_idx, row_cells in enumerate(input_sheet.iter_rows(min_row=2), 2):
            if konten_col_index >= len(row_cells): # Skip if row is too short to contain KONTEN
                continue

            konten_cell = row_cells[konten_col_index]
            konten_value = str(konten_cell.value or '').strip()

            exclude_row = False
            for keyword in keywords_list:
                if str(keyword).strip().lower() in konten_value.lower():
                    exclude_row = True
                    break

            if not exclude_row:
                if re.search(foreign_character_pattern, konten_value):
                    exclude_row = True

            target_worksheets_for_data = []
            if exclude_row:
                target_worksheets_for_data.append(excluded_ws)
            else:
                target_worksheets_for_data.append(cleaned_ws)

            # Copy cells to the target sheets for data rows
            for target_ws in target_worksheets_for_data:
                current_target_row = target_ws.max_row + 1
                target_col_for_data = 1 # Start from column 1 in the new sheet

                for idx in columns_to_copy_indices: # Only iterate over columns we intend to copy
                    original_cell = row_cells[idx]
                    new_cell = target_ws.cell(row=current_target_row, column=target_col_for_data)
                    copy_cell_properties(original_cell, new_cell)
                    target_col_for_data += 1 # Increment only if cell was copied

                # Copy row dimension for data rows
                original_data_row_dim = input_sheet.row_dimensions[r_idx]
                target_ws.row_dimensions[current_target_row].height = original_data_row_dim.height

        # Copy column dimensions (width) from input sheet to the primary processed sheets (SKIPPING UUID)
        for col_dim in input_sheet.column_dimensions.values():
            original_col_idx = openpyxl.utils.column_index_from_string(col_dim.index) - 1 # 0-indexed
            if col_dim.width and original_col_idx != uuid_col_index:
                # Find the corresponding target column letter
                # This assumes columns are contiguous after UUID removal.
                # If UUID was at index 0, then original index 1 maps to target index 0.
                # So target column letter is based on original_col_idx shifted if UUID was before it.
                target_column_letter = openpyxl.utils.get_column_letter(
                    len([c for c_idx, c in enumerate(header_row_cells) if c_idx < original_col_idx and c_idx != uuid_col_index]) + 1
                )
                cleaned_ws.column_dimensions[target_column_letter].width = col_dim.width
                excluded_ws.column_dimensions[target_column_letter].width = col_dim.width


        # Copy all other sheets from the original workbook to both output workbooks (SKIPPING UUID)
        for workbook in [cleaned_workbook, excluded_workbook]:
            for sheet_name in original_workbook.sheetnames:
                # Do not copy the input sheet itself, as its data is handled by the row-by-row logic above
                if sheet_name != input_sheet_name:
                    original_sheet = original_workbook[sheet_name]
                    new_ws = workbook.create_sheet(title=sheet_name)
                    new_ws.sheet_state = original_sheet.sheet_state # Preserve hidden/visible state

                    # Copy cells and properties for other sheets
                    for r_idx_other, row in enumerate(original_sheet.iter_rows(values_only=False), 1):
                        target_other_sheet_col_idx = 1 # Reset column for each row
                        for c_idx, cell in enumerate(row):
                            if c_idx == uuid_col_index: # Skip the UUID column if present in other sheets
                                continue
                            
                            new_cell = new_ws.cell(row=r_idx_other, column=target_other_sheet_col_idx)
                            copy_cell_properties(cell, new_cell)
                            target_other_sheet_col_idx += 1

                    # Copy column dimensions for other sheets (and skip UUID if identified)
                    for col_dim in original_sheet.column_dimensions.values():
                        original_col_idx = openpyxl.utils.column_index_from_string(col_dim.index) - 1
                        if col_dim.width and original_col_idx != uuid_col_index:
                            target_column_letter = openpyxl.utils.get_column_letter(
                                len([c for c_idx, c in enumerate(header_row_cells) if c_idx < original_col_idx and c_idx != uuid_col_index]) + 1
                            )
                            new_ws.column_dimensions[target_column_letter].width = col_dim.width
                    # Copy row dimensions for other sheets
                    for row_dim in original_sheet.row_dimensions.values():
                        if row_dim.height:
                            new_ws.row_dimensions[row_dim.index].height = row_dim.height


        # Auto-fit column widths for the main processed sheets for better readability
        for ws in [cleaned_ws, excluded_ws]:
            for col_idx_new, column in enumerate(ws.columns):
                max_length = 0
                for cell in column:
                    try:
                        if cell.value is not None:
                            current_length = len(str(cell.value))
                            if current_length > max_length:
                                max_length = current_length
                    except TypeError:
                        pass
                column_letter = openpyxl.utils.get_column_letter(col_idx_new + 1)
                adjusted_width = min((max_length + 2), 75)
                if adjusted_width > 0:
                    ws.column_dimensions[column_letter].width = adjusted_width

        # Save the workbooks
        cleaned_workbook.save(cleaned_output_filepath)
        excluded_workbook.save(excluded_output_filepath)

        print(f"Cleaned data saved to {cleaned_output_filepath} in sheet '{output_sheet_name}' and other sheets preserved.")
        print(f"Excluded items saved to {excluded_output_filepath} in sheet '{output_sheet_name}' and other sheets preserved.")

    except FileNotFoundError:
        raise FileNotFoundError(f"Input file not found at {input_filepath}")
    except ValueError as e:
        raise e
    except Exception as e:
        raise Exception(f"An error occurred during data processing: {e}")
    finally:
        # Close the original workbook to release memory
        if original_workbook:
            try:
                original_workbook.close()
                print(f"Closed original workbook to free memory.")
            except Exception as e:
                print(f"Error closing original workbook: {e}")
        
        time.sleep(0.1) # Small delay to help ensure file handles are released
        if os.path.exists(input_filepath):
            try:
                os.remove(input_filepath)
                print(f"Deleted uploaded input file: {input_filepath}")
            except Exception as e:
                print(f"Error deleting input file {input_filepath}: {e}")
