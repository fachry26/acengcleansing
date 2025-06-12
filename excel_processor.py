import os
import time
import re
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection # Still import, but not used for copying styles
from openpyxl.worksheet.hyperlink import Hyperlink

def copy_cell_properties(source_cell, target_cell):
    """
    Copies essential properties from a source cell to a target cell.
    This version is designed for source_cell coming from a read_only workbook,
    prioritizing robustness for large files.
    """
    target_cell.value = source_cell.value
    
    # Attempt to copy number format, but handle potential errors gracefully.
    try:
        if source_cell.number_format: # Check if there's a format to copy
            target_cell.number_format = source_cell.number_format
    except Exception as e:
        # If number_format causes an error (e.g., invalid format code), skip it.
        # The cell will default to 'General' format in the output.
        print(f"Warning: Could not copy number_format for cell {source_cell.coordinate}: {e}")
        pass # Continue without copying the problematic format

    # Hyperlinks ARE generally available and copyable in read_only mode
    if hasattr(source_cell, 'hyperlink') and source_cell.hyperlink:
        target_cell.hyperlink = Hyperlink(ref=source_cell.hyperlink.ref,
                                          target=source_cell.hyperlink.target,
                                          tooltip=source_cell.hyperlink.tooltip,
                                          display=source_cell.hyperlink.display)

    # Detailed style properties (font, fill, border, alignment, protection) are NOT
    # loaded by openpyxl in read_only mode and are intentionally excluded for speed/stability.
    # Therefore, no attempts are made to copy them here.


def process_data_excel(input_filepath, cleaned_output_filepath, excluded_output_filepath, keywords_list=None, input_sheet_name='Sheet1', output_sheet_name='Processed Data'):

    if keywords_list is None:
        keywords_list = []

    original_workbook = None
    try:
        # Load the original workbook using openpyxl in read_only mode for performance
        original_workbook = openpyxl.load_workbook(input_filepath, read_only=True)

        # In read_only mode, sheet_state cannot be directly changed, but data should be accessible.
        all_sheets_hidden = all(sheet.sheet_state == 'hidden' or sheet.sheet_state == 'veryHidden' for sheet in original_workbook._sheets)
        if all_sheets_hidden:
            if input_sheet_name not in original_workbook.sheetnames:
                raise ValueError(f"Input sheet '{input_sheet_name}' not found and all sheets are hidden.")
            print(f"Input workbook loaded in read_only mode. If '{input_sheet_name}' was hidden, its data should still be accessible.")

        try:
            input_sheet = original_workbook[input_sheet_name]
        except KeyError:
            raise ValueError(f"Input sheet '{input_sheet_name}' not found in the Excel file.")

        # Read the first row to get headers (iter_rows is key for read_only)
        first_row_generator = input_sheet.iter_rows(min_row=1, max_row=1)
        header_row_cells = []
        try:
            header_row_cells = [cell for cell in next(first_row_generator)]
        except StopIteration:
            raise ValueError(f"Input sheet '{input_sheet_name}' is empty.")
            
        konten_col_index = -1
        # UUID column will now be copied, so we no longer need to find its index for exclusion.
        # uuid_col_index = -1 

        for idx, cell in enumerate(header_row_cells):
            col_val = str(cell.value or '').strip().upper()
            if col_val == 'KONTEN':
                konten_col_index = idx
            # No special handling for 'UUID' anymore, it will be copied by default.
            # elif col_val == 'UUID':
            #     uuid_col_index = idx

        if konten_col_index == -1:
            raise ValueError(
                f"'KONTEN' column not found in sheet '{input_sheet_name}'. "
                f"Available columns (from first row) are: {[cell.value for cell in header_row_cells]}"
            )
        
        # All columns will be copied by default, as UUID is no longer excluded.
        # columns_to_copy_indices = [idx for idx in range(len(header_row_cells)) if idx != uuid_col_index]
        columns_to_copy_indices = list(range(len(header_row_cells)))


        # Create new workbooks for output
        cleaned_workbook = openpyxl.Workbook()
        excluded_workbook = openpyxl.Workbook()

        if 'Sheet' in cleaned_workbook.sheetnames: cleaned_workbook.remove(cleaned_workbook['Sheet'])
        if 'Sheet' in excluded_workbook.sheetnames: excluded_workbook.remove(excluded_workbook['Sheet'])

        cleaned_ws = cleaned_workbook.create_sheet(title=output_sheet_name, index=0)
        excluded_ws = excluded_workbook.create_sheet(title=output_sheet_name, index=0)
        cleaned_ws.sheet_state = 'visible'
        excluded_ws.sheet_state = 'visible'

        # --- Copy HEADER ROW (UUID now included) ---
        current_row_in_cleaned = 1
        current_row_in_excluded = 1
        target_col_for_header_cleaned = 1
        target_col_for_header_excluded = 1

        for idx in columns_to_copy_indices: # Iterate through all columns
            original_cell = header_row_cells[idx]
            new_cleaned_cell = cleaned_ws.cell(row=current_row_in_cleaned, column=target_col_for_header_cleaned)
            copy_cell_properties(original_cell, new_cleaned_cell)
            target_col_for_header_cleaned += 1

            new_excluded_cell = excluded_ws.cell(row=current_row_in_excluded, column=target_col_for_header_excluded)
            copy_cell_properties(original_cell, new_excluded_cell)
            target_col_for_header_excluded += 1
        
        # Row and column dimensions are not available in read_only mode, so skipping copying them.


        # Foreign character pattern
        foreign_character_pattern = r'[\u4E00-\u9FFF\uAC00-\uD7AF\u0900-\u097F\u0600-\u06FF\u0400-\u04FF]'

        # Iterate through DATA ROWS (starting from row 2)
        for r_idx, row_cells in enumerate(input_sheet.iter_rows(min_row=2), 2):
            row_cells_list = list(row_cells) 

            if konten_col_index >= len(row_cells_list):
                continue

            konten_cell = row_cells_list[konten_col_index]
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
                target_col_for_data = 1

                for idx in columns_to_copy_indices: # Iterate through all columns, including where UUID was
                    original_cell = row_cells_list[idx]
                    new_cell = target_ws.cell(row=current_target_row, column=target_col_for_data)
                    copy_cell_properties(original_cell, new_cell)
                    target_col_for_data += 1


        # Copy all other sheets from the original workbook (UUID now included, no dimensions/styles)
        for workbook in [cleaned_workbook, excluded_workbook]:
            for sheet_name in original_workbook.sheetnames:
                if sheet_name != input_sheet_name:
                    original_sheet = original_workbook[sheet_name]
                    new_ws = workbook.create_sheet(title=sheet_name)
                    new_ws.sheet_state = original_sheet.sheet_state

                    for r_idx_other, row in enumerate(original_sheet.iter_rows(values_only=False), 1):
                        target_other_sheet_col_idx = 1
                        row_cells_other_list = list(row)
                        for c_idx, cell in enumerate(row_cells_other_list):
                            # UUID column is no longer specifically skipped
                            # if c_idx == uuid_col_index:
                            #    continue
                            new_cell = new_ws.cell(row=r_idx_other, column=target_other_sheet_col_idx)
                            copy_cell_properties(cell, new_cell)
                            target_other_sheet_col_idx += 1


        # Auto-fit column widths for the main processed sheets
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
        # Catch any other unexpected errors during data processing
        raise Exception(f"An error occurred during data processing: {e}")
    finally:
        if original_workbook:
            try:
                original_workbook.close()
                print(f"Closed original workbook to free memory.")
            except Exception as e:
                print(f"Error closing original workbook: {e}")
        
        time.sleep(0.1)
        if os.path.exists(input_filepath):
            try:
                os.remove(input_filepath)
                print(f"Deleted uploaded input file: {input_filepath}")
            except Exception as e:
                print(f"Error deleting input file {input_filepath}: {e}")
