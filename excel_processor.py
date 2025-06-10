import pandas as pd
import os

def process_data_excel(input_filepath, cleaned_output_filepath, excluded_output_filepath, keywords_list=None):
    
    if keywords_list is None or not keywords_list:
        keywords_list = [] # Default keywords

    try:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(input_filepath)

        # --- Debugging: Print all columns found in the Excel file ---
        print(f"Columns found in {os.path.basename(input_filepath)}: {df.columns.tolist()}")

        # Try to find the 'KONTEN' column robustly
        konten_col_name = None
        for col in df.columns:
            # Normalize column name for comparison (strip spaces, convert to uppercase)
            normalized_col = str(col).strip().upper()
            if normalized_col == 'KONTEN':
                konten_col_name = col
                break # Found exact or normalized match

        if konten_col_name is None:
            # If 'KONTEN' column is still not found after robust search
            raise ValueError(
                f"'KONTEN' column not found in {os.path.basename(input_filepath)}. "
                f"Available columns are: {df.columns.tolist()}"
            )

        # Convert the identified 'KONTEN' column to string type and fill NaN values
        df[konten_col_name] = df[konten_col_name].astype(str).fillna('')

        # Create a combined boolean mask for rows that contain ANY of the keywords
        # Initialize an empty mask that is all False
        combined_mask_exclude = pd.Series([False] * len(df), index=df.index)

        for keyword in keywords_list:
            # For each keyword, create a mask and combine it with the existing mask using OR
            combined_mask_exclude = combined_mask_exclude | \
                                    df[konten_col_name].str.contains(str(keyword).strip(), case=False, na=False)

        # Create two DataFrames:
        # 1. Cleaned data: rows where the combined_mask_exclude is False (do NOT contain any keywords)
        df_cleaned = df[~combined_mask_exclude]
        # 2. Excluded data: rows where the combined_mask_exclude is True (DO contain any keywords)
        df_excluded = df[combined_mask_exclude]

        # Save the cleaned DataFrame
        df_cleaned.to_excel(cleaned_output_filepath, index=False)
        print(f"Cleaned data saved to {cleaned_output_filepath}")

        # Save the excluded DataFrame
        df_excluded.to_excel(excluded_output_filepath, index=False)
        print(f"Excluded items saved to {excluded_output_filepath}")

    except FileNotFoundError:
        raise FileNotFoundError(f"Input file not found at {input_filepath}")
    except pd.errors.EmptyDataError:
        raise ValueError(f"No columns to parse from file {input_filepath}. Is it empty?")
    except Exception as e:
        # Catch any other unexpected errors during pandas operations
        raise Exception(f"An error occurred during data processing: {e}")

# This part is typically not run when imported by app.py,
# but can be useful for standalone testing of the processor script.
if __name__ == "__main__":
    print("Running excel_processor.py in standalone test mode.")
    # Example usage for testing purposes (replace with actual file paths for testing)
    # create dummy excel file for testing
    dummy_data = {
        'UUID': [1, 2, 3, 4, 5, 6, 7],
        'KONTEN': [
            'This is some content.',
            'Gopay transaction details.',
            'Selling an item (dijual cepat).',
            'Another normal entry.',
            'Limited time promo, grab it now!',
            'Check out our latest product.',
            'Digital wallet gopay'
        ]
    }
    dummy_df = pd.DataFrame(dummy_data)
    test_input_file = 'test_data_dynamic.xlsx'
    dummy_df.to_excel(test_input_file, index=False)

    test_cleaned_output = 'test_cleaned_data_dynamic.xlsx'
    test_excluded_output = 'test_excluded_items_dynamic.xlsx'

    # Test with custom keywords
    test_keywords_custom = ['promo', 'digital wallet']
    try:
        print(f"\n--- Testing with custom keywords: {test_keywords_custom} ---")
        process_data_excel(test_input_file, test_cleaned_output, test_excluded_output, test_keywords_custom)
        print("Standalone test with custom keywords completed successfully.")
    except Exception as e:
        print(f"Standalone test with custom keywords failed: {e}")

    # Test with default keywords (if the function were called without providing a list)
    test_keywords_default = ['gopay', 'dijual'] # Explicitly setting defaults for test clarity
    test_cleaned_output_default = 'test_cleaned_data_default.xlsx'
    test_excluded_output_default = 'test_excluded_items_default.xlsx'
    try:
        print(f"\n--- Testing with default keywords: {test_keywords_default} ---")
        process_data_excel(test_input_file, test_cleaned_output_default, test_excluded_output_default, test_keywords_default)
        print("Standalone test with default keywords completed successfully.")
    except Exception as e:
        print(f"Standalone test with default keywords failed: {e}")

    # Clean up dummy file
    if os.path.exists(test_input_file):
        os.remove(test_input_file)
    if os.path.exists(test_cleaned_output):
        os.remove(test_cleaned_output)
    if os.path.exists(test_excluded_output):
        os.remove(test_excluded_output)
    if os.path.exists(test_cleaned_output_default):
        os.remove(test_cleaned_output_default)
    if os.path.exists(test_excluded_output_default):
        os.remove(test_excluded_output_default)
