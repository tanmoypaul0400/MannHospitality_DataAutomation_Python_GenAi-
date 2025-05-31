import pandas as pd
import os
import re

def excel_cell_to_indices(cell_ref):
    """
    Convert an Excel cell reference (e.g. 'A1') to zero-based row and column indices.
    Returns (row_index, col_index).
    """
    match = re.match(r"([A-Za-z]+)([0-9]+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    col_letters = match.group(1).upper()
    row_number = int(match.group(2))
    col_number = 0
    for char in col_letters:
        col_number = col_number * 26 + (ord(char) - ord('A') + 1)
    return row_number - 1, col_number - 1

def extract_cells_from_sheet(df, cells):
    """
    Extract specific cells or ranges from the DataFrame.
    Returns list of DataFrames with unique columns per extraction.
    """
    extracted_blocks = []
    for cell_ref in cells:
        cell_ref = cell_ref.strip()
        if ':' in cell_ref:  # Range e.g., 'A1:C3'
            start_cell, end_cell = cell_ref.split(':')
            start_row, start_col = excel_cell_to_indices(start_cell)
            end_row, end_col = excel_cell_to_indices(end_cell)
            row_start, row_end = min(start_row, end_row), max(start_row, end_row)
            col_start, col_end = min(start_col, end_col), max(start_col, end_col)
            block = df.iloc[row_start:row_end+1, col_start:col_end+1].copy().reset_index(drop=True)
            col_count = block.shape[1]
            col_names = []
            for i in range(col_count):
                col_names.append(f"{cell_ref}_Col{i+1}")
            block.columns = col_names
            extracted_blocks.append(block)
        else:  # Single cell e.g., 'A1'
            row, col = excel_cell_to_indices(cell_ref)
            try:
                value = df.iat[row, col]
            except IndexError:
                value = None
            block = pd.DataFrame({cell_ref: [value]})
            extracted_blocks.append(block)
    return extracted_blocks

def consolidate_cells_from_workbooks(folder_path, specific_cells, output_file):
    """
    Extract specified cells from specified sheets across multiple Excel workbooks.
    Consolidate into one Excel file avoiding pandas InvalidIndexError.
    """
    all_data = []

    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith(('~$', '.'))]
    if not files:
        print(f"No Excel files found in folder: {folder_path}")
        return

    for filename in files:
        filepath = os.path.join(folder_path, filename)
        print(f"Processing workbook: {filename}")
        try:
            xls = pd.ExcelFile(filepath)
        except Exception as e:
            print(f"Could not open {filename}: {e}")
            continue

        sheets_to_process = [s for s in specific_cells if s in xls.sheet_names]
        if not sheets_to_process:
            print(f"No matching sheets in {filename} for given specific_cells keys.")
            continue

        for sheet_name in sheets_to_process:
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            except Exception as e:
                print(f"Could not read sheet {sheet_name} in {filename}: {e}")
                continue

            extracted_blocks = extract_cells_from_sheet(df, specific_cells[sheet_name])

            for idx, block in enumerate(extracted_blocks):
                block.insert(0, 'Source Workbook', filename)
                block.insert(1, 'Source Sheet', sheet_name)
                block.insert(2, 'Extraction Block', idx + 1)  # Keep it for internal use
                all_data.append(block)

    if not all_data:
        print("No data extracted.")
        return

    try:
        consolidated_df = pd.concat(all_data, ignore_index=True)

        # Drop the 'Extraction Block' column to hide it from the output
        if 'Extraction Block' in consolidated_df.columns:
            consolidated_df = consolidated_df.drop(columns=['Extraction Block'])

        consolidated_df.to_excel(output_file, index=True)
        print(f"Data successfully consolidated into {output_file}")
    except Exception as e:
        print(f"Failed to write output file: {e}")

if __name__ == '__main__':
    folder_path = 'C:\\Invoice_Annexure'

    specific_cells = {
        'Payout Breakup': ['C4', 'D4', 'E4', 'F4'],
        'Summary': ['B5', 'B8', 'C12'],
        # add more sheets and their cell ranges/cells as needed
    }

    output_file = 'C:\\Invoice_Annexure\\Payment_breakup.xlsx'

    consolidate_cells_from_workbooks(folder_path, specific_cells, output_file)

