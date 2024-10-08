# named_ranges.py

import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from utils import parse_cell

def add_named_ranges(wb, ws, cell_range, search_columns, prefix=None):
    """
    Adds named ranges based on the provided cell range and search columns.
    If a prefix is provided, it will be used for the named ranges and tax code validation will be applied.
    If no prefix is provided, any non-empty cell value will be used for the named range.
    
    Parameters:
    - wb: openpyxl Workbook object
    - ws: openpyxl Worksheet object
    - cell_range: String representing the cell range (e.g., 'L200:L408')
    - search_columns: List of column letters to search for values (e.g., ['J', 'K'])
    - prefix: Optional string prefix for the named ranges
    """
    # Extract start and end cells
    try:
        start_cell, end_cell = cell_range.split(':')
    except ValueError:
        print(f"Error: Invalid cell range format '{cell_range}'. Please use format like 'L200:L408'.")
        return

    start_col_letter, start_row = parse_cell(start_cell)
    end_col_letter, end_row = parse_cell(end_cell)

    # Ensure target_column is correct (should always be the same as start_col_letter)
    target_column = start_col_letter

    # Debugging: Print parsed cell references
    print(f"\nProcessing Range:")
    print(f"Start Column: {start_col_letter}, Start Row: {start_row}")
    print(f"End Column: {end_col_letter}, End Row: {end_row}")
    print(f"Target Column: {target_column}")

    # Iterate over the range and find values in the specified columns
    for row in range(start_row, end_row + 1):
        value = None

        # Loop through each specified column to search for a value
        for col_letter in search_columns:
            cell_address = f'{col_letter}{row}'
            cell_value = ws[cell_address].value

            # Debugging: Uncomment the next line to see which cells are being checked
            print(f"Checking cell {cell_address} with value: {cell_value}")

            # Handle possible empty cells
            if cell_value is None:
                continue

            if prefix:
                # If prefix is provided, apply tax code validation
                if isinstance(cell_value, str) and cell_value.isdigit() and len(cell_value) == 4:
                    value = cell_value
                    break  # Found a valid tax code
                elif isinstance(cell_value, int) and 1000 <= cell_value <= 9999:
                    value = str(cell_value)
                    break  # Found a valid tax code
            else:
                # If no prefix, use any non-empty value
                value = str(cell_value).strip()
                if value:
                    break  # Found a non-empty value

        # If no valid value is found, skip this row
        if value is None:
            continue

        # Check if the target cell has a valid value
        target_cell_address = f'{target_column}{row}'
        target_cell_value = ws[target_cell_address].value
        if target_cell_value is None:
            continue  # Skip if the target cell is empty

        # Create a named range for the target cell
        # Ensure sheet name is properly quoted (handles single quotes in sheet name)
        sheet_name_quoted = ws.title.replace("'", "''")
        target_cell_ref = f"'{sheet_name_quoted}'!${target_column}${row}"
        
        if prefix:
            named_range = f"{prefix}{value}"
        else:
            named_range = value

        # Debugging: Print the named range details
        print(f"Creating named range '{named_range}' referring to '{target_cell_ref}'.")

        # Remove existing named range if it exists
        if named_range in wb.defined_names:
            del wb.defined_names[named_range]
            print(f"Removed existing named range '{named_range}'.")

        # Add the named range using DefinedName and append to defined_names
        try:
            new_defined_name = DefinedName(name=named_range, attr_text=target_cell_ref)
            wb.defined_names[named_range] = new_defined_name  
            print(f"Named range '{named_range}' created successfully.")
        except Exception as e:
            print(f"Error creating named range '{named_range}': {e}")
            continue