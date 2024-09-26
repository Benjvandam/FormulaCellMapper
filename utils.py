# utils.py

def get_user_input(prompt, default):
    """Helper function to get user input with a default value."""
    user_input = input(f"{prompt} (default: '{default}'): ")
    return user_input.strip() or default

def parse_cell(cell):
    """Parses a cell reference into column letter and row number."""
    col_letter = ''.join(filter(str.isalpha, cell)).upper()
    row_number = int(''.join(filter(str.isdigit, cell)))
    return col_letter, row_number


# utils.py

def get_user_input(prompt, default=""):
    """
    Prompts the user for input, displaying a default value.

    Parameters:
    - prompt: The message displayed to the user.
    - default: The default value if the user provides no input.

    Returns:
    - The user's input or the default value.
    """
    if default:
        user_input = input(f"{prompt} (default: '{default}'): ").strip()
        return user_input if user_input else default
    else:
        return input(f"{prompt}: ").strip()
