"""Excel utility functions for column conversion and formatting."""


def col_to_index(col_ref):
    """
    Convert Excel column reference to zero-based index.

    Args:
        col_ref: Column reference ('A', 'B', 'AB', etc.) or numeric string

    Returns:
        int: Zero-based column index (A=0, B=1, etc.)

    Examples:
        >>> col_to_index('A')
        0
        >>> col_to_index('Z')
        25
        >>> col_to_index('AA')
        26
    """
    if not col_ref:
        return 0

    col_ref = str(col_ref).strip().upper()

    # Handle numeric input
    if col_ref.isdigit():
        return int(col_ref) - 1

    # Convert letters to index
    result = 0
    for char in col_ref:
        if 'A' <= char <= 'Z':
            result = result * 26 + (ord(char) - ord('A') + 1)
    return result - 1


def index_to_col(index):
    """
    Convert zero-based column index to Excel column letter.

    Args:
        index: Zero-based column index

    Returns:
        str: Excel column letter (0='A', 1='B', 26='AA', etc.)

    Examples:
        >>> index_to_col(0)
        'A'
        >>> index_to_col(25)
        'Z'
        >>> index_to_col(26)
        'AA'
    """
    if index < 0:
        return 'A'

    col = ''
    index += 1  # Convert from 0-based to 1-based

    while index > 0:
        index -= 1
        col = chr(index % 26 + ord('A')) + col
        index //= 26

    return col if col else 'A'


def format_share_for_excel(share):
    """
    Format a share value for Excel storage.

    Args:
        share: Float share value (0.0 to 1.0)

    Returns:
        str: Formatted share string
    """
    if share is None:
        return ""
    try:
        return str(float(share))
    except (ValueError, TypeError):
        return ""


def parse_share_from_excel(share_str):
    """
    Parse share value from Excel cell.

    Args:
        share_str: String representation of share from Excel

    Returns:
        float: Share value, or 0.0 if invalid
    """
    if not share_str or str(share_str).strip() == '':
        return 0.0

    try:
        # Replace comma with dot and parse
        numeric_str = str(share_str).replace(',', '.').strip()
        return float(numeric_str)
    except (ValueError, TypeError):
        return 0.0
