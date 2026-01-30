"""Data parsing and formatting utilities for Excel I/O."""

import re
import pandas as pd
from typing import List, Dict, Any, Optional


def parse_numeric_value(value) -> float:
    """
    Parse numeric value from Excel cell.

    Args:
        value: Value from Excel cell (can be string, number, or None)

    Returns:
        float: Numeric value or 0.0 if parsing fails
    """
    if not value or pd.isna(value) or str(value).strip() == '' or str(value) == 'nan':
        return 0.0

    try:
        # Replace comma with dot and convert to float
        numeric_str = str(value).replace(',', '.').strip()
        return float(numeric_str)
    except (ValueError, TypeError):
        return 0.0


def parse_share_from_excel(share_str) -> float:
    """
    Parse share value from Excel (handles comma/dot decimal separators).

    Args:
        share_str: Share string from Excel

    Returns:
        float: Share value (0.0 to 1.0)
    """
    if not share_str:
        return 0.0
    try:
        normalized_str = str(share_str).replace(',', '.')
        return float(normalized_str)
    except (ValueError, TypeError):
        return 0.0


def format_share_for_excel(share: float) -> str:
    """
    Format share value for Excel storage.

    Args:
        share: Share value (0.0 to 1.0)

    Returns:
        str: Formatted share string (uses comma as decimal separator)
    """
    if abs(share - 1.0) < 0.0001:
        return "1"
    else:
        return str(share).replace('.', ',')


def parse_prg_bindings(binding_string: str) -> List[Dict[str, Any]]:
    """
    Parse PRG binding string from consumer code column.

    Format: "PRG_ID1|share1|GRS_Name1;PRG_ID2|share2|GRS_Name2"

    Args:
        binding_string: Semicolon-separated binding string

    Returns:
        List of binding dictionaries with keys: prg_id, share, grs_name
    """
    if not binding_string or binding_string.strip() == '':
        return []

    bindings = []
    parts = binding_string.split(';')

    for part in parts:
        part = part.strip()
        if not part:
            continue

        components = part.split('|')
        if len(components) >= 3:
            try:
                prg_id = components[0].strip()
                share_str = components[1].strip()
                grs_name = '|'.join(components[2:]).strip()  # Handle pipes in GRS name

                share = parse_share_from_excel(share_str)

                bindings.append({
                    'prg_id': prg_id,
                    'share': share,
                    'grs_name': grs_name
                })
            except (ValueError, IndexError) as e:
                print(f"[WARNING] Failed to parse binding: {part} - {e}")
                continue

    return bindings


def format_prg_bindings(bindings: List[Dict[str, Any]]) -> str:
    """
    Format list of bindings to Excel string.

    Args:
        bindings: List of binding dictionaries

    Returns:
        str: Semicolon-separated binding string
    """
    if not bindings:
        return ''

    formatted_parts = []
    for binding in bindings:
        share_str = format_share_for_excel(binding['share'])
        formatted_parts.append(f"{binding['prg_id']}|{share_str}|{binding['grs_name']}")

    return ';'.join(formatted_parts)


def calculate_total_share(bindings: List[Dict[str, Any]]) -> float:
    """
    Calculate total share from list of bindings.

    Args:
        bindings: List of binding dictionaries

    Returns:
        float: Sum of all shares
    """
    return sum(binding['share'] for binding in bindings)


def parse_grs_id_column(grs_id_value) -> Optional[str]:
    """
    Parse GRS ID from Excel column (extracts first non-zero number).

    Args:
        grs_id_value: Value from GRS ID column

    Returns:
        str: GRS ID or None if not found
    """
    if not grs_id_value or pd.isna(grs_id_value):
        return None

    grs_str = str(grs_id_value).strip()
    if not grs_str or grs_str == 'nan':
        return None

    # Extract all numbers and return first non-zero
    numbers = re.findall(r'\d+', grs_str)
    for num_str in numbers:
        try:
            num = int(num_str)
            if num != 0:
                return str(num)
        except ValueError:
            continue

    return None


def extract_grs_name_from_id(grs_id_value: str) -> str:
    """
    Extract GRS name from GRS ID value.

    Format: "ГРС Название_ГРС"

    Args:
        grs_id_value: GRS ID value from Excel

    Returns:
        str: GRS name (empty if not found)
    """
    if not grs_id_value:
        return ""

    value = grs_id_value.strip()

    # Look for pattern "ГРС " (with space)
    if value.lower().startswith('грс '):
        return value[4:].strip()  # Return everything after "ГРС "

    return ""


def extract_grs_name_from_code(code_value: str) -> str:
    """
    Extract GRS name from consumer code column.

    Format: "PRG_ID|share|ГРС Название_ГРС" or multiple bindings

    Args:
        code_value: Code value from consumer binding column

    Returns:
        str: First GRS name found (empty if none)
    """
    if not code_value:
        return ""

    # Split by semicolons (multiple bindings possible)
    bindings = code_value.split(';')

    for binding in bindings:
        binding = binding.strip()
        if not binding:
            continue

        # Split by pipe
        parts = binding.split('|')
        if len(parts) >= 3:
            # Third component should be GRS name
            grs_part = '|'.join(parts[2:]).strip()
            if grs_part.lower().startswith('грс '):
                return grs_part[4:].strip()
            return grs_part

    return ""


def normalize_string(value) -> str:
    """
    Normalize a string value from Excel.

    Args:
        value: Any value from Excel cell

    Returns:
        str: Normalized string (empty if invalid)
    """
    if value is None or str(value).strip() == '' or str(value) == 'nan' or pd.isna(value):
        return ""
    return str(value).strip()
