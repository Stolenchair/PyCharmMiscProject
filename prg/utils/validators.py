"""Validation utility functions."""


def validate_share(share):
    """
    Validate that a share value is valid.

    Args:
        share: Share value to validate

    Returns:
        bool: True if valid share (0.0 to 1.0)
    """
    try:
        share_float = float(share)
        return 0.0 <= share_float <= 1.0
    except (ValueError, TypeError):
        return False


def validate_numeric(value):
    """
    Check if a value can be converted to a number.

    Args:
        value: Value to check

    Returns:
        bool: True if value is numeric
    """
    if value is None or str(value).strip() == '' or str(value) == 'nan':
        return False

    try:
        float(str(value).replace(',', '.').strip())
        return True
    except (ValueError, TypeError):
        return False
