"""String utility functions."""


def normalize_string(value):
    """
    Normalize a string value from Excel.

    Args:
        value: Any value from Excel cell

    Returns:
        str: Normalized string, empty if invalid
    """
    if value is None or str(value).strip() == '' or str(value) == 'nan':
        return ""
    return str(value).strip()


def get_expenses_symbol(has_yearly, has_hourly):
    """
    Get visual symbol for expense status.

    Args:
        has_yearly: Whether consumer has yearly expenses
        has_hourly: Whether consumer has hourly expenses

    Returns:
        str: Symbol representing expense status
    """
    if not has_yearly and not has_hourly:
        return "🚫"
    return ""
