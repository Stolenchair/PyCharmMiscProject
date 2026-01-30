"""Default configuration settings for PRG Pipeline Manager."""

from typing import Dict, Any


def get_default_settings() -> Dict[str, Dict[str, Any]]:
    """
    Get default column configuration settings.

    Returns:
        dict: Default settings for all table types (prg, grs, population, organizations)
    """
    return {
        'prg': {
            'sheet': '',
            'start_row': '10',
            'mo_col': 'A',
            'settlement_col': 'B',
            'prg_id_col': 'C',
            'grs_id_col': 'D',
            # Load columns (v7.4)
            'qy_pop_col': 'E',  # QY_pop - Population yearly volume
            'qh_pop_col': 'F',  # QH_pop - Population hourly rate
            'qy_ind_col': 'G',  # QY_ind - Industry yearly volume
            'qh_ind_col': 'H',  # QH_ind - Industry hourly rate
            'year_volume_col': 'I',  # Year_volume - Total yearly
            'max_hour_col': 'J'  # Max_hour - Total hourly max
        },
        'grs': {
            'sheet': '',
            'start_row': '10',
            'mo_col': 'A',
            'grs_id_col': 'B',
            'grs_name_col': 'C'
        },
        'population': {
            'sheet': '',
            'start_row': '10',
            'mo_col': 'A',
            'settlement_col': 'B',
            'code_col': 'M',  # PRG binding column
            'expenses_col': 'N',  # Yearly expenses
            'hourly_expenses_col': 'O'  # Hourly expenses (v7.4)
        },
        'organizations': {
            'sheet': '',
            'start_row': '10',
            'name_col': 'D',
            'mo_col': 'A',
            'settlement_col': 'B',
            'code_col': 'M',  # PRG binding column
            'expenses_col': 'N',  # Yearly expenses
            'hourly_expenses_col': 'O',  # Hourly expenses (v7.4)
            'grs_id_col': 'L'
        }
    }


# Field labels for UI display
FIELD_LABELS = {
    'prg': {
        'sheet': 'Название листа',
        'start_row': 'Начальная строка данных',
        'mo_col': 'Колонка МО (район)',
        'settlement_col': 'Колонка НП (населенный пункт)',
        'prg_id_col': 'Колонка ПРГ ID',
        'grs_id_col': 'Колонка ГРС ID',
        'qy_pop_col': 'QY_pop (годовой объем население)',
        'qh_pop_col': 'QH_pop (часовой расход население)',
        'qy_ind_col': 'QY_ind (годовой объем организации)',
        'qh_ind_col': 'QH_ind (часовой расход организации)',
        'year_volume_col': 'Year_volume (годовой объем всего)',
        'max_hour_col': 'Max_Hour (максимальный часовой расход)'
    },
    'grs': {
        'sheet': 'Название листа',
        'start_row': 'Начальная строка данных',
        'mo_col': 'Колонка МО (район)',
        'grs_id_col': 'Колонка ГРС ID',
        'grs_name_col': 'Колонка Название ГРС'
    },
    'population': {
        'sheet': 'Название листа',
        'start_row': 'Начальная строка данных',
        'mo_col': 'Колонка МО (район)',
        'settlement_col': 'Колонка НП',
        'code_col': 'Колонка Привязки ПРГ',
        'expenses_col': 'Колонка Годовые расходы',
        'hourly_expenses_col': 'Колонка Часовые расходы'
    },
    'organizations': {
        'sheet': 'Название листа',
        'start_row': 'Начальная строка данных',
        'name_col': 'Колонка Название организации',
        'mo_col': 'Колонка МО (район)',
        'settlement_col': 'Колонка НП',
        'code_col': 'Колонка Привязки ПРГ',
        'expenses_col': 'Колонка Годовые расходы',
        'hourly_expenses_col': 'Колонка Часовые расходы',
        'grs_id_col': 'Колонка ГРС ID'
    }
}
