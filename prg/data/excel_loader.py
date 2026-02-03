"""Excel data loading operations."""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Any
from ..config import SettingsManager
from ..utils import col_to_index
from .parsers import parse_numeric_value, parse_grs_id_column, normalize_string


class ExcelLoader:
    """
    Loads data from Excel files into dictionaries.

    Handles loading of PRG, GRS, and consumer data based on configuration settings.
    """

    def __init__(self, settings_manager: SettingsManager):
        """
        Initialize Excel loader.

        Args:
            settings_manager: SettingsManager instance with column mappings
        """
        self.settings_manager = settings_manager

    def load_prg_data(self, excel_path: Path) -> List[Dict[str, Any]]:
        """
        Load PRG (pipeline) data from Excel.

        Args:
            excel_path: Path to Excel file

        Returns:
            List of PRG dictionaries with structure, loads, and Excel metadata

        Raises:
            Exception: If loading fails
        """
        try:
            settings = self.settings_manager.get_table_settings('prg')
            df = pd.read_excel(excel_path, sheet_name=settings['sheet'], header=None)
            start_row = int(settings['start_row']) - 1

            if start_row > 0:
                df = df.iloc[start_row:].reset_index(drop=True)

            # Get column indices
            mo_col = col_to_index(settings['mo_col'])
            settlement_col = col_to_index(settings['settlement_col'])
            prg_id_col = col_to_index(settings['prg_id_col'])
            grs_id_col = col_to_index(settings['grs_id_col'])

            # Load columns (v7.4)
            qy_pop_col = col_to_index(settings.get('qy_pop_col', 'E'))
            qh_pop_col = col_to_index(settings.get('qh_pop_col', 'F'))
            qy_ind_col = col_to_index(settings.get('qy_ind_col', 'G'))
            qh_ind_col = col_to_index(settings.get('qh_ind_col', 'H'))
            year_volume_col = col_to_index(settings.get('year_volume_col', 'I'))
            max_hour_col = col_to_index(settings.get('max_hour_col', 'J'))

            prg_data = []
            for idx, row in df.iterrows():
                try:
                    # Check required columns exist
                    if mo_col >= len(row) or settlement_col >= len(row) or \
                       prg_id_col >= len(row) or grs_id_col >= len(row):
                        continue

                    mo = normalize_string(row.iloc[mo_col])
                    settlement = normalize_string(row.iloc[settlement_col])
                    prg_id = normalize_string(row.iloc[prg_id_col])
                    grs_id_raw = row.iloc[grs_id_col] if pd.notna(row.iloc[grs_id_col]) else ""

                    grs_id = parse_grs_id_column(grs_id_raw)

                    # Load values from Excel
                    qy_pop = parse_numeric_value(row.iloc[qy_pop_col] if qy_pop_col < len(row) else "")
                    qh_pop = parse_numeric_value(row.iloc[qh_pop_col] if qh_pop_col < len(row) else "")
                    qy_ind = parse_numeric_value(row.iloc[qy_ind_col] if qy_ind_col < len(row) else "")
                    qh_ind = parse_numeric_value(row.iloc[qh_ind_col] if qh_ind_col < len(row) else "")
                    year_volume = parse_numeric_value(
                        row.iloc[year_volume_col] if year_volume_col < len(row) else "")
                    max_hour = parse_numeric_value(row.iloc[max_hour_col] if max_hour_col < len(row) else "")

                    if mo and settlement and prg_id and grs_id:
                        prg_data.append({
                            'id': f"prg_{idx}",
                            'mo': mo,
                            'settlement': settlement,
                            'prg_id': prg_id,
                            'grs_id': grs_id,
                            # Load values
                            'QY_pop': qy_pop,
                            'QH_pop': qh_pop,
                            'QY_ind': qy_ind,
                            'QH_ind': qh_ind,
                            'Year_volume': year_volume,
                            'Max_Hour': max_hour,
                            # Excel metadata for persistence
                            'sheet_name': settings['sheet'],
                            'excel_row': start_row + idx,
                            'qy_pop_col': qy_pop_col,
                            'qh_pop_col': qh_pop_col,
                            'qy_ind_col': qy_ind_col,
                            'qh_ind_col': qh_ind_col,
                            'year_volume_col': year_volume_col,
                            'max_hour_col': max_hour_col
                        })
                except Exception:
                    continue

            print(f"[OK] Loaded PRG: {len(prg_data)}")
            return prg_data

        except Exception as e:
            raise Exception(f"PRG loading error: {str(e)}")

    def load_grs_data(self, excel_path: Path) -> List[Dict[str, Any]]:
        """
        Load GRS (Gas Reduction Station) reference data from Excel.

        Args:
            excel_path: Path to Excel file

        Returns:
            List of GRS dictionaries

        Raises:
            Exception: If loading fails
        """
        try:
            settings = self.settings_manager.get_table_settings('grs')
            df = pd.read_excel(excel_path, sheet_name=settings['sheet'], header=None)
            start_row = int(settings['start_row']) - 1

            if start_row > 0:
                df = df.iloc[start_row:].reset_index(drop=True)

            mo_col = col_to_index(settings['mo_col'])
            grs_id_col = col_to_index(settings['grs_id_col'])
            grs_name_col = col_to_index(settings['grs_name_col'])

            grs_data = []
            for idx, row in df.iterrows():
                try:
                    if mo_col >= len(row) or grs_id_col >= len(row) or grs_name_col >= len(row):
                        continue

                    mo = normalize_string(row.iloc[mo_col])
                    grs_id = normalize_string(row.iloc[grs_id_col])
                    grs_name = normalize_string(row.iloc[grs_name_col])

                    if mo and grs_id and grs_name:
                        grs_data.append({
                            'id': f"grs_{idx}",
                            'mo': mo,
                            'grs_id': grs_id,
                            'grs_name': grs_name,
                            'sheet_name': settings['sheet'],
                            'excel_row': start_row + idx
                        })
                except Exception:
                    continue

            print(f"[OK] Loaded GRS: {len(grs_data)}")
            return grs_data

        except Exception as e:
            raise Exception(f"GRS loading error: {str(e)}")

    def load_population_data(self, excel_path: Path) -> List[Dict[str, Any]]:
        """
        Load population consumer data from Excel.

        Args:
            excel_path: Path to Excel file

        Returns:
            List of population consumer dictionaries

        Raises:
            Exception: If loading fails
        """
        try:
            settings = self.settings_manager.get_table_settings('population')
            df = pd.read_excel(excel_path, sheet_name=settings['sheet'], header=None)
            start_row = int(settings['start_row']) - 1

            if start_row > 0:
                df = df.iloc[start_row:].reset_index(drop=True)

            mo_col = col_to_index(settings['mo_col'])
            settlement_col = col_to_index(settings['settlement_col'])
            code_col = col_to_index(settings['code_col'])
            expenses_col = col_to_index(settings['expenses_col'])
            hourly_expenses_col = col_to_index(settings.get('hourly_expenses_col', 'O'))

            population_data = []
            for idx, row in df.iterrows():
                try:
                    if mo_col >= len(row) or settlement_col >= len(row):
                        continue

                    mo = normalize_string(row.iloc[mo_col])
                    settlement = normalize_string(row.iloc[settlement_col])
                    code = normalize_string(row.iloc[code_col]) if code_col < len(row) else ""

                    # Yearly expenses - parse as numeric
                    yearly_expenses = parse_numeric_value(
                        row.iloc[expenses_col] if expenses_col < len(row) else "")

                    # Hourly expenses (v7.4) - parse as numeric
                    hourly_expenses = parse_numeric_value(
                        row.iloc[hourly_expenses_col] if hourly_expenses_col < len(row) else "")

                    if mo and settlement:
                        population_data.append({
                            'id': f"pop_{settings['sheet']}_{start_row + idx}",
                            'type': 'Население',
                            'consumer_type': 'population',
                            'mo': mo,
                            'settlement': settlement,
                            'name': f"Население {settlement}",
                            'code': code if code else '',
                            'yearly_expenses': yearly_expenses,
                            'hourly_expenses': hourly_expenses,
                            'sheet_name': settings['sheet'],
                            'excel_row': start_row + idx,
                            'code_col': code_col,
                            'expenses_col': expenses_col,
                            'hourly_expenses_col': hourly_expenses_col
                        })
                except Exception:
                    continue

            print(f"[OK] Loaded population: {len(population_data)}")
            return population_data

        except Exception as e:
            raise Exception(f"Population loading error: {str(e)}")

    def load_organization_data(self, excel_path: Path) -> List[Dict[str, Any]]:
        """
        Load organization consumer data from Excel.

        Args:
            excel_path: Path to Excel file

        Returns:
            List of organization consumer dictionaries

        Raises:
            Exception: If loading fails
        """
        try:
            settings = self.settings_manager.get_table_settings('organizations')
            df = pd.read_excel(excel_path, sheet_name=settings['sheet'], header=None)
            start_row = int(settings['start_row']) - 1

            if start_row > 0:
                df = df.iloc[start_row:].reset_index(drop=True)

            name_col = col_to_index(settings['name_col'])
            mo_col = col_to_index(settings['mo_col'])
            settlement_col = col_to_index(settings['settlement_col'])
            code_col = col_to_index(settings['code_col'])
            expenses_col = col_to_index(settings['expenses_col'])
            hourly_expenses_col = col_to_index(settings.get('hourly_expenses_col', 'O'))
            grs_id_col = col_to_index(settings['grs_id_col'])

            organization_data = []
            for idx, row in df.iterrows():
                try:
                    if name_col >= len(row) or mo_col >= len(row) or settlement_col >= len(row):
                        continue

                    name = normalize_string(row.iloc[name_col])
                    mo = normalize_string(row.iloc[mo_col])
                    settlement = normalize_string(row.iloc[settlement_col])
                    code = normalize_string(row.iloc[code_col]) if code_col < len(row) else ""

                    # Yearly expenses - parse as numeric
                    yearly_expenses = parse_numeric_value(
                        row.iloc[expenses_col] if expenses_col < len(row) else "")

                    # Hourly expenses (v7.4) - parse as numeric
                    hourly_expenses = parse_numeric_value(
                        row.iloc[hourly_expenses_col] if hourly_expenses_col < len(row) else "")

                    grs_id = normalize_string(row.iloc[grs_id_col]) if grs_id_col < len(row) else ""

                    if name and mo and settlement:
                        organization_data.append({
                            'id': f"org_{settings['sheet']}_{start_row + idx}",
                            'type': 'Организация',
                            'consumer_type': 'organization',
                            'mo': mo,
                            'settlement': settlement,
                            'name': name,
                            'code': code if code else '',
                            'grs_id': grs_id if grs_id else '',
                            'grs_id_col': grs_id_col,
                            'yearly_expenses': yearly_expenses,
                            'hourly_expenses': hourly_expenses,
                            'sheet_name': settings['sheet'],
                            'excel_row': start_row + idx,
                            'code_col': code_col,
                            'expenses_col': expenses_col,
                            'hourly_expenses_col': hourly_expenses_col
                        })
                except Exception:
                    continue

            print(f"[OK] Loaded organizations: {len(organization_data)}")
            return organization_data

        except Exception as e:
            raise Exception(f"Organization loading error: {str(e)}")

    def load_all_data(self, excel_path: Path) -> Dict[str, List[Dict[str, Any]]]:
        """
        Load all data from Excel file.

        Args:
            excel_path: Path to Excel file

        Returns:
            Dictionary with keys: 'prg', 'grs', 'consumers'

        Raises:
            Exception: If any loading operation fails
        """
        print(f"\n[INFO] Loading data from: {excel_path}")

        prg_data = self.load_prg_data(excel_path)
        grs_data = self.load_grs_data(excel_path)
        population_data = self.load_population_data(excel_path)
        organization_data = self.load_organization_data(excel_path)

        consumer_data = population_data + organization_data

        print(f"\n[OK] Total loaded: PRG={len(prg_data)}, GRS={len(grs_data)}, Consumers={len(consumer_data)}\n")

        return {
            'prg': prg_data,
            'grs': grs_data,
            'consumers': consumer_data
        }
