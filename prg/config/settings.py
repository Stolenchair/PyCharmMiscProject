"""Settings management for PRG Pipeline Manager."""

import json
from pathlib import Path
from typing import Dict, Any, Optional
from .defaults import get_default_settings


class SettingsManager:
    """
    Manages application settings persistence to/from prg_settings.json.

    Settings include Excel sheet names and column mappings for:
    - PRG (pipeline data)
    - GRS (gas reduction stations)
    - Population (consumer data)
    - Organizations (consumer data)
    """

    def __init__(self, settings_file: str = "prg_settings.json"):
        """
        Initialize settings manager.

        Args:
            settings_file: Path to settings JSON file
        """
        self.settings_file = Path(settings_file)
        self.settings = self.load()

    def load(self) -> Dict[str, Dict[str, Any]]:
        """
        Load settings from file, merging with defaults.

        Returns:
            dict: Settings for all table types
        """
        default_settings = get_default_settings()

        # Add UI preferences section
        if 'ui_preferences' not in default_settings:
            default_settings['ui_preferences'] = {
                'theme': 'light',
                'window_geometry': '1500x900'
            }

        try:
            if self.settings_file.exists():
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    saved_settings = json.load(f)

                # Merge saved settings with defaults
                for table_type in default_settings:
                    if table_type in saved_settings:
                        default_settings[table_type].update(saved_settings[table_type])

                print(f"[OK] Settings loaded from {self.settings_file}")
            else:
                print(f"[INFO] No settings file found, using defaults")

        except Exception as e:
            print(f"[WARNING] Error loading settings: {e}")
            print("[INFO] Using default settings")

        return default_settings

    def save(self, settings: Optional[Dict[str, Dict[str, Any]]] = None) -> bool:
        """
        Save settings to file.

        Args:
            settings: Settings to save (uses current settings if None)

        Returns:
            bool: True if save successful
        """
        if settings is None:
            settings = self.settings

        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)

            print(f"[OK] Settings saved to {self.settings_file}")
            return True

        except Exception as e:
            print(f"[ERROR] Failed to save settings: {e}")
            return False

    def get_table_settings(self, table_type: str) -> Dict[str, Any]:
        """
        Get settings for a specific table type.

        Args:
            table_type: One of 'prg', 'grs', 'population', 'organizations'

        Returns:
            dict: Settings for the table type

        Raises:
            ValueError: If table_type is invalid
        """
        if table_type not in self.settings:
            raise ValueError(f"Invalid table type: {table_type}")

        return self.settings[table_type]

    def update_table_settings(self, table_type: str, updates: Dict[str, Any]) -> None:
        """
        Update settings for a specific table type.

        Args:
            table_type: One of 'prg', 'grs', 'population', 'organizations'
            updates: Dictionary of settings to update
        """
        if table_type not in self.settings:
            raise ValueError(f"Invalid table type: {table_type}")

        self.settings[table_type].update(updates)

    def reset_to_defaults(self) -> None:
        """Reset all settings to defaults."""
        self.settings = get_default_settings()

    def get_all_settings(self) -> Dict[str, Dict[str, Any]]:
        """
        Get all settings.

        Returns:
            dict: Complete settings dictionary
        """
        return self.settings.copy()

    def get_ui_preference(self, key: str, default: Any = None) -> Any:
        """
        Get a UI preference value.

        Args:
            key: Preference key (e.g., 'theme', 'window_geometry')
            default: Default value if not found

        Returns:
            The preference value or default
        """
        if 'ui_preferences' not in self.settings:
            return default
        return self.settings['ui_preferences'].get(key, default)

    def set_ui_preference(self, key: str, value: Any) -> None:
        """
        Set a UI preference value.

        Args:
            key: Preference key (e.g., 'theme', 'window_geometry')
            value: Value to set
        """
        if 'ui_preferences' not in self.settings:
            self.settings['ui_preferences'] = {}
        self.settings['ui_preferences'][key] = value

    def __repr__(self):
        return f"SettingsManager(file='{self.settings_file}')"
