"""Configuration management for PRG Pipeline Manager."""

from .settings import SettingsManager
from .defaults import get_default_settings

__all__ = [
    'SettingsManager',
    'get_default_settings',
]
