"""
PRG Pipeline Manager - Main Entry Point

Gas pipeline (PRG) binding management system with modern architecture.
"""

import sys
import tkinter as tk
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from prg.config import SettingsManager
from prg.data import ExcelLoader
from prg.business import (
    ValidationService,
    CalculationService,
    BindingService,
    SearchService
)
from prg.ui import StyleManager, PRGPipelineManager


def main():
    """
    Main application entry point with dependency injection.

    Sets up all services and dependencies before creating the UI.
    """
    print("[INFO] Initializing PRG Pipeline Manager v7.4...")

    # Initialize configuration
    settings_manager = SettingsManager()
    print("[OK] Settings loaded")

    # Initialize data layer
    excel_loader = ExcelLoader(settings_manager)
    print("[OK] Excel loader initialized")

    # Initialize business services
    validation_service = ValidationService()
    calculation_service = CalculationService(validation_service)
    binding_service = BindingService(validation_service)
    search_service = SearchService(validation_service)
    print("[OK] Business services initialized")

    # Initialize UI styling with saved theme preference
    saved_theme = settings_manager.get_ui_preference('theme', 'light')
    style_manager = StyleManager(theme=saved_theme)
    print(f"[OK] Style manager initialized with {saved_theme} theme")

    # Create main window
    root = tk.Tk()
    root.title("PRG Pipeline Manager v7.4 - Professional Edition")

    # Create application with dependency injection
    app = PRGPipelineManager(
        root=root,
        settings_manager=settings_manager,
        excel_loader=excel_loader,
        validation_service=validation_service,
        calculation_service=calculation_service,
        binding_service=binding_service,
        search_service=search_service,
        style_manager=style_manager
    )

    print("[OK] Application initialized")
    print("[INFO] Starting main loop...\n")

    # Run application
    app.run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[INFO] Application closed by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n[ERROR] Application error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
