"""Settings dialog for configuring Excel column mappings."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any, Optional


class SettingsDialog:
    """
    Dialog for configuring Excel column mappings for all data types.

    Allows user to configure:
    - PRG: sheet name, start row, column mappings
    - GRS: sheet name, start row, column mappings
    - Population: sheet name, start row, column mappings
    - Organizations: sheet name, start row, column mappings
    """

    def __init__(self, parent, settings_manager, style_manager):
        """
        Initialize settings dialog.

        Args:
            parent: Parent window
            settings_manager: SettingsManager instance
            style_manager: StyleManager instance for theming
        """
        self.parent = parent
        self.settings_manager = settings_manager
        self.style_manager = style_manager
        self.colors = style_manager.colors

        self.result = None  # Will be set to True if user saves

        # Create dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Настройки столбцов Excel")
        self.dialog.geometry("800x700")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=self.colors['bg'])

        # Center dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - self.dialog.winfo_width()) // 2
        y = (self.dialog.winfo_screenheight() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

        # Load current settings
        self.settings = {
            'prg': self.settings_manager.get_table_settings('prg').copy(),
            'grs': self.settings_manager.get_table_settings('grs').copy(),
            'population': self.settings_manager.get_table_settings('population').copy(),
            'organizations': self.settings_manager.get_table_settings('organizations').copy()
        }

        self.create_ui()

        # Bind keyboard shortcuts
        self.dialog.bind('<Escape>', lambda e: self.cancel())

        self.dialog.wait_window()

    def create_ui(self):
        """Create UI elements."""
        main_frame = tk.Frame(self.dialog, padx=20, pady=20, bg=self.colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = tk.Label(
            main_frame,
            text="⚙️ НАСТРОЙКИ СТОЛБЦОВ EXCEL",
            font=('Segoe UI', 14, 'bold'),
            fg=self.colors['primary'],
            bg=self.colors['bg']
        )
        title_label.pack(pady=(0, 10))

        # Info
        info_label = tk.Label(
            main_frame,
            text="Укажите названия листов и буквы колонок для каждого типа данных.\n"
                 "Колонки указываются буквами: A, B, C, ... Z, AA, AB, ...",
            font=('Segoe UI', 9),
            fg=self.colors['text_secondary'],
            bg=self.colors['bg'],
            justify=tk.LEFT
        )
        info_label.pack(pady=(0, 15))

        # Notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Create tabs
        self.create_prg_tab()
        self.create_grs_tab()
        self.create_population_tab()
        self.create_organizations_tab()

        # Buttons
        button_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        button_frame.pack(fill=tk.X)

        save_btn = self.style_manager.create_button(
            button_frame,
            text="Сохранить",
            command=self.save,
            color='success',
            width=15
        )
        save_btn.pack(side=tk.RIGHT, padx=(10, 0))

        cancel_btn = self.style_manager.create_button(
            button_frame,
            text="Отмена (Esc)",
            command=self.cancel,
            color='text_secondary',
            width=15
        )
        cancel_btn.config(bg=self.colors['text_secondary'])
        self.style_manager.add_button_hover(
            cancel_btn,
            self.colors['text_secondary'],
            self.colors['text_muted']
        )
        cancel_btn.pack(side=tk.RIGHT)

    def create_settings_frame(self, parent, table_type: str, fields: Dict[str, str]) -> Dict[str, tk.StringVar]:
        """
        Create settings frame for a table type.

        Args:
            parent: Parent widget
            table_type: Type of table ('prg', 'grs', etc.)
            fields: Dictionary of field_key: label_text

        Returns:
            Dictionary of field_key: StringVar
        """
        frame = tk.Frame(parent, bg=self.colors['bg'], padx=10, pady=10)
        frame.pack(fill=tk.BOTH, expand=True)

        # Create scrollable frame
        canvas = tk.Canvas(frame, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.colors['bg'])

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create entry fields
        vars_dict = {}
        current_settings = self.settings[table_type]

        for field_key, label_text in fields.items():
            row_frame = tk.Frame(scrollable_frame, bg=self.colors['bg'])
            row_frame.pack(fill=tk.X, pady=5, padx=10)

            label = tk.Label(
                row_frame,
                text=label_text,
                font=('Segoe UI', 10),
                bg=self.colors['bg'],
                fg=self.colors['text'],
                width=40,
                anchor=tk.W
            )
            label.pack(side=tk.LEFT, padx=(0, 10))

            var = tk.StringVar(value=current_settings.get(field_key, ''))
            entry = tk.Entry(
                row_frame,
                textvariable=var,
                font=('Segoe UI', 10),
                width=15,
                bg=self.colors['bg_panel'],
                fg=self.colors['text']
            )
            entry.pack(side=tk.LEFT)

            vars_dict[field_key] = var

        return vars_dict

    def create_prg_tab(self):
        """Create PRG settings tab."""
        tab = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(tab, text="🏭 ПРГ")

        fields = {
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
        }

        self.prg_vars = self.create_settings_frame(tab, 'prg', fields)

    def create_grs_tab(self):
        """Create GRS settings tab."""
        tab = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(tab, text="🏢 ГРС")

        fields = {
            'sheet': 'Название листа',
            'start_row': 'Начальная строка данных',
            'mo_col': 'Колонка МО (район)',
            'grs_id_col': 'Колонка ГРС ID',
            'grs_name_col': 'Колонка Название ГРС'
        }

        self.grs_vars = self.create_settings_frame(tab, 'grs', fields)

    def create_population_tab(self):
        """Create Population settings tab."""
        tab = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(tab, text="👥 Население")

        fields = {
            'sheet': 'Название листа',
            'start_row': 'Начальная строка данных',
            'mo_col': 'Колонка МО (район)',
            'settlement_col': 'Колонка НП',
            'code_col': 'Колонка Привязки ПРГ',
            'expenses_col': 'Колонка Годовые расходы',
            'hourly_expenses_col': 'Колонка Часовые расходы'
        }

        self.population_vars = self.create_settings_frame(tab, 'population', fields)

    def create_organizations_tab(self):
        """Create Organizations settings tab."""
        tab = tk.Frame(self.notebook, bg=self.colors['bg'])
        self.notebook.add(tab, text="🏢 Организации")

        fields = {
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

        self.organizations_vars = self.create_settings_frame(tab, 'organizations', fields)

    def validate_settings(self) -> bool:
        """
        Validate all settings.

        Returns:
            True if valid, False otherwise
        """
        # Validate PRG
        if not self.prg_vars['sheet'].get().strip():
            messagebox.showerror("Ошибка", "Укажите название листа для ПРГ", parent=self.dialog)
            self.notebook.select(0)
            return False

        try:
            start_row = int(self.prg_vars['start_row'].get())
            if start_row < 1:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Ошибка", "Начальная строка ПРГ должна быть числом >= 1", parent=self.dialog)
            self.notebook.select(0)
            return False

        # Validate GRS
        if not self.grs_vars['sheet'].get().strip():
            messagebox.showerror("Ошибка", "Укажите название листа для ГРС", parent=self.dialog)
            self.notebook.select(1)
            return False

        # Validate Population
        if not self.population_vars['sheet'].get().strip():
            messagebox.showerror("Ошибка", "Укажите название листа для Население", parent=self.dialog)
            self.notebook.select(2)
            return False

        # Validate Organizations
        if not self.organizations_vars['sheet'].get().strip():
            messagebox.showerror("Ошибка", "Укажите название листа для Организации", parent=self.dialog)
            self.notebook.select(3)
            return False

        return True

    def save(self):
        """Save settings."""
        if not self.validate_settings():
            return

        # Collect all settings
        new_settings = {
            'prg': {key: var.get().strip() for key, var in self.prg_vars.items()},
            'grs': {key: var.get().strip() for key, var in self.grs_vars.items()},
            'population': {key: var.get().strip() for key, var in self.population_vars.items()},
            'organizations': {key: var.get().strip() for key, var in self.organizations_vars.items()}
        }

        # Ask for confirmation
        response = messagebox.askyesno(
            "Сохранение настроек",
            "Сохранить настройки столбцов?\n\n"
            "Настройки будут применены при следующей загрузке файла.\n"
            "Текущие данные останутся без изменений.",
            parent=self.dialog
        )

        if not response:
            return

        try:
            # Update settings manager
            for table_type, settings in new_settings.items():
                for key, value in settings.items():
                    self.settings_manager.settings[table_type][key] = value

            # Save to file
            self.settings_manager.save()

            messagebox.showinfo(
                "Успех",
                "Настройки успешно сохранены!\n\n"
                "Перезагрузите файл Excel, чтобы применить новые настройки.",
                parent=self.dialog
            )

            self.result = True
            self.dialog.destroy()

        except Exception as e:
            messagebox.showerror(
                "Ошибка",
                f"Ошибка сохранения настроек:\n\n{str(e)}",
                parent=self.dialog
            )

    def cancel(self):
        """Cancel and close dialog."""
        self.dialog.destroy()
