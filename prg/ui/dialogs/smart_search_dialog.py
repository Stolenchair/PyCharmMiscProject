"""Smart search dialog for finding organizations by location and street."""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, List, Optional, Any


class SmartSearchDialog:
    """Диалог умного поиска с выпадающими списками"""

    def __init__(self, parent, districts: List[str], settlements: List[str],
                 prg_ids: List[str], selected_prg: Dict[str, Any], style_manager=None):
        """
        Initialize smart search dialog.

        Args:
            parent: Parent tkinter window
            districts: List of districts for dropdown
            settlements: List of settlements for dropdown
            prg_ids: List of PRG IDs for dropdown
            selected_prg: Currently selected PRG dictionary
            style_manager: StyleManager instance for theming
        """
        self.result: Optional[Dict[str, Any]] = None
        self.style_manager = style_manager

        # Get colors from style manager or use defaults
        if style_manager:
            colors = style_manager.colors
        else:
            colors = {
                'bg': '#F5F7FA',
                'bg_panel': '#FFFFFF',
                'text': '#1A1A1A',
                'primary': '#1565C0',
                'success': '#2E7D32',
                'danger': '#C62828'
            }

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Умный поиск")
        self.dialog.geometry("800x750")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg=colors['bg'])

        # Центрируем
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - self.dialog.winfo_width()) // 2
        y = (self.dialog.winfo_screenheight() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

        self.create_dialog_content(districts, settlements, prg_ids, selected_prg, colors)

        # Ожидание результата
        self.dialog.wait_window()

    def create_dialog_content(self, districts, settlements, prg_ids, selected_prg, colors):
        """Создание содержимого диалога"""
        main_frame = tk.Frame(self.dialog, padx=30, pady=30, bg=colors['bg'])
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = tk.Label(main_frame, text="УМНЫЙ ПОИСК v7.4",
                               font=('Segoe UI', 18, 'bold'), fg=colors['primary'],
                               bg=colors['bg'])
        title_label.pack(pady=(0, 20))

        # Информация о выбранном ПРГ
        selected_info_frame = tk.LabelFrame(main_frame, text="Выбранный ПРГ (автозаполнение)",
                                            font=('Segoe UI', 11, 'bold'), fg=colors['success'],
                                            bg=colors['bg'], borderwidth=1, relief='solid')
        selected_info_frame.pack(fill=tk.X, pady=(0, 20))

        selected_info = tk.Frame(selected_info_frame, bg=colors['bg'])
        selected_info.pack(fill=tk.X, padx=20, pady=15)

        tk.Label(selected_info, text=f"ПРГ ID: {selected_prg['prg_id']}",
                 font=('Segoe UI', 11, 'bold'), fg=colors['primary'],
                 bg=colors['bg']).pack(anchor=tk.W)
        tk.Label(selected_info, text=f"Район: {selected_prg['mo']}",
                 font=('Segoe UI', 10), fg=colors['text'],
                 bg=colors['bg']).pack(anchor=tk.W, pady=(5, 0))
        tk.Label(selected_info, text=f"НП: {selected_prg['settlement']}",
                 font=('Segoe UI', 10), fg=colors['text'],
                 bg=colors['bg']).pack(anchor=tk.W, pady=(5, 0))

        # Описание
        desc_frame = tk.LabelFrame(main_frame, text="Функции v7.4",
                                   font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                                   bg=colors['bg'], borderwidth=1, relief='solid')
        desc_frame.pack(fill=tk.X, pady=(0, 25))

        desc_text = tk.Text(desc_frame, height=4, wrap=tk.WORD, font=('Segoe UI', 10),
                           bg=colors['bg_panel'], fg=colors['text'], borderwidth=0)
        desc_text.pack(fill=tk.X, padx=20, pady=15)

        desc_content = """ВЫПАДАЮЩИЕ СПИСКИ: Район, НП, ПРГ ID заполняются из данных
АВТОЗАПОЛНЕНИЕ: Поля заполняются из выбранного в интерфейсе ПРГ
РУЧНОЙ ВВОД: Только поле "улица" требует ручного ввода
УМНЫЙ ПОИСК: Ищет организации по 4 критериям + проверяет расходы"""

        desc_text.insert(tk.END, desc_content)
        desc_text.config(state=tk.DISABLED)

        # Поля ввода
        input_frame = tk.LabelFrame(main_frame, text="Параметры умного поиска",
                                    font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                                    bg=colors['bg'], borderwidth=1, relief='solid')
        input_frame.pack(fill=tk.X, pady=(0, 25))

        fields_frame = tk.Frame(input_frame, bg=colors['bg'])
        fields_frame.pack(fill=tk.X, padx=25, pady=20)

        # Район организации (выпадающий список)
        tk.Label(fields_frame, text="1. Район организации:",
                 font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                 bg=colors['bg']).grid(row=0, column=0, sticky=tk.W, pady=12)
        self.mo_var = tk.StringVar()
        self.mo_combo = ttk.Combobox(fields_frame, textvariable=self.mo_var,
                                     values=districts, font=('Segoe UI', 10), width=25,
                                     state="readonly", style='Modern.TCombobox')
        if selected_prg['mo'] in districts:
            self.mo_combo.set(selected_prg['mo'])
        elif districts:
            self.mo_combo.set(districts[0])
        self.mo_combo.grid(row=0, column=1, padx=(20, 0), pady=12, sticky=tk.W)

        tk.Label(fields_frame, text="Выпадающий список",
                 font=('Segoe UI', 9), fg=colors['success'],
                 bg=colors['bg']).grid(row=0, column=2, padx=(10, 0), pady=12, sticky=tk.W)

        # НП организации (выпадающий список)
        tk.Label(fields_frame, text="2. Населенный пункт:",
                 font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                 bg=colors['bg']).grid(row=1, column=0, sticky=tk.W, pady=12)
        self.settlement_var = tk.StringVar()
        self.settlement_combo = ttk.Combobox(fields_frame, textvariable=self.settlement_var,
                                             values=settlements, font=('Segoe UI', 10), width=25,
                                             state="readonly", style='Modern.TCombobox')
        if selected_prg['settlement'] in settlements:
            self.settlement_combo.set(selected_prg['settlement'])
        elif settlements:
            self.settlement_combo.set(settlements[0])
        self.settlement_combo.grid(row=1, column=1, padx=(20, 0), pady=12, sticky=tk.W)

        tk.Label(fields_frame, text="Выпадающий список",
                 font=('Segoe UI', 9), fg=colors['success'],
                 bg=colors['bg']).grid(row=1, column=2, padx=(10, 0), pady=12, sticky=tk.W)

        # Улица (ручной ввод)
        tk.Label(fields_frame, text="3. Улица (без 'ул.'):",
                 font=('Segoe UI', 11, 'bold'), fg=colors['danger'],
                 bg=colors['bg']).grid(row=2, column=0, sticky=tk.W, pady=12)
        self.street_var = tk.StringVar()
        self.street_entry = tk.Entry(fields_frame, textvariable=self.street_var,
                                     font=('Segoe UI', 10), width=27,
                                     bg=colors['bg_panel'], fg=colors['text'])
        self.street_entry.grid(row=2, column=1, padx=(20, 0), pady=12, sticky=tk.W)

        tk.Label(fields_frame, text="Ручной ввод",
                 font=('Segoe UI', 9), fg=colors['danger'],
                 bg=colors['bg']).grid(row=2, column=2, padx=(10, 0), pady=12, sticky=tk.W)

        # ПРГ ID (выпадающий список, автозаполнен)
        tk.Label(fields_frame, text="4. ПРГ ID:",
                 font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                 bg=colors['bg']).grid(row=3, column=0, sticky=tk.W, pady=12)
        self.prg_id_var = tk.StringVar()
        self.prg_id_combo = ttk.Combobox(fields_frame, textvariable=self.prg_id_var,
                                         values=prg_ids, font=('Segoe UI', 10), width=25,
                                         state="readonly", style='Modern.TCombobox')
        self.prg_id_combo.set(selected_prg['prg_id'])
        self.prg_id_combo.grid(row=3, column=1, padx=(20, 0), pady=12, sticky=tk.W)

        tk.Label(fields_frame, text="Автозаполнение",
                 font=('Segoe UI', 9), fg=colors['primary'],
                 bg=colors['bg']).grid(row=3, column=2, padx=(10, 0), pady=12, sticky=tk.W)

        # Доля
        tk.Label(fields_frame, text="5. Доля для привязки:",
                 font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                 bg=colors['bg']).grid(row=4, column=0, sticky=tk.W, pady=12)
        self.share_var = tk.StringVar()
        self.share_var.set("1.0")
        self.share_entry = tk.Entry(fields_frame, textvariable=self.share_var,
                                    font=('Segoe UI', 10), width=27,
                                    bg=colors['bg_panel'], fg=colors['text'])
        self.share_entry.grid(row=4, column=1, padx=(20, 0), pady=12, sticky=tk.W)

        tk.Label(fields_frame, text="Стандартная доля",
                 font=('Segoe UI', 9), fg=colors['text_secondary'],
                 bg=colors['bg']).grid(row=4, column=2, padx=(10, 0), pady=12, sticky=tk.W)

        # Пример
        example_frame = tk.LabelFrame(main_frame, text="Пример умного поиска",
                                      font=('Segoe UI', 11, 'bold'), fg=colors['text'],
                                      bg=colors['bg'], borderwidth=1, relief='solid')
        example_frame.pack(fill=tk.X, pady=(0, 25))

        example_text = tk.Text(example_frame, height=6, wrap=tk.WORD, font=('Segoe UI', 10),
                              bg=colors['bg_panel'], fg=colors['text'], borderwidth=0)
        example_text.pack(fill=tk.X, padx=20, pady=15)

        example_content = f"""ПРИМЕР (данные автоматически заполнены):
Район: {selected_prg['mo']} (из выбранного ПРГ)
НП: {selected_prg['settlement']} (из выбранного ПРГ)
Улица: Ленина (единственное поле для ручного ввода)
ПРГ ID: {selected_prg['prg_id']} (из выбранного ПРГ)

РЕЗУЛЬТАТ: Найдутся все организации с "ул.Ленина" в названии
в районе "{selected_prg['mo']}", НП "{selected_prg['settlement']}" и привяжутся к ПРГ {selected_prg['prg_id']}."""

        example_text.insert(tk.END, example_content)
        example_text.config(state=tk.DISABLED)

        # Кнопки
        button_frame = tk.Frame(main_frame, bg=colors['bg'])
        button_frame.pack(fill=tk.X)

        if self.style_manager:
            search_btn = self.style_manager.create_button(
                button_frame, text="Найти и привязать",
                command=self.ok_clicked, color='secondary', width=18
            )
            search_btn.pack(side=tk.RIGHT, padx=(20, 0))

            cancel_btn = self.style_manager.create_button(
                button_frame, text="Отмена",
                command=self.cancel_clicked, color='danger', width=12
            )
            cancel_btn.pack(side=tk.RIGHT)
        else:
            tk.Button(button_frame, text="Найти и привязать",
                      command=self.ok_clicked,
                      bg=colors['primary'], fg='white', font=('Segoe UI', 12, 'bold'),
                      width=18, relief='flat').pack(side=tk.RIGHT, padx=(20, 0))
            tk.Button(button_frame, text="Отмена", command=self.cancel_clicked,
                      bg=colors['danger'], fg='white', font=('Segoe UI', 12),
                      width=12, relief='flat').pack(side=tk.RIGHT)

        # Устанавливаем фокус на поле улицы (единственное для ручного ввода)
        self.street_entry.focus()

        # Привязки клавиш
        self.dialog.bind('<Return>', lambda e: self.ok_clicked())
        self.dialog.bind('<Escape>', lambda e: self.cancel_clicked())

    def ok_clicked(self):
        """Обработка нажатия OK"""
        try:
            mo_district = self.mo_var.get().strip()
            settlement = self.settlement_var.get().strip()
            street = self.street_var.get().strip()
            prg_id = self.prg_id_var.get().strip()
            share_str = self.share_var.get().strip()

            # Проверяем заполнение полей
            if not mo_district:
                messagebox.showerror("Ошибка", "Выберите район организации из списка")
                return

            if not settlement:
                messagebox.showerror("Ошибка", "Выберите населенный пункт из списка")
                return

            if not street:
                messagebox.showerror("Ошибка", "Введите улицу (без 'ул.')\n\nЭто единственное поле для ручного ввода!")
                self.street_entry.focus()
                return

            if not prg_id:
                messagebox.showerror("Ошибка", "Выберите ПРГ ID из списка")
                return

            # Проверяем долю
            try:
                share = float(share_str.replace(',', '.'))
                if share <= 0 or share > 1:
                    messagebox.showerror("Ошибка", "Доля должна быть от 0 до 1")
                    self.share_entry.focus()
                    return
            except ValueError:
                messagebox.showerror("Ошибка", "Введите корректную долю (например: 0.5)")
                self.share_entry.focus()
                return

            self.result = {
                'mo_district': mo_district,
                'settlement': settlement,
                'street': street,
                'prg_id': prg_id,
                'share': share
            }

            self.dialog.destroy()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка ввода данных: {str(e)}")

    def cancel_clicked(self):
        """Обработка отмены"""
        self.result = None
        self.dialog.destroy()
