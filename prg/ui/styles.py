"""Modern UI styling for PRG Pipeline Manager with theme support."""

import tkinter as tk
from tkinter import ttk
from typing import Dict


class StyleManager:
    """
    Manages modern UI styling for the application with theme support.

    Provides:
    - Light and dark theme color palettes
    - ttk.Style configuration
    - Button hover effects
    - Theme switching
    """

    def __init__(self, theme: str = 'light'):
        """
        Initialize style manager with theme.

        Args:
            theme: 'light' or 'dark'
        """
        self.current_theme = theme
        self._setup_themes()
        self.apply_theme(theme)

    def _setup_themes(self):
        """Setup color palettes for both themes."""
        # Light theme - Professional business style
        self.light_theme = {
            'bg': '#F5F7FA',
            'bg_secondary': '#E8ECF1',
            'bg_panel': '#FFFFFF',
            'card': '#FFFFFF',
            'primary': '#1565C0',
            'primary_hover': '#0D47A1',
            'primary_light': '#1976D2',
            'success': '#2E7D32',
            'success_hover': '#1B5E20',
            'warning': '#EF6C00',
            'warning_hover': '#E65100',
            'danger': '#C62828',
            'danger_hover': '#B71C1C',
            'secondary': '#00838F',
            'secondary_hover': '#006064',
            'purple': '#6A1B9A',
            'purple_hover': '#4A148C',
            'text': '#1A1A1A',
            'text_secondary': '#5F6368',
            'text_muted': '#80868B',
            'border': '#DADCE0',
            'border_light': '#E8EAED',
            'hover': '#F1F3F4',
            'selected': '#E8F0FE',
            'shadow': '#00000015'
        }

        # Dark theme - Professional dark gray (not pure black)
        self.dark_theme = {
            'bg': '#1E1E1E',
            'bg_secondary': '#2B2B2B',
            'bg_panel': '#252525',
            'card': '#2D2D2D',
            'primary': '#4A9EFF',
            'primary_hover': '#6BB1FF',
            'primary_light': '#3B8FE8',
            'success': '#4CAF50',
            'success_hover': '#66BB6A',
            'warning': '#FF9800',
            'warning_hover': '#FFB74D',
            'danger': '#F44336',
            'danger_hover': '#EF5350',
            'secondary': '#26C6DA',
            'secondary_hover': '#4DD0E1',
            'purple': '#AB47BC',
            'purple_hover': '#BA68C8',
            'text': '#E8EAED',
            'text_secondary': '#9AA0A6',
            'text_muted': '#5F6368',
            'border': '#3C3C3C',
            'border_light': '#444444',
            'hover': '#383838',
            'selected': '#2B3E50',
            'shadow': '#00000040'
        }

        # Set initial colors
        self.colors = self.light_theme.copy()

        # Tree row alternating colors
        self._update_tree_tags()

    def _update_tree_tags(self):
        """Update tree row colors based on current theme."""
        self.tree_tags = {
            'evenrow': self.colors['card'],
            'oddrow': self.colors['bg_secondary']
        }

    def apply_theme(self, theme: str):
        """
        Apply a theme to the application.

        Args:
            theme: 'light' or 'dark'
        """
        if theme not in ['light', 'dark']:
            theme = 'light'

        self.current_theme = theme
        self.colors = self.light_theme.copy() if theme == 'light' else self.dark_theme.copy()
        self._update_tree_tags()

    def apply(self):
        """Apply modern style to all ttk widgets."""
        style = ttk.Style()
        style.theme_use('clam')

        # Modern Treeview style
        style.configure('Modern.Treeview',
                       background=self.colors['card'],
                       foreground=self.colors['text'],
                       fieldbackground=self.colors['card'],
                       borderwidth=0,
                       rowheight=28,
                       font=('Segoe UI', 10))

        style.configure('Modern.Treeview.Heading',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['text'],
                       borderwidth=1,
                       relief='flat',
                       font=('Segoe UI', 10, 'bold'))

        style.map('Modern.Treeview.Heading',
                 background=[('active', self.colors['hover'])])

        style.map('Modern.Treeview',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', 'white')])

        # LabelFrame style
        style.configure('Modern.TLabelframe',
                       background=self.colors['bg_panel'],
                       borderwidth=1,
                       relief='solid',
                       bordercolor=self.colors['border'])

        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['bg_panel'],
                       foreground=self.colors['text'],
                       font=('Segoe UI', 11, 'bold'))

        # Frame style
        style.configure('Modern.TFrame',
                       background=self.colors['bg'])

        # Scrollbar style
        style.configure('Modern.Vertical.TScrollbar',
                       background=self.colors['bg_secondary'],
                       troughcolor=self.colors['card'],
                       borderwidth=0,
                       arrowsize=14)

        # Combobox style
        style.configure('Modern.TCombobox',
                       fieldbackground=self.colors['card'],
                       background=self.colors['card'],
                       foreground=self.colors['text'],
                       borderwidth=1,
                       relief='solid')

        style.map('Modern.TCombobox',
                 fieldbackground=[('readonly', self.colors['card'])],
                 selectbackground=[('readonly', self.colors['primary'])],
                 selectforeground=[('readonly', 'white')])

    def add_button_hover(self, button: tk.Button, base_color: str, hover_color: str):
        """
        Add hover effect to button.

        Args:
            button: tkinter Button widget
            base_color: Base background color
            hover_color: Hover background color
        """
        def on_enter(e):
            if button['state'] != 'disabled':
                button['background'] = hover_color

        def on_leave(e):
            if button['state'] != 'disabled':
                button['background'] = base_color

        button.bind('<Enter>', on_enter)
        button.bind('<Leave>', on_leave)

    def create_button(self, parent, text: str, command,
                     color: str = 'primary', **kwargs) -> tk.Button:
        """
        Create a styled button with hover effect.

        Args:
            parent: Parent widget
            text: Button text
            command: Button command
            color: Color key from palette ('primary', 'success', 'danger', etc.)
            **kwargs: Additional button options

        Returns:
            tk.Button: Styled button with hover effect
        """
        base_color = self.colors.get(color, self.colors['primary'])
        hover_color = self.colors.get(f'{color}_hover', base_color)

        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=base_color,
            fg='white',
            font=('Segoe UI', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=20,
            pady=10,
            borderwidth=0,
            **kwargs
        )

        self.add_button_hover(button, base_color, hover_color)
        return button

    def toggle_theme(self):
        """Toggle between light and dark themes."""
        new_theme = 'dark' if self.current_theme == 'light' else 'light'
        self.apply_theme(new_theme)
        return new_theme

    def get_theme(self) -> str:
        """Get current theme name."""
        return self.current_theme

    def create_card_frame(self, parent, **kwargs) -> tk.Frame:
        """
        Create a card-style frame with shadow effect.

        Args:
            parent: Parent widget
            **kwargs: Additional frame options

        Returns:
            tk.Frame: Card-style frame
        """
        return tk.Frame(
            parent,
            bg=self.colors['card'],
            relief='solid',
            borderwidth=1,
            highlightbackground=self.colors['border'],
            highlightthickness=1,
            **kwargs
        )

    def get_color(self, key: str) -> str:
        """
        Get color from palette.

        Args:
            key: Color key

        Returns:
            str: Color hex code
        """
        return self.colors.get(key, '#000000')
