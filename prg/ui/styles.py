"""Modern UI styling for PRG Pipeline Manager."""

import tkinter as tk
from tkinter import ttk
from typing import Dict


class StyleManager:
    """
    Manages modern UI styling for the application.

    Provides:
    - Color palette definition
    - ttk.Style configuration
    - Button hover effects
    """

    def __init__(self):
        """Initialize style manager with modern color palette."""
        self.colors = {
            'bg': '#FAFAFA',
            'bg_dark': '#F5F5F5',
            'card': '#FFFFFF',
            'primary': '#1976D2',
            'primary_hover': '#1565C0',
            'success': '#43A047',
            'warning': '#FB8C00',
            'danger': '#E53935',
            'secondary': '#00ACC1',
            'purple': '#8E24AA',
            'text': '#212121',
            'text_secondary': '#757575',
            'border': '#E0E0E0',
            'shadow': '#00000010'
        }

        # Tree row alternating colors
        self.tree_tags = {
            'evenrow': self.colors['card'],
            'oddrow': self.colors['bg_dark']
        }

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
                       background=self.colors['bg_dark'],
                       foreground=self.colors['text'],
                       borderwidth=0,
                       relief='flat',
                       font=('Segoe UI', 10, 'bold'))

        style.map('Modern.Treeview.Heading',
                 background=[('active', self.colors['border'])])

        style.map('Modern.Treeview',
                 background=[('selected', self.colors['primary'])],
                 foreground=[('selected', 'white')])

        # LabelFrame style
        style.configure('Modern.TLabelframe',
                       background=self.colors['card'],
                       borderwidth=1,
                       relief='solid')

        style.configure('Modern.TLabelframe.Label',
                       background=self.colors['card'],
                       foreground=self.colors['text'],
                       font=('Segoe UI', 11, 'bold'))

        # Scrollbar style
        style.configure('Modern.Vertical.TScrollbar',
                       background=self.colors['bg_dark'],
                       troughcolor=self.colors['card'],
                       borderwidth=0,
                       arrowsize=14)

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
            **kwargs
        )

        self.add_button_hover(button, base_color, hover_color)
        return button

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
