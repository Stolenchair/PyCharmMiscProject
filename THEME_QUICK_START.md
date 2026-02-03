# Theme System - Quick Start Guide

## For Users

### Switching Themes

1. **Open the Application**
   ```bash
   python main.py
   ```

2. **Change Theme**
   - Click on the menu: **Вид** → **Темная тема** (if in light mode)
   - Or: **Вид** → **Светлая тема** (if in dark mode)

3. **Restart Application**
   - Close and reopen the app for full effect
   - Your theme choice is saved automatically

### Theme Locations

Your theme preference is saved in:
```
C:\Users\Gurbanov\PyCharmMiscProject\prg_settings.json
```

Example content:
```json
{
  "ui_preferences": {
    "theme": "dark",
    "window_geometry": "1500x900"
  }
}
```

### Choosing the Right Theme

**Light Theme** - Best for:
- Bright office environments
- Daytime work
- Presentations
- Printed screenshots

**Dark Theme** - Best for:
- Low-light environments
- Night work
- Reduced eye strain
- Extended screen time

---

## For Developers

### Adding Theme Support to New Components

#### Basic Widget Theming

```python
def create_my_widget(self):
    colors = self.style_manager.colors

    # Frame
    frame = tk.Frame(parent, bg=colors['bg'])

    # Label
    label = tk.Label(frame,
                    text="My Label",
                    bg=colors['bg'],
                    fg=colors['text'],
                    font=('Segoe UI', 10))

    # Entry
    entry = tk.Entry(frame,
                    bg=colors['bg_panel'],
                    fg=colors['text'],
                    font=('Segoe UI', 10))
```

#### Using StyleManager Buttons

```python
# Success button (green)
btn = self.style_manager.create_button(
    parent,
    text="Save",
    command=self.save_action,
    color='success'
)

# Danger button (red)
btn = self.style_manager.create_button(
    parent,
    text="Delete",
    command=self.delete_action,
    color='danger'
)

# Primary button (blue)
btn = self.style_manager.create_button(
    parent,
    text="OK",
    command=self.ok_action,
    color='primary'
)
```

#### Creating Themed Dialogs

```python
def show_dialog(self):
    colors = self.style_manager.colors

    dialog = tk.Toplevel(self.root)
    dialog.configure(bg=colors['bg'])

    main_frame = tk.Frame(dialog, bg=colors['bg'], padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)

    # Title
    tk.Label(main_frame,
            text="Dialog Title",
            font=('Segoe UI', 14, 'bold'),
            bg=colors['bg'],
            fg=colors['primary']).pack(pady=(0, 15))

    # Content
    content_frame = tk.LabelFrame(main_frame,
                                 text="Content",
                                 font=('Segoe UI', 11, 'bold'),
                                 bg=colors['bg'],
                                 fg=colors['text'],
                                 borderwidth=1,
                                 relief='solid')
    content_frame.pack(fill=tk.BOTH, expand=True)

    # Buttons
    button_frame = tk.Frame(main_frame, bg=colors['bg'])
    button_frame.pack(fill=tk.X, pady=(15, 0))

    ok_btn = self.style_manager.create_button(
        button_frame, text="OK", command=dialog.destroy, color='primary'
    )
    ok_btn.pack(side=tk.RIGHT)
```

### Available Colors

Access via `self.style_manager.colors`:

```python
# Backgrounds
colors['bg']              # Main background
colors['bg_secondary']    # Secondary background
colors['bg_panel']        # Panel background
colors['card']            # Card background

# Actions
colors['primary']         # Primary actions (blue)
colors['success']         # Success actions (green)
colors['warning']         # Warning actions (orange)
colors['danger']          # Danger actions (red)
colors['secondary']       # Secondary actions (teal)
colors['purple']          # Special actions (purple)

# Text
colors['text']            # Primary text
colors['text_secondary']  # Secondary text
colors['text_muted']      # Muted text

# UI Elements
colors['border']          # Borders
colors['border_light']    # Light borders
colors['hover']           # Hover state
colors['selected']        # Selected state
colors['shadow']          # Shadow effects
```

### Hover Effects

```python
# Manual hover effect
button = tk.Button(parent, text="Click", bg=colors['primary'], fg='white')
self.style_manager.add_button_hover(
    button,
    base_color=colors['primary'],
    hover_color=colors['primary_hover']
)

# Or use create_button which includes hover automatically
button = self.style_manager.create_button(
    parent, text="Click", command=action, color='primary'
)
```

### Styled ttk Widgets

```python
# TreeView
tree = ttk.Treeview(parent, style='Modern.Treeview')

# Scrollbar
scroll = ttk.Scrollbar(parent, style='Modern.Vertical.TScrollbar')

# Combobox
combo = ttk.Combobox(parent, style='Modern.TCombobox')

# Frame
frame = ttk.Frame(parent, style='Modern.TFrame')

# LabelFrame
lframe = ttk.LabelFrame(parent, text="Title", style='Modern.TLabelframe')
```

### Creating a New Theme

Edit `prg/ui/styles.py`:

```python
def _setup_themes(self):
    # Add new theme
    self.my_custom_theme = {
        'bg': '#YOUR_COLOR',
        'bg_secondary': '#YOUR_COLOR',
        'bg_panel': '#YOUR_COLOR',
        'card': '#YOUR_COLOR',
        'primary': '#YOUR_COLOR',
        'primary_hover': '#YOUR_COLOR',
        # ... more colors
    }

    # Update apply_theme method
    def apply_theme(self, theme: str):
        if theme == 'my_custom':
            self.colors = self.my_custom_theme.copy()
        # ... rest of logic
```

### Best Practices

1. **Always use theme colors**
   ```python
   # ✓ Good
   label = tk.Label(frame, bg=colors['bg'], fg=colors['text'])

   # ✗ Bad
   label = tk.Label(frame, bg='#FFFFFF', fg='#000000')
   ```

2. **Use semantic colors**
   ```python
   # ✓ Good - semantic
   btn = style_manager.create_button(parent, color='success')

   # ✗ Bad - arbitrary
   btn = tk.Button(parent, bg='#4CAF50')
   ```

3. **Pass style_manager to dialogs**
   ```python
   # ✓ Good
   dialog = MyDialog(parent, style_manager=self.style_manager)

   # ✗ Bad
   dialog = MyDialog(parent)  # No theme support
   ```

4. **Apply theme on initialization**
   ```python
   # In __init__
   self.style_manager.apply()  # Apply ttk styles
   self._apply_window_theme()   # Apply tk styles
   ```

### Testing Both Themes

```python
# Test script
def test_themes():
    from prg.ui import StyleManager

    # Test light theme
    style_light = StyleManager(theme='light')
    print("Light theme colors:", style_light.colors)

    # Test dark theme
    style_dark = StyleManager(theme='dark')
    print("Dark theme colors:", style_dark.colors)

    # Test toggle
    current = style_light.toggle_theme()
    print(f"Toggled to: {current}")

if __name__ == '__main__':
    test_themes()
```

---

## Common Issues

### Issue: Colors not applying

**Solution:**
```python
# Make sure to call apply() after StyleManager initialization
style_manager.apply()

# And apply window theme
colors = style_manager.colors
root.configure(bg=colors['bg'])
```

### Issue: Theme not persisting

**Solution:**
```python
# Make sure to save settings on theme change
settings_manager.set_ui_preference('theme', new_theme)
settings_manager.save()
```

### Issue: Some widgets not themed

**Solution:**
```python
# For tk widgets, manually apply colors
widget.config(bg=colors['bg'], fg=colors['text'])

# For ttk widgets, use styled versions
widget = ttk.Widget(parent, style='Modern.WidgetType')
```

### Issue: Hover effects not working

**Solution:**
```python
# Use StyleManager's create_button method
btn = style_manager.create_button(parent, text="Click", color='primary')

# Or manually add hover
button = tk.Button(parent, ...)
style_manager.add_button_hover(button, base_color, hover_color)
```

---

## Migration Guide

### Migrating Old Code to New Theme System

**Before:**
```python
button = tk.Button(parent, text="Save",
                  bg='#4CAF50', fg='white',
                  font=('Arial', 11, 'bold'))
```

**After:**
```python
button = self.style_manager.create_button(
    parent, text="Save", color='success'
)
```

**Before:**
```python
frame = tk.Frame(parent, bg='#f0f0f0')
label = tk.Label(frame, text="Text", bg='#f0f0f0', fg='black')
```

**After:**
```python
colors = self.style_manager.colors
frame = tk.Frame(parent, bg=colors['bg'])
label = tk.Label(frame, text="Text", bg=colors['bg'], fg=colors['text'])
```

---

## Resources

- **Full Documentation**: `THEME_REDESIGN.md`
- **Design Comparison**: `DESIGN_COMPARISON.md`
- **Source Code**:
  - `prg/ui/styles.py` - StyleManager implementation
  - `prg/ui/main_window.py` - Example usage
  - `prg/config/settings.py` - Theme persistence

---

## Support

For questions or issues:
1. Check the documentation files
2. Review example code in `main_window.py`
3. Examine `styles.py` for color definitions
4. Test with both themes to ensure compatibility
