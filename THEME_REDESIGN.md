# PRG Pipeline Manager - Professional UI Redesign with Theme Support

## Overview

The PRG Pipeline Manager application has been redesigned with a professional, business-like appearance and includes full support for light and dark themes that can be toggled at runtime.

## What Was Changed

### 1. Enhanced StyleManager (`prg/ui/styles.py`)

**New Features:**
- **Dual Theme System**: Complete light and dark theme color palettes
- **Professional Color Schemes**:
  - Light theme: Clean blues (#1565C0), whites, and light grays (#F5F7FA)
  - Dark theme: Professional dark grays (#1E1E1E, #2B2B2B, #2D2D2D) with bright accent colors
- **Theme Toggle**: Dynamic switching between themes
- **Comprehensive Color Palette**:
  - Background colors (bg, bg_secondary, bg_panel, card)
  - Interactive colors (primary, success, warning, danger, secondary, purple)
  - Text colors (text, text_secondary, text_muted)
  - UI elements (border, hover, selected, shadow)
  - All colors have hover variants

**Key Methods:**
- `apply_theme(theme)`: Switch to light or dark theme
- `toggle_theme()`: Toggle between themes
- `get_theme()`: Get current theme name

### 2. Settings Persistence (`prg/config/settings.py`)

**New Features:**
- Theme preference saved to `prg_settings.json`
- Window geometry saved
- New `ui_preferences` section in settings
- Methods: `get_ui_preference()`, `set_ui_preference()`

### 3. Main Window Redesign (`prg/ui/main_window.py`)

**Professional Business UI:**
- Modern, clean interface suitable for enterprise software
- Consistent color scheme throughout
- Professional fonts (Segoe UI)
- Flat design with subtle borders

**Updated Components:**

**Menu Bar:**
- New "Вид" (View) menu with theme toggle option
- Menu colors respect current theme
- Clean, professional menu styling

**Top Panel:**
- Styled with theme colors
- Professional button styling using StyleManager
- Consistent spacing and padding
- Theme-aware labels and backgrounds

**Main Area:**
- Three-column layout with proper spacing
- **Left Panel (PRG Tree)**:
  - Modern Treeview styling
  - Professional heading colors
  - Theme-aware selection colors
- **Center Panel (Action Buttons)**:
  - Redesigned with StyleManager
  - Color-coded by function:
    - Green: Success actions (bind)
    - Blue: Primary actions (manual bind, edit)
    - Orange: Warning actions (unbind settlement)
    - Red: Danger actions (unbind consumer)
    - Purple: Special actions (auto-bind, calculate)
  - Hover effects on all buttons
- **Right Panel (Consumer Tree)**:
  - Matches PRG tree styling
  - Consistent theming

**Status Panel:**
- Modern, themed design
- Professional information display
- Themed text widget with proper contrast
- Context menu respects theme

**Dialogs:**
- Smart Search Dialog: Fully themed
- Manual Binding Dialog: Professional appearance
- Edit Shares Dialog: Consistent styling
- All dialogs use theme colors

### 4. Smart Search Dialog (`prg/ui/dialogs/smart_search_dialog.py`)

**Updates:**
- Accepts `style_manager` parameter
- Uses theme colors throughout
- Professional fonts and spacing
- Themed comboboxes and inputs
- Styled buttons via StyleManager

### 5. Main Application Entry Point (`main.py`)

**Changes:**
- Loads saved theme preference on startup
- Initializes StyleManager with saved theme
- Professional window title

## Usage

### Running the Application

```bash
python main.py
```

The application will:
1. Load the saved theme preference (defaults to light)
2. Apply the theme to all UI components
3. Display with professional, business-like appearance

### Toggling Themes

**Option 1: Menu**
- Go to **Вид** → **Темная тема** (or **Светлая тема**)
- Theme preference is saved automatically
- Restart recommended for full effect

**Option 2: Programmatic**
```python
app.toggle_theme()
```

### Theme Persistence

Themes are automatically saved to `prg_settings.json`:
```json
{
  "ui_preferences": {
    "theme": "dark",
    "window_geometry": "1500x900"
  }
}
```

## Color Schemes

### Light Theme
- **Background**: #F5F7FA (light blue-gray)
- **Panel**: #FFFFFF (white)
- **Primary**: #1565C0 (professional blue)
- **Text**: #1A1A1A (dark gray)
- **Success**: #2E7D32 (green)
- **Danger**: #C62828 (red)
- **Warning**: #EF6C00 (orange)

### Dark Theme
- **Background**: #1E1E1E (dark gray, not pure black)
- **Panel**: #2D2D2D (medium dark gray)
- **Primary**: #4A9EFF (bright blue)
- **Text**: #E8EAED (light gray)
- **Success**: #4CAF50 (bright green)
- **Danger**: #F44336 (bright red)
- **Warning**: #FF9800 (bright orange)

## Design Principles

1. **Professional Appearance**: Suitable for enterprise gas pipeline management
2. **Business-Like**: Clean, no-nonsense design with proper hierarchy
3. **Accessibility**: High contrast in both themes for readability
4. **Consistency**: All UI elements follow the same design language
5. **Modern**: Flat design with subtle shadows and borders
6. **Dark Theme Philosophy**: Dark gray (not black) for reduced eye strain

## File Structure

```
prg/
├── ui/
│   ├── styles.py           # Enhanced with theme support
│   ├── main_window.py      # Redesigned with professional UI
│   └── dialogs/
│       ├── smart_search_dialog.py  # Themed dialog
│       └── __init__.py
├── config/
│   └── settings.py         # Added UI preferences
└── ...

main.py                     # Loads theme on startup
prg_settings.json          # Stores theme preference
```

## Benefits

1. **Professional Look**: Suitable for business presentations and daily use
2. **User Choice**: Light theme for bright environments, dark for low-light
3. **Eye Comfort**: Dark theme reduces eye strain during long work sessions
4. **Persistent Preferences**: Theme choice remembered between sessions
5. **Consistent Experience**: All dialogs and windows respect the theme
6. **Modern Design**: Follows contemporary UI/UX best practices

## Technical Details

### Theme Application Flow

1. **Startup**:
   ```
   main.py → SettingsManager.load() → get_ui_preference('theme')
   → StyleManager(theme=saved_theme) → apply_theme()
   ```

2. **Toggle**:
   ```
   Menu Click → toggle_theme() → StyleManager.toggle_theme()
   → set_ui_preference('theme') → save()
   ```

3. **Component Rendering**:
   ```
   Each component → style_manager.colors → apply colors to widgets
   ```

### Color Usage Pattern

```python
# In any UI component:
colors = self.style_manager.colors

# Use themed colors:
frame = tk.Frame(parent, bg=colors['bg'])
label = tk.Label(frame, bg=colors['bg'], fg=colors['text'])
button = self.style_manager.create_button(frame, color='primary')
```

## Future Enhancements

Potential improvements for future versions:
1. Additional themes (e.g., high contrast, blue theme)
2. Custom accent color picker
3. Font size preferences
4. Icon themes
5. Theme preview before applying
6. Export/import theme configurations

## Notes

- **Restart Recommended**: While the theme toggle works, some UI elements may require an application restart for full effect
- **Settings File**: Located at `prg_settings.json` in the project root
- **Backwards Compatible**: Works with existing data files and configurations
- **No Data Impact**: Theme changes only affect UI appearance, not data or functionality

## Version

- **Version**: 7.4 Professional Edition
- **Theme System**: 1.0
- **Date**: 2026-02-03
