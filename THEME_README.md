# PRG Pipeline Manager v7.4 - Professional Edition

## Theme System Overview

The PRG Pipeline Manager now features a **professional, business-like interface** with full support for **light and dark themes**.

---

## Screenshots (Conceptual)

### Light Theme
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ File  View  Settings  Tools                                        [─][□][×]│
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                               │
│  [Открыть Excel файл]  📄 pipeline_data.xlsx      [Сохранить изменения]    │
│                                                                               │
├────────────────────────┬────────────┬──────────────────────────────────────┤
│  PRG (Tree)            │  Actions   │  Consumers (Tree)                     │
│  ├─ District 1         │            │  ├─ Population                        │
│  │  ├─ Settlement A    │ [Bind to   │  │  ├─ District 1                    │
│  │  │  └─ PRG-001      │  全 НП]     │  │  │  └─ Settlement A              │
│  │  └─ Settlement B    │            │  │  └─ District 2                    │
│  └─ District 2         │ [Search    │  └─ Organizations                    │
│     └─ Settlement C    │  Bind]     │     └─ District 1                    │
│                        │            │                                       │
│                        │ [Manual    │                                       │
│                        │  Bind]     │                                       │
│                        │            │                                       │
│                        │ [Unbind    │                                       │
│                        │  НП]       │                                       │
│                        │            │                                       │
│                        │ [Auto      │                                       │
│                        │  Bind]     │                                       │
│                        │            │                                       │
│                        │ [Edit      │                                       │
│                        │  Shares]   │                                       │
│                        │            │                                       │
│                        │ [Unbind    │                                       │
│                        │  Consumer] │                                       │
│                        │            │                                       │
│                        │ [Calculate │                                       │
│                        │  Load]     │                                       │
├────────────────────────┴────────────┴──────────────────────────────────────┤
│ v7.4 Professional Edition - Modular Architecture                            │
│ PRG: 150 | GRS: 12 | Consumers: 3,421                                      │
│ ┌─ Detail Information ────────────────────────────────────────────────────┐ │
│ │ Selected PRG:                                                            │ │
│ │ PRG ID: PRG-001                                                          │ │
│ │ District: District 1                                                     │ │
│ └──────────────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────┘

Colors: Light blue-gray background (#F5F7FA)
        White panels (#FFFFFF)
        Blue buttons (#1565C0)
        Dark text (#1A1A1A)
```

### Dark Theme
```
┌─────────────────────────────────────────────────────────────────────────────┐
│ File  View  Settings  Tools                                        [─][□][×]│
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                               │
│  [Открыть Excel файл]  📄 pipeline_data.xlsx      [Сохранить изменения]    │
│                                                                               │
├────────────────────────┬────────────┬──────────────────────────────────────┤
│  PRG (Tree)            │  Actions   │  Consumers (Tree)                     │
│  ├─ District 1         │            │  ├─ Population                        │
│  │  ├─ Settlement A    │ [Bind to   │  │  ├─ District 1                    │
│  │  │  └─ PRG-001      │  全 НП]     │  │  │  └─ Settlement A              │
│  │  └─ Settlement B    │            │  │  └─ District 2                    │
│  └─ District 2         │ [Search    │  └─ Organizations                    │
│     └─ Settlement C    │  Bind]     │     └─ District 1                    │
│                        │            │                                       │
│                        │ [Manual    │                                       │
│                        │  Bind]     │                                       │
│                        │            │                                       │
│                        │ [Unbind    │                                       │
│                        │  НП]       │                                       │
│                        │            │                                       │
│                        │ [Auto      │                                       │
│                        │  Bind]     │                                       │
│                        │            │                                       │
│                        │ [Edit      │                                       │
│                        │  Shares]   │                                       │
│                        │            │                                       │
│                        │ [Unbind    │                                       │
│                        │  Consumer] │                                       │
│                        │            │                                       │
│                        │ [Calculate │                                       │
│                        │  Load]     │                                       │
├────────────────────────┴────────────┴──────────────────────────────────────┤
│ v7.4 Professional Edition - Modular Architecture                            │
│ PRG: 150 | GRS: 12 | Consumers: 3,421                                      │
│ ┌─ Detail Information ────────────────────────────────────────────────────┐ │
│ │ Selected PRG:                                                            │ │
│ │ PRG ID: PRG-001                                                          │ │
│ │ District: District 1                                                     │ │
│ └──────────────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────┘

Colors: Dark gray background (#1E1E1E)
        Medium gray panels (#2D2D2D)
        Bright blue buttons (#4A9EFF)
        Light text (#E8EAED)
```

---

## Features

### 🎨 Professional Design
- Clean, business-like interface
- Modern flat design
- Consistent spacing and alignment
- Professional color scheme

### 🌓 Dual Themes
- **Light Theme**: Bright, professional blue-gray
- **Dark Theme**: Dark gray (not pure black) for reduced eye strain
- Easy toggle via menu
- Preference saved automatically

### 🎯 Semantic Colors
Each color has meaning:
- 🔵 **Primary** (Blue): Main actions
- 🟢 **Success** (Green): Positive actions
- 🟠 **Warning** (Orange): Caution actions
- 🔴 **Danger** (Red): Destructive actions
- 🔷 **Secondary** (Teal): Alternative actions
- 🟣 **Purple**: Special operations

### ⚡ Interactive Elements
- Hover effects on all buttons
- Visual feedback on interactions
- Smooth color transitions
- Professional appearance

---

## Quick Start

### Installation
```bash
# Install dependencies
pip install pandas openpyxl

# Run application
python main.py
```

### Toggle Theme
1. Click menu: **View** → **Dark Theme** (or **Light Theme**)
2. Theme is saved automatically
3. Restart for full effect

### Settings File
Theme saved to `prg_settings.json`:
```json
{
  "ui_preferences": {
    "theme": "dark"
  }
}
```

---

## Color Palettes

### Light Theme Colors
| Purpose | Color | Hex |
|---------|-------|-----|
| Background | 🔲 Light Blue-Gray | `#F5F7FA` |
| Panel | ⬜ White | `#FFFFFF` |
| Primary | 🔵 Professional Blue | `#1565C0` |
| Success | 🟢 Business Green | `#2E7D32` |
| Warning | 🟠 Attention Orange | `#EF6C00` |
| Danger | 🔴 Alert Red | `#C62828` |
| Text | ⬛ Dark Gray | `#1A1A1A` |

### Dark Theme Colors
| Purpose | Color | Hex |
|---------|-------|-----|
| Background | ⬛ Dark Gray | `#1E1E1E` |
| Panel | 🔲 Medium Gray | `#2D2D2D` |
| Primary | 🔵 Bright Blue | `#4A9EFF` |
| Success | 🟢 Bright Green | `#4CAF50` |
| Warning | 🟠 Bright Orange | `#FF9800` |
| Danger | 🔴 Bright Red | `#F44336` |
| Text | ⬜ Light Gray | `#E8EAED` |

---

## Button Guide

### Action Buttons (Center Panel)
```
┌────────────────────────┐
│  Bind to Settlement    │ 🟢 Green (Success)
├────────────────────────┤
│  Bind by Search        │ 🔷 Teal (Secondary)
├────────────────────────┤
│  Manual Bind           │ 🔵 Blue (Primary)
├────────────────────────┤
│  Unbind Settlement     │ 🟠 Orange (Warning)
├────────────────────────┤
│  Auto-Bind             │ 🟣 Purple (Special)
├────────────────────────┤
│  Edit Shares           │ 🔵 Blue (Primary)
├────────────────────────┤
│  Unbind Consumer       │ 🔴 Red (Danger)
├────────────────────────┤
│  Calculate Load        │ 🟣 Purple (Special)
└────────────────────────┘
```

---

## Architecture

### Style Management Flow
```
main.py
  │
  ├─> SettingsManager
  │     └─> Load theme from prg_settings.json
  │
  ├─> StyleManager(theme='light' or 'dark')
  │     ├─> Light theme colors
  │     ├─> Dark theme colors
  │     └─> Button styling methods
  │
  └─> PRGPipelineManager
        ├─> Apply theme to window
        ├─> Apply theme to menus
        ├─> Apply theme to panels
        ├─> Apply theme to buttons
        └─> Apply theme to dialogs
```

### Code Example
```python
# In any UI component:
colors = self.style_manager.colors

# Create themed widgets
frame = tk.Frame(parent, bg=colors['bg'])
label = tk.Label(frame, bg=colors['bg'], fg=colors['text'])

# Create styled button
button = self.style_manager.create_button(
    parent,
    text="Action",
    command=callback,
    color='primary'  # or 'success', 'danger', etc.
)
```

---

## Documentation

### Complete Guides
1. **THEME_REDESIGN.md** - Full technical documentation
2. **DESIGN_COMPARISON.md** - Before/after analysis
3. **THEME_QUICK_START.md** - Developer quick start
4. **IMPLEMENTATION_SUMMARY.md** - Project summary

### Code Documentation
- **prg/ui/styles.py** - StyleManager implementation
- **prg/ui/main_window.py** - Theme application examples
- **prg/config/settings.py** - Theme persistence

---

## Benefits

### For Users
✅ Professional appearance
✅ Comfortable viewing in any environment
✅ Reduced eye strain with dark mode
✅ Persistent theme preference
✅ Easy theme switching

### For Developers
✅ Centralized style management
✅ Easy to maintain
✅ Consistent design system
✅ Simple to extend
✅ Well-documented

---

## System Requirements

- **Python**: 3.7+
- **OS**: Windows, Linux, macOS
- **Dependencies**: pandas, openpyxl, tkinter (built-in)
- **Disk Space**: ~50 MB
- **Memory**: ~100 MB

---

## Version History

### v7.4 Professional Edition (2026-02-03)
- ✨ Added light and dark theme support
- 🎨 Redesigned entire UI with professional appearance
- 🎯 Implemented semantic color system
- 💾 Added theme persistence
- 📚 Created comprehensive documentation

### v7.3 (Previous)
- Modular architecture with dependency injection
- Smart search functionality
- Binding operations
- Load calculations

---

## Support

### Documentation
- Read `THEME_REDESIGN.md` for complete details
- Check `THEME_QUICK_START.md` for code examples
- Review `DESIGN_COMPARISON.md` for visual changes

### Common Issues
1. **Theme not changing**: Restart application
2. **Colors look wrong**: Check `prg_settings.json` theme value
3. **Buttons not hovering**: StyleManager not initialized

---

## Future Plans

Potential enhancements:
- [ ] Additional themes (high contrast, solarized)
- [ ] Custom theme creator
- [ ] Font size preferences
- [ ] Time-based theme switching
- [ ] Theme preview

---

## License & Credits

**Project**: PRG Pipeline Manager
**Version**: 7.4 Professional Edition
**Theme System**: v1.0
**Architecture**: Python + Tkinter
**Design**: Professional business UI with theme support

---

## Contact

For questions, issues, or suggestions about the theme system, refer to the documentation files in the project root.

---

**Built with ❤️ for professional gas pipeline management**
