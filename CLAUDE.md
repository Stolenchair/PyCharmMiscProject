# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PRG Pipeline Manager (v7.4 Professional Edition) - A desktop application for managing gas pipeline (ПРГ/PRG) bindings to consumers (population and organizations). Built with Python/Tkinter using a modular layered architecture with dependency injection. Features a professional UI with light/dark theme support. Works with Excel files containing pipeline and consumer data.

## Key Dependencies

- **pandas**: Data manipulation and Excel I/O
- **openpyxl**: Excel file reading/writing
- **tkinter**: GUI framework (built-in with Python)

Install with: `python -m pip install pandas openpyxl`

## Running the Application

```bash
python main.py
```

The application will:
1. Load settings from `prg_settings.json` if it exists (including theme preference)
2. Launch the Tkinter GUI with professional themed interface
3. Wait for user to open an Excel file with PRG/consumer data
4. Toggle themes via View menu (Вид → Темная тема / Светлая тема)

## Code Architecture

The application uses a **layered architecture with dependency injection** for modularity and testability.

### Project Structure

```
prg/                             # Main package
├── models/                      # Data models (dataclasses)
│   ├── prg.py                   # PRGData - pipeline data
│   ├── consumer.py              # ConsumerData - consumer data
│   ├── binding.py               # PRGBinding - binding data
│   ├── grs.py                   # GRSData - GRS reference data
│   └── change.py                # Change tracking models
├── config/                      # Settings management
│   ├── settings.py              # SettingsManager
│   └── defaults.py              # Default column mappings
├── data/                        # Excel I/O operations
│   ├── excel_loader.py          # ExcelLoader - loads Excel data
│   └── parsers.py               # Parsing utilities
├── business/                    # Business logic services
│   ├── validation_service.py    # ValidationService
│   ├── binding_service.py       # BindingService
│   ├── calculation_service.py   # CalculationService
│   └── search_service.py        # SearchService
├── ui/                          # UI components and styling
│   ├── styles.py                # StyleManager - modern UI styling
│   └── dialogs/                 # UI dialog components
└── utils/                       # Utility functions
    ├── excel_utils.py           # Excel column conversions
    ├── parsers.py               # Data parsing
    ├── string_utils.py          # String utilities
    └── validators.py            # Validation utilities
```

### Main Classes and Responsibilities

| Class | Location | Purpose |
|-------|----------|---------|
| **SettingsManager** | `prg/config/settings.py` | Loads/saves configuration from `prg_settings.json` |
| **ExcelLoader** | `prg/data/excel_loader.py` | Loads Excel data into Python dictionaries with metadata |
| **ValidationService** | `prg/business/validation_service.py` | Validates data consistency, expense presence, location compatibility |
| **BindingService** | `prg/business/binding_service.py` | Manages PRG-consumer binding operations with change tracking |
| **CalculationService** | `prg/business/calculation_service.py` | Calculates PRG load metrics from consumer bindings |
| **SearchService** | `prg/business/search_service.py` | Filters and searches consumers by location/criteria |
| **StyleManager** | `prg/ui/styles.py` | Manages professional UI themes (light/dark), color palettes, and button styling |

### Data Models (Dataclasses)

All models are **immutable dataclasses** with Excel metadata:

- **PRGData** (`prg/models/prg.py`): Pipeline with location, GRS ID, and load metrics (QY_pop, QH_pop, QY_ind, QH_ind, Year_volume, Max_Hour)
- **ConsumerData** (`prg/models/consumer.py`): Consumer with type, name, location, PRG bindings, expenses (yearly/hourly), and GRS info (organizations only)
- **PRGBinding** (`prg/models/binding.py`): Single binding with PRG ID, share (0.0-1.0), and GRS name
- **GRSData** (`prg/models/grs.py`): GRS reference data with ID, name, and district
- **Change Models** (`prg/models/change.py`): Change tracking for audit trails (PRGLoadChange, ConsumerBindingChange)

### Dependency Injection Setup

Services are instantiated in `main.py` and injected into the main window:

```python
# Initialize services
settings_manager = SettingsManager()
excel_loader = ExcelLoader(settings_manager)
validation_service = ValidationService()
calculation_service = CalculationService(validation_service)
binding_service = BindingService(validation_service)
search_service = SearchService(validation_service)
style_manager = StyleManager()

# Inject into main window
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
```

### Configuration System

Settings stored in `prg_settings.json` define Excel sheet names and column mappings:
- `prg`: PRG sheet configuration
- `grs`: GRS sheet configuration
- `population`: Population consumers sheet configuration
- `organizations`: Organization consumers sheet configuration

Each section specifies:
- Sheet name
- Start row for data
- Column letters for each field (A, B, C, etc.)

Settings dialog accessible via menu allows customization without editing JSON.

### Key Operations

**Binding Operations** (BindingService)
- `bind_prg_to_settlement(prg_id, settlement, mo, consumers)`: Bind PRG to all consumers in a settlement with expense validation
- `bind_single_consumer(prg_id, consumer, force=False)`: Bind single consumer (force=True bypasses validation)
- `unbind_single_consumer(consumer)`: Remove all bindings from consumer
- `unbind_entire_settlement(settlement, mo, consumers)`: Remove bindings from all consumers in settlement
- `remove_prg_from_consumer(prg_id, consumer)`: Remove specific PRG binding while keeping others
- Returns `BindingResult` with success/skip/error counts and change list

**Load Calculations** (CalculationService)
- `calculate_prg_loads(prg_data, consumer_data, grs_data)`: Calculate load metrics for all PRGs
  - Iterates all consumers with valid expenses
  - Parses PRG bindings and weights by share percentage
  - Accumulates loads: QY_pop/QH_pop (population), QY_ind/QH_ind (organizations)
  - Calculates totals: Year_volume, Max_Hour
- `apply_loads_to_prg_data(prg_data, prg_loads)`: Updates PRG objects with computed loads
- Returns `CalculationResult` with prg_loads dict and statistics

**Search Operations** (SearchService)
- `smart_search_organizations(prg_id, mo, settlement, street_filter, consumers)`: Search by location and street pattern
- `find_consumers_by_location(mo, settlement, consumers)`: Find all consumers in location
- `find_prg_by_id(prg_id, prg_data)`: Direct PRG ID lookup
- `get_unique_districts(data)`: Extract all district names
- `get_settlements_by_district(district, data)`: Extract settlements for district
- `get_prg_ids_by_location(mo, settlement, prg_data)`: Extract PRG IDs for location
- `filter_consumers_by_criteria(consumers, has_binding, has_expenses, consumer_type)`: Multi-criteria filtering

**Validation Operations** (ValidationService)
- `has_expenses(consumer)`: Check if consumer has valid expense data
- `get_consumer_expenses(consumer)`: Returns dict with yearly/hourly expenses
- `find_unbound_prg(prg_data, consumer_data)`: Find PRGs without consumers in same location
- `find_unbound_consumers(consumer_data)`: Find consumers without PRG bindings
- `validate_binding_compatibility(prg, consumer)`: Verify district/settlement match
- `check_organization_grs_mismatches(consumer_data, grs_data)`: Detect GRS ID inconsistencies

**Data Persistence**
- `save_changes_to_excel(workbook_path, changes)`: Batch-writes all tracked changes to Excel
- Change tracking: Each operation creates change_dict with metadata (change_id, type, sheet, row, col, old_value, new_value, description)
- Changes accumulated in memory then saved in batch with error reporting per change

### Column Reference Format

Excel columns specified as letters (A, B, C, etc.) are converted to zero-based indices using `col_to_index`.

### Binding Format

Consumer bindings stored as strings in Excel:
- Format: `"PRG_ID1|share1|GRS_Name1;PRG_ID2|share2|GRS_Name2"` (pipe-separated within binding, semicolon-separated between bindings)
- Example: `"928|1|ГРС Прогресс-2;867|0.5|ГРС Хаштук"`
- Parsed by `parse_prg_bindings()` in `prg/data/parsers.py` → returns list of binding dictionaries
- Formatted by `format_prg_bindings()` in `prg/data/parsers.py` → converts list to string
- Shares should sum to ≤1.0 for validated binding, but force binding allows exceeding this
- Share format: Supports both comma and dot decimal separators (e.g., "0.5" or "0,5")

### UI State Management

- Tree views save/restore expanded state between refreshes
- `save_tree_state`/`restore_tree_state` preserve user's view
- Status bar shows: file path, data counts, unsaved changes
- Detail panel shows selected item info (right-click for copy/select all)

### Theme System (v7.4 Professional Edition)

The application features a comprehensive theme system with light and dark modes:

**StyleManager** (`prg/ui/styles.py`):
- Manages two professional themes: light (default) and dark
- Provides centralized color palettes with semantic naming
- Creates styled buttons with hover effects
- Applies consistent styling to all UI components

**Light Theme**:
- Professional blue-gray backgrounds (#F5F7FA, #E8ECF1)
- Clean white panels (#FFFFFF)
- Business blue primary (#1565C0)
- High contrast for readability

**Dark Theme**:
- Professional dark gray backgrounds (#1E1E1E, #2B2B2B, #2D2D2D)
- Bright accent colors for visibility (#4A9EFF, #4CAF50, #FF9800)
- Reduced eye strain for extended use
- High contrast text (#E8EAED)

**Theme Persistence**:
- Theme preference saved to `prg_settings.json` under `ui_preferences.theme`
- Automatically loads saved theme on startup
- Toggle via menu: View → Dark/Light Theme
- Changes require restart for full effect

**Usage in Code**:
```python
# Access colors
colors = self.style_manager.colors
frame = tk.Frame(parent, bg=colors['bg'])
label = tk.Label(frame, bg=colors['bg'], fg=colors['text'])

# Create themed buttons
button = self.style_manager.create_button(
    parent, text="Action", command=callback, color='primary'
)

# Toggle theme
new_theme = self.style_manager.toggle_theme()  # Returns 'light' or 'dark'
```

**Color Semantics**:
- `primary` (blue): Main actions, selections
- `success` (green): Positive actions, confirmations
- `warning` (orange): Caution actions, modifications
- `danger` (red): Destructive actions, deletions
- `secondary` (teal): Alternative actions, search
- `purple`: Special operations, calculations

**Documentation**:
- Full guide: `THEME_REDESIGN.md`
- Comparison: `DESIGN_COMPARISON.md`
- Quick start: `THEME_QUICK_START.md`

## Development Notes

### Excel File Structure Expected

The application expects Excel workbooks with specific sheets:
- PRG sheet: Pipeline data with district, settlement, PRG ID columns
- GRS sheet: Reference data for gas reduction stations
- Population sheet: Consumer data for population
- Organizations sheet: Consumer data for organizations

Column mappings are flexible via settings but data structure is rigid.

### Search Logic

Smart search (v7.4):
- Requires PRG selection in UI first
- Auto-fills district/settlement from selected PRG
- Dropdown filters: district → settlements → PRG IDs (cascading dropdowns)
- Manual input only for street name (uses regex pattern matching)
- Returns filtered consumer list for binding
- Implemented in `SearchService.smart_search_organizations()`

### Data Flow

**Loading Excel Data**:
```
Excel File → ExcelLoader.load_all_data()
  ├→ load_prg_data() → List of PRG dicts with metadata
  ├→ load_grs_data() → List of GRS dicts
  ├→ load_population_data() → List of consumer dicts
  └→ load_organization_data() → List of consumer dicts
Returns: {'prg': [...], 'grs': [...], 'consumers': [...]}
```

**Binding Flow**:
```
User Action → BindingService.bind_*()
  ├→ Parse existing bindings: parse_prg_bindings()
  ├→ Validate (unless force=True): ValidationService
  ├→ Calculate share: Check available capacity (1.0 - total_existing_shares)
  ├→ Create new binding: format_prg_bindings()
  ├→ Track change: Create change_dict with metadata
  └→ Return: BindingResult with changes[]
```

**Calculation Flow**:
```
User Action → CalculationService.calculate_prg_loads()
  ├→ For each consumer with expenses:
  │   ├→ Parse bindings: parse_prg_bindings()
  │   └→ For each binding:
  │       ├→ Load = expense × share
  │       └→ Accumulate by PRG and type (pop/org)
  ├→ Apply loads: apply_loads_to_prg_data()
  └→ Return: CalculationResult with prg_loads and stats
```

### Change Tracking

All modifications tracked before save:
- Change ID format: timestamp-based for uniqueness
- Types: PRG load changes vs regular consumer binding changes
- Batch save to Excel with error reporting per change

### Common Issues

**No expenses symbol (🚫)**: Consumer missing expense data, cannot be bound with validation (force=True bypasses)
**Yellow highlight (🟡)**: PRG has no consumers in same district+settlement, or consumer has no PRG binding
**Manual/force binding**: Bypasses all validation - use carefully for edge cases

## Design Patterns

The codebase uses several design patterns:
- **Dependency Injection**: Services injected at construction for loose coupling and testability
- **Service Layer Pattern**: Business logic isolated in services (validation, binding, calculation, search)
- **Data Transfer Objects**: Dataclasses for immutable data models with metadata
- **Result Pattern**: Operations return Result objects (BindingResult, CalculationResult, SearchResult) with aggregated outcomes
- **Repository Pattern**: ExcelLoader acts as data access abstraction
- **Change Log Pattern**: Change tracking with metadata for audit trails and batch persistence

## Important Implementation Details

### Excel Column Conversion
- Columns specified as letters (A, Z, AA, etc.) in settings
- `col_to_index()` in `prg/utils/excel_utils.py` converts to 0-based indices
- `index_to_col()` converts back for display/logging

### Numeric Parsing
- `parse_numeric_value()` handles both comma and dot decimal separators
- Excel data may have mixed formats (e.g., "1234,56" or "1234.56")
- All parsing functions in `prg/data/parsers.py` normalize to Python floats

### Change Tracking Format
Each change tracked as dictionary:
```python
{
  'change_id': 'unique_timestamp_id',
  'type': 'prg_load' or 'consumer_binding',
  'sheet_name': 'Excel sheet name',
  'row': Excel row number (1-based),
  'col': Excel column index (0-based),
  'old_value': previous value,
  'new_value': new value,
  'description': 'Human-readable description'
}
```

### Result Objects Pattern
Services return structured result objects instead of raw data:
- `BindingResult`: `success_count`, `skipped_count`, `error_count`, `changes[]`, `details[]`
- `CalculationResult`: `prg_loads{}`, `total_prg_count`, `processed_consumer_count`, `total_yearly_load`, `total_hourly_load`
- `SearchResult`: `matches[]`, `match_count`, `details{}`

This pattern provides:
- Consistent return format across services
- Aggregated statistics for UI display
- Detailed information for logging/debugging
- Type-safe return values

