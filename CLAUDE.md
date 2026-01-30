# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

PRG Pipeline Manager (v7.3 FINAL) - A desktop application for managing gas pipeline (ПРГ/PRG) bindings to consumers (population and organizations). Built with Python/Tkinter and works with Excel files containing pipeline and consumer data.

## Key Dependencies

- **pandas**: Data manipulation and Excel I/O
- **openpyxl**: Excel file reading/writing
- **tkinter**: GUI framework (built-in with Python)

Install with: `python -m pip install pandas openpyxl`

## Running the Application

```bash
python script.py
```

The application will:
1. Load settings from `prg_settings.json` if it exists
2. Launch the Tkinter GUI
3. Wait for user to open an Excel file with PRG/consumer data

## Code Architecture

### Main Classes

**PRGPipelineManager** (lines 18-4776)
- Primary application controller and UI manager
- Manages all data structures: `prg_data`, `grs_data`, `consumer_data`, `changes`
- Contains 100+ methods organized by functionality
- Lifecycle: `__init__` → `setup_ui` → `run` (mainloop)

**SmartSearchDialog** (lines 4780-5012)
- Modal dialog for smart search with dropdown filters
- Used to find consumers based on district/settlement/PRG ID
- Returns search criteria for binding operations

### Data Model

The application works with three main data types loaded from Excel sheets:

1. **PRG Data** (load_prg_data:1575)
   - Pipeline/gas reduction station data
   - Key fields: district (MO), settlement, PRG ID, GRS ID
   - Load columns: QY_pop, QH_pop, QY_ind, QH_ind, Year_volume, Max_hour

2. **GRS Data** (load_grs_data:1867)
   - Gas Reduction Station reference data
   - Key fields: GRS ID, GRS name

3. **Consumer Data** (load_population_data:2074, load_organization_data:2132)
   - Two types: population and organizations
   - Key fields: district (MO), settlement, code, expenses (yearly/hourly)
   - Organizations also have GRS ID and name
   - Each consumer can have bindings to multiple PRGs with shares

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

**Binding Operations**
- `bind_prg_to_settlement`: Bind PRG to all consumers in a settlement (with expense validation)
- `bind_by_search`: Smart search-based binding using filters
- `bind_manually`: Manual forced binding without validation
- `unbind_single_consumer`: Remove single consumer binding
- `unbind_entire_settlement`: Remove all consumer bindings for a settlement

**Load Calculations**
- `calculate_prg_load`: Calculate pipeline load from consumer bindings
- Uses yearly expenses and shares to compute PRG loads
- Saves results back to Excel PRG sheet

**Data Persistence**
- `save_changes_to_excel`: Writes all tracked changes back to Excel
- `save_prg_load_change`: Handles PRG load data updates
- `save_regular_change`: Handles consumer binding updates
- Changes tracked in `self.changes` dictionary with change IDs

**GRS Validation**
- `check_organization_grs`: Validates organization GRS references
- Checks for empty GRS fields and mismatches with actual bindings
- Offers CSV export of mismatches

### Column Reference Format

Excel columns specified as letters (A, B, C, etc.) are converted to zero-based indices using `col_to_index`.

### Binding Format

Consumer bindings stored as strings:
- Format: `"PRG_ID1:0.5;PRG_ID2:0.3"` (PRG ID:share pairs separated by semicolons)
- Parsed by `parse_prg_bindings`
- Formatted by `format_prg_bindings`
- Shares should sum to ≤1.0 but manual binding allows exceeding this

### UI State Management

- Tree views save/restore expanded state between refreshes
- `save_tree_state`/`restore_tree_state` preserve user's view
- Status bar shows: file path, data counts, unsaved changes
- Detail panel shows selected item info (right-click for copy/select all)

## Development Notes

### Excel File Structure Expected

The application expects Excel workbooks with specific sheets:
- PRG sheet: Pipeline data with district, settlement, PRG ID columns
- GRS sheet: Reference data for gas reduction stations
- Population sheet: Consumer data for population
- Organizations sheet: Consumer data for organizations

Column mappings are flexible via settings but data structure is rigid.

### Search Logic

Smart search (v7.3):
- Requires PRG selection in UI first
- Auto-fills district/settlement from selected PRG
- Dropdown filters: district → settlements → PRG IDs
- Manual input only for street name
- Returns filtered consumer list for binding

### Change Tracking

All modifications tracked before save:
- Change ID format: timestamp-based for uniqueness
- Types: PRG load changes vs regular consumer binding changes
- Batch save to Excel with error reporting per change

### Common Issues

**No expenses symbol (🚫)**: Consumer missing expense data, cannot be bound with validation
**Yellow highlight (🟡)**: PRG has no consumers in same district+settlement, or consumer has no PRG binding
**Manual binding**: Bypasses all validation - use carefully for edge cases

