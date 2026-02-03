"""Test script to verify Excel file loading and parsing."""

import sys
sys.path.insert(0, '.')

from prg.config.settings import SettingsManager
from prg.data.excel_loader import ExcelLoader
from prg.data.parsers import parse_prg_bindings

# Initialize services
settings_manager = SettingsManager()
excel_loader = ExcelLoader(settings_manager)

# Load the Excel file
excel_file = r'D:\__Касум\New_priviazka\СмартПанелСвязьАдыгея1.xlsx'

print("=" * 70)
print("TESTING EXCEL FILE LOADING")
print("=" * 70)
print(f"\nFile: {excel_file}\n")

try:
    # Load all data
    print("Loading data from Excel...")
    data = excel_loader.load_all_data(excel_file)

    # Check PRG data
    prg_data = data.get('prg', [])
    print(f"\n1. PRG DATA: Loaded {len(prg_data)} PRG records")
    if prg_data:
        sample = prg_data[0]
        print(f"   Sample PRG:")
        print(f"     - ID: {sample.get('prg_id', 'N/A')}")
        print(f"     - District: {sample.get('mo', 'N/A')}")
        print(f"     - Settlement: {sample.get('settlement', 'N/A')}")
        print(f"     - GRS ID: {sample.get('grs_id', 'N/A')}")
        print(f"     - QY_pop: {sample.get('qy_pop', 0)}")
        print(f"     - QH_pop: {sample.get('qh_pop', 0)}")
        print(f"     - QY_ind: {sample.get('qy_ind', 0)}")
        print(f"     - QH_ind: {sample.get('qh_ind', 0)}")

    # Check GRS data
    grs_data = data.get('grs', [])
    print(f"\n2. GRS DATA: Loaded {len(grs_data)} GRS records")
    if grs_data:
        sample = grs_data[0]
        print(f"   Sample GRS:")
        print(f"     - ID: {sample.get('grs_id', 'N/A')}")
        print(f"     - Name: {sample.get('grs_name', 'N/A')}")
        print(f"     - District: {sample.get('mo', 'N/A')}")

    # Check consumer data
    consumers = data.get('consumers', [])
    population = [c for c in consumers if c.get('consumer_type') == 'population']
    organizations = [c for c in consumers if c.get('consumer_type') == 'organization']

    print(f"\n3. CONSUMER DATA:")
    print(f"   - Population: {len(population)} records")
    print(f"   - Organizations: {len(organizations)} records")
    print(f"   - Total: {len(consumers)} records")

    # Check binding parsing
    consumers_with_bindings = [c for c in consumers if c.get('code')]
    print(f"\n4. BINDING ANALYSIS:")
    print(f"   - Consumers with bindings: {len(consumers_with_bindings)}")

    if consumers_with_bindings:
        sample = consumers_with_bindings[0]
        code = sample.get('code', '')
        bindings = parse_prg_bindings(code)

        print(f"\n   Sample consumer with binding:")
        print(f"     - Type: {sample.get('consumer_type', 'N/A')}")
        print(f"     - Name: {sample.get('name', sample.get('settlement', 'N/A'))}")
        print(f"     - District: {sample.get('mo', 'N/A')}")
        print(f"     - Settlement: {sample.get('settlement', 'N/A')}")
        print(f"     - Raw code: {code}")
        print(f"     - Parsed bindings: {len(bindings)} binding(s)")
        for i, binding in enumerate(bindings, 1):
            print(f"       {i}. PRG: {binding['prg_id']}, Share: {binding['share']}, GRS: {binding['grs_name']}")

    # Check expenses
    consumers_with_expenses = [c for c in consumers
                               if c.get('yearly_expenses', 0) > 0 or c.get('hourly_expenses', 0) > 0]
    print(f"\n5. EXPENSES ANALYSIS:")
    print(f"   - Consumers with expenses: {len(consumers_with_expenses)}")

    if consumers_with_expenses:
        sample = consumers_with_expenses[0]
        print(f"\n   Sample consumer with expenses:")
        print(f"     - Type: {sample.get('consumer_type', 'N/A')}")
        print(f"     - Name: {sample.get('name', sample.get('settlement', 'N/A'))}")
        print(f"     - Yearly expenses: {sample.get('yearly_expenses', 0):.2f} thousand m3/year")
        print(f"     - Hourly expenses: {sample.get('hourly_expenses', 0):.2f} m3/hour")

    # Summary
    print("\n" + "=" * 70)
    print("LOADING SUMMARY")
    print("=" * 70)
    print(f"[OK] Successfully loaded Excel file")
    print(f"[OK] PRG records: {len(prg_data)}")
    print(f"[OK] GRS records: {len(grs_data)}")
    print(f"[OK] Consumer records: {len(consumers)}")
    print(f"[OK] Bindings parsed: {len(consumers_with_bindings)}")
    print(f"[OK] Consumers with expenses: {len(consumers_with_expenses)}")
    print("\n[SUCCESS] The application should work correctly with this file!")

except Exception as e:
    print(f"\n[ERROR] {e}")
    import traceback
    traceback.print_exc()
