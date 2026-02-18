# talen_extractor.py - Extract ALL "Talen" data from EIA-860 ZIP files
# Save as talen_extractor.py, then: python talen_extractor.py

import zipfile
import pandas as pd
import os
from pathlib import Path

# YOUR FILE PATH
ZIP_PATH = r"C:\Users\LENOVO\OneDrive\Máy tính\NAPE\eia8602024.zip"
OUTPUT_FILE = "talen_energy_data.xlsx"

def extract_talen_from_zip(zip_path):
    """Extract all Talen data from EIA-860 Excel files in ZIP"""
    
    print("🔍 Searching for Talen Energy in EIA-860 files...")
    all_talen_data = []
    
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        excel_files = [f for f in zip_ref.namelist() if f.endswith('.xlsx')]
        
        print(f"📁 Found {len(excel_files)} Excel files")
        
        for excel_file in excel_files:
            print(f"📖 Processing: {excel_file}")
            
            try:
                # Read Excel from ZIP directly
                with zip_ref.open(excel_file) as file:
                    xl = pd.ExcelFile(file)
                    
                    for sheet_name in xl.sheet_names:
                        try:
                            df = pd.read_excel(file, sheet_name=sheet_name)
                            
                            # Search ALL columns for "Talen" (case-insensitive)
                            mask = df.astype(str).apply(
                                lambda col: col.str.contains('Talen', case=False, na=False)
                            ).any(axis=1)
                            
                            talen_rows = df[mask]
                            if not talen_rows.empty:
                                talen_rows['File'] = excel_file
                                talen_rows['Sheet'] = sheet_name
                                all_talen_data.append(talen_rows)
                                print(f"   ✅ Found {len(talen_rows)} Talen rows in '{sheet_name}'")
                                
                        except Exception as e:
                            print(f"   ⚠️  Skipped sheet '{sheet_name}': {e}")
                            
            except Exception as e:
                print(f"   ❌ Error processing {excel_file}: {e}")
    
    return all_talen_data

def save_results(all_data, output_file):
    """Save all Talen findings to Excel with separate sheets"""
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, df in enumerate(all_data):
            sheet_name = f"Talen_Data_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"\n🎉 SAVED {len(all_data)} sheets to {output_file}")
    print("\n📊 Columns found with Talen: Operator_Name, Owner_Name, Operator_ID expected")

# MAIN EXECUTION
if __name__ == "__main__":
    if not os.path.exists(ZIP_PATH):
        print(f"❌ ZIP file not found: {ZIP_PATH}")
        exit()
    
    talen_data = extract_talen_from_zip(ZIP_PATH)
    
    if talen_data:
        save_results(talen_data, OUTPUT_FILE)
        print(f"\n✅ SUCCESS! Check '{OUTPUT_FILE}' for all Talen Energy data")
        print("\nKey files to check first: 2___Plant2024, 3_1_Generator2024, 4___Owner2024")
    else:
        print("❌ No Talen data found - check ZIP path or spelling")
