"""
generate_datasets.py — Dataset Generator

Reads xlsx files from EIA 860/923 zip archives, filters rows by company
operator names, and saves categorized datasets per company folder.
"""

import glob
import io
import os
import re
import zipfile

import pandas as pd

# ---------------------------------------------------------------------------
# Company keyword mappings (case-insensitive partial match)
# ---------------------------------------------------------------------------
COMPANY_KEYWORDS = {
    "Talen_Energy": [
        "TalenEnergy Susquehanna LLC",
        "TalenEnergy Martins Creek LLC",
        "Talen Montana LLC",
        "TalenEnergy Montour LLC",
        "Talen Energy",
        "Talen",
    ],
    "Vistra_Corp": [
        "Vistra Corp",
        "Luminant Generation Company LLC",
        "Luminant Energy",
        "Luminant",
        "Vistra",
    ],
    "Constellation_Energy": [
        "Constellation Energy Generation, LLC",
        "Constellation Energy",
        "Calpine",
    ],
    "NRG_Energy": [
        "NRG Energy Inc",
        "NRG Texas Power LLC",
        "NRG Generation",
        "NRG Power Marketing LLC",
    ],
}

# Filenames to skip (metadata / form templates)
SKIP_PATTERNS = [
    re.compile(r"EIA-860 Form\.xlsx", re.IGNORECASE),
    re.compile(r"Layout.*\.xlsx", re.IGNORECASE),
]

# ---------------------------------------------------------------------------
# Zip scanning
# ---------------------------------------------------------------------------

def scan_zips(base_dir):
    """Find all zip files and their xlsx contents.

    Returns: dict mapping zip_path -> list of xlsx entry names
    """
    pattern = os.path.join(base_dir, "*.zip")
    zip_paths = sorted(glob.glob(pattern))
    zip_contents = {}
    for zpath in zip_paths:
        try:
            with zipfile.ZipFile(zpath, "r") as zf:
                xlsx_names = [
                    n for n in zf.namelist()
                    if n.lower().endswith(".xlsx") and not _should_skip(n)
                ]
                zip_contents[zpath] = xlsx_names
        except (zipfile.BadZipFile, Exception) as e:
            print(f"  Warning: Could not read {zpath}: {e}")
    return zip_contents


def _should_skip(filename):
    """Check if a filename matches any skip pattern."""
    basename = os.path.basename(filename)
    return any(pat.search(basename) for pat in SKIP_PATTERNS)


# ---------------------------------------------------------------------------
# Category parsing
# ---------------------------------------------------------------------------

def parse_category_from_filename(xlsx_filename):
    """Extract a category name from an EIA-style xlsx filename.

    Patterns handled:
      - N___Name_Y20XX  -> "Name"   (e.g. 2___Plant_Y2024)
      - N_N_Name_Y20XX  -> "Name"   (e.g. 3_1_Generator_Y2024)
      - Keyword fallback for EIA 923 files
    """
    basename = os.path.basename(xlsx_filename)
    name_no_ext = os.path.splitext(basename)[0]

    # Pattern: digits___Name_Y20XX
    m = re.match(r"^\d+_{2,3}([A-Za-z]+(?:[A-Za-z_]*)?)_Y\d{4}$", name_no_ext)
    if m:
        return m.group(1).replace("_", "")

    # Pattern: digits_digits_Name_Y20XX
    m = re.match(r"^\d+_\d+_([A-Za-z]+(?:[A-Za-z_]*)?)_Y\d{4}$", name_no_ext)
    if m:
        return m.group(1).replace("_", "")

    # Keyword fallback for EIA 923 and other naming conventions
    fallback_keywords = [
        "FuelReceipts", "Generation", "Boiler", "Cooling",
        "Plant", "Generator", "Owner", "Utility",
        "Wind", "Solar", "EnergyStorage", "Multifuel",
        "Emissions", "Water", "FGD",
    ]
    for kw in fallback_keywords:
        if kw.lower() in name_no_ext.lower():
            return kw

    # Last resort: use the whole name
    return name_no_ext


# ---------------------------------------------------------------------------
# Reading xlsx from zip
# ---------------------------------------------------------------------------

def read_xlsx_from_zip(zip_ref, xlsx_name):
    """Read an xlsx file from inside a zip archive.

    Returns: list of (sheet_name, DataFrame) tuples
    """
    data = zip_ref.read(xlsx_name)
    buf = io.BytesIO(data)

    results = []
    try:
        xls = pd.ExcelFile(buf, engine="openpyxl")
        for sheet in xls.sheet_names:
            df = _read_sheet_with_header_detection(xls, sheet)
            if df is not None and not df.empty:
                results.append((sheet, df))
    except Exception as e:
        print(f"    Warning: Could not read {xlsx_name}: {e}")

    return results


def _read_sheet_with_header_detection(xls, sheet_name):
    """Read a sheet, detecting whether row 0 is a descriptive header."""
    try:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
    except Exception:
        return None

    if df.empty:
        return df

    # Heuristic: if every column name looks like a long sentence or
    # the first row has the same pattern as subsequent rows while column
    # names are descriptive, retry with header=1.
    col_str = " ".join(str(c) for c in df.columns)
    if _looks_like_description_row(col_str):
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=1)
        except Exception:
            pass

    return df


def _looks_like_description_row(col_str):
    """Heuristic check if column header string looks like a descriptive row."""
    # If the average word count per column name is high, it's likely descriptive
    parts = col_str.split()
    if len(parts) > 30:
        return True
    # Check for typical EIA description patterns
    description_signals = [
        "This field", "Enter the", "Report the", "Specify the",
        "Unnamed:", "described in",
    ]
    return any(sig.lower() in col_str.lower() for sig in description_signals)


# ---------------------------------------------------------------------------
# Row filtering
# ---------------------------------------------------------------------------

def filter_rows_for_company(df, keywords):
    """Filter DataFrame rows where any column contains any keyword (case-insensitive).

    Concatenates all column values per row into a single search string,
    then checks for keyword matches.
    """
    if df.empty:
        return df

    # Build a single search string per row from all columns
    str_df = df.astype(str)
    combined = str_df.apply(lambda row: " ".join(row.values), axis=1).str.lower()

    mask = pd.Series(False, index=df.index)
    for kw in keywords:
        mask = mask | combined.str.contains(kw.lower(), regex=False)

    return df[mask]


# ---------------------------------------------------------------------------
# Saving
# ---------------------------------------------------------------------------

def save_company_data(base_dir, company, category, df_list):
    """Concatenate DataFrames, drop duplicates, and save to company folder."""
    if not df_list:
        return

    combined = pd.concat(df_list, ignore_index=True)
    combined = combined.drop_duplicates()

    if combined.empty:
        return

    folder = os.path.join(base_dir, company)
    os.makedirs(folder, exist_ok=True)

    out_path = os.path.join(folder, f"{company}_{category}.xlsx")
    combined.to_excel(out_path, index=False, engine="openpyxl")
    print(f"    Saved {out_path} ({len(combined)} rows)")


# ---------------------------------------------------------------------------
# Main processing
# ---------------------------------------------------------------------------

def process_all_zips(base_dir, zip_contents):
    """Process all xlsx files across all zips, filter by company, and save."""
    # Accumulate: company -> category -> [DataFrames]
    accumulator = {
        company: {} for company in COMPANY_KEYWORDS
    }

    for zpath, xlsx_names in zip_contents.items():
        zip_name = os.path.basename(zpath)
        print(f"\nProcessing {zip_name} ({len(xlsx_names)} xlsx files)...")

        try:
            zf = zipfile.ZipFile(zpath, "r")
        except Exception as e:
            print(f"  Error opening {zip_name}: {e}")
            continue

        with zf:
            for xlsx_name in xlsx_names:
                category = parse_category_from_filename(xlsx_name)
                basename = os.path.basename(xlsx_name)
                print(f"  Reading {basename} -> category: {category}")

                sheets = read_xlsx_from_zip(zf, xlsx_name)

                for sheet_name, df in sheets:
                    for company, keywords in COMPANY_KEYWORDS.items():
                        filtered = filter_rows_for_company(df, keywords)
                        if not filtered.empty:
                            if category not in accumulator[company]:
                                accumulator[company][category] = []
                            accumulator[company][category].append(filtered)

    # Save accumulated data
    print("\nSaving company datasets...")
    for company, categories in accumulator.items():
        for category, df_list in categories.items():
            save_company_data(base_dir, company, category, df_list)

    # Summary
    print("\nSummary:")
    for company in COMPANY_KEYWORDS:
        cats = accumulator[company]
        total = sum(len(dfs) for dfs in cats.values())
        if total > 0:
            print(f"  {company}: {len(cats)} categories, {total} data chunks")
        else:
            print(f"  {company}: No matching data found")


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__)) or "."

    print("Scanning for ZIP files...")
    zip_contents = scan_zips(base_dir)

    total_xlsx = sum(len(v) for v in zip_contents.values())
    print(f"Found {len(zip_contents)} ZIP file(s) with {total_xlsx} XLSX file(s) total.\n")

    if total_xlsx == 0:
        print("No XLSX files found inside any ZIP. Nothing to process.")
        return

    process_all_zips(base_dir, zip_contents)
    print("\nDone.")


if __name__ == "__main__":
    main()
