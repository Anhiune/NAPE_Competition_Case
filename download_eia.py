"""
download_eia.py — EIA 860/923 Zip File Downloader

Downloads EIA Form 860 and Form 923 zip files (containing xlsx spreadsheets)
for a configurable year range. These are direct downloads from eia.gov — no
API key is needed.

Also includes an optional function to query the EIA API v2 for international
electricity data (requires a free API key from https://www.eia.gov/opendata/).

Usage:
    python download_eia.py                     # Download all years (2014-2024)
    python download_eia.py --start 2020        # Download 2020-2024
    python download_eia.py --start 2018 --end 2022  # Download 2018-2022
    python download_eia.py --api-key YOUR_KEY  # Also fetch international electricity data
"""

import argparse
import os
import sys
import time
import urllib.request
import urllib.error
import json

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Year boundaries
DEFAULT_START_YEAR = 2014
DEFAULT_END_YEAR = 2024

# The "current" year threshold — files at or above this year use the non-archive URL.
# EIA typically keeps only the latest year in the non-archive path.
EIA_860_CURRENT_YEAR = 2024
EIA_923_CURRENT_YEAR = 2024


def build_eia860_url(year):
    """Build download URL for EIA Form 860."""
    if year >= EIA_860_CURRENT_YEAR:
        return f"https://www.eia.gov/electricity/data/eia860/xls/eia860{year}.zip"
    return f"https://www.eia.gov/electricity/data/eia860/archive/xls/eia860{year}.zip"


def build_eia923_url(year):
    """Build download URL for EIA Form 923."""
    if year >= EIA_923_CURRENT_YEAR:
        return f"https://www.eia.gov/electricity/data/eia923/xls/f923_{year}.zip"
    return f"https://www.eia.gov/electricity/data/eia923/archive/xls/f923_{year}.zip"


# ---------------------------------------------------------------------------
# Downloading
# ---------------------------------------------------------------------------

def download_file(url, dest_path):
    """Download a file from url to dest_path. Returns True on success."""
    if os.path.exists(dest_path):
        size_mb = os.path.getsize(dest_path) / (1024 * 1024)
        print(f"  Already exists: {os.path.basename(dest_path)} ({size_mb:.1f} MB) — skipping")
        return True

    print(f"  Downloading: {url}")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (NAPE-EIA-Downloader)"})
        with urllib.request.urlopen(req, timeout=120) as response:
            data = response.read()

        with open(dest_path, "wb") as f:
            f.write(data)

        size_mb = len(data) / (1024 * 1024)
        print(f"  Saved: {os.path.basename(dest_path)} ({size_mb:.1f} MB)")
        return True

    except urllib.error.HTTPError as e:
        print(f"  HTTP Error {e.code}: {url}")
        return False
    except urllib.error.URLError as e:
        print(f"  URL Error: {e.reason}")
        return False
    except Exception as e:
        print(f"  Error: {e}")
        return False


def download_eia_zips(base_dir, start_year, end_year):
    """Download EIA 860 and 923 zip files for the given year range."""
    years = list(range(start_year, end_year + 1))
    print(f"Downloading EIA data for years {start_year}-{end_year}...\n")

    success_count = 0
    fail_count = 0

    # --- EIA Form 860 ---
    print("=== EIA Form 860 (Annual Electric Generator Report) ===")
    for year in years:
        url = build_eia860_url(year)
        filename = f"eia860{year}.zip"
        dest = os.path.join(base_dir, filename)
        if download_file(url, dest):
            success_count += 1
        else:
            fail_count += 1
        time.sleep(0.5)  # Be polite to EIA servers

    print()

    # --- EIA Form 923 ---
    print("=== EIA Form 923 (Power Plant Operations Report) ===")
    for year in years:
        url = build_eia923_url(year)
        filename = f"f923_{year}.zip"
        dest = os.path.join(base_dir, filename)
        if download_file(url, dest):
            success_count += 1
        else:
            fail_count += 1
        time.sleep(0.5)

    print(f"\nDownload complete: {success_count} succeeded, {fail_count} failed")
    return fail_count == 0


# ---------------------------------------------------------------------------
# EIA API v2 — International Electricity (optional)
# ---------------------------------------------------------------------------

def fetch_international_electricity(api_key, base_dir):
    """Query EIA API v2 for international electricity data and save as JSON.

    API route: /v2/international/data
    Filter: productId facet for electricity

    Get a free API key at: https://www.eia.gov/opendata/
    """
    print("\n=== EIA API v2 — International Electricity Data ===")

    # First, discover the electricity product ID by querying the facet
    facet_url = (
        f"https://api.eia.gov/v2/international?api_key={api_key}"
    )
    print(f"  Querying international route metadata...")

    try:
        req = urllib.request.Request(facet_url, headers={"User-Agent": "NAPE-EIA-Client"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            meta = json.loads(resp.read().decode("utf-8"))

        # Show available sub-routes / facets
        response_data = meta.get("response", {})
        routes = response_data.get("routes", [])
        if routes:
            print("  Available sub-routes under /international:")
            for r in routes:
                print(f"    /{r['id']} — {r.get('name', '')}")

    except Exception as e:
        print(f"  Error querying metadata: {e}")

    # Fetch international electricity generation data
    # productId=2 is typically electricity; activityId=1 is generation
    data_url = (
        f"https://api.eia.gov/v2/international/data"
        f"?api_key={api_key}"
        f"&data[]=value"
        f"&facets[productId][]=2"
        f"&facets[activityId][]=1"
        f"&frequency=annual"
        f"&sort[0][column]=period"
        f"&sort[0][direction]=desc"
        f"&length=5000"
    )

    print(f"  Fetching international electricity generation data...")

    try:
        req = urllib.request.Request(data_url, headers={"User-Agent": "NAPE-EIA-Client"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode("utf-8"))

        records = result.get("response", {}).get("data", [])
        total = result.get("response", {}).get("total", 0)
        print(f"  Received {len(records)} records (total available: {total})")

        # Save to JSON file
        out_path = os.path.join(base_dir, "international_electricity.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2)
        print(f"  Saved to: {out_path}")

    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"  HTTP Error {e.code}: {body[:300]}")
    except Exception as e:
        print(f"  Error: {e}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Download EIA 860/923 zip files and optionally query the EIA API."
    )
    parser.add_argument(
        "--start", type=int, default=DEFAULT_START_YEAR,
        help=f"Start year (default: {DEFAULT_START_YEAR})"
    )
    parser.add_argument(
        "--end", type=int, default=DEFAULT_END_YEAR,
        help=f"End year (default: {DEFAULT_END_YEAR})"
    )
    parser.add_argument(
        "--api-key", type=str, default=None,
        help="EIA API key (for international electricity query). Get one at https://www.eia.gov/opendata/"
    )
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.abspath(__file__)) or "."

    # Download EIA 860/923 zips
    download_eia_zips(base_dir, args.start, args.end)

    # Optionally query international electricity data
    if args.api_key:
        fetch_international_electricity(args.api_key, base_dir)
    else:
        print("\nTip: Pass --api-key YOUR_KEY to also fetch international electricity data.")
        print("     Get a free key at: https://www.eia.gov/opendata/")

    print("\nDone.")


if __name__ == "__main__":
    main()
