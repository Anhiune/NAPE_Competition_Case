"""
target.py — Zip Scanner & Report Generator

Scans the current directory for .zip files, identifies .xlsx contents
inside each zip, prints a summary, and writes a structured Similarname.md report.
"""

import glob
import os
import zipfile


def scan_zip_files(base_dir):
    """Find all .zip files in the given directory."""
    pattern = os.path.join(base_dir, "*.zip")
    return sorted(glob.glob(pattern))


def list_xlsx_in_zips(zip_paths):
    """For each zip file, list the .xlsx entries with their compressed sizes.

    Returns: dict mapping zip_path -> list of (xlsx_name, size_bytes)
    """
    results = {}
    for zpath in zip_paths:
        xlsx_entries = []
        try:
            with zipfile.ZipFile(zpath, "r") as zf:
                for info in zf.infolist():
                    if info.filename.lower().endswith(".xlsx") and not info.is_dir():
                        xlsx_entries.append((info.filename, info.file_size))
        except (zipfile.BadZipFile, Exception) as e:
            print(f"  Warning: Could not read {zpath}: {e}")
        results[zpath] = xlsx_entries
    return results


def format_size(size_bytes):
    """Format byte size into a human-readable string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.1f} MB"


def write_report(base_dir, results):
    """Write Similarname.md with a structured report of zip contents."""
    report_path = os.path.join(base_dir, "Similarname.md")
    total_zips = len(results)

    lines = []
    lines.append("# ZIP File Contents Report\n")
    lines.append(f"**Total ZIP files found:** {total_zips}\n")

    for zpath, xlsx_entries in results.items():
        zip_name = os.path.basename(zpath)
        lines.append(f"## {zip_name}\n")

        if not xlsx_entries:
            lines.append("_No XLSX files found in this archive._\n")
            continue

        lines.append(f"**XLSX files:** {len(xlsx_entries)}\n")
        lines.append("| # | XLSX File | Size |")
        lines.append("|---|-----------|------|")
        for i, (fname, size) in enumerate(xlsx_entries, 1):
            lines.append(f"| {i} | {fname} | {format_size(size)} |")
        lines.append("")

    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    return report_path


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__)) or "."

    print("Scanning for ZIP files...")
    zip_paths = scan_zip_files(base_dir)
    print(f"Found {len(zip_paths)} ZIP file(s).\n")

    if not zip_paths:
        print("No ZIP files found. Nothing to report.")
        return

    results = list_xlsx_in_zips(zip_paths)

    # Print console summary
    for zpath, xlsx_entries in results.items():
        zip_name = os.path.basename(zpath)
        print(f"  {zip_name}: {len(xlsx_entries)} XLSX file(s)")

    print()

    # Write report
    report_path = write_report(base_dir, results)
    print(f"Report written to: {report_path}")


if __name__ == "__main__":
    main()
