"""
Input files from ExpoM-4 RF sniffer

The script converts tab-separated sensor log files (e.g., .tsv, .txt, or .csv with tabs and LF 0x0A newlines) into clean, properly formatted Excel (.xlsx) files.
It standardizes headers, cleans corrupted or null data, parses GPS coordinates into decimal degrees, and removes invalid or placeholder entries.

1. File Selection (Interactive): Opens a file dialog to let you select one or more tab-separated input files. 
Prompts you to choose a destination folder for the converted Excel files.

2. Reading and Preserving Rows: Reads the entire file as raw text, splitting on \n to ensure that empty lines are preserved.
Splits each line on \t (tab) — commas are treated as part of the data, not as delimiters.
This preserves the exact row numbering, including blank lines, matching what you see in a text editor.

3. Header Construction (from lines 12 and 13): Builds the Excel column headers as follows:
The first two header names come from line 13 (row index 12).
From the third column onward: Combines the string in line 12 (row index 11) with the last substring in parentheses from line 13.
Example: Row 12: "Amplitude" Row 13: "Ch1 (mV)" → Header: "Amplitude (mV)"
If the cell in line 12 is empty, it uses the string from line 13 instead. 
If multiple headers end up identical, the script makes them unique by appending .1, .2, etc.

4. Data Range Selection: Skips lines 1 through 14 (header/metadata region).
Excludes the last three lines of the file (which for ExpoM files typically contain summary or junk data).
Only the middle section (line 15 to end -3) is written to Excel.

5. Data Cleaning & Conversion: For each cell in each data row: Replaces any \x00 (null byte) or carriage return with nothing.
Trims whitespace; treats truly empty strings as empty Excel cells (None).
Escapes any value starting with = to prevent Excel from misinterpreting it as a formula.
Converts numbers: Pure digits → integers. Decimals or scientific notation → floats. Anything else remains as text.

6. GPS Coordinate Parsing: Detects columns named “GPS Lat” and “GPS Lon” (case-insensitive).
Converts their values using the following logic: Compact format like DDMM.mmmmN/S or DDDMM.mmmmE/W.
Decimal degrees. Converts to signed decimal degrees: South and West are negative. Returns None (empty cell) if 
the field is blank or corrupted or the result is exactly 0.0 (invalid placeholder coordinate).

7. Output Writing: Creates a new Excel workbook (.xlsx) for each input file.
Writes: A single header row (as built in step 3). Cleaned, converted data rows beneath it.
Saves the result in the chosen destination folder with the same base filename as the input.

8. Output Characteristics: Tab-separated input handled precisely.Excel output free of repair errors or text-number warnings. True numeric and floating-point cells.
All nulls, zeros, and null bytes properly represented as empty. Properly parsed GPS coordinates in decimal degrees. Header naming fully standardized and unique.

Script created with ChatGPT by RToledo-Crow, ASRC CUNY, Oct 2025
"""

import os
import re
import pandas as pd
from tkinter import Tk, filedialog
from openpyxl import Workbook

# ---------- GPS Parsers ----------

def _parse_lat(value):
    """Parse latitude DDMM.mmmmN/S → signed decimal degrees."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "" or s == "\x00":
        return None
    s = s.replace(" ", "")
    last = s[-1].upper()
    if last in ("N", "S") and len(s) >= 4:
        body = s[:-1]
        try:
            deg = int(body[:2])
            mins = float(body[2:])
            sign = -1.0 if last == "S" else 1.0
            return sign * (deg + mins / 60.0)
        except Exception:
            pass
    try:
        return float(s)
    except Exception:
        return None


def _parse_lon(value):
    """Parse longitude DDDMM.mmmmE/W → signed decimal degrees."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if s == "" or s == "\x00":
        return None
    s = s.replace(" ", "")
    last = s[-1].upper()
    if last in ("E", "W") and len(s) >= 5:
        body = s[:-1]
        try:
            deg = int(body[:3])
            mins = float(body[3:])
            sign = -1.0 if last == "W" else 1.0
            return sign * (deg + mins / 60.0)
        except Exception:
            pass
    try:
        return float(s)
    except Exception:
        return None


# ---------- Utility Helpers ----------

def make_unique_headers(headers):
    seen = {}
    unique = []
    for h in headers:
        key = h if h is not None else ""
        if key in seen:
            seen[key] += 1
            unique.append(f"{key}.{seen[key]}")
        else:
            seen[key] = 0
            unique.append(key)
    return unique


def extract_parenthetical(text):
    if not text:
        return ""
    matches = re.findall(r'\(([^()]*)\)', text)
    return matches[-1] if matches else ""


def clean_cell(raw):
    if raw is None:
        return None
    s = str(raw)
    s = s.replace("\x00", "").replace("\r", "").strip()
    if s == "":
        return None
    if s.startswith("="):
        return "'" + s
    try:
        if s.isdigit():
            return int(s)
        return float(s)
    except ValueError:
        return s


def normalize_row_length(values, target_len):
    if len(values) < target_len:
        return values + [None] * (target_len - len(values))
    elif len(values) > target_len:
        return values[:target_len]
    return values


# ---------- Main ----------

def main():
    Tk().withdraw()
    print("Select one or more tab-separated files to process...")
    tsv_files = filedialog.askopenfilenames(
        title="Select Tab-Separated Files",
        filetypes=[("Tab-separated files", "*.tsv *.txt *.csv"), ("All files", "*.*")]
    )
    if not tsv_files:
        print("No files selected. Exiting.")
        return
    print(f"Selected {len(tsv_files)} files.")

    print("Select the destination folder...")
    dest_folder = filedialog.askdirectory(title="Select Destination Folder")
    if not dest_folder:
        print("No destination folder selected. Exiting.")
        return
    print(f"Destination folder: {dest_folder}")

    for tsv_file in tsv_files:
        try:
            print(f"\nProcessing: {tsv_file}")
            with open(tsv_file, "r", encoding="utf-8", newline="") as f:
                lines = f.read().split("\n")
            rows = [line.split("\t") for line in lines]

            if len(rows) < 15 + 2:
                print(f"⚠️ {tsv_file} has too few rows to process. Skipping.")
                continue

            row12 = rows[11] if len(rows) > 11 else []
            row13 = rows[12] if len(rows) > 12 else []

            hdr = []
            for i in range(0, 2):
                val13 = row13[i] if i < len(row13) else ""
                hdr.append((val13 or "").replace("\x00", "").strip())

            max_cols = max(len(row12), len(row13))
            for i in range(2, max_cols):
                val12 = (row12[i] if i < len(row12) else "").replace("\x00", "").strip()
                val13 = (row13[i] if i < len(row13) else "").replace("\x00", "").strip()
                paren = extract_parenthetical(val13)
                if val12:
                    hdr_val = f"{val12} ({paren})" if paren else val12
                else:
                    hdr_val = val13
                hdr.append(hdr_val)

            headers = make_unique_headers(hdr)
            data_rows = rows[14:-3]

            wb = Workbook()
            ws = wb.active
            ws.append(headers)
            ncols = len(headers)

            gps_lat_idx = next((i for i, h in enumerate(headers) if h.strip().lower() == "gps lat"), None)
            gps_lon_idx = next((i for i, h in enumerate(headers) if h.strip().lower() == "gps lon"), None)

            for raw_row in data_rows:
                cleaned = [clean_cell(cell) for cell in raw_row]
                cleaned = normalize_row_length(cleaned, ncols)

                # Handle GPS Lat/Lon parsing and skip zero values
                if gps_lat_idx is not None:
                    lat = _parse_lat(cleaned[gps_lat_idx])
                    cleaned[gps_lat_idx] = None if (lat is None or abs(lat) < 1e-9) else lat
                if gps_lon_idx is not None:
                    lon = _parse_lon(cleaned[gps_lon_idx])
                    cleaned[gps_lon_idx] = None if (lon is None or abs(lon) < 1e-9) else lon

                # Remove cells containing \x00
                cleaned = [None if (isinstance(v, str) and "\x00" in v) else v for v in cleaned]

                ws.append(cleaned)

            base_name = os.path.splitext(os.path.basename(tsv_file))[0]
            output_file = os.path.join(dest_folder, f"{base_name}.xlsx")
            wb.save(output_file)
            print(f"✅ Converted -> {output_file}")

        except Exception as e:
            print(f"❌ Error converting {tsv_file}: {e}")

    print("\nAll conversions complete.")


if __name__ == "__main__":
    main()
