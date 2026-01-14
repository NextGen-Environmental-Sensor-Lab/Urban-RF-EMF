#!/usr/bin/env python3
"""
Combine multiple RF-EMF Excel path files into ONE CSV for QGIS.

- Select multiple Excel files (interactive)
- Reads ONLY requested columns (exact names)
- Keeps ONLY rows with valid GPS Lat & GPS Lon
- Adds LayerName = filename without extension
- Writes ONE combined CSV
"""

import os
import sys
import pandas as pd

# Exact column names as they appear in your files:
KEEP_COLS = [
    "Date&Time",
    "Broadcast",
    "Downlink",
    "Uplink",
    "WLAN",            # <-- FIXED (was LAN)
    "TDD",
    "Total",
    "Total (RMS)",
    "GPS Lat",
    "GPS Lon",
    "GPS Altitude",
    "GPS HDOP",
    "GPS# Satellites",
    "GPS Speed",
    "Marker",
]

LAT_COL = "GPS Lat"
LON_COL = "GPS Lon"

HEADER_SCAN_MAX = 30

def pick_files_and_output():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except Exception as e:
        print("❌ Tkinter not available:", e)
        sys.exit(1)

    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="Select one or more Excel source files",
        filetypes=[("Excel files", "*.xlsx *.xls")],
    )
    if not files:
        messagebox.showinfo("No files selected", "No source files were selected. Exiting.")
        sys.exit(0)

    out_csv = filedialog.asksaveasfilename(
        title="Save combined CSV as…",
        defaultextension=".csv",
        filetypes=[("CSV", "*.csv")],
        initialfile="combined_paths_qgis_points.csv",
    )
    if not out_csv:
        messagebox.showinfo("No output selected", "No output file was selected. Exiting.")
        sys.exit(0)

    return list(files), out_csv

def to_numeric(s):
    return pd.to_numeric(s, errors="coerce")

def valid_lat_lon(df):
    lat = to_numeric(df[LAT_COL])
    lon = to_numeric(df[LON_COL])

    ok = (
        lat.notna()
        & lon.notna()
        & (lat != 0)
        & (lon != 0)
        & (lat >= -90) & (lat <= 90)
        & (lon >= -180) & (lon <= 180)
    )
    return ok, lat, lon

def find_header_row(excel_path, sheet_name=0, scan_max=HEADER_SCAN_MAX):
    """Try header rows 0..scan_max and return the first that contains all KEEP_COLS."""
    for hdr in range(scan_max + 1):
        try:
            df0 = pd.read_excel(excel_path, sheet_name=sheet_name, header=hdr, nrows=1, engine="openpyxl")
            cols = list(df0.columns)
            if all(c in cols for c in KEEP_COLS):
                return hdr
        except Exception:
            continue
    return None

def main():
    files, out_csv = pick_files_and_output()
    combined = []
    skipped = []

    print(f"✅ Selected {len(files)} file(s)\n")

    for path in files:
        layer_name = os.path.splitext(os.path.basename(path))[0]

        try:
            hdr = find_header_row(path, sheet_name=0)
            if hdr is None:
                # show columns at header=0 for debugging
                try:
                    cols0 = list(pd.read_excel(path, header=0, nrows=1, engine="openpyxl").columns)
                except Exception:
                    cols0 = []
                skipped.append((layer_name, f"Header row not found (0..{HEADER_SCAN_MAX}). Columns at header=0: {cols0}"))
                print(f"❌ {layer_name}: header row not found")
                continue

            df = pd.read_excel(path, header=hdr, usecols=KEEP_COLS, engine="openpyxl")

            total_rows = len(df)
            ok, lat, lon = valid_lat_lon(df)

            df = df.loc[ok].copy()
            kept_rows = len(df)

            if kept_rows == 0:
                skipped.append((layer_name, f"0 valid GPS points after filtering (read {total_rows} rows)."))
                print(f"⚠️  {layer_name}: read {total_rows} rows (header={hdr}), kept 0 valid GPS points")
                continue

            df[LAT_COL] = lat.loc[ok]
            df[LON_COL] = lon.loc[ok]

            # optional: parse Date&Time
            try:
                df["Date&Time"] = pd.to_datetime(df["Date&Time"], errors="coerce")
            except Exception:
                pass

            df.insert(0, "LayerName", layer_name)

            combined.append(df)
            print(f"✅ {layer_name}: read {total_rows} rows (header={hdr}), kept {kept_rows} point(s)")

        except Exception as e:
            skipped.append((layer_name, f"Exception: {e}"))
            print(f"❌ {layer_name}: exception: {e}")

    if not combined:
        print("\n❌ No valid data found in any selected files. Nothing written.")
        if skipped:
            print("\nDetails (why files were skipped):")
            for name, reason in skipped:
                print(f" - {name}: {reason}")
        sys.exit(1)

    out = pd.concat(combined, ignore_index=True)
    out.to_csv(out_csv, index=False, encoding="utf-8")
    print(f"\n✅ Wrote combined CSV with {len(out)} point(s): {out_csv}")

    if skipped:
        print("\n⚠️ Some files were skipped:")
        for name, reason in skipped:
            print(f" - {name}: {reason}")

if __name__ == "__main__":
    main()
