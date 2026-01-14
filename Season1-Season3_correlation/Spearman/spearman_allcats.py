#!/usr/bin/env python3
"""
RF-EMF NYC — Spearman correlation matrix across multiple path Excel files

UPDATE (per your request):
- Output table is now **3 x 14**
  Rows = [census_block_population, census_block_population_wneighborblks, Pedestrian_mobility]
  Columns = for each of the 7 RF metrics: two adjacent columns [rho, p]
    e.g. "Broadcast rho", "Broadcast p", "Downlink rho", "Downlink p", ... (14 total)

What this script does
1) Interactively select multiple source Excel files.
2) Select (or create) an output Excel file.
3) Concatenate needed columns across ALL files (row-wise).
4) Compute Spearman rho + p-value for all X vs Y pairs (pairwise dropna).
5) Write the 3x14 table into a new sheet in the output workbook.

Requires: pandas, scipy, openpyxl
    pip install pandas scipy openpyxl
"""

import os
import re
from datetime import datetime
from typing import List, Tuple, Optional

import numpy as np
import pandas as pd
from scipy.stats import spearmanr

import tkinter as tk
from tkinter import filedialog


# -----------------------------
# Column lists
# -----------------------------
X_COLS = ["Broadcast", "Downlink", "Uplink", "WLAN", "TDD", "Total", "Total (RMS)"]
Y_COLS = ["census_block_population", "census_block_population_wneighborblks", "Pedestrian_mobility"]


def _safe_sheet_name(base: str, existing: List[str]) -> str:
    """Excel sheet names must be <=31 chars and cannot contain []:*?/\\ ."""
    name = re.sub(r"[\[\]\:\*\?\/\\]", "_", base)[:31].strip() or "Spearman"
    if name not in existing:
        return name
    for i in range(2, 1000):
        trial = (name[: (31 - len(f"_{i}"))] + f"_{i}")[:31]
        if trial not in existing:
            return trial
    return (name[:28] + "_999")[:31]


def _pick_source_files() -> List[str]:
    root = tk.Tk()
    root.withdraw()
    root.update()
    files = filedialog.askopenfilenames(
        title="Select source Excel files (RF-EMF paths)",
        filetypes=[("Excel files", "*.xlsx *.xlsm *.xls")],
    )
    root.destroy()
    return list(files) if files else []


def _pick_output_file() -> Optional[str]:
    root = tk.Tk()
    root.withdraw()
    root.update()
    out_path = filedialog.asksaveasfilename(
        title="Choose output Excel file",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="rfemf_spearman_matrix.xlsx",
    )
    root.destroy()
    return out_path if out_path else None


def _read_one_file(path: str) -> Tuple[pd.DataFrame, List[str]]:
    warnings = []
    try:
        df = pd.read_excel(path, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), [f"Failed to read '{os.path.basename(path)}': {e}"]

    df.columns = [str(c).strip() for c in df.columns]

    needed = [c for c in (X_COLS + Y_COLS) if c in df.columns]
    missing = [c for c in (X_COLS + Y_COLS) if c not in df.columns]

    if missing:
        warnings.append(
            f"{os.path.basename(path)}: missing columns -> {missing} (will still use available columns)"
        )

    if not needed:
        warnings.append(f"{os.path.basename(path)}: none of the required columns were found; skipping.")
        return pd.DataFrame(), warnings

    sub = df[needed].copy()

    for c in needed:
        sub[c] = pd.to_numeric(sub[c], errors="coerce")

    return sub, warnings


def _concat_sources(paths: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    all_warnings = []
    frames = []
    for p in paths:
        sub, warns = _read_one_file(p)
        all_warnings.extend(warns)
        if not sub.empty:
            frames.append(sub)

    if not frames:
        return pd.DataFrame(), all_warnings

    combined = pd.concat(frames, axis=0, ignore_index=True, sort=False)
    return combined, all_warnings


def _spearman_matrix_3x14(df: pd.DataFrame) -> pd.DataFrame:
    """
    Build a 3 x 14 table:
      rows = Y_COLS
      cols = for each X in X_COLS: two adjacent columns [X rho, X p]
    """
    # Create the column order: 14 columns total
    out_cols = []
    for x in X_COLS:
        out_cols.extend([f"{x} rho", f"{x} p"])

    rows = []
    for y in Y_COLS:
        row = {"Y": y}
        for x in X_COLS:
            rho_col = f"{x} rho"
            p_col = f"{x} p"

            if x not in df.columns or y not in df.columns:
                row[rho_col] = np.nan
                row[p_col] = np.nan
                continue

            pair = df[[x, y]].dropna()
            n = int(pair.shape[0])
            if n < 3:
                row[rho_col] = np.nan
                row[p_col] = np.nan
                continue

            rho, p = spearmanr(pair[x].to_numpy(), pair[y].to_numpy(), nan_policy="omit")
            row[rho_col] = float(rho) if rho is not None else np.nan
            row[p_col] = float(p) if p is not None else np.nan

        rows.append(row)

    result = pd.DataFrame(rows)
    # Put Y label first, then the 14 columns
    result = result[["Y"] + out_cols]
    return result


def _write_to_excel(out_path: str, table_3x14: pd.DataFrame, warnings: List[str], n_total_rows: int) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    base_sheet = f"Spearman_{timestamp}"

    if os.path.exists(out_path):
        with pd.ExcelWriter(out_path, engine="openpyxl", mode="a", if_sheet_exists="new") as writer:
            sheet_name = _safe_sheet_name(base_sheet, writer.book.sheetnames)
            table_3x14.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            start = len(table_3x14) + 3
            notes = pd.DataFrame(
                {
                    "Info": [
                        "RF-EMF NYC Spearman correlation table (combined across all selected files)",
                        f"Total concatenated rows (before pairwise dropna): {n_total_rows}",
                        f"X columns: {', '.join(X_COLS)}",
                        f"Y columns: {', '.join(Y_COLS)}",
                        "Each cell reports Spearman rho and p-value (pairwise non-NaN rows).",
                        "",
                        "Warnings / missing-columns notes:",
                    ]
                }
            )
            notes.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start, startcol=0)

            if warnings:
                pd.DataFrame({"Warning": warnings}).to_excel(
                    writer, index=False, sheet_name=sheet_name, startrow=start + len(notes) + 1
                )
    else:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            sheet_name = _safe_sheet_name(base_sheet, [])
            table_3x14.to_excel(writer, index=False, sheet_name=sheet_name, startrow=0)

            start = len(table_3x14) + 3
            notes = pd.DataFrame(
                {
                    "Info": [
                        "RF-EMF NYC Spearman correlation table (combined across all selected files)",
                        f"Total concatenated rows (before pairwise dropna): {n_total_rows}",
                        f"X columns: {', '.join(X_COLS)}",
                        f"Y columns: {', '.join(Y_COLS)}",
                        "Each cell reports Spearman rho and p-value (pairwise non-NaN rows).",
                        "",
                        "Warnings / missing-columns notes:",
                    ]
                }
            )
            notes.to_excel(writer, index=False, sheet_name=sheet_name, startrow=start, startcol=0)

            if warnings:
                pd.DataFrame({"Warning": warnings}).to_excel(
                    writer, index=False, sheet_name=sheet_name, startrow=start + len(notes) + 1
                )


def main() -> None:
    src_files = _pick_source_files()
    if not src_files:
        print("No source files selected. Exiting.")
        return

    out_path = _pick_output_file()
    if not out_path:
        print("No output file selected. Exiting.")
        return

    combined, warns = _concat_sources(src_files)
    if combined.empty:
        print("No usable data found across selected files.")
        if warns:
            print("\nWarnings:")
            for w in warns:
                print(" -", w)
        return

    table_3x14 = _spearman_matrix_3x14(combined)

    _write_to_excel(out_path, table_3x14, warns, n_total_rows=int(combined.shape[0]))

    print(f"✅ Wrote 3x14 Spearman table to: {out_path}")
    print(table_3x14.to_string(index=False))

    if warns:
        print("\n⚠️ Warnings:")
        for w in warns[:50]:
            print(" -", w)
        if len(warns) > 50:
            print(f" ... ({len(warns) - 50} more)")


if __name__ == "__main__":
    main()
