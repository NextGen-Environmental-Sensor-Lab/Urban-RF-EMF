import os
import math
import pandas as pd
import numpy as np
from tkinter import Tk, filedialog, messagebox

# -----------------------------
# User configuration
# -----------------------------
CATEGORY_MAP = {
    "Broadcast": ["FM Radio (RMS)", "VHF 1, 2, 3 (RMS)", "UHF1 (RMS)", "UHF2 (RMS)", "UHF3 (RMS)"],
    "Downlink":  ["Mobile DL (RMS)", "Mobile DL (RMS).1", "Mobile DL (RMS).2", "Mobile DL (RMS).3", "Mobile DL (RMS).4"],
    "Uplink":    ["Mobile UL (RMS).1", "Mobile UL (RMS).2", "Mobile UL (RMS).3", "Mobile UL (RMS).4", "Mobile UL (RMS).5"],
    "WLAN":      ["ISM (RMS)", "WLAN (RMS)", "WLAN (RMS).1", "WLAN (RMS).2", "WLAN (RMS).3", "WLAN (RMS).4",
                  "WLAN (RMS).5", "WLAN (RMS).6", "WLAN (RMS).7", "WLAN (RMS).8", "WLAN (RMS).9", "WLAN (RMS).10"],
    "TDD":       ["TDD (RMS)", "TDD (RMS).1", "TDD (RMS).2", "TDD (RMS).3", "TDD (RMS).4",
                  "TDD (RMS).5", "TDD (RMS).6", "TDD (RMS).7", "TDD (RMS).8"],
}
CATEGORY_MAP["Total"] = sum(CATEGORY_MAP.values(), [])

CATS = ["Broadcast", "Downlink", "Uplink", "WLAN", "TDD", "Total"]

# -----------------------------
# Helpers
# -----------------------------
def excel_col_to_index(col_letters: str) -> int:
    """Convert Excel column letters (e.g., 'A', 'DP') to 0-based index."""
    col_letters = col_letters.strip().upper()
    n = 0
    for ch in col_letters:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid Excel column letters: {col_letters}")
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1  # 0-based

def safe_numeric_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Return numeric series for df[col]; missing column -> zeros; non-numeric -> NaN then filled 0."""
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    s = pd.to_numeric(df[col], errors="coerce")
    return s.fillna(0.0)

def rss_for_category(df: pd.DataFrame, cols: list[str]) -> pd.Series:
    """Row-wise RSS = sqrt(sum(x^2)) across given cols; missing cols treated as zeros."""
    # Accumulate sum of squares
    ss = pd.Series(0.0, index=df.index)
    for c in cols:
        s = safe_numeric_series(df, c)
        ss = ss + (s * s)
    return np.sqrt(ss)

# -----------------------------
# Main
# -----------------------------
def main():
    root = Tk()
    root.withdraw()

    src_files = filedialog.askopenfilenames(
        title="Select source path Excel files",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not src_files:
        return

    out_dir = filedialog.askdirectory(title="Select destination folder for output files")
    if not out_dir:
        return

    # Excel column slice DP..EA (inclusive)
    dp_idx = excel_col_to_index("DP")
    ea_idx = excel_col_to_index("EA")

    errors = []
    for src in src_files:
        try:
            df = pd.read_excel(src, engine="openpyxl")

            if df.shape[1] < 2:
                raise ValueError("Source file has fewer than 2 columns; cannot copy [A,B].")

            # [A, B] by position
            ab = df.iloc[:, 0:2].copy()

            # Aggregations
            agg = pd.DataFrame(index=df.index)
            for cat in CATS:
                agg[cat] = rss_for_category(df, CATEGORY_MAP[cat])

            # [DP..EA] by position if present; if not, create empty frame
            if df.shape[1] > dp_idx:
                tail = df.iloc[:, dp_idx:min(ea_idx + 1, df.shape[1])].copy()
            else:
                tail = pd.DataFrame(index=df.index)

            out_df = pd.concat([ab, agg, tail], axis=1)

            base = os.path.basename(src)
            name, ext = os.path.splitext(base)
            out_path = os.path.join(out_dir, f"{name}_agg.xlsx")

            # Write
            out_df.to_excel(out_path, index=False, engine="openpyxl")

        except Exception as e:
            errors.append(f"{os.path.basename(src)}: {e}")

    if errors:
        messagebox.showwarning(
            "Completed with warnings",
            "Some files could not be processed:\n\n" + "\n".join(errors)
        )
    else:
        messagebox.showinfo("Done", f"Processed {len(src_files)} file(s).\nOutput folder:\n{out_dir}")

if __name__ == "__main__":
    main()
