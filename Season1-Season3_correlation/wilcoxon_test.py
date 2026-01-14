import os
import numpy as np
import pandas as pd
from tkinter import Tk, filedialog

try:
    from scipy.stats import wilcoxon
except ImportError as e:
    raise SystemExit("Missing dependency: scipy\nInstall with: pip install scipy") from e


# -----------------------------
# File picker
# -----------------------------
def pick_excel_file(title="Select Excel file"):
    root = Tk()
    root.withdraw()
    root.update()
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All files", "*.*")]
    )
    root.destroy()
    return path


# -----------------------------
# Stats
# -----------------------------
def paired_wilcoxon(x1, x3, alternative="two-sided", zero_method="wilcox"):
    """
    Wilcoxon signed-rank on paired values (measurement3 - measurement1).
    Returns dict with W and p plus diagnostics.
    """
    x1 = pd.to_numeric(pd.Series(x1), errors="coerce").to_numpy(dtype=float)
    x3 = pd.to_numeric(pd.Series(x3), errors="coerce").to_numpy(dtype=float)

    mask = np.isfinite(x1) & np.isfinite(x3)
    x1 = x1[mask]
    x3 = x3[mask]

    diffs = x3 - x1
    n_pairs = len(diffs)
    n_zero = int(np.sum(diffs == 0))

    if n_pairs == 0 or np.all(diffs == 0):
        return {
            "n_pairs": int(n_pairs),
            "n_zero_diff": int(n_zero),
            "median_diff_(3-1)": (0.0 if n_pairs > 0 else np.nan),
            "mean_diff_(3-1)": (0.0 if n_pairs > 0 else np.nan),
            "wilcoxon_W": np.nan,
            "p_value": np.nan,
            "note": "No valid non-zero paired differences (or no valid pairs)."
        }

    W, p = wilcoxon(x1, x3, alternative=alternative, zero_method=zero_method)

    return {
        "n_pairs": int(n_pairs),
        "n_zero_diff": int(n_zero),
        "median_diff_(3-1)": float(np.nanmedian(diffs)),
        "mean_diff_(3-1)": float(np.nanmean(diffs)),
        "wilcoxon_W": float(W),
        "p_value": float(p),
        "note": ""
    }


def require_columns(df, required):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "Missing required columns:\n  - " + "\n  - ".join(missing) +
            "\n\nColumns found:\n  - " + "\n  - ".join(map(str, df.columns))
        )


def run_tests(df, label="ALL"):
    """
    Runs Wilcoxon for mean, median, gmean comparing 1 vs 3.
    Returns rows for a results table.
    """
    rows = []
    for metric in ["mean", "median", "gmean"]:
        col1 = f"{metric}1"
        col3 = f"{metric}3"
        res = paired_wilcoxon(df[col1], df[col3], alternative="two-sided", zero_method="wilcox")
        res.update({
            "group": label,
            "metric": metric,
            "col_1": col1,
            "col_3": col3
        })
        rows.append(res)
    return rows


# -----------------------------
# Main
# -----------------------------
def main():
    in_path = pick_excel_file("Select Excel file with columns: mean1/gmean1/median1 ... mean3/gmean3/median3")
    if not in_path:
        print("No file selected. Exiting.")
        return

    # Read first sheet by default
    xls = pd.ExcelFile(in_path)
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(in_path, sheet_name=sheet_name)

    # Expected headers (exact)
    required = [
        "path location", "borough", "env type",
        "mean1", "gmean1", "median1", "n1",
        "mean3", "gmean3", "median3", "n3"
    ]
    require_columns(df, required)

    # Normalize the key text fields a bit (optional but helps grouping)
    for c in ["path location", "borough", "env type"]:
        df[c] = df[c].astype(str).str.strip()

    # Overall tests (paired by row/path)
    results_rows = []
    results_rows.extend(run_tests(df, label="ALL"))

    # Optional: subgroup tests (still paired within subgroup = rows belonging to that group)
    # Only meaningful if each subgroup has enough paired paths.
    for b in sorted(df["borough"].dropna().unique()):
        sub = df[df["borough"] == b]
        results_rows.extend(run_tests(sub, label=f"borough={b}"))

    for e in sorted(df["env type"].dropna().unique()):
        sub = df[df["env type"] == e]
        results_rows.extend(run_tests(sub, label=f"env type={e}"))

    results = pd.DataFrame(results_rows, columns=[
        "group", "metric", "col_1", "col_3",
        "n_pairs", "n_zero_diff",
        "median_diff_(3-1)", "mean_diff_(3-1)",
        "wilcoxon_W", "p_value", "note"
    ])

    # Create paired differences sheet for transparency
    diffs = df[["path location", "borough", "env type", "n1", "n3"]].copy()
    for metric in ["mean", "median", "gmean"]:
        diffs[f"{metric}1"] = pd.to_numeric(df[f"{metric}1"], errors="coerce")
        diffs[f"{metric}3"] = pd.to_numeric(df[f"{metric}3"], errors="coerce")
        diffs[f"{metric}_diff_(3-1)"] = diffs[f"{metric}3"] - diffs[f"{metric}1"]

    # Write output
    base, _ = os.path.splitext(in_path)
    out_path = base + "_wilcoxon_1vs3_bias.xlsx"

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        results.to_excel(writer, sheet_name="Wilcoxon_results", index=False)
        diffs.to_excel(writer, sheet_name="Paired_differences", index=False)

        readme = pd.DataFrame({
            "Item": [
                "What is being tested?",
                "Pairing assumption",
                "Direction of bias",
                "Interpretation"
            ],
            "Explanation": [
                "Wilcoxon signed-rank tests on paired path summaries: mean1 vs mean3, median1 vs median3, gmean1 vs gmean3.",
                "Rows are treated as matched paths (same route) measured twice; exact GPS point co-location is not required.",
                "Diff is (measurement3 - measurement1). Positive median_diff means measurement3 tends to be higher.",
                "Small p-value suggests a systematic shift across paths (overall bias). Consider also Spearman for rank stability."
            ]
        })
        readme.to_excel(writer, sheet_name="README", index=False)

    print("✅ Done.")
    print(f"Input:  {in_path} (sheet: {sheet_name})")
    print(f"Output: {out_path}")
    print("\nPreview:")
    print(results.head(20).to_string(index=False))


if __name__ == "__main__":
    main()
