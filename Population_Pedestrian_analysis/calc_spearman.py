"""
interactively opens many source excel fls
calculates Spearman rho and p-values by: total, by borough, by environment
output to file 'spearman_analysis.xlsx'

"""


import os
import re
import pandas as pd
from scipy.stats import spearmanr
from tkinter import Tk, filedialog

# Column names in the Excel files
TOTAL_COL = 'Total (RMS)'
CB_POP = 'census_block_population'
CBN_POP = 'census_block_population_wneighborblks'
PM = 'Pedestrian_mobility'

# Codes to detect in filenames
BORO_CODES = {'BK', 'BX', 'M', 'Q', 'SI'}
ENV_CODES = {'C', 'R', 'G', 'T', 'I'}  # you mainly care about C, R, G, T

def compute_spearman(x, y):
    """
    Compute Spearman correlation between two Pandas Series.
    Returns (N_valid, rho, pval) where N_valid is number of non-NaN pairs.
    """
    valid = pd.concat([x, y], axis=1).dropna()
    N_valid = len(valid)
    if N_valid < 2:
        return N_valid, None, None
    rho, pval = spearmanr(valid.iloc[:, 0], valid.iloc[:, 1])
    return N_valid, rho, pval

def parse_filename_tokens(name_no_ext):
    """
    Parse borough and environment codes from filename (without extension).
    We split on underscores or spaces, then look for known codes.
    """
    tokens = re.split(r'[_\s]+', name_no_ext)
    boro = None
    env = None
    for t in tokens:
        if t in BORO_CODES and boro is None:
            boro = t
        if t in ENV_CODES and env is None:
            env = t
    return boro, env

def main():
    # --- 1. Select source Excel files interactively ---
    root = Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="Select Excel files for Spearman analysis",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_paths:
        print("No files selected. Exiting.")
        return

    first_dir = os.path.dirname(file_paths[0])
    output_path = os.path.join(first_dir, "spearman_analysis.xlsx")

    print("Selected files:")
    for fp in file_paths:
        print("  ", fp)
    print("Output will be written to:", output_path)
    print("-" * 80)

    # Containers
    path_results = []   # Sheet1: per file (path)
    all_rows = []       # For borough/env/all aggregation

    # --- 2. Process each file ---
    for path in file_paths:
        fname = os.path.basename(path)
        name_no_ext, _ = os.path.splitext(fname)
        print(f"Processing: {fname}")

        try:
            df = pd.read_excel(path)
        except Exception as e:
            print(f"  ❌ Could not read '{fname}': {e}")
            continue

        N_rows = len(df)
        if TOTAL_COL not in df.columns:
            print(f"  ⚠ Missing '{TOTAL_COL}' in '{fname}', skipping file.")
            continue

        # Make a uniform subset for aggregation, using .get so missing columns become NaN
        df_sub = pd.DataFrame({
            TOTAL_COL: df.get(TOTAL_COL),
            CB_POP: df.get(CB_POP),
            CBN_POP: df.get(CBN_POP),
            PM: df.get(PM),
        })

        boro, env = parse_filename_tokens(name_no_ext)
        df_sub['borough'] = boro
        df_sub['environment'] = env
        df_sub['source_file'] = name_no_ext
        all_rows.append(df_sub)

        # ---------- Path-level stats (Sheet1) ----------
        x = df_sub[TOTAL_COL]
        row = {
            'filename': name_no_ext,
            'N_rows': N_rows
        }

        # census_block_population
        if CB_POP in df_sub.columns:
            N_valid, rho, pval = compute_spearman(x, df_sub[CB_POP])
            row['N_valid_points'] = N_valid  # use this as N_valid_points
            row['cb_pop_rho'] = rho
            row['cb_pop_pval'] = pval
        else:
            print(f"  ⚠ '{CB_POP}' missing in '{fname}'")
            row['N_valid_points'] = 0
            row['cb_pop_rho'] = None
            row['cb_pop_pval'] = None

        # census_block_population_wneighborblks
        if CBN_POP in df_sub.columns:
            _, rho, pval = compute_spearman(x, df_sub[CBN_POP])
            row['cbn_pop_rho'] = rho
            row['cbn_pop_pval'] = pval
        else:
            print(f"  ⚠ '{CBN_POP}' missing in '{fname}'")
            row['cbn_pop_rho'] = None
            row['cbn_pop_pval'] = None

        # Pedestrian_mobility
        if PM in df_sub.columns:
            _, rho, pval = compute_spearman(x, df_sub[PM])
            row['pm_rho'] = rho
            row['pm_pval'] = pval
        else:
            print(f"  ⚠ '{PM}' missing in '{fname}'")
            row['pm_rho'] = None
            row['pm_pval'] = None

        path_results.append(row)

        print(f"  ✔ Completed {fname}")
        print("-" * 80)

    if not path_results:
        print("No valid path-level results produced. Exiting.")
        return

    # ---------- 3. Build combined DataFrame for aggregation ----------
    all_data = pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()

    # Helper to compute group stats
    def group_stats(df_group, label_value, label_name):
        """
        df_group: subset DataFrame
        label_value: borough/env/all_points
        label_name: 'borough' or 'environment' or 'group'
        Returns a dict.
        """
        N_rows = len(df_group)
        x = df_group[TOTAL_COL]

        row = {
            label_name: label_value,
            'N_rows': N_rows
        }

        # cb_pop
        if CB_POP in df_group.columns:
            N_valid, rho, pval = compute_spearman(x, df_group[CB_POP])
            row['N_valid_points'] = N_valid
            row['cb_pop_rho'] = rho
            row['cb_pop_pval'] = pval
        else:
            row['N_valid_points'] = 0
            row['cb_pop_rho'] = None
            row['cb_pop_pval'] = None

        # cbn_pop
        if CBN_POP in df_group.columns:
            _, rho, pval = compute_spearman(x, df_group[CBN_POP])
            row['cbn_pop_rho'] = rho
            row['cbn_pop_pval'] = pval
        else:
            row['cbn_pop_rho'] = None
            row['cbn_pop_pval'] = None

        # pm
        if PM in df_group.columns:
            _, rho, pval = compute_spearman(x, df_group[PM])
            row['pm_rho'] = rho
            row['pm_pval'] = pval
        else:
            row['pm_rho'] = None
            row['pm_pval'] = None

        return row

    # ---------- 4. Borough-level stats (Sheet2) ----------
    borough_results = []
    if not all_data.empty:
        # drop rows where borough is None
        for boro_code, df_boro in all_data.dropna(subset=['borough']).groupby('borough'):
            borough_results.append(
                group_stats(df_boro, boro_code, 'borough')
            )

    # ---------- 5. Environment-level stats (Sheet3) ----------
    env_results = []
    if not all_data.empty:
        for env_code, df_env in all_data.dropna(subset=['environment']).groupby('environment'):
            env_results.append(
                group_stats(df_env, env_code, 'environment')
            )

    # ---------- 6. All-points stats (Sheet4) ----------
    all_results = []
    if not all_data.empty:
        all_results.append(
            group_stats(all_data, 'all_points', 'group')
        )

    # ---------- 7. Write to Excel with 4 sheets ----------
    path_df = pd.DataFrame(path_results, columns=[
        'filename',
        'N_rows',
        'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ])

    borough_df = pd.DataFrame(borough_results, columns=[
        'borough',
        'N_rows',
        'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ]) if borough_results else pd.DataFrame(columns=[
        'borough', 'N_rows', 'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ])

    env_df = pd.DataFrame(env_results, columns=[
        'environment',
        'N_rows',
        'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ]) if env_results else pd.DataFrame(columns=[
        'environment', 'N_rows', 'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ])

    all_df = pd.DataFrame(all_results, columns=[
        'group',
        'N_rows',
        'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ]) if all_results else pd.DataFrame(columns=[
        'group', 'N_rows', 'N_valid_points',
        'cb_pop_rho', 'cb_pop_pval',
        'cbn_pop_rho', 'cbn_pop_pval',
        'pm_rho', 'pm_pval'
    ])

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            path_df.to_excel(writer, sheet_name='Sheet1', index=False)
            borough_df.to_excel(writer, sheet_name='Sheet2', index=False)
            env_df.to_excel(writer, sheet_name='Sheet3', index=False)
            all_df.to_excel(writer, sheet_name='Sheet4', index=False)

        print("✅ Spearman analysis saved to:", output_path)
    except Exception as e:
        print("❌ Error writing output Excel:", e)


if __name__ == "__main__":
    main()
