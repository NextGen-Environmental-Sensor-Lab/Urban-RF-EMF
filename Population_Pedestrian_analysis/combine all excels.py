import pandas as pd
from tkinter import Tk, filedialog
import sys
import os

def main():
    # Hide the root Tk window
    root = Tk()
    root.withdraw()

    # 1. Select multiple source Excel files
    file_paths = filedialog.askopenfilenames(
        title="Select source Excel files",
        filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")]
    )

    if not file_paths:
        print("No files selected. Exiting.")
        sys.exit(0)

    # 2. Ask where to save total_paths.xlsx (default name)
    save_path = filedialog.asksaveasfilename(
        title="Save combined file as...",
        defaultextension=".xlsx",
        initialfile="total_paths.xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")]
    )

    if not save_path:
        print("No output file chosen. Exiting.")
        sys.exit(0)

    combined_df = None

    # for idx, path in enumerate(file_paths):
    #     print(f"Reading file {idx+1}/{len(file_paths)}: {os.path.basename(path)}")
    #     df = pd.read_excel(path)

    #     if combined_df is None:
    #         # First file: keep everything as is
    #         combined_df = df.copy()
    #     else:
    #         # Subsequent files:
    #         # "minus the first row" – here interpreted as:
    #         # skip the first data row (often the header row if reading differently).
    #         # If you want to include ALL rows, change df = df.iloc[1:] to df = df.
    #         df = df.iloc[1:]  # <-- this is where we drop the first row of EACH subsequent file
    #         combined_df = pd.concat([combined_df, df], ignore_index=True)

    for idx, path in enumerate(file_paths):
        print(f"Reading file {idx+1}/{len(file_paths)}: {os.path.basename(path)}")
        df = pd.read_excel(path)

        # Drop the first data row from EACH file
        df = df.iloc[1:]

        if combined_df is None:
            combined_df = df.copy()
        else:
            combined_df = pd.concat([combined_df, df], ignore_index=True)


    # 3. Save to total_paths.xlsx
    combined_df.to_excel(save_path, index=False)
    print(f"Done! Combined file saved to:\n{save_path}")

if __name__ == "__main__":
    main()
