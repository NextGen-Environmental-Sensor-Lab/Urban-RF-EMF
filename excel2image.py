import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import LogNorm
from tkinter import Tk, filedialog
import numpy as np
import os
import re

# --- Step 1: Select multiple Excel files ---
Tk().withdraw()
file_paths = filedialog.askopenfilenames(
    title="Select one or more Excel files",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if not file_paths:
    print("No files selected. Exiting.")
    exit()

# --- Step 2: Compute global min/max for shared log scale ---
global_min, global_max = np.inf, -np.inf
valid_files = []

for file_path in file_paths:
    df = pd.read_excel(file_path).iloc[:, 2:41]  # Columns 3–41
    df = df.apply(pd.to_numeric, errors="coerce").mask(df <= 0)

    if df.isna().all().all():
        print(f"⚠️ Skipping {os.path.basename(file_path)} (no valid data)")
        continue

    vmin = np.nanmin(df.values)
    vmax = np.nanmax(df.values)
    global_min = min(global_min, vmin)
    global_max = max(global_max, vmax)
    valid_files.append(file_path)

if not valid_files:
    print("No valid numeric data found.")
    exit()

print(f"Global color scale: min={global_min:.3g}, max={global_max:.3g}")

# --- Step 3: Generate heatmap for each file (no text/ticks) ---
for file_path in valid_files:
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_dir = os.path.dirname(file_path)
    output_path = os.path.join(output_dir, base_name + ".svg")

    df = pd.read_excel(file_path).iloc[:, 2:41]
    df = df.apply(pd.to_numeric, errors="coerce").mask(df <= 0)

    # Plot heatmap
    fig, ax = plt.subplots(figsize=(10, 3))
    ax.imshow(
        df.values,
        cmap="viridis",
        aspect="auto",
        interpolation="nearest",
        norm=LogNorm(vmin=global_min, vmax=global_max)
    )

    # Remove ticks, labels, borders, and spines
    ax.axis("off")

    # Save as SVG
    plt.savefig(output_path, format="svg", bbox_inches="tight", pad_inches=0)
    plt.close()
    print(f"✅ Saved heatmap: {output_path}")

# --- Step 4: Generate single color key SVG ---
color_key_path = os.path.join(os.path.dirname(valid_files[0]), "color_key.svg")
fig, ax = plt.subplots(figsize=(1.2, 4))
norm = LogNorm(vmin=global_min, vmax=global_max)
cbar = fig.colorbar(
    plt.cm.ScalarMappable(norm=norm, cmap="viridis"),
    cax=ax, orientation="vertical"
)
cbar.set_label("Value (log scale)", fontsize=9)
plt.savefig(color_key_path, format="svg", bbox_inches="tight", pad_inches=0.1)
plt.close()
print(f"✅ Saved color key: {color_key_path}")
