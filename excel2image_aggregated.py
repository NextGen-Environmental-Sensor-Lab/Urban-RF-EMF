import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import LogNorm
from tkinter import Tk, filedialog
import numpy as np
import os

# --- CATEGORY DEFINITIONS ---
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

LABEL_FONT_SIZE_PT = 18
LABEL_Y_OFFSET_AX = 1.005  # just above the top edge
LABEL_X_OFFSET_AX = 0.07 
LABEL_LINESPACING_AX = 3.750

def main():
    """Generate per-file heatmaps with single rotated, line-spaced x-axis label block aligned to columns."""
    # --- Pick files & destination ---
    Tk().withdraw()
    file_paths = filedialog.askopenfilenames(
        title="Select one or more Excel files",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_paths:
        print("No files selected. Exiting.")
        return

    dest_dir = filedialog.askdirectory(title="Select destination folder for SVG output")
    if not dest_dir:
        print("No destination folder selected. Exiting.")
        return

    # --- Global color scale ---
    global_min, global_max = np.inf, -np.inf
    valid_files = []

    for file_path in file_paths:
        df = pd.read_excel(file_path)
        df = df.apply(pd.to_numeric, errors="coerce")

        # Category RSS
        agg = pd.DataFrame()
        for cat, cols in CATEGORY_MAP.items():
            cols_present = [c for c in cols if c in df.columns]
            if cols_present:
                agg[cat] = np.sqrt((df[cols_present] ** 2).sum(axis=1))

        agg = agg.replace([np.inf, -np.inf], np.nan).dropna(how="all")
        agg = agg.mask(agg <= 0)

        if agg.isna().all().all():
            print(f"⚠️ Skipping {os.path.basename(file_path)} (no valid category data)")
            continue

        vmin, vmax = np.nanmin(agg.values), np.nanmax(agg.values)
        global_min, global_max = min(global_min, vmin), max(global_max, vmax)
        valid_files.append((file_path, agg))

    if not valid_files:
        print("No valid numeric data found.")
        return

    print(f"Global color scale (log): min={global_min:.3g}, max={global_max:.3g}")

    # --- Per-file plots ---
    for file_path, agg in valid_files:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        output_path = os.path.join(dest_dir, base_name + ".svg")

        # Rows: first at top
        data = agg[CATS].iloc[::-1].values
        nrows, ncols = data.shape

        # Dynamic height; width fixed
        SCALE_FACTOR = 0.05  # inches per row
        # fig_height = max(1.5, nrows * SCALE_FACTOR)
        fig_height = max(1.0, nrows * SCALE_FACTOR)
        fig_width = 6

        fig, ax = plt.subplots(figsize=(fig_width, fig_height))
        ax.imshow(
            data,
            cmap="viridis",
            aspect="auto",
            interpolation="nearest",
            norm=LogNorm(vmin=global_min, vmax=global_max)
        )

        # Clean axes
        ax.set_xticks([])
        ax.set_yticks([])
        for s in ax.spines.values():
            s.set_visible(False)

        # --- Compute line spacing so each label aligns with a column center ---
        # Need renderer to get axes width in pixels
        fig.canvas.draw()  # ensure a renderer exists
        renderer = fig.canvas.get_renderer()
        bbox = ax.get_window_extent(renderer=renderer)
        axes_width_px = bbox.width

        # Column center spacing in pixels:
        col_spacing_px = axes_width_px / ncols

        # Convert font size points to pixels
        pixels_per_point = fig.dpi / 72.0
        # linespacing is a multiplier on font size; choose so that:
        # (font_size_points * pixels_per_point * linespacing) ~= col_spacing_px
        computed_linespacing = col_spacing_px / (LABEL_FONT_SIZE_PT * pixels_per_point)
        # Keep it sane
        linespacing = max(0.6, min(3.0, computed_linespacing))

        # One text object: stacked labels with newlines, left-aligned, rotated -90°
        labels_str = "\n".join(CATS)
        ax.text(
            LABEL_X_OFFSET_AX, # left/top in axes coords
            LABEL_Y_OFFSET_AX,  # left/top in axes coords
            labels_str,
            fontsize=LABEL_FONT_SIZE_PT,
            rotation=90,
            # linespacing=linespacing,
            linespacing=LABEL_LINESPACING_AX,
            transform=ax.transAxes,
            ha="left",
            va="bottom"
        )

        plt.savefig(output_path, format="svg", bbox_inches="tight", pad_inches=0)
        plt.close()
        print(f"✅ Saved aggregated heatmap: {output_path}")

    # --- Color key (no caption text) ---
    color_key_path = os.path.join(dest_dir, "color_key.svg")
    fig, ax = plt.subplots(figsize=(1.2, 4))
    norm = LogNorm(vmin=global_min, vmax=global_max)
    fig.colorbar(
        plt.cm.ScalarMappable(norm=norm, cmap="viridis"),
        cax=ax, orientation="vertical"
    )
    # No label text next to numbers
    plt.savefig(color_key_path, format="svg", bbox_inches="tight", pad_inches=0.1)
    plt.close()
    print(f"✅ Saved color key: {color_key_path}")
    print("\n🎉 All heatmaps and color key saved to:")
    print(dest_dir)


if __name__ == "__main__":
    main()
