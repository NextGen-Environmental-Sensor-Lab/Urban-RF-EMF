#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel MULTILINESTRING -> KMZ (Rank-inverted heat colormap)

Steps:
1. Interactively select a source Excel file.
2. For each row:
   - Read 'the_geom' column (MULTILINESTRING WKT).
   - Read corresponding 'Rank' value.
3. Generate a KMZ file:
   - Each polyline is drawn from the MULTILINESTRING.
   - Line color from a heat colormap where:
       - LOWER Rank = hotter color
       - HIGHER Rank = cooler color
   - Line width = 2.0
"""

import os
from tkinter import Tk, filedialog

import pandas as pd
import simplekml
from shapely import wkt
import matplotlib.cm as cm
import matplotlib.colors as mcolors


# ----------------- CONFIGURATION -----------------

GEOM_COL = "the_geom"   # WKT MULTILINESTRING column
RANK_COL = "Rank"       # numeric rank column
LINE_WIDTH = 2.0        # KML line width
COLORMAP_NAME = "turbo" # blue -> green -> yellow -> red


# ----------------- HELPER FUNCTIONS --------------

def pick_excel_file():
    """Open a file dialog to pick an Excel file and return its path (or None)."""
    root = Tk()
    root.withdraw()
    root.update()
    file_path = filedialog.askopenfilename(
        title="Select Excel file with MULTILINESTRING data",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.destroy()
    return file_path or None


def normalize_values(values):
    """Normalize a 1D array-like of numeric values to [0, 1]."""
    vmin = float(values.min())
    vmax = float(values.max())
    if vmax == vmin:  # all equal
        return [0.5] * len(values)
    return [(float(v) - vmin) / (vmax - vmin) for v in values]


def rgba_to_kml_color(rgba):
    """
    Convert an RGBA (0-1 range) tuple to KML color string 'aabbggrr'.

    KML uses ABGR (alpha, blue, green, red) in hex, little-endian-like order.
    """
    r, g, b, a = rgba
    r8 = int(round(r * 255))
    g8 = int(round(g * 255))
    b8 = int(round(b * 255))
    a8 = int(round(a * 255))
    return f"{a8:02x}{b8:02x}{g8:02x}{r8:02x}"


def get_color_for_rank(rank_norm, cmap):
    """
    Map a normalized rank value in [0, 1] to a KML color string
    using a matplotlib colormap.
    """
    rgba = cmap(rank_norm)  # (r, g, b, a)
    return rgba_to_kml_color(rgba)


# ----------------- MAIN LOGIC --------------------

def main():
    # 1. Pick Excel file
    excel_path = pick_excel_file()
    if not excel_path:
        print("No file selected. Exiting.")
        return

    print(f"Reading Excel file: {excel_path}")
    df = pd.read_excel(excel_path)

    # Basic column checks
    if GEOM_COL not in df.columns:
        raise ValueError(f"Column '{GEOM_COL}' not found in Excel file.")
    if RANK_COL not in df.columns:
        raise ValueError(f"Column '{RANK_COL}' not found in Excel file.")

    # Drop rows with missing geometry or rank
    df = df.dropna(subset=[GEOM_COL, RANK_COL]).copy()

    if df.empty:
        raise ValueError("No valid rows with both geometry and rank found.")

    # Extract rank values
    ranks = df[RANK_COL].astype(float)

    # --- KEY CHANGE: normalize then invert so higher Rank = cooler color ---
    norm_ranks = normalize_values(ranks)          # 0 = lowest rank, 1 = highest rank
    inv_norm_ranks = [1.0 - v for v in norm_ranks]  # 0 = highest rank (cool), 1 = lowest (hot)
    # ----------------------------------------------------------------------

    # Prepare colormap
    cmap = cm.get_cmap(COLORMAP_NAME)

    # 3. Create KML/KMZ
    kml = simplekml.Kml()
    folder = kml.newfolder(name="Polylines by Inverted Rank")

    # Optional: cache styles to avoid duplication for same color
    style_cache = {}

    for (idx, row), rnorm in zip(df.iterrows(), inv_norm_ranks):
        wkt_geom = row[GEOM_COL]
        rank_val = float(row[RANK_COL])

        # Parse WKT MULTILINESTRING (or LINESTRING)
        try:
            geom = wkt.loads(wkt_geom)
        except Exception as e:
            print(f"Row {idx}: failed to parse geometry WKT. Skipping. Error: {e}")
            continue

        # Determine color based on *inverted* normalized rank
        kml_color = get_color_for_rank(rnorm, cmap)

        # Reuse style if same color already created
        if kml_color in style_cache:
            line_style = style_cache[kml_color]
        else:
            # Create style
            style = simplekml.Style()
            style.linestyle.color = kml_color
            style.linestyle.width = LINE_WIDTH
            style_cache[kml_color] = style
            line_style = style

        # Handle both MultiLineString and LineString
        if geom.geom_type == "MultiLineString":
            lines = geom.geoms
        elif geom.geom_type == "LineString":
            lines = [geom]
        else:
            print(f"Row {idx}: geometry type '{geom.geom_type}' not supported. Skipping.")
            continue

        # Create KML line(s)
        for line in lines:
            coords = [(float(x), float(y)) for x, y in line.coords]
            ls = folder.newlinestring()
            ls.coords = coords
            ls.style = line_style
            ls.name = f"Rank {rank_val:.3f}"

    # Save as KMZ next to source file
    base, _ = os.path.splitext(excel_path)
    kmz_path = base + "_polylines_by_inverted_rank.kmz"
    kml.savekmz(kmz_path)
    print(f"KMZ saved to: {kmz_path}")


if __name__ == "__main__":
    main()
