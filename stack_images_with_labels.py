import os
import io
import re
import base64
import numpy as np
from tkinter import Tk, filedialog
from PIL import Image
import svgwrite

# Optional: for rendering pure vector SVGs to PNGs before embedding
try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False
    print("⚠️ CairoSVG not found — vector-only SVGs may not render. Install via: pip install cairosvg")

# --- Lookup tables for boroughs and environments ---
BOROUGH_FULL = {
    "M": "Manhattan",
    "BK": "Brooklyn",
    "Q": "Queens",
    "BX": "Bronx",
    "SI": "Staten Island",
    "FERRY": "Ferry"
}

ENV_FULL = {
    "C": "Commercial",
    "R": "Residential",
    "G": "Greenery",
    "T": "Transportation",
    "I": "Indoors"
}


def parse_filename(path):
    """Parse filename 'YYYY-MM-DD_hh.mm.ss E B location'."""
    name = os.path.splitext(os.path.basename(path))[0]
    m = re.match(r"(\d{4}-\d{2}-\d{2})(?:[_T]\d{2}\.\d{2}\.\d{2})?\s+([CRGTI])\s+(M|BK|Q|BX|SI|FERRY)\s+(.+)", name)
    if not m:
        return "", "", "", name
    date_str, env_type, borough, location = m.groups()
    return date_str, env_type, borough, location.strip()


def select_svg_files():
    root = Tk()
    root.withdraw()
    return list(filedialog.askopenfilenames(
        title="Select one or more SVG image files",
        filetypes=[("SVG files", "*.svg")])
    )


def extract_image_from_svg(svg_path):
    with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
        svg_text = f.read()

    match = re.search(r'data:image/(png|jpeg);base64,([A-Za-z0-9+/=]+)', svg_text)
    if match:
        fmt, b64_data = match.groups()
        img_bytes = base64.b64decode(b64_data)
        return Image.open(io.BytesIO(img_bytes)).convert("RGB")

    if CAIROSVG_AVAILABLE:
        png_bytes = cairosvg.svg2png(url=svg_path)
        return Image.open(io.BytesIO(png_bytes)).convert("RGB")
    else:
        raise RuntimeError(f"No embedded image found and CairoSVG not installed for: {svg_path}")


def compute_pixel_sum(image):
    arr = np.array(image, dtype=np.float32)
    return arr.sum()


def make_composite_svg(entries):
    """Create an SVG stacking the ranked images on the right, with color key and parsed metadata labels."""
    if not entries:
        raise ValueError("No images to compose")

    n = len(entries)
    img_w, img_h = entries[0]['image'].size

    # Try to find color key
    key_path = None
    base_dir = os.path.dirname(entries[0]['path'])
    base_name = os.path.splitext(os.path.basename(entries[0]['path']))[0]
    for ext in [".png", ".jpg", ".jpeg", ".tif", ".tiff"]:
        candidate = os.path.join(base_dir, base_name + "_key" + ext)
        if os.path.exists(candidate):
            key_path = candidate
            break

    key_img = None
    key_w = 0
    key_margin = int(img_w * 0.05)
    if key_path:
        try:
            key_img = Image.open(key_path).convert("RGB")
            key_w = key_img.width
        except Exception as e:
            print(f"⚠️ Could not load color key: {e}")

    # Artboard
    art_w = int(img_w * 1.5)
    art_h = int((n + 1) * img_h)
    total_stack_h = n * img_h
    y_offset = int((art_h - total_stack_h) / 2)

    # Right-side placement
    right_x = art_w - img_w - key_margin
    if key_img:
        right_x -= key_w + key_margin

    # Prepare SVG
    dwg = svgwrite.Drawing(size=(art_w, art_h))

    # Helper to embed Pillow image as base64 PNG in SVG
    def embed_image(img, x, y):
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        href = f"data:image/png;base64,{b64}"
        dwg.add(dwg.image(href=href, insert=(x, y), size=(img.width, img.height)))

    # Paste color key on left, vertically centered
    if key_img:
        key_y = int((art_h - key_img.height) / 2)
        embed_image(key_img, key_margin, key_y)

    # Font sizes
    font_large = int(img_h * 0.12)
    font_small = int(img_h * 0.10)
    line_spacing = int(font_large * 1.15)

    # Stack images + labels
    y = y_offset
    for entry in entries:
        img = entry['image']
        embed_image(img, right_x, y)

        # Parse metadata from filename
        date_str, env_type, borough, location = parse_filename(entry['path'])
        borough_full = BOROUGH_FULL.get(borough, borough)
        env_full = ENV_FULL.get(env_type, env_type)

        # Compute left margin for text
        text_x = right_x - key_margin * 2
        base_y = y + img_h / 2 - (1.5 * line_spacing)  # vertically center text block

        # 1. Date
        dwg.add(dwg.text(date_str,
                         insert=(text_x, base_y),
                         font_size=font_large,
                         fill="black",
                         text_anchor="end"))
        # 2. Location
        dwg.add(dwg.text(location,
                         insert=(text_x, base_y + line_spacing),
                         font_size=font_large,
                         fill="black",
                         text_anchor="end"))
        # 3. Borough (full)
        dwg.add(dwg.text(borough_full,
                         insert=(text_x, base_y + 2 * line_spacing),
                         font_size=font_small,
                         fill="black",
                         text_anchor="end"))
        # 4. Environment type (full)
        dwg.add(dwg.text(env_full,
                         insert=(text_x, base_y + 3 * line_spacing),
                         font_size=font_small,
                         fill="black",
                         text_anchor="end"))

        y += img_h

    return dwg


def main():
    svg_files = select_svg_files()
    if not svg_files:
        print("No files selected.")
        return

    entries = []
    for path in svg_files:
        try:
            img = extract_image_from_svg(path)
            rank = compute_pixel_sum(img)
            entries.append({"path": path, "image": img, "rank": rank})
            print(f"{os.path.basename(path)} → rank {rank:,.0f}")
        except Exception as e:
            print(f"⚠️ Skipped {path}: {e}")

    if not entries:
        print("No valid images extracted.")
        return

    entries.sort(key=lambda e: e['rank'], reverse=True)
    dwg = make_composite_svg(entries)

    out_path = os.path.join(os.path.dirname(svg_files[0]), "stacked_ranked_images.svg")
    dwg.saveas(out_path)
    print(f"✅ Saved composite SVG:\n{out_path}")


if __name__ == "__main__":
    main()
