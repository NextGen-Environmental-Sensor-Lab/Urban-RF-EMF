import os
import io
import re
import base64
import numpy as np
from tkinter import Tk, filedialog
from PIL import Image
import svgwrite
import xml.etree.ElementTree as ET

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

def extract_xlabels_from_svg(svg_path):
    """
    Extract x-axis category labels from comment nodes in the SVG.
    Example comment: <!-- Broadcast -->
    Returns a list of label strings in order of appearance.
    """
    import xml.etree.ElementTree as ET
    labels = []
    try:
        with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()

        # Capture comment nodes like <!-- Broadcast -->
        for m in re.finditer(r"<!--\s*([^>]+?)\s*-->", text):
            label = m.group(1).strip()
            # Filter out irrelevant comments
            if len(label) > 0 and not any(x in label.lower() for x in ["svg", "created", "metadata"]):
                labels.append(label)
    except Exception as e:
        print(f"⚠️ Failed to parse x-axis labels: {e}")

    # Remove duplicates while preserving order
    seen = set()
    uniq_labels = []
    for lbl in labels:
        if lbl not in seen:
            uniq_labels.append(lbl)
            seen.add(lbl)

    return uniq_labels


def extract_image_from_svg(svg_path):
    """
    Extract ONLY the embedded raster image from <image> tags that use a data: URL.
    - Supports xlink:href or href
    - Tolerates newlines/whitespace inside the base64 payload
    - Ignores all vector/text primitives
    """
    # Namespaces commonly used in SVGs
    NS = {
        "svg": "http://www.w3.org/2000/svg",
        "xlink": "http://www.w3.org/1999/xlink",
    }

    # 1) Try proper XML parsing first
    try:
        tree = ET.parse(svg_path)
        root = tree.getroot()

        # Find all <image> elements (both with and without explicit namespace)
        # This covers: <image>, <svg:image>
        image_elems = list(root.findall(".//{http://www.w3.org/2000/svg}image"))
        if not image_elems:
            image_elems = list(root.findall(".//image"))

        for el in image_elems:
            # Try xlink:href first, then href
            href = el.get(f"{{{NS['xlink']}}}href") or el.get("href")
            if not href:
                continue

            # Expect data URL: data:image/(png|jpeg|jpg);base64,<data...>
            if href.lower().startswith("data:image/") and ";base64," in href.lower():
                # Split only once (handles newlines/whitespace after comma)
                header, b64_payload = href.split(",", 1)
                # Remove any whitespace/newlines in the base64 payload
                b64_payload = re.sub(r"\s+", "", b64_payload)
                try:
                    img_bytes = base64.b64decode(b64_payload, validate=False)
                    return Image.open(io.BytesIO(img_bytes)).convert("RGB")
                except Exception as e:
                    raise RuntimeError(
                        f"Error decoding embedded raster image in {os.path.basename(svg_path)}: {e}"
                    )

        # If we got here, no suitable <image> data URLs were found
        # Fall through to the regex approach (handles some malformed SVGs)
    except Exception:
        # If XML parsing fails (malformed SVG), we’ll try a regex below
        pass

    # 2) Regex fallback for very odd SVGs or ones with inline namespaces/attrs in unexpected orders.
    #    - Handles xlink:href OR href
    #    - Tolerates newlines after the comma and anywhere in the base64 block
    with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
        svg_text = f.read()

    # Look for either href="data:image/...;base64,..." or xlink:href="..."
    m = re.search(
        r'''(?:xlink:href|href)\s*=\s*["']\s*data:image/(png|jpeg|jpg)\s*;\s*base64\s*,\s*([A-Za-z0-9+/=\s]+?)\s*["']''',
        svg_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if m:
        _fmt, b64_payload = m.groups()
        b64_payload = re.sub(r"\s+", "", b64_payload)
        try:
            img_bytes = base64.b64decode(b64_payload, validate=False)
            return Image.open(io.BytesIO(img_bytes)).convert("RGB")
        except Exception as e:
            raise RuntimeError(
                f"Error decoding embedded raster image in {os.path.basename(svg_path)}: {e}"
            )

    # 3) No embedded raster found → skip (do NOT rasterize entire SVG via CairoSVG)
    raise RuntimeError(
        f"❌ No embedded raster image found in: {os.path.basename(svg_path)} (skipping)"
    )


def compute_pixel_sum(image):
    arr = np.array(image, dtype=np.float32)
    return arr.sum()


def extract_ticks_from_color_key(svg_path):
    """
    Extract tick labels and their y-positions from color_key.svg.
    Returns a list of (fractional_y, label_text) sorted from bottom to top.
    fractional_y is between 0 (bottom) and 1 (top).
    """
    try:
        tree = ET.parse(svg_path)
        root = tree.getroot()

        ticks = []
        # Iterate through all text elements
        for text_elem in root.findall(".//{http://www.w3.org/2000/svg}text"):
            txt = "".join(text_elem.itertext()).strip()
            # Identify numeric labels
            if re.match(r"^[-+]?\d*\.?\d+(e[-+]?\d+)?$", txt):
                y = float(text_elem.attrib.get("y", "0"))
                ticks.append((y, txt))

        if not ticks:
            # Try unnamespaced <text> elements
            for text_elem in root.findall(".//text"):
                txt = "".join(text_elem.itertext()).strip()
                if re.match(r"^[-+]?\d*\.?\d+(e[-+]?\d+)?$", txt):
                    y = float(text_elem.attrib.get("y", "0"))
                    ticks.append((y, txt))

        if not ticks:
            return []

        # Normalize positions: 0 at bottom, 1 at top
        ys = [y for y, _ in ticks]
        y_min, y_max = min(ys), max(ys)
        ticks_normalized = [((y - y_min) / (y_max - y_min), label) for y, label in ticks]
        # Sort from bottom (0) to top (1)
        ticks_normalized.sort(key=lambda x: x[0])
        return ticks_normalized

    except Exception as e:
        print(f"⚠️ Failed to parse ticks from color key: {e}")
        return []


def make_composite_svg(entries):
    """Create an SVG stacking the ranked images with spacing, grouped text labels,
       rotated top x-axis labels, and a stretched color key on the right using true tick labels.
    """
    if not entries:
        raise ValueError("No images to compose")

    n = len(entries)
    img_w, img_h = entries[0]['image'].size
    spacing = img_h * 0.01  # 1% vertical spacing
    total_stack_h = n * img_h + (n - 1) * spacing

    # --- Locate color key SVG ---
    base_dir = os.path.dirname(entries[0]['path'])
    color_key_path = os.path.join(base_dir, "color_key.svg")

    key_img = None
    tick_info = []
    if os.path.exists(color_key_path):
        try:
            key_img = extract_image_from_svg(color_key_path)
            tick_info = extract_ticks_from_color_key(color_key_path)
        except Exception as e:
            print(f"⚠️ Could not load color key SVG: {e}")

    # --- Layout geometry ---
    art_w = int(img_w * 1.7)
    art_h = int(total_stack_h * 1.25)
    y_offset = int((art_h - total_stack_h) / 2)
    right_margin = int(img_w * 0.05)
    right_x = art_w - img_w - (right_margin * 3)

    dwg = svgwrite.Drawing(size=(art_w, art_h))

    # --- Helper to embed Pillow images ---
    def embed_image(img, x, y, width=None, height=None):
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        href = f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode('ascii')}"
        if width is None:
            width, height = img.width, img.height
        dwg.add(dwg.image(href=href, insert=(x, y), size=(width, height)))

    # --- Font parameters ---
    font_pt = 18
    line_spacing = font_pt * 1.2
    text_color = "black"

    # --- Stack the images ---
    y = y_offset
    for entry in entries:
        img = entry['image']
        embed_image(img, right_x, y)

        # Metadata text block
        date_str, env_type, borough, location = parse_filename(entry['path'])
        borough_full = BOROUGH_FULL.get(borough, borough)
        env_full = ENV_FULL.get(env_type, env_type)

        text_x = right_x - right_margin * 2
        base_y = y + img_h / 2 - (1.5 * line_spacing)
        g = dwg.g(text_anchor="end", fill=text_color, font_size=f"{font_pt}px")
        g.add(dwg.text(date_str, insert=(text_x, base_y)))
        g.add(dwg.text(location, insert=(text_x, base_y + line_spacing)))
        g.add(dwg.text(borough_full, insert=(text_x, base_y + 2 * line_spacing)))
        g.add(dwg.text(env_full, insert=(text_x, base_y + 3 * line_spacing)))
        dwg.add(g)

        y += img_h + spacing

    top_y = y_offset
    bottom_y = y - spacing

    # --- X-axis labels (rotated) ---
    xlabels = extract_xlabels_from_svg(entries[0]['path'])
    if xlabels:
        step = img_w / max(1, len(xlabels))
        x_start = right_x + step / 2
        label_y = top_y - (font_pt * 1.5)
        for i, label in enumerate(xlabels):
            dwg.add(
                dwg.text(
                    label,
                    insert=(x_start + i * step, label_y),
                    font_size=f"{font_pt}px",
                    fill=text_color,
                    text_anchor="middle",
                    transform=f"rotate(-90,{x_start + i * step},{label_y})",
                )
            )

    # --- Color key (stretched to full height) ---
    if key_img:
        key_width = int(img_w * 0.12)
        key_height = int(bottom_y - top_y)
        key_x = right_x + img_w + right_margin
        key_y = top_y
        embed_image(key_img.resize((key_width, key_height)), key_x, key_y)

        # --- True tick marks and labels from color_key.svg ---
        for frac, label in tick_info:
            ty = key_y + key_height * (1 - frac)
            tick_len = int(key_width * 0.3)
            dwg.add(
                dwg.line(
                    start=(key_x + key_width, ty),
                    end=(key_x + key_width + tick_len, ty),
                    stroke=text_color,
                    stroke_width=1,
                )
            )
            dwg.add(
                dwg.text(
                    label,
                    insert=(key_x + key_width + tick_len + 5, ty + font_pt / 3),
                    font_size=f"{font_pt}px",
                    fill=text_color,
                )
            )

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
