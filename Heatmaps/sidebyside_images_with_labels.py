import os
import io
import re
import base64
import numpy as np
from tkinter import Tk, filedialog
from PIL import Image
import svgwrite
import xml.etree.ElementTree as ET

# --- Optional CairoSVG support for vector fallback ---
try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False
    print("⚠️ CairoSVG not found — vector-only SVGs may not render. Install via: pip install cairosvg")


# --- File selection dialog ---
def select_svg_files():
    root = Tk()
    root.withdraw()
    return list(filedialog.askopenfilenames(
        title="Select one or more SVG image files",
        filetypes=[("SVG files", "*.svg")]
    ))


# --- Extract embedded raster image from SVG ---
def extract_image_from_svg(svg_path):
    NS = {"svg": "http://www.w3.org/2000/svg", "xlink": "http://www.w3.org/1999/xlink"}
    try:
        tree = ET.parse(svg_path)
        root = tree.getroot()
        image_elems = list(root.findall(".//{http://www.w3.org/2000/svg}image")) or list(root.findall(".//image"))
        for el in image_elems:
            href = el.get(f"{{{NS['xlink']}}}href") or el.get("href")
            if href and href.lower().startswith("data:image/") and ";base64," in href.lower():
                _, b64_payload = href.split(",", 1)
                b64_payload = re.sub(r"\s+", "", b64_payload)
                img_bytes = base64.b64decode(b64_payload, validate=False)
                return Image.open(io.BytesIO(img_bytes)).convert("RGB")
    except Exception:
        pass

    # Regex fallback for malformed SVGs
    with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
        svg_text = f.read()
    m = re.search(
        r'''(?:xlink:href|href)\s*=\s*["']\s*data:image/(png|jpeg|jpg)\s*;\s*base64\s*,\s*([A-Za-z0-9+/=\s]+?)\s*["']''',
        svg_text, flags=re.IGNORECASE | re.DOTALL,
    )
    if m:
        _, b64_payload = m.groups()
        b64_payload = re.sub(r"\s+", "", b64_payload)
        img_bytes = base64.b64decode(b64_payload, validate=False)
        return Image.open(io.BytesIO(img_bytes)).convert("RGB")

    raise RuntimeError(f"No embedded raster image found in {os.path.basename(svg_path)}")


# --- Extract x-axis labels from SVG comments ---
def extract_xlabels_from_svg(svg_path):
    labels = []
    try:
        with open(svg_path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
        for m in re.finditer(r"<!--\s*([^>]+?)\s*-->", text):
            label = m.group(1).strip()
            if len(label) > 0 and not any(x in label.lower() for x in ["svg", "created", "metadata"]):
                labels.append(label)
    except Exception as e:
        print(f"⚠️ Failed to parse x-axis labels in {os.path.basename(svg_path)}: {e}")

    # Remove duplicates preserving order
    seen = set()
    uniq_labels = []
    for lbl in labels:
        if lbl not in seen:
            uniq_labels.append(lbl)
            seen.add(lbl)
    return uniq_labels


# --- Compute rank divided by number of rows in image ---
def compute_adjusted_rank(image):
    arr = np.array(image, dtype=np.float32)
    height = arr.shape[0]  # number of rows (image height)
    return arr.sum() / max(1, height)


# --- Build composite SVG (side by side layout) ---
def make_side_by_side_svg(entries):
    n = len(entries)
    img_w, img_h = entries[0]['image'].size
    spacing = img_w * 0.05
    font_pt = 18
    label_font_pt = 14
    text_h = font_pt * 2
    x_label_space = label_font_pt * 4
    total_width = n * img_w + (n - 1) * spacing
    total_height = img_h + text_h + x_label_space

    dwg = svgwrite.Drawing(size=(total_width, total_height))

    def embed_image(img, x, y):
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        href = f"data:image/png;base64,{base64.b64encode(buf.getvalue()).decode('ascii')}"
        dwg.add(dwg.image(href=href, insert=(x, y), size=(img.width, img.height)))

    for i, entry in enumerate(entries):
        img = entry['image']
        x = i * (img_w + spacing)
        y = text_h
        embed_image(img, x, y)

        # --- Filename above each image ---
        filename = os.path.splitext(os.path.basename(entry['path']))[0]
        dwg.add(dwg.text(
            filename,
            insert=(x + img_w / 2, font_pt),
            font_size=f"{font_pt}px",
            text_anchor="middle",
            fill="black"
        ))

        # --- X-axis labels below each image ---
        xlabels = extract_xlabels_from_svg(entry['path'])
        if xlabels:
            step = img_w / max(1, len(xlabels))
            x_start = x + step / 2
            label_y = y + img_h + label_font_pt * 1.5
            for j, label in enumerate(xlabels):
                dwg.add(dwg.text(
                    label,
                    insert=(x_start + j * step, label_y),
                    font_size=f"{label_font_pt}px",
                    fill="black",
                    text_anchor="middle"
                ))

    return dwg


# --- Main ---
def main():
    svg_files = select_svg_files()
    if not svg_files:
        print("No files selected.")
        return

    entries = []
    for path in svg_files:
        try:
            img = extract_image_from_svg(path)
            rank = compute_adjusted_rank(img)
            entries.append({"path": path, "image": img, "rank": rank})
            print(f"{os.path.basename(path)} → adjusted rank {rank:,.2f}")
        except Exception as e:
            print(f"⚠️ Skipped {path}: {e}")

    if not entries:
        print("No valid images extracted.")
        return

    # Sort: highest rank (left) to lowest (right)
    entries.sort(key=lambda e: e['rank'], reverse=True)

    dwg = make_side_by_side_svg(entries)
    out_path = os.path.join(os.path.dirname(svg_files[0]), "ranked_side_by_side_heatmap.svg")
    dwg.saveas(out_path)
    print(f"✅ Saved composite SVG:\n{out_path}")


if __name__ == "__main__":
    main()
