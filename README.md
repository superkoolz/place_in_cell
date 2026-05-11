# Excel Place In Cell
#### Place Images Inside Excel Cells with Python · by koolz

> Bulk-embed images **inside** cells in `.xlsx` files — not floating shapes, not drawing objects — using Excel 365's native **"Place in Cell"** feature. Pure Python, zero dependencies, no Excel installation required.

If you have ever tried to programmatically place an image *inside* an Excel cell (like the Insert → Picture → Place in Cell option in Microsoft 365) and found that openpyxl, xlsxwriter, and the stable Office.js API all lack support — this library solves that problem.

---

## Why this exists

Microsoft introduced **in-cell images** (`LocalImageCellValue` / `WebImageCellValue`) as part of the Excel 365 rich data types system. No mainstream Python Excel library supports it yet. This library reverse-engineers the underlying OOXML `richData` format and writes it directly, giving you full programmatic control.

Typical use cases:
- Product catalogs with a thumbnail per row
- Automated report generation with embedded charts or logos
- Data pipelines that output Excel files with visual content
- Apps that let users download Excel files with images already placed in cells

---

## Requirements

- **Python 3.9+**
- **Microsoft 365** (Excel for Windows, Mac, or Web) — this feature does not exist in Excel 2019 or earlier
- No pip dependencies for the core library
- [Pillow](https://pypi.org/project/Pillow/) only to run the built-in demo

---

## Installation

No package yet — just drop `place_images_in_cells.py` into your project:

```bash
curl -O https://raw.githubusercontent.com/superkoolz/place_in_cell/main/place_images_in_cells.py
```

---

## Quick start

```python
from place_images_in_cells import create_excel_with_cell_images

create_excel_with_cell_images(
    output_path="catalog.xlsx",
    images=[
        {"cell": "A1", "image": "photos/shoe_001.png"},
        {"cell": "A2", "image": "photos/shoe_002.jpg"},
        {"cell": "A3", "url": "https://example.com/shoe_003.webp"},  # Excel fetches at open time
    ]
)
```

Open `catalog.xlsx` in **Excel 365** — each image sits inside its cell, resizes with the cell, and behaves like a native in-cell image inserted by hand.

---

## API reference

```python
create_excel_with_cell_images(
    output_path,          # str   — path for the output .xlsx file
    images,               # list  — one dict per image cell (see below)
    sheet_name="Sheet1",  # str   — worksheet tab name
    row_height=97.2,      # float — row height in points  (97.2 pt ≈ 130 px)
    col_width=18.0,       # float — column width in char units (18 ≈ 130 px)
)
```

**Image entry keys:**

| Key | Required | Description |
|-----|----------|-------------|
| `cell` | yes | Cell address, e.g. `"B3"` or `"AA10"` |
| `image` | one of | Path to a local image file |
| `url` | one of | HTTPS URL — Excel downloads the image on open |
| `alt_text` | no | Accessibility label shown in screen readers |

**Supported local formats:** PNG · JPEG · GIF · BMP · TIFF · WEBP

---

## Building apps on top of this

Because `create_excel_with_cell_images` takes a plain list of dicts and writes a standard `.xlsx`, it fits naturally into any backend:

```python
# Flask / FastAPI — serve a generated xlsx as a download
from flask import Flask, send_file
import io, zipfile
from place_images_in_cells import create_excel_with_cell_images

app = Flask(__name__)

@app.route("/export")
def export():
    create_excel_with_cell_images(
        output_path="/tmp/export.xlsx",
        images=[
            {"cell": f"A{i}", "image": f"products/{sku}.png"}
            for i, sku in enumerate(get_skus(), start=1)
        ]
    )
    return send_file("/tmp/export.xlsx", as_attachment=True)
```

```python
# Batch pipeline — process a folder of images into one spreadsheet
import os
from place_images_in_cells import create_excel_with_cell_images

images = [
    {"cell": f"A{i}", "image": os.path.join("photos", fname)}
    for i, fname in enumerate(sorted(os.listdir("photos")), start=1)
]
create_excel_with_cell_images("batch_output.xlsx", images)
```

---

## How it works (OOXML internals)

Excel's "Place in Cell" images are stored via the **OOXML rich data** system — the same format used for linked data types and `IMAGE()` formula results. Inside the `.xlsx` ZIP:

```
xl/
├── metadata.xml                          # maps each cell's vm= index → rich value index
├── richData/
│   ├── rdrichvaluestructure.xml          # declares _localImage / _webImage schemas
│   ├── rdrichvalue.xml                   # one <rv> entry per image cell
│   ├── richValueRel.xml                  # maps rich value index → relationship rId
│   └── _rels/richValueRel.xml.rels       # rId → xl/media/imageN.ext
└── media/
    ├── image1.png
    └── image2.jpg
```

Each image cell is written as `t="e" vm="N"` (error value type + value-metadata index). Excel follows the metadata chain at load time and renders the image inside the cell boundary.

The `_localImage` rich value structure:

```xml
<s t="_localImage">
  <k n="_rvRel:LocalImageIdentifier" t="i"/>   <!-- index into richValueRel.xml -->
  <k n="CalcOrigin" t="i"/>                    <!-- always 5 for embedded images -->
</s>
```

---

## Compatibility

| Environment | Status |
|---|---|
| Microsoft 365 — Excel for Windows | ✅ Supported |
| Microsoft 365 — Excel for Mac | ✅ Supported |
| Microsoft 365 — Excel for Web | ✅ Supported |
| Excel 2021 and earlier (perpetual licence) | ❌ No in-cell image feature |
| LibreOffice Calc | ❌ Not supported |
| Google Sheets | ❌ Not supported |

> **This is a Microsoft 365-exclusive feature.** The `.xlsx` file opens without errors in older Excel versions, but image cells will show `#VALUE!` instead of the image.

---

## Demo

```bash
pip install Pillow
python place_images_in_cells.py
# → generates test_cell_images.xlsx with 6 colour-block images in A1:B3
```

---

## Credits

Made by superkoolz
