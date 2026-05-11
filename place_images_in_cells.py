#!/usr/bin/env python3
"""
place_images_in_cells.py

Bulk-place images inside Excel cells using the OOXML "Place in Cell" feature
(LocalImageCellValue). No Excel installation required. Works on any platform.

Supports PNG, JPEG, GIF, BMP, TIFF, WEBP for local files.
Supports HTTPS URLs via WebImageCellValue (Excel fetches the image at open time).

Usage:
    from place_images_in_cells import create_excel_with_cell_images

    create_excel_with_cell_images(
        output_path="output.xlsx",
        images=[
            {"cell": "A1", "image": "photos/cat.png"},
            {"cell": "A2", "image": "photos/dog.jpg"},
            {"cell": "B1", "url": "https://example.com/bird.png"},
        ]
    )
"""

import io
import os
import re
import zipfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_EXT_TO_MIME = {
    "png":  "image/png",
    "jpg":  "image/jpeg",
    "jpeg": "image/jpeg",
    "gif":  "image/gif",
    "bmp":  "image/bmp",
    "tiff": "image/tiff",
    "tif":  "image/tiff",
    "webp": "image/webp",
}

def _col_num(col: str) -> int:
    """'A'->1, 'Z'->26, 'AA'->27 …"""
    n = 0
    for c in col.upper():
        n = n * 26 + ord(c) - 64
    return n

def _parse_cell(ref: str):
    m = re.fullmatch(r"([A-Za-z]+)(\d+)", ref.strip())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return m.group(1).upper(), int(m.group(2))

def _ext(path: str) -> str:
    return Path(path).suffix.lstrip(".").lower()

def _mime(ext: str) -> str:
    return _EXT_TO_MIME.get(ext, "image/png")


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def create_excel_with_cell_images(
    output_path: str,
    images: list,
    sheet_name: str = "Sheet1",
    row_height: float = 97.2,   # points — ~130 px, fits a medium thumbnail
    col_width: float = 18.0,    # char units — ~130 px
):
    """
    Create an xlsx file with images placed *inside* cells (not floating).

    Parameters
    ----------
    output_path : str
        Destination .xlsx file path.
    images : list[dict]
        Each dict must have 'cell' and either 'image' (local file path) or
        'url' (HTTPS URL for a WebImageCellValue). Optional 'alt_text'.
        Examples:
            {"cell": "A1", "image": "cat.png"}
            {"cell": "A2", "url": "https://example.com/dog.png", "alt_text": "Dog"}
    sheet_name : str
        Name shown on the sheet tab.
    row_height : float
        Default height (points) applied to every row that has an image.
        97.2 pt ≈ 130 px at 96 dpi.
    col_width : float
        Default width (character units) applied to every column with an image.
        18 ≈ 130 px with Calibri 11pt.
    """
    local_entries = []   # dicts for _localImage cells
    web_entries   = []   # dicts for _webImage cells

    for item in images:
        col, row = _parse_cell(item["cell"])
        alt = item.get("alt_text", "")
        if "image" in item:
            path = item["image"]
            if not os.path.exists(path):
                raise FileNotFoundError(f"Image not found: {path}")
            ext = _ext(path)
            local_entries.append({"col": col, "row": row, "path": path, "ext": ext, "alt": alt})
        elif "url" in item:
            web_entries.append({"col": col, "row": row, "url": item["url"], "alt": alt})
        else:
            raise ValueError(f"Each image entry needs 'image' or 'url': {item}")

    # Merge into a single ordered list keeping original order for vm assignment
    all_entries = []
    for item in images:
        col, row = _parse_cell(item["cell"])
        if "image" in item:
            # Find matching local entry
            for e in local_entries:
                if e["col"] == col and e["row"] == row:
                    all_entries.append(("local", e))
                    break
        else:
            for e in web_entries:
                if e["col"] == col and e["row"] == row:
                    all_entries.append(("web", e))
                    break

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",              _content_types(all_entries))
        zf.writestr("_rels/.rels",                      _ROOT_RELS)
        zf.writestr("xl/workbook.xml",                  _workbook(sheet_name))
        zf.writestr("xl/_rels/workbook.xml.rels",       _workbook_rels(bool(local_entries), bool(web_entries)))
        zf.writestr("xl/styles.xml",                    _STYLES)
        zf.writestr("xl/worksheets/sheet1.xml",         _sheet(all_entries, row_height, col_width))
        zf.writestr("xl/metadata.xml",                  _metadata(len(all_entries)))
        zf.writestr("xl/richData/rdrichvaluestructure.xml", _rich_struct(bool(local_entries), bool(web_entries)))
        zf.writestr("xl/richData/rdRichValueTypes.xml", _RICH_TYPES)
        zf.writestr("xl/richData/rdrichvalue.xml",      _rich_values(all_entries))

        if local_entries:
            # richValueRel only needed for local images (file relationships)
            local_indices = [i for i, (kind, _) in enumerate(all_entries) if kind == "local"]
            zf.writestr("xl/richData/richValueRel.xml",
                        _rich_value_rel_xml(len(local_entries)))
            zf.writestr("xl/richData/_rels/richValueRel.xml.rels",
                        _rich_value_rel_rels(all_entries))

            # Embed image files
            for rel_idx, glob_idx in enumerate(local_indices):
                _, entry = all_entries[glob_idx]
                media_name = f"xl/media/image{rel_idx + 1}.{entry['ext']}"
                with open(entry["path"], "rb") as f:
                    zf.writestr(media_name, f.read())

    buf.seek(0)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Path(output_path).write_bytes(buf.read())
    print(f"Saved {output_path}  ({len(all_entries)} in-cell image(s))")


# ---------------------------------------------------------------------------
# XML builders
# ---------------------------------------------------------------------------

def _content_types(all_entries):
    exts = {e["ext"] for kind, e in all_entries if kind == "local"}
    ext_defaults = "\n".join(
        f'  <Default Extension="{ext}" ContentType="{_mime(ext)}"/>'
        for ext in sorted(exts)
    )
    richvalrel_parts = ""
    if any(k == "local" for k, _ in all_entries):
        richvalrel_parts = '  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>\n'

    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
{ext_defaults}
  <Override PartName="/xl/workbook.xml"                   ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"          ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml"                     ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/metadata.xml"                   ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
{richvalrel_parts}  <Override PartName="/xl/richData/rdrichvalue.xml"          ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
  <Override PartName="/xl/richData/rdRichValueTypes.xml"  ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>
</Types>"""


_ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""


def _workbook(sheet_name: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192"/></bookViews>
  <sheets>
    <sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""


def _workbook_rels(has_local: bool, has_web: bool) -> str:
    lines = [
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"    Target="worksheets/sheet1.xml"/>',
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/>',
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"        Target="styles.xml"/>',
        '  <Relationship Id="rId4" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue"             Target="richData/rdrichvalue.xml"/>',
        '  <Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure"    Target="richData/rdrichvaluestructure.xml"/>',
        '  <Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes"        Target="richData/rdRichValueTypes.xml"/>',
    ]
    if has_local:
        lines.append('  <Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/>')
    body = "\n".join(lines)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{body}
</Relationships>"""


_STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"""


def _sheet(all_entries, row_height: float, col_width: float) -> str:
    if not all_entries:
        return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>"""

    entries_data = [(kind, e) for kind, e in all_entries]

    all_cols = [_col_num(e["col"]) for _, e in entries_data]
    all_rows = [e["row"] for _, e in entries_data]
    min_col_n, max_col_n = min(all_cols), max(all_cols)
    min_row,   max_row   = min(all_rows), max(all_rows)

    def num_to_col(n: int) -> str:
        s = ""
        while n > 0:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    min_col_l = num_to_col(min_col_n)
    max_col_l = num_to_col(max_col_n)

    # Column width overrides
    unique_col_nums = sorted(set(all_cols))
    col_defs = "".join(
        f'<col min="{cn}" max="{cn}" width="{col_width}" customWidth="1"/>'
        for cn in unique_col_nums
    )

    # Group cells by row
    rows_map: dict = {}
    for vm_idx, (kind, e) in enumerate(entries_data, start=1):
        row_num = e["row"]
        rows_map.setdefault(row_num, []).append((e["col"], vm_idx))

    rows_xml = ""
    for row_num in sorted(rows_map):
        cells = sorted(rows_map[row_num], key=lambda x: _col_num(x[0]))
        col_nums_in_row = [_col_num(c) for c, _ in cells]
        spans = f"{min(col_nums_in_row)}:{max(col_nums_in_row)}"
        cells_xml = "".join(
            f'<c r="{col}{row_num}" t="e" vm="{vm}"><v>#VALUE!</v></c>'
            for col, vm in cells
        )
        rows_xml += f'<row r="{row_num}" spans="{spans}" ht="{row_height}" customHeight="1">{cells_xml}</row>\n'

    dim = f"{min_col_l}{min_row}:{max_col_l}{max_row}"
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
           xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
           mc:Ignorable="x14ac"
           xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
  <dimension ref="{dim}"/>
  <sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="14.4" x14ac:dyDescent="0.3"/>
  <cols>{col_defs}</cols>
  <sheetData>
{rows_xml}  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>"""


def _metadata(n: int) -> str:
    future = "\n".join(
        f'    <bk><extLst><ext uri="{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}">'
        f'<xlrd:rvb i="{i}" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"/>'
        f'</ext></extLst></bk>'
        for i in range(n)
    )
    values = "\n".join(f'    <bk><rc t="1" v="{i}"/></bk>' for i in range(n))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata">
  <metadataTypes count="1">
    <metadataType name="XLRICHVALUE" minSupportedVersion="120000"
                  copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1"
                  rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>
  </metadataTypes>
  <futureMetadata name="XLRICHVALUE" count="{n}">
{future}
  </futureMetadata>
  <valueMetadata count="{n}">
{values}
  </valueMetadata>
</metadata>"""


def _rich_struct(has_local: bool, has_web: bool) -> str:
    structs = []
    if has_local:
        structs.append(
            '  <s t="_localImage">\n'
            '    <k n="_rvRel:LocalImageIdentifier" t="i"/>\n'
            '    <k n="CalcOrigin" t="i"/>\n'
            '  </s>'
        )
    if has_web:
        # WebImage: URL stored as a string value, no file relationship needed
        structs.append(
            '  <s t="_webImage">\n'
            '    <k n="address" t="s"/>\n'
            '    <k n="CalcOrigin" t="i"/>\n'
            '  </s>'
        )
    count = len(structs)
    body = "\n".join(structs)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="{count}">
{body}
</rvStructures>"""


_RICH_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2"
             xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
             mc:Ignorable="x"
             xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <global>
    <keyFlags>
      <key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key>
      <key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key>
    </keyFlags>
  </global>
</rvTypesInfo>"""


def _rich_values(all_entries: list) -> str:
    """
    Build rdrichvalue.xml.

    Structure indices:
      0 = _localImage  (if any local entries exist)
      1 = _webImage    (if any web entries exist, 0 if no local)

    For _localImage: <v>{rel_index}</v><v>5</v>
    For _webImage:   <v>{url}</v><v>1</v>
    """
    has_local = any(k == "local" for k, _ in all_entries)
    has_web   = any(k == "web"   for k, _ in all_entries)
    local_struct_idx = 0
    web_struct_idx   = 1 if has_local else 0

    local_rel_counter = 0
    rvs = []
    for kind, entry in all_entries:
        if kind == "local":
            rvs.append(f'  <rv s="{local_struct_idx}"><v>{local_rel_counter}</v><v>5</v></rv>')
            local_rel_counter += 1
        else:
            url = entry["url"].replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            rvs.append(f'  <rv s="{web_struct_idx}"><v>{url}</v><v>1</v></rv>')

    body = "\n".join(rvs)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="{len(all_entries)}">
{body}
</rvData>"""


def _rich_value_rel_xml(n_local: int) -> str:
    rels = "\n".join(f'  <rel r:id="rId{i+1}"/>' for i in range(n_local))
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"
               xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
{rels}
</richValueRels>"""


def _rich_value_rel_rels(all_entries: list) -> str:
    local_entries = [(i, e) for i, (kind, e) in enumerate(all_entries) if kind == "local"]
    lines = []
    for rel_idx, (_, entry) in enumerate(local_entries):
        lines.append(
            f'  <Relationship Id="rId{rel_idx+1}"'
            f' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"'
            f' Target="../media/image{rel_idx+1}.{entry["ext"]}"/>'
        )
    body = "\n".join(lines)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
{body}
</Relationships>"""


# ---------------------------------------------------------------------------
# CLI / demo
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    # Quick self-test: create coloured PNG tiles on the fly (requires Pillow)
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        print("Pillow is not installed. Run:  pip install Pillow")
        sys.exit(1)

    colours = [
        (220, 60,  60,  "Red"),
        (60,  180, 60,  "Green"),
        (60,  100, 220, "Blue"),
        (200, 160, 40,  "Yellow"),
        (140, 60,  180, "Purple"),
        (40,  180, 180, "Teal"),
    ]

    cells   = ["A1", "A2", "A3", "B1", "B2", "B3"]
    entries = []

    for cell, (r, g, b, label) in zip(cells, colours):
        img_path = f"_test_{label.lower()}.png"
        img = Image.new("RGB", (300, 300), color=(r, g, b))
        d   = ImageDraw.Draw(img)
        d.rectangle([10, 10, 289, 289], outline=(255, 255, 255), width=4)
        d.text((90, 130), label, fill=(255, 255, 255))
        img.save(img_path)
        entries.append({"cell": cell, "image": img_path, "alt_text": label})

    create_excel_with_cell_images(
        output_path="test_cell_images.xlsx",
        images=entries,
        sheet_name="Cell Images",
        row_height=120,
        col_width=22,
    )

    # Cleanup temp PNGs
    for e in entries:
        os.remove(e["image"])

    print("Open test_cell_images.xlsx in Excel to verify.")
