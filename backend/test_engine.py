"""
test_engine.py — proves the generalized engine works for all three
current use cases before we build any API/UI on top of it.
"""

import io
import json
import os
import re
import zipfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

from schemas import DocumentConfig, ColumnConfig, GroupedReportConfig
import engine

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
OUT = os.path.join(HERE, "output")
os.makedirs(OUT, exist_ok=True)


def make_dummy_image_bytes(color):
    img = PILImage.new("RGB", (80, 80), color)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def build_client_catalog_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["-- title row --"])
    ws.append(["Select", "Product Name", "Dimension (Spec)", "Supply Advantage", "Price Range (THB/kg)"])
    rows = [
        ["Yes", "Frozen Red Bean", "whole, 1kgx10pk", "Year-round supply, cold chain verified", "100-110"],
        ["Yes", "Frozen Lotus Roots", "1kg x 10 pcs", "China sourced, 1 week lead time", "80-100"],
        ["No", "Excluded Item", "n/a", "n/a", "n/a"],
        ["Yes", "Frozen Mango", "500g x 12 pk", "", "125-140"],
    ]
    for r in rows:
        ws.append(r)

    for row_idx, color in [(3, (200, 80, 80)), (4, (80, 160, 90))]:
        img = XLImage(make_dummy_image_bytes(color))
        img.width, img.height = 80, 80
        ws.add_image(img, f"A{row_idx}")

    wb.save(path)


def build_quotation_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["-- title row --"])
    ws.append(["Select", "Product Name", "Quantity", "Price Range (THB/kg)", "Remarks"])
    rows = [
        ["Yes", "Frozen Strawberry", "500", "12.50", "Bulk discount applied"],
        ["Yes", "Frozen Blueberry", "50", "45.00", ""],
        ["No", "Skip Me", "1", "1", ""],
    ]
    for r in rows:
        ws.append(r)
    wb.save(path)


def run_case(name, xlsx_builder, preset_file):
    print(f"\n=== {name} ===")
    xlsx_path = os.path.join(OUT, f"{name}_input.xlsx")
    xlsx_builder(xlsx_path)

    with open(os.path.join(HERE, "presets", preset_file)) as f:
        config = DocumentConfig(**json.load(f))

    headers = engine.detect_headers(xlsx_path, config.header_row)
    print(f"Detected headers: {headers}")

    items = engine.read_excel(xlsx_path, config)
    print(f"Items included: {len(items)}")
    for it in items:
        print(f"  {it}")

    out_pdf = os.path.join(OUT, f"{name}.pdf")
    engine.generate_pdf(items, config, out_pdf, ASSETS)
    size_kb = os.path.getsize(out_pdf) / 1024
    print(f"PDF generated: {out_pdf} ({size_kb:.1f} KB)")


def test_truncated_dimension_metadata():
    """
    Regression test: files exported from Google Sheets/Lark can have a
    <dimension> tag in the sheet XML that under-reports the real column
    range. detect_headers() must not silently drop columns because of this
    (found via a real Cogistics export — see engine.py detect_headers docstring).
    """
    print("\n=== truncated_dimension_metadata (regression) ===")
    src = os.path.join(OUT, "dim_regression_source.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["title"])
    ws.append(["Select", "Product Name", "Dimension (Spec)", "Supply Advantage", "Price Range (THB/kg)"])
    ws.append(["Yes", "Item A", "10x10", "adv", "100"])
    wb.save(src)

    with zipfile.ZipFile(src, "r") as z:
        sheet_xml = z.read("xl/worksheets/sheet1.xml").decode()
    bad_xml = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="A1:A3"/>', sheet_xml)

    corrupted = os.path.join(OUT, "dim_regression_corrupted.xlsx")
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(corrupted, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet1.xml":
                data = bad_xml.encode()
            zout.writestr(item, data)

    headers = engine.detect_headers(corrupted, header_row=2)
    print(f"Detected headers: {headers}")
    assert len(headers) == 5, f"Expected 5 headers, got {len(headers)}: {headers}"
    print("PASS")


def test_signature_only_on_last_page():
    """
    Regression test: the quotation preset's signature block must appear
    exactly once, pinned near the bottom of the LAST page — never on
    earlier pages, and never flowing loose wherever the table happens to end.
    """
    print("\n=== signature_only_on_last_page (regression) ===")
    with open(os.path.join(HERE, "presets", "quotation_sheet.json")) as f:
        config = DocumentConfig(**json.load(f))

    xlsx_path = os.path.join(OUT, "sig_regression_input.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["title"])
    ws.append(["Select", "Product Name", "Quantity", "Price Range (THB/kg)", "Remarks"])
    for i in range(40):  # enough rows to force multiple pages
        ws.append(["Yes", f"Item {i+1:02d}", str(10 * (i + 1)), f"{5.5+i}", "remark"])
    wb.save(xlsx_path)

    items = engine.read_excel(xlsx_path, config)
    out_pdf = os.path.join(OUT, "sig_regression_output.pdf")
    engine.generate_pdf(items, config, out_pdf, ASSETS)

    import subprocess
    result = subprocess.run(["pdftotext", "-layout", out_pdf, "-"], capture_output=True, text=True)
    pages = result.stdout.split("\x0c")  # form-feed separates pages in pdftotext output
    pages = [p for p in pages if p.strip()]
    print(f"Generated {len(pages)} page(s)")
    assert len(pages) >= 2, "Test setup should force multiple pages — check row count"

    occurrences = [i for i, p in enumerate(pages) if "Prepared by" in p]
    print(f"'Prepared by' found on page index(es): {occurrences}")
    assert occurrences == [len(pages) - 1], (
        f"Expected signature only on the last page (index {len(pages)-1}), found on {occurrences}"
    )
    print("PASS")


def test_thai_text_renders():
    """
    Regression test: Thai text must render as real embedded text (not boxes,
    not silently dropped). Found via a real Cogistics export with a Thai
    company name — Helvetica has no Thai glyphs and reportlab silently drew
    black boxes instead of raising an error. Extracts text back out of the
    generated PDF with pdftotext to prove the actual Thai characters are
    embedded, not just visually present in a screenshot.
    """
    print("\n=== thai_text_renders (regression) ===")
    xlsx_path = os.path.join(OUT, "thai_regression_input.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["title"])
    ws.append(["Select", "Product Name", "Dimension (Spec)", "Supply Advantage", "Price Range (THB/kg)"])
    thai_name = "บริษัท ออน-กรีน โปรดิวส์ จำกัด"
    ws.append(["Yes", thai_name, "10x10", "adv", "100"])
    wb.save(xlsx_path)

    with open(os.path.join(HERE, "presets", "client_catalog.json")) as f:
        config = DocumentConfig(**json.load(f))

    items = engine.read_excel(xlsx_path, config)
    out_pdf = os.path.join(OUT, "thai_regression_output.pdf")
    engine.generate_pdf(items, config, out_pdf, ASSETS)

    import subprocess
    result = subprocess.run(["pdftotext", "-layout", out_pdf, "-"], capture_output=True, text=True)
    extracted = result.stdout
    # Check each word independently rather than requiring one exact contiguous
    # match: pdftotext -layout interleaves table columns in reading order, so
    # when a long cell value wraps onto two lines, other columns' content from
    # the same row gets extracted in between — a text-extraction-order quirk,
    # not a rendering defect (confirmed correct visually via screenshot).
    words = thai_name.replace("-", " ").split()
    missing = [w for w in words if w not in extracted]
    assert not missing, (
        f"Thai word(s) not found as real text in the generated PDF (rendered "
        f"as boxes or dropped): {missing}. Extracted text was: {extracted!r}"
    )
    print("PASS: Thai text extracted correctly from the generated PDF")


def test_quotation_manual_subtotal_and_total():
    """
    Regression test: quotation_sheet's Subtotal is now a MANUALLY typed,
    optional column (real-world Quantity values are things like "2 tons" or
    "500 kg", so auto-calculating price x qty didn't work). Verifies that
    typed-in subtotals flow through and that the grand Total row still sums
    them, skipping any row left blank or typed as non-numeric text.
    """
    print("\n=== quotation_manual_subtotal_and_total (regression) ===")
    with open(os.path.join(HERE, "presets", "quotation_sheet.json")) as f:
        raw = json.load(f)
    for col in raw["columns"]:
        if col["key"] == "qty":
            col["source_header"] = "Quantity"  # map it, as the UI would require
    config = DocumentConfig(**raw)

    subtotal_col = next(c for c in config.columns if c.key == "subtotal")
    assert subtotal_col.source == "manual", "Subtotal should be a manual-entry column"
    assert subtotal_col.optional is True, "Subtotal should be optional"

    xlsx_path = os.path.join(OUT, "subtotal_regression_input.xlsx")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["title"])
    ws.append(["Select", "Product Name", "Price Range (THB/kg)", "Quantity", "Remarks"])
    ws.append(["Yes", "Box", "12.50", "2 tons", ""])
    ws.append(["Yes", "Wrap", "45.00", "500 kg", ""])
    ws.append(["Yes", "No-Subtotal Item", "10.00", "1 pallet", ""])
    wb.save(xlsx_path)

    items = engine.read_excel(xlsx_path, config)
    # Simulate what the UI sends: subtotals typed in per row, third left blank.
    engine.apply_manual_values(items, {"subtotal": ["6,250.00", "2,250.00", ""]})
    engine.apply_computed_values(items, config)  # no-op now, but must not clobber manual values

    print(f"Items: {items}")
    assert items[0]["subtotal"] == "6,250.00", f"Got {items[0]['subtotal']!r}"
    assert items[1]["subtotal"] == "2,250.00", f"Got {items[1]['subtotal']!r}"
    assert items[2]["subtotal"] == "", f"Expected blank, got {items[2]['subtotal']!r}"
    # Non-numeric Quantity must survive untouched — it's just text now.
    assert items[0]["qty"] == "2 tons", f"Got {items[0]['qty']!r}"

    out_pdf = os.path.join(OUT, "subtotal_regression_output.pdf")
    engine.generate_pdf(items, config, out_pdf, ASSETS)

    import subprocess
    result = subprocess.run(["pdftotext", "-layout", out_pdf, "-"], capture_output=True, text=True)
    assert "8,500.00" in result.stdout, f"Expected grand total 8,500.00, got: {result.stdout!r}"
    assert "2 tons" in result.stdout, "Non-numeric Quantity should render as-is"
    print("PASS: manual subtotals render and grand Total sums them correctly")


def test_computed_columns_still_work():
    """
    The "computed" column type is no longer used by any shipped preset (the
    quotation switched to manual subtotals), but the feature remains
    supported for any future preset with genuinely numeric source columns.
    Tested directly against a synthetic config so the capability doesn't
    silently rot.
    """
    print("\n=== computed_columns_still_work ===")
    config = DocumentConfig(
        id="synthetic", name="Synthetic", document_title="Synthetic",
        columns=[
            ColumnConfig(key="name", label="Item", source_header="Product Name"),
            ColumnConfig(key="price", label="Price", source_header="Price"),
            ColumnConfig(key="qty", label="Qty", source_header="Qty"),
            ColumnConfig(key="total", label="Line Total", source="computed",
                         compute_operation="multiply", compute_operands=["price", "qty"]),
        ],
    )
    items = [
        {"name": "A", "price": "10.00", "qty": "3"},
        {"name": "B", "price": "7.25", "qty": "4"},
        {"name": "C", "price": "5.00", "qty": ""},      # missing operand -> blank
        {"name": "D", "price": "180-220", "qty": "2"},  # range, not a number -> blank
    ]
    engine.apply_computed_values(items, config)
    print(f"Items: {items}")
    assert items[0]["total"] == "30.00", f"Got {items[0]['total']!r}"
    assert items[1]["total"] == "29.00", f"Got {items[1]['total']!r}"
    assert items[2]["total"] == "", f"Got {items[2]['total']!r}"
    assert items[3]["total"] == "", f"Got {items[3]['total']!r}"
    print("PASS: computed columns calculate correctly and fail soft on bad operands")


def test_excel_format_sniffing():
    """Regression: the TMS export is named .xls but is really an .xlsx inside.
    The upload must judge files by content, not by extension."""
    import io, zipfile
    print("\n=== excel_format_sniffing (regression) ===")

    # A real workbook is recognised as xlsx...
    buf = io.BytesIO()
    wb = openpyxl.Workbook()
    wb.active.append(["x"])
    wb.save(buf)
    assert engine.detect_excel_format(buf.getvalue()) == "xlsx", "real .xlsx not recognised"

    # ...a genuine legacy binary .xls is identified as such (needs a clear message, not a crash)...
    legacy = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\x00" * 64
    assert engine.detect_excel_format(legacy) == "xls", "legacy .xls not recognised"

    # ...and a zip that is NOT a workbook (e.g. a .docx) is rejected, as is garbage.
    docx_like = io.BytesIO()
    with zipfile.ZipFile(docx_like, "w") as z:
        z.writestr("word/document.xml", "<x/>")
    assert engine.detect_excel_format(docx_like.getvalue()) == "unknown", ".docx-like zip accepted"
    assert engine.detect_excel_format(b"%PDF-1.7 not a spreadsheet") == "unknown", "PDF accepted"
    assert engine.detect_excel_format(b"") == "unknown", "empty file accepted"
    print("PASS: workbook detected by content; legacy .xls, non-Excel zips and junk rejected")


def _page_boxes(pdf_path):
    """[(width, height), ...] of every page, read straight from the PDF bytes."""
    data = open(pdf_path, "rb").read().decode("latin-1")
    return [(round(float(a)), round(float(b)))
            for a, b in re.findall(r"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]", data)]


def test_landscape_orientation():
    """The orientation setting changes page size, and ONLY when asked.
    Existing (portrait) presets must be unaffected."""
    print("\n=== landscape_orientation ===")
    xlsx = os.path.join(OUT, "_orient.xlsx")
    build_quotation_xlsx(xlsx)
    raw = json.load(open(os.path.join(HERE, "presets", "quotation_sheet.json")))

    for orientation, expected in [(None, (595, 842)), ("portrait", (595, 842)), ("landscape", (842, 595))]:
        d = dict(raw)
        if orientation:
            d["orientation"] = orientation
        cfg = DocumentConfig(**d)
        items = engine.read_excel(xlsx, cfg)
        pdf = os.path.join(OUT, "_orient.pdf")
        engine.generate_pdf(items, cfg, pdf, ASSETS)
        boxes = _page_boxes(pdf)
        assert boxes and all(b == expected for b in boxes), f"orientation={orientation!r}: got {boxes}, wanted {expected}"
    print("PASS: default and 'portrait' stay A4 portrait; 'landscape' gives A4 landscape")


def test_grouped_report_config():
    """Every shipped report JSON validates, and the schema rejects the mistakes
    that would otherwise fail confusingly at generation time."""
    print("\n=== grouped_report_config ===")
    reports_dir = os.path.join(HERE, "reports")
    files = sorted(f for f in os.listdir(reports_dir) if f.endswith(".json"))
    assert files, "no report configs found in backend/reports/"
    for f in files:
        raw = json.load(open(os.path.join(reports_dir, f), encoding="utf-8"))
        cfg = GroupedReportConfig(**raw)
        assert cfg.id == os.path.splitext(f)[0], f"{f}: id {cfg.id!r} must match the filename"

    good = json.load(open(os.path.join(reports_dir, files[0]), encoding="utf-8"))

    def must_reject(mutate, why):
        bad = json.loads(json.dumps(good))
        mutate(bad)
        try:
            GroupedReportConfig(**bad)
        except ValueError:
            return
        raise AssertionError(f"schema accepted a bad config: {why}")

    must_reject(lambda c: c["sections"][1].update(key=c["sections"][0]["key"]), "duplicate section keys")
    must_reject(lambda c: c["sections"][1].update(sheet_name=c["sections"][0]["sheet_name"]), "two sections on one sheet")
    must_reject(lambda c: c["sections"][0]["columns"][0].update(source="manual"), "manual column in a report")
    must_reject(lambda c: c["sections"][0]["columns"][0].update(source_header=None), "column with no source_header")
    must_reject(lambda c: c["sections"][0]["columns"].append(dict(c["sections"][0]["columns"][0])), "duplicate column keys")
    must_reject(lambda c: c.update(status_header=None), "expected_status without status_header")
    must_reject(lambda c: c.update(sections=[]), "no sections")
    print(f"PASS: {len(files)} report config(s) valid; 7 kinds of bad config rejected")


if __name__ == "__main__":
    run_case("client_catalog", build_client_catalog_xlsx, "client_catalog.json")
    run_case("quotation_sheet", build_quotation_xlsx, "quotation_sheet.json")
    test_truncated_dimension_metadata()
    test_signature_only_on_last_page()
    test_thai_text_renders()
    test_quotation_manual_subtotal_and_total()
    test_computed_columns_still_work()
    test_excel_format_sniffing()
    test_landscape_orientation()
    test_grouped_report_config()
    print("\nAll cases passed.")
