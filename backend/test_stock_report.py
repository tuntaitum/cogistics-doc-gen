"""
test_stock_report.py — tests for stock_report.py (the 3PL daily stock movement summary PDF).

Synthetic data only. Reads the generated PDFs back with pdftotext (poppler), as the other test files do.
Run:  python test_stock_report.py
"""

import os
import re
import shutil
import subprocess
import tempfile
from datetime import date

import stock_ledger as sl
import stock_master as sm
import stock_report as sr
import test_stock as ts
from schemas import StockReportConfig

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
TMP = tempfile.mkdtemp(prefix="stock_report_test_")
CFG = sr.load_stock_config()
D = date(2026, 10, 6)


def product(sku, name, per, client="Eurosia", **kw):
    return sm.Product(sku, name, [client], per, **kw)


MASTER = sm.ProductMaster(products={
    "3001868": product("3001868", "Mango bar 78g", 30),
    "3001970": product("3001970", "Fruit tart 26g", 20),
    "9000001": product("9000001", "Bulk item", None),                       # no carton size
    "R0201": product("R0201", "Smoothie", None, "Real Simple", storage_unit="Crate", unit="Bottle"),
}, updated_at="2026-10-16T01:00:00+00:00")


def pdf_text(path):
    out = subprocess.run(["pdftotext", path, "-"], capture_output=True, text=True, check=True).stdout
    return " ".join(out.split())                       # join lines: long cells may wrap


def page_boxes(path):
    data = open(path, "rb").read().decode("latin-1")
    return [(round(float(a)), round(float(b))) for a, b in
            re.findall(r"/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)\s*\]", data)]


def make(rows, client="Eurosia", when=D, master=MASTER, name="ledger.xlsx", **kw):
    out = os.path.join(TMP, "out_" + name)
    shutil.rmtree(out, ignore_errors=True)
    return sr.generate_stock_summary(ts.write(name, rows), client, when, master, CFG, out, ASSETS, **kw), out


def test_full_report():
    print("\n=== full_report ===")
    res, out = make(ts.CLEAN)
    assert res.filename == "Eurosia - Daily Stock Movement - 2026-10-06.pdf" and os.path.exists(res.path)
    assert page_boxes(res.path) and all(b == (842, 595) for b in page_boxes(res.path))      # landscape A4
    t = pdf_text(res.path)

    assert "Daily Stock Movement Summary" in t and "Eurosia — 06-10-2026" in t
    assert "Warehouses: CGT1, CGT6" in t and "cartons + loose pieces" in t
    for sku in ("3001868", "3001970", "0200639", "9000001"):
        assert sku in t
    assert "Mango bar 78g" in t and "Fruit tart 26g" in t and "(not in product master data)" in t

    # cartons + leftover pieces, from the master's carton sizes
    assert "3 ctn + 10 pcs" in t                      # 3001868 opening: 100 pcs at 30/ctn
    assert "1 ctn + 10 pcs" in t                      # 3001868 out: 40
    assert "2 ctn" in t and "60 pcs" in t             # 3001868 closing: 60 = 2 cartons exactly
    assert "8 ctn + 10 pcs" in t and "170 pcs" in t   # 3001970 closing: 170 at 20/ctn
    assert "CGT1: 5 ctn + 1 pcs" in t and "CGT6: 3 ctn + 9 pcs" in t      # warehouses kept apart (101 and 69)
    assert "1,134 pcs" in t and "carton size not on file" in t and "have no carton size in the product master data" in t

    # totals are in pieces (SKUs have different carton sizes); batch text is exactly as the WMS wrote it
    assert "1,374 pcs" in t and "1,400 pcs" in t and "46 pcs" in t and "4 SKUs" in t
    assert ts.BATCH_THAI in t
    assert "product master data of 16-10-2026" in t
    assert "does not add up" not in t                  # no reconciliation flag on a healthy file

    assert (res.sku_count, res.batch_count, res.closing_qty, res.in_qty, res.out_qty) == (4, 4, 1374, 20, 46)
    assert res.all_reconciled and res.check.verdict == "warn"
    assert any("0200639" in w and "not in the product master data" in w for w in res.warnings)
    assert not any("You entered" in w for w in res.warnings)
    print("PASS: layout, cartons + pieces, warehouses, totals, batch text, footnotes and warnings")


def test_wrong_client_writes_nothing():
    print("\n=== wrong_client ===")
    out = os.path.join(TMP, "out_wrong")
    shutil.rmtree(out, ignore_errors=True)
    try:
        sr.generate_stock_summary(ts.write("wrong.xlsx", ts.CLEAN), "Real Simple", D, MASTER, CFG, out, ASSETS)
    except sr.ClientMismatchError as e:
        assert "Eurosia" in str(e) and e.check.blocking
        assert not os.path.exists(out) or not os.listdir(out), "a PDF was written despite the wrong client"
        print("PASS: the wrong client is refused, names the real owner, and writes nothing")
        return
    raise AssertionError("a file for another client was accepted")


def test_date_and_reconciliation_flags():
    print("\n=== date_and_reconciliation ===")
    res, _ = make(ts.CLEAN, when=date(2026, 10, 7), name="wrongdate.xlsx")
    assert any("You entered 07-10-2026" in w and "06-10-2026" in w for w in res.warnings)
    assert "07-10-2026" in pdf_text(res.path), "the PDF should carry the date the person confirmed"

    rows = [ts.opening("3001868", "b", 162), ts.total("3001868", "b", 150)]                 # WMS says 150, maths says 162
    res, _ = make(rows, name="mismatch.xlsx")
    t = pdf_text(res.path)
    assert not res.all_reconciled and "closing balance differs from opening" in t            # the footnote
    assert t.count("does not add up") >= 3, "flag must be in the SKU row, the batch row AND the footnote"
    assert "5 ctn" in t and "150 pcs" in t                                  # closing is the WMS figure: 150 pcs = 5 ctn
    assert res.closing_qty == 150 and "162 pcs" in t                        # (opening was 162, so the 162 is the opening total)
    assert any("does not add up" in w for w in res.warnings)
    print("PASS: a mismatched date and a batch that doesn't add up are flagged, and the WMS figure is what's printed")


def test_single_warehouse_and_empty_cases():
    print("\n=== single_warehouse_and_errors ===")
    rows = [ts.opening("3001868", "b", 100), ts.move("3001868", "b", "W1", out=10), ts.total("3001868", "b", 90)]
    res, _ = make(rows, name="single.xlsx")
    t = pdf_text(res.path)
    assert "Warehouse: CGT1." in t and "Closing by warehouse" not in t and "carton size not on file" not in t

    for fragment, ledger_rows, name in [("no stock rows", [], "empty.xlsx")]:
        try:
            make(ledger_rows, name=name)
        except sr.StockReportError as e:
            assert fragment in str(e)
        else:
            raise AssertionError("an empty ledger produced a report")
    try:
        sr.generate_stock_summary(ts.write("x.xlsx", [], header=["a", "b"]), "Eurosia", D, MASTER, CFG, TMP, ASSETS)
    except sl.LedgerError:
        pass
    else:
        raise AssertionError("a non-ledger file produced a report")
    print("PASS: one warehouse hides the split column; empty and non-ledger files are refused")


def test_special_characters_and_filenames():
    print("\n=== special_characters ===")
    master = sm.ProductMaster(products={"3001868": product("3001868", "Mango <special> & co", 30, "A & B <Foods>")})
    rows = [ts.opening("3001868", "b<1>&", 100), ts.total("3001868", "b<1>&", 100)]
    res, _ = make(rows, client="A & B <Foods>", master=master, name="special.xlsx")
    t = pdf_text(res.path)
    assert "A & B <Foods> — 06-10-2026" in t and "Mango <special> & co" in t and "b<1>&" in t
    assert not re.search(r'[\\/:*?"<>|]', res.filename), res.filename
    assert sr.stock_filename(CFG, 'We/ird: "Name"', D) == "We-ird- -Name- - Daily Stock Movement - 2026-10-06.pdf"
    print("PASS: '&' and '<' in client, product and batch text survive; file names are safe")


def test_many_skus_paginate():
    print("\n=== many_skus ===")
    rows = [ts.opening("3001868", "b", 100), ts.total("3001868", "b", 100)]
    for i in range(1, 91):
        rows += [ts.opening(f"S{i:04d}", "b", 100 + i), ts.total(f"S{i:04d}", "b", 100 + i)]
    res, _ = make(rows, name="many.xlsx")
    t = pdf_text(res.path)
    assert len(page_boxes(res.path)) >= 4, "91 SKUs plus 91 batches should run onto several pages"
    assert all(f"S{i:04d}" in t for i in (1, 45, 90)) and "91 SKUs" in t
    print(f"PASS: {len(page_boxes(res.path))} pages, every SKU present, totals row at the end")


def test_config():
    print("\n=== config ===")
    raw = CFG.model_dump()
    assert CFG.orientation == "landscape" and CFG.id == "stock_summary"
    for bad in ("{clientt} - x", "{date_iso"):
        try:
            StockReportConfig(**dict(raw, filename_template=bad))
        except ValueError:
            continue
        raise AssertionError(f"accepted a bad template {bad!r}")
    assert sr.stock_filename(CFG.model_copy(update={"filename_template": "{date_iso} {client}"}), "Acme", D) == "2026-10-06 Acme.pdf"
    print("PASS: shipped config is valid; placeholder typos are rejected when it loads")


if __name__ == "__main__":
    test_full_report()
    test_wrong_client_writes_nothing()
    test_date_and_reconciliation_flags()
    test_single_warehouse_and_empty_cases()
    test_special_characters_and_filenames()
    test_many_skus_paginate()
    test_config()
    print("\nAll stock report cases passed.")
