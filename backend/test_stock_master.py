"""
test_stock_master.py — tests for stock_master.py (product master data for the 3PL stock report).

Synthetic data only (names and IDs invented), laid out like the real Lark Base export.
Run:  python test_stock_master.py
"""

import json
import os
import tempfile
from datetime import date

import openpyxl

import stock_ledger as sl
import stock_master as sm
import test_stock as ts

TMP = tempfile.mkdtemp(prefix="master_test_")

HEADER = ["Product Name", "Product ID", "Alt ID", "รับเข้าเมื่อ", "Photo", "Clients", "Storage Temp", "Unit",
          "Storage Unit", "Carton L (cm)", "Unit per Storage Unit", "Unit per Delivery Unit", "น้ำหนักต่อชิ้น (kg)"]


def prod(name, sku, clients, per="30 pieces per carton", storage="Carton", unit="Piece", wt=0.1, alt=None):
    # other columns are present in the real export and must simply be ignored
    return [name, sku, alt, "2026-01-01", None, clients, "Chilled", unit, storage, 40, per, "80 pieces per basket", wt]


def write_master(name, rows, header=HEADER):
    path = os.path.join(TMP, name)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Products"
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


ROWS = [
    prod("  Mango bar 78g  ", "3001868", "Eurosia", "30 pieces per carton", wt=0.078),     # padded name
    prod("Cheese slice", "3001869", "Eurosia", "36 pieces per carton", wt=0.12),
    prod("Kimchi", "0200639", "CJ Foods Thailand", "24 pieces per carton", alt="7-Eleven: 0200639\nCJ More: 893"),   # leading zero
    prod("Hot kimchi", None, "CJ Foods Thailand", "24 pieces per carton", alt="CJ More: 8935297106567"),            # no ID
    prod("Smoothie A", "R0201", "Real Simple", None, "Crate", "Bottle"),                  # no carton size, crates/bottles
    prod("Smoothie B", "R0202", "Real Simple", None, "Crate", "Bottle"),
    prod("Dup of mango", "3001868", "Eurosia", "25 pieces per carton"),                   # duplicate ID, different size
    prod("Odd size", "R0203", "Real Simple", "varies"),                                   # can't read the size
    prod("Zero size", "R0204", "Real Simple", "0 pieces per carton"),
    prod("Shared item", "M0001", "Eurosia, Real Simple", "12 pieces per carton"),         # two owners in one cell
    prod("Numeric id", 3002088, "Eurosia", "14 pieces per carton"),                       # ID stored as a number
    prod("Orphan", "N0001", None, "10 pieces per carton"),                                # no client
    [None] * len(HEADER),                                                                 # blank row
]


def test_parse_master():
    print("\n=== parse_master ===")
    M = sm.parse_master_workbook(write_master("m.xlsx", ROWS), "m.xlsx")
    P = M.products
    assert "0200639" in P and P["0200639"].pieces_per_carton == 24            # leading zero kept
    assert P["3001868"].name == "Mango bar 78g" and P["3001868"].pieces_per_carton == 30   # first row wins, name trimmed
    assert P["3001868"].weight_per_piece == 0.078 and "3002088" in P          # numeric ID cell -> text
    assert P["R0201"].pieces_per_carton is None and P["R0201"].storage_unit == "Crate" and P["R0201"].unit == "Bottle"
    assert P["M0001"].clients == ["Eurosia", "Real Simple"]
    assert M.clients() == ["CJ Foods Thailand", "Eurosia", "Real Simple"]     # sorted; 'Orphan' adds no client
    assert M.lookup(" 3001869 ").name == "Cheese slice" and M.lookup("nope") is None

    text = "\n".join(M.problems)
    assert "no Product ID (Alt ID: CJ More: 8935297106567)" in text            # unmatched product reported
    assert "2 product(s) have no carton size" in text and "R0201, R0202" in text   # grouped into ONE message
    assert text.count("no carton size") == 1
    assert "Product ID 3001868 is listed twice" in text and "30" in text and "25" in text   # duplicate, both sizes shown
    assert "carton size 'varies' can't be read" in text and "must be more than 0" in text
    assert "SKU N0001: no client" in text
    assert M.source_filename == "m.xlsx" and M.updated_at.endswith("+00:00")
    print(f"PASS: {len(P)} products parsed; leading zero, numeric IDs, two-owner cell, and {len(M.problems)} problems handled")


def test_unreadable_masters():
    print("\n=== unreadable_masters ===")
    def expect(fragment, path):
        try:
            sm.parse_master_workbook(path)
        except sm.MasterError as e:
            assert fragment in str(e), str(e)
            return
        raise AssertionError(f"expected MasterError containing {fragment!r}")
    expect("doesn't look like the product master data", write_master("wrong.xlsx", [["a", "b"]], header=["x", "y"]))
    noper = [h for h in HEADER if h != "Unit per Storage Unit"]
    expect("missing column(s): Unit per Storage Unit", write_master("nocol.xlsx", [], header=noper))
    junk = os.path.join(TMP, "junk.xlsx")
    open(junk, "wb").write(b"not a workbook")
    expect("Could not open", junk)
    print("PASS: a wrong file, a missing column and a junk file each get a clear message")


def test_cartons_conversion():
    print("\n=== cartons_conversion ===")
    # (pieces, per carton, expected cartons, expected leftover pieces) — the first rows are your real figures
    for pieces, per, ctn, pcs in [(1061, 30, 35, 11), (2942, 36, 81, 26), (3521, 20, 176, 1), (7737, 20, 386, 17),
                                  (178, 14, 12, 10), (21, 20, 1, 1), (0, 30, 0, 0), (30, 30, 1, 0), (29, 30, 0, 29),
                                  (5, 1, 5, 0), (59.9999999, 30, 2, 0), (2.5, 30, 0, 3)]:
        q = sm.to_cartons(pieces, per)
        assert (q.cartons, q.pieces) == (ctn, pcs), (pieces, per, q)
        assert q.cartons * per + q.pieces == q.total, "cartons and pieces must add back to the total"
    q = sm.to_cartons(-65, 30)
    assert (q.cartons, q.pieces, q.total) == (-2, -5, -65)
    for unknown in (None, 0):
        q = sm.to_cartons(178, unknown)
        assert (q.cartons, q.pieces, q.total) == (None, 178, 178)

    F = sm.format_cartons
    assert F(sm.to_cartons(1061, 30)) == "35 ctn + 11 pcs"
    assert F(sm.to_cartons(360, 30)) == "12 ctn" and F(sm.to_cartons(9, 14)) == "9 pcs" and F(sm.to_cartons(0, 14)) == "0"
    assert F(sm.to_cartons(178, None)) == "178 pcs" and F(sm.to_cartons(12345, 30)) == "411 ctn + 15 pcs"
    assert F(sm.to_cartons(-65, 30)) == "-2 ctn + 5 pcs"
    assert F(sm.to_cartons(70, 20), "crate", "btl") == "3 crate + 10 btl"

    M = sm.parse_master_workbook(write_master("u.xlsx", ROWS))
    assert sm.unit_labels(M.lookup("3001868")) == ("ctn", "pcs") and sm.unit_labels(M.lookup("R0201")) == ("crate", "btl")
    assert sm.unit_labels(None) == ("ctn", "pcs")                      # no product: sensible defaults
    print("PASS: cartons + leftover always add back to the total; zero, exact, negative, fractional and unknown sizes handled")


def test_client_check():
    print("\n=== client_check ===")
    M = sm.parse_master_workbook(write_master("c.xlsx", ROWS))
    ok = sm.check_client(M, "Eurosia", ["3001868", "3001869", "3002088"])
    assert ok.verdict == "ok" and not ok.messages and not ok.blocking and ok.matched == ["3001868", "3001869", "3002088"]

    # another client's SKU mixed in -> warning, still allowed
    mixed = sm.check_client(M, "Eurosia", ["3001868", "0200639"])
    assert mixed.verdict == "warn" and "different client" in mixed.messages[0] and "CJ Foods Thailand" in mixed.messages[0]
    assert mixed.other_clients == {"0200639": ["CJ Foods Thailand"]}

    # WRONG CLIENT: nothing in the file belongs to the one selected -> blocked, and says who it belongs to
    wrong = sm.check_client(M, "Real Simple", ["3001868", "3001869", "3002088"])
    assert wrong.verdict == "stop" and wrong.blocking
    assert "None of the SKUs in this file belong to Real Simple" in wrong.messages[0] and "Eurosia (3 SKUs)" in wrong.messages[0]

    # a SKU owned by two clients is fine for either
    assert sm.check_client(M, "Eurosia", ["M0001"]).verdict == "ok" and sm.check_client(M, "Real Simple", ["M0001"]).verdict == "ok"

    # unknown SKUs: warning when the rest match; a stop when NOTHING is known (can't verify the client at all)
    some = sm.check_client(M, "Eurosia", ["3001868", "9999999"])
    assert some.verdict == "warn" and some.unknown == ["9999999"] and "pieces only" in some.messages[0]
    none = sm.check_client(M, "Eurosia", ["9999999"])
    assert none.verdict == "stop" and "can't be checked" in none.messages[0]

    # a lost leading zero is hinted, never matched
    hint = sm.check_client(M, "CJ Foods Thailand", ["200639", "3001868"])
    assert M.lookup("200639") is None and "did you mean 0200639" in " ".join(hint.messages)
    assert sm.check_client(M, "CJ Foods Thailand", ["0200639"]).verdict == "ok"

    # matched but no carton size -> warning naming them
    nocarton = sm.check_client(M, "Real Simple", ["R0201", "R0202"])
    assert nocarton.verdict == "warn" and nocarton.no_carton_size == ["R0201", "R0202"]

    # client names that don't exist, and duplicate / blank SKUs in the input
    assert sm.check_client(M, "Nobody", ["3001868"]).blocking
    assert sm.check_client(M, "Eurosia", ["3001868", "3001868", "", "  "]).matched == ["3001868"]
    print("PASS: ok / warn / stop verdicts; a wrong client is blocked and the real owner is named")


def test_carton_view_on_a_real_shaped_ledger():
    print("\n=== carton_view ===")
    L = sl.read_stock_ledger(ts.write("cv.xlsx", ts.CLEAN))
    master = sm.ProductMaster(products={
        "3001868": sm.Product("3001868", "Mango", ["Eurosia"], 30),
        "3001970": sm.Product("3001970", "Tart", ["Eurosia"], 20),
        "0200639": sm.Product("0200639", "Kimchi", ["CJ Foods Thailand"], None)})
    s = {x.sku: x for x in L.sku_summaries()}
    a = sm.carton_view(s["3001868"], master.lookup("3001868"))          # opening 100, out 40, closing 60 at 30/ctn
    assert (a["opening"].cartons, a["opening"].pieces) == (3, 10) and (a["out"].cartons, a["out"].pieces) == (1, 10)
    assert (a["closing"].cartons, a["closing"].pieces) == (2, 0) and a["in"].total == 0
    b = sm.carton_view(s["3001970"], master.lookup("3001970"))          # receipt of 20 is included; closing 170 at 20/ctn
    assert (b["in"].cartons, b["closing"].cartons, b["closing"].pieces) == (1, 8, 10)
    c = sm.carton_view(s["0200639"], master.lookup("0200639"))          # no carton size -> pieces only
    assert c["closing"].cartons is None and c["closing"].total == 10
    d = sm.carton_view(s["9000001"], None)                              # not in the master at all
    assert d["closing"].cartons is None and d["closing"].total == 1134
    print("PASS: ledger quantities convert correctly, including receipts, unknown carton sizes and unknown SKUs")


def test_store_and_diff():
    print("\n=== store_and_diff ===")
    store = sm.MasterStore(os.path.join(TMP, "masterdata"))
    assert store.load() is None and store.restore_previous() is False        # nothing uploaded yet

    first = sm.parse_master_workbook(write_master("s1.xlsx", ROWS), "first.xlsx")
    store.save(first)
    back = store.load()
    assert list(back.products) == list(first.products) and back.clients() == first.clients()
    assert back.products["0200639"].pieces_per_carton == 24 and back.problems == first.problems
    assert back.source_filename == "first.xlsx" and store.load_previous() is None

    newer_rows = [r for r in ROWS if r[1] != "3001869"] + [prod("New item", "3009999", "Eurosia", "50 pieces per carton")]
    newer_rows[0] = prod("Mango bar 78g", "3001868", "Eurosia", "32 pieces per carton")          # carton size changed
    second = sm.parse_master_workbook(write_master("s2.xlsx", newer_rows), "second.xlsx")
    d = sm.diff_masters(store.load(), second)
    assert d["added"] == ["3009999"] and d["removed"] == ["3001869"] and d["carton_changed"] == [("3001868", 30, 32)]
    assert sm.diff_masters(None, second)["added"] == list(second.products)

    store.save(second)
    assert store.load().source_filename == "second.xlsx" and store.load_previous().source_filename == "first.xlsx"
    assert store.restore_previous() is True and store.load().source_filename == "first.xlsx"   # undo a bad upload
    assert not [f for f in os.listdir(store.directory) if f.endswith(".tmp")], "temp files left behind"

    open(store.path, "w").write("{ this is not json")
    try:
        store.load()
    except sm.MasterError as e:
        assert "Upload it again" in str(e)
    else:
        raise AssertionError("a corrupt file was accepted")
    open(store.path, "w").write(json.dumps({"products": [{"sku": "x"}]}))
    try:
        store.load()
    except sm.MasterError as e:
        assert "damaged" in str(e)
    else:
        raise AssertionError("an incomplete file was accepted")
    print("PASS: saved data round-trips; previous copy kept and restorable; corrupt files give a clear error; diff shows changes")


if __name__ == "__main__":
    test_parse_master()
    test_unreadable_masters()
    test_cartons_conversion()
    test_client_check()
    test_carton_view_on_a_real_shaped_ledger()
    test_store_and_diff()
    print("\nAll stock master cases passed.")
