"""
test_api.py — end-to-end tests of the grouped-report HTTP API.

Separate from test_engine.py because it needs one extra package just for tests:
    pip install httpx
(see requirements-dev.txt). Run with:  python test_api.py

Runs against a throwaway DATA_DIR, so it never touches real uploads/output.
"""

import io
import os
import tempfile
import time
import uuid
import zipfile
from urllib.parse import quote

# Must be set BEFORE main is imported: main reads DATA_DIR at import time.
_TMP = tempfile.mkdtemp(prefix="codocs_test_")
os.environ["DATA_DIR"] = _TMP

from fastapi.testclient import TestClient   # noqa: E402

import main                                  # noqa: E402
from test_engine import _load_report_cfg, build_tms_report_xlsx   # noqa: E402

client = TestClient(main.app)
CFG = _load_report_cfg()
REPORT = CFG.id
WORKBOOK = os.path.join(_TMP, "tms.xlsx")


def upload(path=WORKBOOK, name="ReportDailyDeliveryByShippingNoteList.xls"):
    """Upload like the browser does — under the TMS's real (misleading) .xls name."""
    with open(path, "rb") as f:
        r = client.post("/api/upload", files={"file": (name, f.read())}, data={"header_row": 2})
    assert r.status_code == 200, r.text
    return r.json()["session_id"]


def generate(session_id, report=REPORT):
    return client.post(f"/api/reports/{report}/generate", data={"session_id": session_id})


def test_list_reports():
    print("\n=== list_reports ===")
    r = client.get("/api/reports")
    assert r.status_code == 200
    ids = [x["id"] for x in r.json()]
    assert REPORT in ids, ids
    entry = next(x for x in r.json() if x["id"] == REPORT)
    assert [s["key"] for s in entry["sections"]] == ["delivered", "in_transit", "failed"]
    # the veggie preset picker must not see report configs
    assert REPORT not in [p["id"] for p in client.get("/api/presets").json()]
    print("PASS: report types listed; veggie preset picker unaffected")


def test_generate_and_download():
    print("\n=== generate_and_download ===")
    build_tms_report_xlsx(WORKBOOK, CFG)
    r = generate(upload())
    assert r.status_code == 200, r.text
    s = r.json()
    assert s["report_date"] == "06-10-2026" and s["client_count"] == 4
    assert len(s["warnings"]) >= 3, s["warnings"]
    assert {c["name"]: c["total"] for c in s["clients"]}["Alpha Foods Co., Ltd."] == 5

    # zip: one valid PDF per client, names intact (Thai, '&', spaces)
    z = client.get(s["zip_url"])
    assert z.status_code == 200 and z.headers["content-type"] == "application/zip"
    assert "2026-10-06" in z.headers["content-disposition"]
    with zipfile.ZipFile(io.BytesIO(z.content)) as zf:
        names = sorted(zf.namelist())
        assert names == sorted(c["filename"] for c in s["clients"]), names
        assert all(zf.read(n).startswith(b"%PDF") for n in names)

    # every per-client link works, including Thai / special-character names
    for c in s["clients"]:
        f = client.get(c["download_url"])
        assert f.status_code == 200 and f.content.startswith(b"%PDF"), c["filename"]
        with zipfile.ZipFile(io.BytesIO(z.content)) as zf:
            assert f.content == zf.read(c["filename"]), "single download differs from zip copy"
    print("PASS: 4 clients, warnings returned, zip and per-client downloads all serve valid PDFs")


def test_bad_requests():
    print("\n=== bad_requests ===")
    build_tms_report_xlsx(WORKBOOK, CFG)
    sid = upload()
    assert generate(sid, "no_such_report").status_code == 404
    assert generate(str(uuid.uuid4())).status_code == 400                    # valid shape, never uploaded
    assert generate("../../etc/passwd").status_code == 400                   # not a uuid -> never reaches a path
    # Traversal-style report ids must never reach a file path. An encoded slash becomes a real
    # path separator, so it matches no report route (404/405); a dots-only id is rejected by name.
    assert client.post("/api/reports/..%2Fx/generate", data={"session_id": sid}).status_code in (400, 404, 405)
    r = client.post("/api/reports/%2e%2e/generate", data={"session_id": sid})
    assert r.status_code == 400 and "Invalid report id" in r.json()["detail"], r.text

    # a workbook missing a sheet -> clear, user-facing 422 (not a 500)
    build_tms_report_xlsx(WORKBOOK, CFG, drop_sheet="failed")
    r = generate(upload())
    assert r.status_code == 422 and "จัดส่งไม่สำเร็จ" in r.json()["detail"], r.text
    # ...and a failed run leaves no folder behind
    leftovers = [p for p in main.OUTPUT_DIR.iterdir() if p.is_dir() and p.name.startswith("report-")]
    n_before = len(leftovers)
    generate(upload())
    assert len([p for p in main.OUTPUT_DIR.iterdir() if p.is_dir() and p.name.startswith("report-")]) == n_before

    # download endpoints only serve what a real run produced
    build_tms_report_xlsx(WORKBOOK, CFG)
    s = generate(upload()).json()
    rid = s["run_id"]
    assert client.get("/api/reports/runs/not-a-uuid/zip").status_code == 400
    assert client.get(f"/api/reports/runs/{uuid.uuid4()}/zip").status_code == 404
    assert client.get(f"/api/reports/runs/{rid}/files/..%2F..%2Fmain.py").status_code in (400, 404)
    assert client.get(f"/api/reports/runs/{rid}/files/{quote('nope.pdf')}").status_code == 404
    assert client.get(f"/api/reports/runs/{rid}/files/notes.txt").status_code == 400
    print("PASS: unknown report/session, malformed ids, missing sheet, path traversal all rejected cleanly")


def test_empty_day():
    print("\n=== empty_day ===")
    build_tms_report_xlsx(WORKBOOK, CFG, rows=False)
    r = generate(upload())
    assert r.status_code == 200
    s = r.json()
    assert s["clients"] == [] and s["zip_url"] is None and any("No orders" in w for w in s["warnings"])
    print("PASS: a day with no orders returns a clear warning and nothing to download")


def test_cleanup_removes_old_report_folders():
    print("\n=== cleanup ===")
    old, fresh, other = (main.OUTPUT_DIR / n for n in (f"report-{uuid.uuid4()}", f"report-{uuid.uuid4()}", "not-a-report"))
    for d in (old, fresh, other):
        d.mkdir()
        (d / "x.pdf").write_bytes(b"%PDF")
    stale = time.time() - (main.RETENTION_HOURS + 1) * 3600
    os.utime(old, (stale, stale))
    os.utime(other, (stale, stale))
    main._cleanup_old_files()
    assert not old.exists(), "stale report folder was not cleaned up"
    assert fresh.exists(), "fresh report folder was wrongly deleted"
    assert other.exists(), "cleanup must only touch report-* folders"
    print("PASS: stale report folders swept; fresh ones and unrelated folders left alone")


if __name__ == "__main__":
    test_list_reports()
    test_generate_and_download()
    test_bad_requests()
    test_empty_day()
    test_cleanup_removes_old_report_folders()
    print("\nAll API cases passed.")
