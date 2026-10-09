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
    names = (f"report-{uuid.uuid4()}", f"report-{uuid.uuid4()}", f"report-{uuid.uuid4()}", "not-a-report")
    old, middle, fresh, other = (main.OUTPUT_DIR / n for n in names)
    for d in (old, middle, fresh, other):
        d.mkdir()
        (d / "x.pdf").write_bytes(b"%PDF")
    plain = main.OUTPUT_DIR / "plain.pdf"                       # an ordinary (veggie) output file
    plain.write_bytes(b"%PDF")

    def age(path, hours):
        t = time.time() - hours * 3600
        os.utime(path, (t, t))

    age(old, main.RETENTION_HOURS + 1)
    age(other, main.RETENTION_HOURS + 1)
    age(middle, main.REPORT_RETENTION_HOURS + 0.5)              # past the report limit, inside the 6h limit
    age(plain, main.REPORT_RETENTION_HOURS + 0.5)
    main._cleanup_old_files()
    assert not old.exists(), "stale report folder was not cleaned up"
    assert not middle.exists(), "report folders must expire after REPORT_RETENTION_HOURS, not the longer 6h"
    assert plain.exists(), "ordinary output files keep the longer retention"
    assert fresh.exists(), "fresh report folder was wrongly deleted"
    assert other.exists(), "cleanup must only touch report-* folders"
    print("PASS: report folders expire on their own (shorter) clock; other files and folders untouched")


def test_uploads_are_bounded_and_tidy():
    print("\n=== uploads_bounded_and_tidy ===")
    def leftovers():
        return sorted(p.name for p in main.UPLOADS_DIR.iterdir() if p.name != ".gitkeep")

    before = leftovers()
    # rejected files must not stay on disk
    r = client.post("/api/upload", files={"file": ("x.xlsx", b"%PDF-1.7 nope")}, data={"header_row": 2})
    assert r.status_code == 400 and leftovers() == before, "rejected upload left a file behind"
    r = client.post("/api/upload", files={"file": ("old.xls", b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" + b"\0" * 64)}, data={"header_row": 2})
    assert r.status_code == 400 and "old-format" in r.json()["detail"] and leftovers() == before

    # an oversized upload is refused (413) and removed, not loaded whole into memory
    original = main.MAX_UPLOAD_MB
    main.MAX_UPLOAD_MB = 1
    try:
        r = client.post("/api/upload", files={"file": ("big.xlsx", b"PK\x03\x04" + b"0" * (2 * 1024 * 1024))}, data={"header_row": 2})
    finally:
        main.MAX_UPLOAD_MB = original
    assert r.status_code == 413 and "too large" in r.json()["detail"] and leftovers() == before
    print("PASS: rejected and oversized uploads are refused and leave nothing on disk")


def test_report_upload_deleted_after_use():
    print("\n=== report_upload_deleted_after_use ===")
    build_tms_report_xlsx(WORKBOOK, CFG)
    sid = upload()
    path = main.UPLOADS_DIR / f"{sid}.xlsx"
    assert path.exists()
    assert generate(sid).status_code == 200
    assert not path.exists(), "the uploaded workbook (every client's data) should be deleted once the reports exist"
    assert generate(sid).status_code == 400            # and can't be reused
    # but a FAILED run keeps it, so nothing is lost on an error
    build_tms_report_xlsx(WORKBOOK, CFG, drop_sheet="failed")
    sid = upload()
    assert generate(sid).status_code == 422 and (main.UPLOADS_DIR / f"{sid}.xlsx").exists()
    print("PASS: workbook deleted after a successful run, kept after a failed one")


def test_static_files_revalidate():
    print("\n=== static_files_revalidate ===")
    r = client.get("/style.css")
    assert r.status_code == 200 and r.headers["cache-control"] == "no-cache", r.headers
    etag = r.headers.get("etag")
    assert etag, "static files should carry an ETag so revalidation is cheap"
    r2 = client.get("/style.css", headers={"If-None-Match": etag})
    assert r2.status_code == 304 and r2.headers["cache-control"] == "no-cache" and not r2.content
    print("PASS: static files must be revalidated and answer 304 (no body) when unchanged")


def test_memory_helpers():
    print("\n=== memory_helpers ===")
    import inspect
    import memory
    memory.release_memory()                              # must never raise, on any platform
    # the decorator keeps FastAPI-visible parameters, and releases memory even when the endpoint raises
    assert "session_id" in inspect.signature(main.generate).parameters
    calls = []
    original = memory.release_memory
    memory.release_memory = lambda: calls.append(1)
    try:
        @memory.releases_memory
        def boom(x):
            raise ValueError("fail")
        try:
            boom(1)
        except ValueError:
            pass
    finally:
        memory.release_memory = original
    assert calls == [1], "memory was not released after a failing endpoint"
    print("PASS: release_memory is safe; decorator keeps signatures and runs on errors too")


if __name__ == "__main__":
    test_list_reports()
    test_generate_and_download()
    test_bad_requests()
    test_empty_day()
    test_cleanup_removes_old_report_folders()
    test_uploads_are_bounded_and_tidy()
    test_report_upload_deleted_after_use()
    test_static_files_revalidate()
    test_memory_helpers()
    print("\nAll API cases passed.")
