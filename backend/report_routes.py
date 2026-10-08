"""
report_routes.py — HTTP API for grouped reports (e.g. Food Pipeline's daily
delivery report: one TMS workbook in, one PDF per client out).

Kept out of main.py on purpose: main.py already carries the whole veggie flow,
and this is a self-contained feature. main.py just calls make_router(...) and
includes the result (before the static-files mount, which would otherwise
shadow it).

Flow the frontend follows:
  1. POST /api/upload                       (existing) -> session_id
  2. POST /api/reports/{report_id}/generate (form: session_id) -> summary JSON
  3. GET  the zip_url, or any client's download_url, from that summary

Each generate call writes into its own folder, OUTPUT_DIR/report-<run_id>/,
holding the PDFs and one zip. main.py's cleanup sweep deletes old report-*
folders along with everything else.
"""

import json
import logging
import re
import shutil
import uuid
import zipfile
from pathlib import Path
from typing import Optional
from urllib.parse import quote

from fastapi import APIRouter, Form, HTTPException
from fastapi.responses import FileResponse
from pydantic import ValidationError

import grouped_report
from schemas import GroupedReportConfig

logger = logging.getLogger("codocuments")

_UUID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
RUN_DIR_PREFIX = "report-"      # main.py's cleanup sweep relies on this prefix


def _require_uuid(value: str, what: str) -> str:
    """Ids end up inside file paths, so only accept exactly the shape we generate."""
    if not _UUID_RE.match(value or ""):
        raise HTTPException(400, f"Invalid {what}")
    return value


def make_router(*, reports_dir: Path, uploads_dir: Path, output_dir: Path, assets_dir: Path) -> APIRouter:
    router = APIRouter(prefix="/api/reports", tags=["reports"])

    def load_config(report_id: str) -> GroupedReportConfig:
        safe_id = "".join(c for c in report_id if c.isalnum() or c in ("-", "_"))
        if safe_id != report_id:
            raise HTTPException(400, "Invalid report id")
        path = reports_dir / f"{safe_id}.json"
        if not path.exists():
            raise HTTPException(404, "Report type not found")
        try:
            return GroupedReportConfig(**json.loads(path.read_text(encoding="utf-8")))
        except (ValidationError, json.JSONDecodeError) as e:
            raise HTTPException(500, f"Report config '{report_id}' is invalid: {e}")

    def run_dir(run_id: str) -> Path:
        _require_uuid(run_id, "run id")
        d = output_dir / f"{RUN_DIR_PREFIX}{run_id}"
        if not d.is_dir():
            raise HTTPException(404, "Reports not found. They may have expired; please generate them again.")
        return d

    # ── list ─────────────────────────────────────────────────────────────
    @router.get("")
    def list_reports():
        """Summary of each available report type, for the picker."""
        results = []
        for path in sorted(reports_dir.glob("*.json")):
            cfg = load_config(path.stem)
            results.append({
                "id": cfg.id,
                "name": cfg.name,
                "document_title": cfg.document_title,
                "group_by_header": cfg.group_by_header,
                "orientation": cfg.orientation,
                "sections": [{"key": s.key, "title": s.title, "sheet_name": s.sheet_name, "title_color": s.title_color}
                             for s in cfg.sections],
            })
        return results

    # ── generate ─────────────────────────────────────────────────────────
    @router.post("/{report_id}/generate")
    def generate(report_id: str, session_id: str = Form(...)):
        """
        Build one PDF per group (client) from a workbook already uploaded via
        /api/upload, and zip them. Returns a summary the UI can render as-is:
        the clients with their counts, any warnings to review before sending,
        and download links.
        """
        _require_uuid(session_id, "session_id")
        xlsx_path = uploads_dir / f"{session_id}.xlsx"
        if not xlsx_path.exists():
            raise HTTPException(400, "Unknown session_id — please re-upload the file")

        cfg = load_config(report_id)
        run_id = str(uuid.uuid4())
        out_dir = output_dir / f"{RUN_DIR_PREFIX}{run_id}"

        try:
            result = grouped_report.generate_grouped_reports(str(xlsx_path), cfg, str(out_dir), str(assets_dir))
        except grouped_report.ReportError as e:
            shutil.rmtree(out_dir, ignore_errors=True)
            raise HTTPException(422, str(e))        # user-facing: names the missing sheet/column
        except Exception:
            shutil.rmtree(out_dir, ignore_errors=True)
            logger.exception("Report generation failed")
            raise HTTPException(500, "Something went wrong while building the reports.")

        date_text = result.report_date.strftime("%d-%m-%Y") if result.report_date else None
        summary = {
            "run_id": run_id,
            "report_name": cfg.name,
            "report_date": date_text,
            "client_count": len(result.clients),
            "warnings": result.warnings,
            "clients": [],
            "zip_url": None,
        }
        if not result.clients:
            return summary                          # empty day: nothing to download, warnings explain why

        for c in result.clients:
            summary["clients"].append({
                "name": c.name,
                "filename": c.filename,
                "counts": c.counts,
                "total": c.total,
                "download_url": f"/api/reports/runs/{run_id}/files/{quote(c.filename)}",
            })

        zip_base = f"{cfg.document_title} - {result.report_date.isoformat()}" if result.report_date else cfg.document_title
        zip_name = grouped_report.safe_filename(zip_base, set())[:-len(".pdf")] + ".zip"
        with zipfile.ZipFile(out_dir / zip_name, "w", zipfile.ZIP_DEFLATED) as z:
            for c in result.clients:
                z.write(c.path, arcname=c.filename)
        summary["zip_url"] = f"/api/reports/runs/{run_id}/zip"
        return summary

    # ── downloads ────────────────────────────────────────────────────────
    @router.get("/runs/{run_id}/zip")
    def download_zip(run_id: str):
        zips = sorted(run_dir(run_id).glob("*.zip"))
        if not zips:
            raise HTTPException(404, "No zip for this run")
        return FileResponse(zips[0], media_type="application/zip", filename=zips[0].name)

    @router.get("/runs/{run_id}/files/{filename}")
    def download_file(run_id: str, filename: str):
        d = run_dir(run_id)
        # Only a bare PDF file name that actually sits in this run's folder.
        if Path(filename).name != filename or "\\" in filename or not filename.lower().endswith(".pdf"):
            raise HTTPException(400, "Invalid file name")
        path = d / filename
        if not path.is_file():
            raise HTTPException(404, "File not found")
        return FileResponse(path, media_type="application/pdf", filename=filename)

    return router
