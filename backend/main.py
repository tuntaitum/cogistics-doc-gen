"""
main.py — CoDocuments backend API.

Replaces launcher.py. Four jobs:
  1. Accept an uploaded .xlsx, return detected headers (for the mapping UI)
  2. Serve/save/delete presets (DocumentConfig JSON files)
  3. Generate a PDF from an uploaded file + a config (preset + user's column mapping)
  4. Let the frontend download the finished PDF

Run with:
    uvicorn main:app --reload
"""

import json
import os
import uuid
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import ValidationError

import engine
from schemas import DocumentConfig

BASE_DIR = Path(__file__).parent
PRESETS_DIR = BASE_DIR / "presets"
ASSETS_DIR = BASE_DIR / "assets"
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"

UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

app = FastAPI(title="CoDocuments API")

# Internal tool for now — open CORS. Tighten this once it's deployed
# somewhere with a known frontend origin.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ─────────────────────────────────────────────
#  UPLOAD + HEADER DETECTION
# ─────────────────────────────────────────────

@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...), header_row: int = Form(2)):
    """
    Accept an .xlsx, store it under a session id, and return the headers
    found in the given row. The session id is passed back into /generate
    so we don't need to re-upload the file after mapping columns.
    """
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "Please upload a .xlsx file")

    session_id = str(uuid.uuid4())
    dest = UPLOADS_DIR / f"{session_id}.xlsx"
    with open(dest, "wb") as f:
        f.write(await file.read())

    try:
        headers = engine.detect_headers(str(dest), header_row=header_row)
    except Exception as e:
        dest.unlink(missing_ok=True)
        raise HTTPException(400, f"Could not read file: {e}")

    if not headers:
        dest.unlink(missing_ok=True)
        raise HTTPException(400, f"No headers found in row {header_row}. Try a different header row.")

    return {"session_id": session_id, "filename": file.filename, "headers": headers}


# ─────────────────────────────────────────────
#  PRESETS
# ─────────────────────────────────────────────

def _preset_path(preset_id: str) -> Path:
    safe_id = "".join(c for c in preset_id if c.isalnum() or c in ("-", "_"))
    if safe_id != preset_id:
        raise HTTPException(400, "Invalid preset id")
    return PRESETS_DIR / f"{safe_id}.json"


@app.get("/api/presets")
def list_presets():
    """Summary list for the preset picker — id, name, title, column labels."""
    results = []
    for path in sorted(PRESETS_DIR.glob("*.json")):
        with open(path) as f:
            data = json.load(f)
        results.append({
            "id": data.get("id", path.stem),
            "name": data.get("name", path.stem),
            "document_title": data.get("document_title", ""),
            "column_labels": [c["label"] for c in data.get("columns", [])],
        })
    return results


@app.get("/api/presets/{preset_id}")
def get_preset(preset_id: str):
    path = _preset_path(preset_id)
    if not path.exists():
        raise HTTPException(404, "Preset not found")
    with open(path) as f:
        return json.load(f)


@app.put("/api/presets/{preset_id}")
def save_preset(preset_id: str, config: DocumentConfig):
    if config.id != preset_id:
        raise HTTPException(400, "Body 'id' must match the URL preset id")
    path = _preset_path(preset_id)
    with open(path, "w") as f:
        json.dump(config.model_dump(), f, indent=2)
    return {"status": "saved", "id": preset_id}


@app.delete("/api/presets/{preset_id}")
def delete_preset(preset_id: str):
    path = _preset_path(preset_id)
    if not path.exists():
        raise HTTPException(404, "Preset not found")
    path.unlink()
    return {"status": "deleted", "id": preset_id}


# ─────────────────────────────────────────────
#  GENERATE
# ─────────────────────────────────────────────

@app.post("/api/generate")
def generate(session_id: str = Form(...), config_json: str = Form(...)):
    """
    session_id: from /api/upload
    config_json: the full DocumentConfig (preset + user's column mapping),
                 serialized as a JSON string
    """
    xlsx_path = UPLOADS_DIR / f"{session_id}.xlsx"
    if not xlsx_path.exists():
        raise HTTPException(400, "Unknown session_id — please re-upload the file")

    try:
        config = DocumentConfig(**json.loads(config_json))
    except (ValidationError, json.JSONDecodeError) as e:
        raise HTTPException(400, f"Invalid config: {e}")

    try:
        items = engine.read_excel(str(xlsx_path), config)
    except ValueError as e:
        raise HTTPException(400, str(e))

    if not items:
        raise HTTPException(400, "No rows matched — check the select column and value")

    out_path = OUTPUT_DIR / f"{session_id}.pdf"
    engine.generate_pdf(items, config, str(out_path), str(ASSETS_DIR))

    return {"status": "generated", "download_url": f"/api/download/{session_id}", "item_count": len(items)}


@app.get("/api/download/{session_id}")
def download(session_id: str):
    path = OUTPUT_DIR / f"{session_id}.pdf"
    if not path.exists():
        raise HTTPException(404, "PDF not found — generate it first")
    return FileResponse(path, media_type="application/pdf", filename="document.pdf")


@app.get("/api/health")
def health():
    return {"status": "ok"}
