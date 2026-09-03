# cogistics-doc-gen (CoDocuments)

Web-based successor to [`product-cat-gen`](https://github.com/tuntaitum/product-cat-gen) — generates branded PDF documents (product catalogs, BU sector catalogs, quotation sheets) from Excel input, driven by configurable templates instead of hardcoded per-document code.

See [`docs/PROJECT.md`](docs/PROJECT.md) for the full background, problem statement, and roadmap.

## Status

Steps 1–5 of the roadmap are done, plus early header/width customization with a live preview:
- [x] Template config schema (`backend/schemas.py`)
- [x] Generalized, config-driven engine (`backend/engine.py`)
- [x] Backend API (`backend/main.py`)
- [x] Upload + header-preview UI (`frontend/`) — validated against a real Cogistics export
- [x] Preset picker + column-mapping UI (`frontend/`) — full upload-to-download flow works end to end
- [x] Header/column-width customization with live preview (one-off per generation — saving back to the template is still step 6)
- [ ] Preset management UI (step 6 — next)
- [ ] Validation against real sheets (step 7)
- [ ] Deploy (step 8)

## Repo structure

```
backend/
  schemas.py        # DocumentConfig — the schema every template follows
  engine.py          # Excel reading, image extraction, PDF generation (config-driven)
  presets/            # Example templates: client catalog, BU catalog, quotation sheet
  assets/              # Shared brand assets (banners)
  test_engine.py     # End-to-end test proving the engine against all three preset types
  requirements.txt
docs/
  PROJECT.md                            # Full project doc (background, problem, plan)
  legacy-*.py / legacy-*.spec           # Original product-cat-gen desktop app, kept for reference
```

## Running the engine tests

```bash
cd backend
pip install -r requirements.txt
python test_engine.py
```

This builds dummy spreadsheets for the client catalog and quotation sheet presets, runs them through the full engine, and writes the resulting PDFs to `backend/output/`.

## Running the API

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```

Endpoints (all under `/api`):

| Method | Path | Purpose |
|---|---|---|
| POST | `/upload` | Upload an `.xlsx`, get back a `session_id` + detected column headers |
| GET | `/presets` | List available templates |
| GET | `/presets/{id}` | Fetch a full template config |
| PUT | `/presets/{id}` | Save/update a template config |
| DELETE | `/presets/{id}` | Delete a template |
| POST | `/generate` | Generate a PDF from a `session_id` + config JSON |
| GET | `/download/{session_id}` | Download the generated PDF |

Interactive docs (Swagger UI) are auto-served at `http://127.0.0.1:8000/docs` once the server is running — useful for testing endpoints by hand before the frontend exists.

## Running the frontend

The frontend is plain HTML/CSS/JS — no build step. With the backend running (see above):

```bash
cd frontend
python3 -m http.server 8080
```

Then open `http://127.0.0.1:8080`. It talks to the backend at `http://127.0.0.1:8000` by default; override by setting `window.CODOCS_API_BASE` before `app.js` loads if you're running the backend elsewhere.

Currently implemented: upload an `.xlsx`, detect its headers and preview the first few rows, and retry with a different header row if it looks wrong. The "Continue" button is intentionally disabled — document type selection (step 5) isn't built yet.

## How templates work

Every document type — client catalog, a BU sector catalog, a quotation sheet — is a JSON file in `backend/presets/` conforming to `DocumentConfig` in `schemas.py`. New document types are added by creating a new preset, not by writing new code. See `backend/presets/quotation_sheet.json` for an example with a different column set and a signature footer block.
