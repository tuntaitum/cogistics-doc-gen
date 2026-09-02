# cogistics-doc-gen (CoDocuments)

Web-based successor to [`product-cat-gen`](https://github.com/tuntaitum/product-cat-gen) — generates branded PDF documents (product catalogs, BU sector catalogs, quotation sheets) from Excel input, driven by configurable templates instead of hardcoded per-document code.

See [`docs/PROJECT.md`](docs/PROJECT.md) for the full background, problem statement, and roadmap.

## Status

Steps 1–2 of the roadmap are done and verified:
- [x] Template config schema (`backend/schemas.py`)
- [x] Generalized, config-driven engine (`backend/engine.py`)
- [ ] Backend API (step 3 — in progress)
- [ ] Upload + header-preview UI (step 4)
- [ ] Preset picker + column-mapping UI (step 5)
- [ ] Preset management UI (step 6)
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

## How templates work

Every document type — client catalog, a BU sector catalog, a quotation sheet — is a JSON file in `backend/presets/` conforming to `DocumentConfig` in `schemas.py`. New document types are added by creating a new preset, not by writing new code. See `backend/presets/quotation_sheet.json` for an example with a different column set and a signature footer block.
