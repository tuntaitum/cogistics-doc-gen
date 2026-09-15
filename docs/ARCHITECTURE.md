# CoDocuments — Architecture & Maintenance Guide

A practical map of how this project fits together: what each piece does, what's safe to change, and what to watch out for.

---

## The one idea that explains everything

**Document types are data, not code.**

There is exactly one PDF-generating engine. It has no idea what a "quotation sheet" or a "catalog" is. It just receives a config object describing what to render — which columns, what they're labelled, where the values come from, what goes at the bottom — and renders that.

A "document type" is therefore just a JSON file in `backend/presets/`. Adding a new document type = adding a JSON file. It does **not** require touching Python.

If you ever find yourself about to write `if document_type == "quotation"` in the engine, stop — that's the exact design this project was built to escape (the old desktop app worked that way, which is why every request meant a code edit and a re-distribution).

---

## Request flow, end to end

```
Browser (frontend/)                    Server (backend/)
─────────────────────                  ─────────────────
1. Upload .xlsx          ──────────▶   /api/upload
                                         saves file under a session_id
                         ◀──────────     returns headers + 3 preview rows

2. Pick document type    ──────────▶   /api/presets        (list)
                         ──────────▶   /api/presets/{id}   (full config)
                         ──────────▶   /api/presets/{id}/example  (sample PDF)

3. Map columns,          ──────────▶   /api/items     (rows as JSON, for manual-entry table)
   customize, preview    ──────────▶   /api/preview   (real PDF, rendered inline)

4. Generate + download   ──────────▶   /api/generate
                         ──────────▶   /api/download/{session_id}?filename=...
```

The browser holds all the working state (which columns map where, custom labels, column order, manually-typed values). The server is stateless apart from the uploaded file on disk, keyed by `session_id`. Every preview and the final generate send the **whole config** up fresh.

**Why preview and generate produce identical output:** they call the exact same `engine.generate_pdf()`. The preview is not an approximation or an HTML mock-up — it's the real PDF. Keep it that way; the moment you render previews through a different path, they'll drift out of sync with reality.

---

## File map

### `backend/schemas.py` (~120 lines) — the contract
Pydantic models defining what a valid config looks like: `DocumentConfig`, `ColumnConfig`, `FooterBlock`, `BrandConfig`.

This is the source of truth for what a preset JSON is allowed to contain. Every default lives here.

**Touch when:** you want a new capability that presets can express (a new column behaviour, a new footer block type).
**Careful:** adding a field with a sensible default is safe. Renaming or removing a field breaks every existing preset JSON that uses it.

### `backend/engine.py` (~660 lines) — the actual work
Excel in, PDF out. Roughly in pipeline order:

| Function | Does |
|---|---|
| `detect_headers()` / `preview_excel()` | Read the header row + first rows (powers the upload step) |
| `extract_images_by_row()` | Pull floating/embedded photos out of the sheet and match them to rows |
| `read_excel()` | Read the selected rows into a list of dicts, one per item |
| `apply_manual_values()` | Merge in values the user typed in the UI (manual columns) |
| `apply_computed_values()` | Calculate derived columns (e.g. Subtotal = price × qty) |
| `make_branded_canvas()` | Per-page furniture: banners, accent bar, footer, signature block |
| `build_table()` | The main table, including the optional totals row |
| `build_footer_blocks()` | Free-text notes/terms below the table |
| `generate_pdf()` | Assembles all of the above into the file |

**The order matters:** manual values are applied *before* computed values, because a Subtotal often depends on a Quantity the user just typed in. If you add another transformation step, think about where it belongs in that chain.

**Touch when:** you need a genuinely new rendering capability.
**Careful:** this file is shared by every document type. A change here affects all of them at once. It is also the least forgiving code in the project — the image-anchor handling in particular encodes a lot of hard-won trial and error against messy Google Sheets/Lark exports.

### `backend/main.py` (~400 lines) — the API layer
Thin. It validates input, calls the engine, and returns files or JSON. Almost no business logic lives here by design.

Also contains:
- **The cleanup sweep** — deletes uploads/outputs older than `RETENTION_HOURS` (6). Runs on startup and every 30 min.
- **`DATA_DIR`** — your Railway volume hook. Uploads and generated PDFs go here; falls back to `backend/` locally.
- **The static mount at the bottom** — serves `frontend/` as the website. It's a catch-all, so it **must stay last in the file**; any route registered after it will never be reached.

### `backend/presets/*.json` — the document types
`client_catalog.json`, `bu_catalog.json`, `quotation_sheet.json`.

**This is where you should be spending most of your time.** Editing these is safe, reversible, requires no Python, and takes effect on restart. See "Common tasks" below.

### `backend/presets/examples/*.pdf` — sample outputs
Pre-generated PDFs shown on the "choose document type" page. **These are static files — they do not auto-update when you edit a preset.** If you change a preset's layout meaningfully, regenerate the matching example or the preview will be a lie. (Script for this is in the "Common tasks" section.)

### `backend/fonts/` — Thai + Latin font
Noto Sans Thai, Regular and Bold. Needed because the default PDF fonts render Thai as solid black boxes with no error. Don't delete these. See `fonts/README.md` for provenance and how the static weights were generated.

### `backend/test_engine.py` (~255 lines) — the safety net
Run `python test_engine.py` from `backend/`. Generates real PDFs and inspects them. Covers: both catalog types, the quotation, plus four regressions (truncated sheet metadata, signature-on-last-page-only, Thai text rendering, subtotal/total maths).

**Run this before every push.** It has caught real breakage several times, including breakage caused by preset edits.

### `frontend/` — plain HTML/CSS/JS, no build step
- `index.html` — three "pages" (really three divs that show/hide)
- `style.css` — CSS variables at the top define the whole palette
- `app.js` (~910 lines) — sectioned with `// ---- Page N: ... ----` comments

Key state variables live near the top of `app.js` and are worth understanding before editing:

```js
currentSession   // uploaded file + its headers
selectedConfig   // the chosen preset, as loaded from the server
columnMapping    // { columnKey: excelHeaderName }
customization    // { columnKey: { label, widthPreset } }
manualColumns    // user-added columns with no spreadsheet source
columnOrder      // output order, driven by the up/down buttons
manualColumnData // { columnKey: [value per row] }
```

`buildFinalConfig()` is the function that merges all of that into the config sent to the server. **If you add a new customization option, it has to be merged in there or it will silently have no effect.**

---

## Common tasks

### Change a column label, width, or default mapping
Edit the preset JSON. Restart the server. Done.

- `label` — the header text in the PDF
- `source_header` — which Excel column it pulls from by default (users can override in the UI)
- `optional: true` — column disappears entirely if every row is blank; also means it's not required to be mapped
- `emphasis: true` — bold, for the "headline" column of a row
- `width_mode: "fixed"` + `width_mm`, or `"flex"` + `flex_weight` (flex splits leftover space proportionally)

### Add a whole new document type
Copy an existing preset JSON, change `id` (must match the filename), `name`, `document_title`, and the columns. Restart. It appears in the picker automatically.

Then generate an example PDF for it, or the preview page will 404 for that type (handled gracefully, but users see nothing).

### Regenerate an example PDF after changing a preset
From `backend/`, with the venv active:

```python
import json
from schemas import DocumentConfig
import engine, openpyxl

# build a small sample .xlsx whose headers match the preset's source_headers
# then:
config = DocumentConfig(**json.load(open("presets/YOUR_PRESET.json")))
items = engine.read_excel("sample.xlsx", config)
engine.apply_computed_values(items, config)   # only if it has computed columns
engine.generate_pdf(items, config, "presets/examples/YOUR_PRESET.pdf", "assets")
```

### Change brand colours or banners
The `brand` block in each preset JSON. Note it's duplicated per-preset — change one, you probably want to change all.

### Add terms/conditions to a document by default
Add a `text` footer block to the preset's `footer_blocks`. Users can also type notes ad-hoc in the UI, which overrides/merges at generation time.

---

## Landmines

**1. `API_BASE` is hardcoded to localhost — this likely breaks your Railway deployment.**

`frontend/app.js` line 1:
```js
const API_BASE = window.CODOCS_API_BASE || "http://127.0.0.1:8000";
```
Nothing anywhere sets `window.CODOCS_API_BASE`. Since the backend now serves the frontend from the same origin, a deployed browser will try to call `http://127.0.0.1:8000` — which resolves to *the visitor's own laptop*, not your server. An HTTPS page calling HTTP will also be blocked outright by the browser.

The fix is one line — default to same-origin relative URLs (`|| ""`) — but note that afterwards you must run locally the merged way (just `uvicorn`, open `:8000`), not the old two-terminal setup with a separate `http.server` on `:8080`.

**2. Two identical `requirements.txt` files.** One at the repo root (for Railway/Nixpacks), one in `backend/`. They're currently identical. If you add a dependency to only one, you'll get a works-locally-fails-in-production bug. Update both, or delete one and make sure the deploy still detects Python.

**3. Example PDFs don't auto-regenerate.** Covered above, but it's the easiest thing to forget.

**4. The static mount must stay last in `main.py`.** Any `@app.get` added below it is dead code.

**5. Preset `id` must match its filename.** `quotation_sheet.json` must contain `"id": "quotation_sheet"`. The save endpoint enforces this; hand-editing doesn't.

**6. There are three kinds of column, and they behave differently.** A column's `source` field decides where its value comes from:

| `source` | Value comes from | Shows in mapping UI? |
|---|---|---|
| `"excel"` (default) | A mapped spreadsheet column | Yes — a dropdown |
| `"manual"` | Typed in per-row by the user in the browser | No — an info row; inputs appear under "Custom columns" |
| `"computed"` | Calculated from other columns | No — an info row |

The quotation's **Subtotal is `manual`**, deliberately: real Quantity values are things like `"2 tons"` or `"500 kg"`, so multiplying price × qty didn't work in practice. Someone types the subtotal in.

The `computed` type still works and is still tested, but no shipped preset uses it. If you ever reach for it, remember it needs genuinely numeric sources — it fails soft (blank cell, no error) on anything it can't parse, including price ranges like `"180-220"`.

The **Total row** (`totals_column` in the preset) sums whatever's in the named column, numeric-parsing each cell and skipping what it can't read. So it works over manually-typed subtotals, and quietly ignores a row where someone typed prose. If every cell in that column is blank, the column hides itself (it's `optional`) and the Total row disappears with it.

**7. Don't commit anything to `backend/output/` or `backend/uploads/`.** They're gitignored working directories. In production they live on the Railway volume instead.

**8. Manual column values are positional.** `manualColumnData` is an array aligned to row order. It stays correct because row selection doesn't depend on column mapping — but if you ever change *which rows* get selected after the user has typed values in, the alignment breaks.

---

## Safe-to-edit ranking

| Confidence needed | What |
|---|---|
| **Low — go ahead** | Preset JSONs, brand colours, labels, widths, `optional` flags, footer text, README |
| **Medium — run the tests after** | `frontend/style.css`, adding fields to `schemas.py`, new endpoints in `main.py` |
| **High — understand it first** | `engine.py` (esp. `extract_images_by_row`, `build_table`, `make_branded_canvas`), `buildFinalConfig()` in `app.js`, the state variables at the top of `app.js` |
| **Leave alone unless deliberate** | `backend/fonts/`, the static mount at the bottom of `main.py`, `.gitignore` |

---

## Running it

```bash
# one service, serves both API and frontend
cd backend
source ../.venv/bin/activate
uvicorn main:app --reload
# open http://127.0.0.1:8000
```

Tests:
```bash
cd backend && python test_engine.py
```

Interactive API docs (FastAPI generates these automatically, useful for poking endpoints by hand):
`http://127.0.0.1:8000/docs`

---

## If you get lost

- `docs/PROJECT.md` — the original why: background, problem, roadmap.
- `docs/legacy-*.py` — the old desktop app, kept for reference. The image-extraction logic there is the ancestor of `engine.extract_images_by_row()`.
- Git log — commit messages on this project are unusually detailed, including *why* a fix was made and what was verified. `git log --oneline` then `git show <hash>` is often faster than re-deriving the reasoning.
