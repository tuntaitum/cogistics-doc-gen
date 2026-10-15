# Adding your own tool to CoDocuments

This guide is for anyone who wants to put a new feature (a form, a report, a document generator) on the CoDocuments web app. It explains how the app is organised, the steps to add a tool, and the rules that keep your tool from affecting everyone else's.

You do not need to understand the whole app. Your tool gets its own page and, if it needs one, its own backend file. The rest of this guide is about staying inside that space.

## How the app is organised

```
Landing page (index.html)
  └─ a SERVICE (Food Pipeline, Veggie Solution)            ← listed in services.json
       └─ a menu of its TOOLS (service.html?id=...)        ← listed in services.json
            └─ a TOOL = one page + (optionally) one backend file     ← this is you
```

| Where | What it is | Yours to change? |
|---|---|---|
| `frontend/services.json` | The registry: which services and tools exist, and where they link. | **Add your entry.** |
| `frontend/<your-tool>.html` | Your page. | **Yours.** |
| `backend/<your_tool>_routes.py` | Your server code. | **Yours.** |
| `backend/main.py` | Starts the app. | **Add 2 lines** (see step 3). |
| `frontend/style.css` | The shared look. | Add at the bottom only (see "Style"). |
| `engine.py`, `schemas.py`, `grouped_report.py`, `report_routes.py`, `memory.py`, `services.js`, `service.html`, `index.html`, `app.js`, `supply-knowledge.html`, `presets/`, `reports/` | Shared by other tools. | **Ask first.** |

Changing a shared file can silently change every other document the app produces.

## Adding a tool in five steps

### 1. Create your page

Copy `docs/templates/tool-template.html` to `frontend/<your-tool>.html` (lower case, hyphens: `non-gmo.html`). Change the three marked places: the title, the Back link, and the API path.

### 2. Add it to the menu

Open `frontend/services.json`, find the service your tool belongs to, and add an entry to its `tools` list:

```json
{
  "name": "NON-GMO Letter Form",
  "description": "Fill in a form to create a NON-GMO guarantee letter.",
  "icon": "letter",
  "href": "non-gmo.html"
}
```

`icon` is one of `truck`, `leaf`, `sheet`, `letter`, `report`. If it is missing or unknown, `report` is used. The menu builds itself from this file, so there is no HTML to edit.

Your Back link should point to `service.html?id=<service id>`, for example `service.html?id=veggie-solution`.

### 3. Add your server code (only if your page needs a server)

Copy `docs/templates/tool_routes_template.py` to `backend/<your_tool>_routes.py` and replace the example with your own logic.

Then wire it into `backend/main.py` with two lines. **`main.py` already imports a function called `make_router` for reports, so give yours a different name:**

```python
from your_tool_routes import make_router as make_your_tool_router    # near the other imports

app.include_router(make_your_tool_router())    # next to the existing app.include_router(...) line
```

The `include_router` line **must stay above** `app.mount("/", ...)` at the bottom of the file. That line serves the web pages and would otherwise swallow your routes.

Use `/api/<your-tool>/...` for your paths so they never clash with anyone else's.

### 4. Run it and click through

```bash
cd backend
pip install -r requirements.txt
uvicorn main:app --reload
```

Open <http://localhost:8000>, go landing → your service → your tool, and use it. Then do it again with the browser window narrowed to phone width.

### 5. Test, then ask for a review

```bash
cd backend
python test_engine.py                      # needs: pip install -r requirements.txt
pip install -r requirements-dev.txt        # one-time, for the API tests
python test_api.py
```

Both must end with "All cases passed" / "All API cases passed". The tests use `pdftotext` (from the poppler package) to read PDFs. Add a `test_<your_tool>.py` that checks your tool does its job and rejects bad input. Then use the checklist at the end of this guide.

## Style: making your page look like the rest

Use `style.css` and nothing else. **Don't add a `<style>` block or inline colours**; they drift from the rest of the app and nobody notices until the palette changes.

**Page skeleton:** header with tagline → `main.wrap` → Back button (`btn btn--ghost back-btn`) → `section.panel` containing `h1.panel__title` and `p.panel__hint`. For a wide table use `main.wrap wrap--wide` and `<body class="page--wide">` so the header lines up. The template already has all of this.

| You need | Use |
|---|---|
| A form field | `.field-row` → `label.field-row__label` (+ `.field-row__hint`) + `input.field-row__input`. For **text**, add `field-row__input--wide`; the plain class is a 64px number box. |
| The main button | `.btn .btn--primary` (one per panel) |
| A secondary button | `.btn .btn--ghost` |
| A button that is a link (downloads) | `<a class="btn btn--primary">` |
| Loading / error / success text | `.status` plus `is-loading`, `is-error` or `is-success` |
| Result with a download | `.result` + `.result__text` + a primary button |
| PDF preview | `.preview-frame-wrap` + `.preview-frame-header` + `iframe.preview-frame` |
| Something to review before sending | `.warning-box` |
| A table | `.report-table` inside `.report-table-scroll` (scrolls sideways on phones) |
| Upload a file | `.dropzone`, as in `food-pipeline.html` |

**Colours** are variables, never hex codes: `var(--navy)`, `var(--navy-dark)`, `var(--border)`, `var(--text-dark)`, `var(--text-mid)`, `var(--danger)`, `var(--danger-bg)`, `var(--success)`, `var(--success-bg)`, `var(--warn)`, `var(--warn-bg)`, `var(--radius)`.

**If you really need a new style**, add it at the **bottom** of `style.css` under a comment naming your tool, and prefix every class with your tool's name (`.nongmo-...`). Never edit an existing rule; other pages use them.

**Wording:** English, sentence case ("Create PDF", not "CREATE PDF"). Say what happened and what to do next; for example, "The issue date can't be in the future", not "Invalid input".

**Anything that came from a user or a spreadsheet goes on the page with `textContent`, never `innerHTML`.** A client called `Gamma & Sons <Trading>` has to appear exactly like that, and otherwise one odd name can break your page or inject markup.

## Server rules

**Keep your code in your own file.** Don't add routes, models or helper functions to `main.py`. That file is shared and already long.

**Validate on the server.** The browser's `maxlength` and `required` are conveniences. Define a pydantic model with `Field(min_length=..., max_length=...)`, as in the template, and give `HTTPException` messages in plain language, because they are shown to people as written.

**Prefer stateless.** If your tool can build the result in memory and return it, do that. Nothing is stored, so there is nothing to clean up or to cost money. The NON-GMO letter works this way.

**Files.** Anything you save lives on a Railway volume that is **billed per GB and never cleared by a redeploy**. So:

- Write only under `OUTPUT_DIR` or `UPLOADS_DIR` (both are in `main.py`). Single files there are deleted after 6 hours.
- If you create a **folder**, its name must start with `report-` or it is never deleted. Otherwise, ask for cleanup to be extended.
- Never build a path from user input. Use `grouped_report.safe_filename()` for file names and validate ids as UUIDs the way `report_routes.py` does.
- Delete input files when you have finished with them.

**Uploads.** To accept an Excel file, use the existing `POST /api/upload`. It already streams the file to disk, caps its size, checks that it is really a workbook and refuses everything else. Don't write a second upload endpoint without the same protections.

**Memory.** Railway bills memory, and Python keeps memory it has freed. Put `@releases_memory` (from `memory.py`) under your route decorator on anything that opens images, workbooks or builds large PDFs. Don't keep data in module-level variables between requests. Don't decode full-size photos just to shrink them; copy how `engine.make_thumbnail` does it.

**Don't switch `openpyxl` to `read_only=True`** to save memory. It silently drops columns from Google Sheets and Lark exports; the reason is written in `engine.detect_headers`.

**Thai text.** The standard PDF fonts can't draw Thai. Call `engine._ensure_fonts_registered()` and use `engine.FONT_REGULAR` / `engine.FONT_BOLD`.

**Logging.** Use `logging.getLogger("codocuments")`. Don't log client names, phone numbers or spreadsheet contents.

**There is no login yet.** Anything your tool produces is protected only by an unguessable link. Don't put real client data in code, tests or example files.

## Before you ask for a review

- [ ] My tool appears in the right menu and its Back link returns to that menu.
- [ ] My page has no `<style>` block, no hex colours, and no `innerHTML` with user data.
- [ ] Buttons, fields, messages and the preview use the shared classes above.
- [ ] Errors show in red with a message that tells people what to do.
- [ ] I tried it at phone width, and the page doesn't scroll sideways.
- [ ] My server code is in its own file, and my only edit to `main.py` is the import and `include_router` lines.
- [ ] The server rejects bad input (I tested it) and nothing I save is left behind.
- [ ] `python test_engine.py` and `python test_api.py` both pass, and I added tests for my tool.
- [ ] I haven't edited any "Ask first" file, or I've said which and why.

## Worked example: bringing the NON-GMO letter in line

The NON-GMO letter is a good start. It stores nothing, the server checks the date and products, and it opens as its own page. These are the changes that would make it follow the rest of this guide, in the order I'd do them.

**Page (`non-gmo.html`)**

1. **Remove the `<style>` block.** Replace `.letter-field` with `.field-row`, `.letter-preview` with `.preview-frame-wrap` + `.preview-frame` (which also replaces the hardcoded `#ddd` border), and `.letter-product` with `.field-row` rows using `field-row__input--wide`.
2. **Add the tagline** (`app-header__tagline`) to the header and `back-btn` to the Back link, so it matches every other page.
3. **One main button.** "Preview PDF" and "Download PDF" are both `btn--primary` and stack at full width. Keep Preview as the primary button and show Download in a `.result` block after a successful preview, as `food-pipeline.html` does.
4. **Make errors red.** The status line is a plain `.status`. Set `is-loading` / `is-error` / `is-success` from the script, as the template does.

**Server (`main.py`, `letters.py`)**

5. **Move the route and its request model out of `main.py`** into `letters_routes.py`, following the template. Then `main.py` only imports and includes it.
6. **Add `@releases_memory`** to the route. It's light today, but it costs nothing and protects you if the letter ever gets heavier.

**Later, if needed**

7. The signatory name (`letters.py`, "Ms.Thiaranya Maharn"), the 6-product limit and the product suggestions in `<datalist>` are written into the code. That is fine until someone has to change one. Then they should become settings, so a change doesn't need a code edit.

## Questions

Ask the project owner before changing anything in the "Ask first" list, adding a new library to `requirements.txt`, or doing anything that stores files outside `OUTPUT_DIR` / `UPLOADS_DIR`.
