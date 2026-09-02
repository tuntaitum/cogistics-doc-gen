# CoDocuments — Project Documentation

**Owner:** Tai
**Status:** Planning / pre-development
**Last updated:** September 2026

---

## Executive summary

Cogistics currently generates branded PDF documents (product catalogs, quotation sheets) from Excel data using a desktop app I built and distribute internally. As teams have started requesting more variations — different document types, different columns, different business units — every request has meant manually editing the code and redistributing the app. This doesn't scale.

**CoDocuments** is the planned revision: turning the desktop app into a web app where the document engine is driven by configurable templates instead of hardcoded code. Teammates will be able to upload an Excel file, pick or define a document template, map their spreadsheet's columns to the output, and generate a PDF — without needing me to touch the code for each new request. Updates are also instant for everyone, since there's nothing to redistribute.

This doc covers the current state, the problem, the proposed solution, and the implementation plan.

---

## Background — the original project

The original tool, `product-cat-gen`, is a Python desktop app built for Cogistics (a cold-chain logistics company) to generate a professional PDF product catalog from an Excel input file.

**How it works today:**
- A Tkinter GUI (`launcher.py`) lets a user browse for or drop an `.xlsx` file.
- The engine (`supplyknowledge2productcat.py`) reads the Excel file with `openpyxl`, including pulling out floating/embedded product images and matching them to the correct row.
- Only rows marked `"Yes"` in a `Select` column are included.
- A branded PDF is assembled with `reportlab` — Cogistics header banners, a styled product table (photo, name, dimensions, price, supply advantage), page numbers, and a footer — and saved locally.
- The app is packaged with PyInstaller into a standalone executable and shared with the internal team.

This worked well for its original single purpose: a client-facing product suggestion catalog with a fixed set of columns.

---

## The problem

Over time, requests started coming in to customize the tool for different needs:

- **Different business units** (e.g. bakery, sauce) wanting a catalog for their sector instead of by client — different column headers, different framing.
- **Quotation sheets** — a genuinely different document type, needing columns like price, quantity, and remarks, plus a signature area at the bottom that the catalog format never had.

Each of these has meant manually editing the hardcoded Python (column names, table structure, branding) and rebuilding/redistributing the desktop app. This is slow, doesn't scale as more teams ask for variations, and puts me in the critical path for every small change.

---

## Proposed solution

Move from *one hardcoded app per document type* to *one configurable engine, many templates*.

**Key changes:**

1. **Web app instead of desktop app** — hosted centrally, so future updates apply instantly to everyone with no redistribution step.
2. **Template/preset system** — each document type (client catalog, bakery BU catalog, quotation sheet, etc.) becomes a saved configuration rather than a code branch. A config defines: document title, which columns appear and how they're labeled, optional blocks like a signature area, and branding.
3. **Column-mapping UI** — since every team's Excel file has slightly different headers, users map "this output column" to "this column in my spreadsheet" through the UI at generation time, rather than me hardcoding exact header names.

Most of the existing engine (Excel reading, image extraction, PDF generation via `reportlab`) is reusable as-is — it's pure Python with no dependency on the desktop GUI. The Tkinter interface is the only part being fully replaced; the rest is being generalized to read from a config object instead of hardcoded constants.

### How a document gets generated

1. User uploads an `.xlsx` file.
2. Backend detects the header row and reads the available column names.
3. User picks an existing preset (or starts a new one) and maps output fields to their spreadsheet's columns.
4. The mapping + preset combine into a single config object.
5. The generalized engine (same `read_excel` / `generate_pdf` logic as today) produces the PDF.
6. User downloads the finished PDF.

---

## Implementation roadmap

| # | Step | Outcome |
|---|------|---------|
| 1 | Design the template config schema | A JSON structure that can represent any document type (columns, header row, select column, footer blocks, branding) with no code differences between types |
| 2 | Generalize the existing engine | `read_excel`, `extract_images_by_row`, `build_product_table`, `generate_pdf` all take a config object instead of hardcoded `PATHS`/`BRAND` |
| 3 | Stand up the backend API | Endpoints for upload, header detection, preset CRUD, and PDF generation — replaces `launcher.py` |
| 4 | Build upload + header-preview step | Drag-and-drop/browse upload, file validation, header row detection/confirmation |
| 5 | Build preset picker + column-mapping UI | Pick a template, map spreadsheet columns to output fields, save as a new/updated preset |
| 6 | Build preset management | List, edit, duplicate, delete presets — new document types created without touching code |
| 7 | Validate against real sheets | Run actual client, BU, and quotation files through CoDocuments and compare output against the current app |
| 8 | Deploy and cut over | Go live, point the team at the URL, retire the desktop app |

**Current status:** Planning complete for the architecture and schema approach. Steps 1–2 (config schema + engine refactor) are next, and can start before hosting/infrastructure decisions are finalized since they don't depend on where the app eventually runs.

**Open decisions:**
- Hosting: internal server vs. a simple cloud host — not yet finalized.
- Preset storage: simple JSON files vs. a small database — leaning toward starting simple and revisiting if the number of presets/teams grows significantly.

---

## Why this matters (for stakeholders)

- **Removes the bottleneck**: teams can self-serve new document variations without waiting on engineering time from me.
- **No more redistribution**: updates and fixes apply instantly to every user.
- **Scales to new use cases**: any future document type (new BU, new report format) is a configuration, not a development request.
- **Reuses proven logic**: the core Excel-reading and PDF-generation code — including the trickier parts like floating image extraction — already works and isn't being rewritten from scratch, just made configurable.

---

## Appendix

- **Original repo:** [github.com/tuntaitum/product-cat-gen](https://github.com/tuntaitum/product-cat-gen)
- **Original engine file:** `supplyknowledge2productcat.py`
- **Original launcher:** `launcher.py` (Tkinter GUI, to be retired)
- **Core libraries in use:** `openpyxl` (Excel + image reading), `Pillow` (thumbnails), `reportlab` (PDF generation)
