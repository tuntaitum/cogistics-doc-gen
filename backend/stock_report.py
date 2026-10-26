"""
stock_report.py — the 3PL daily stock movement summary PDF.

    WMS stock ledger (stock_ledger)  +  product master data (stock_master)
        +  the client someone chose  +  the date they confirmed      ->   one PDF

What the PDF shows
  * a header (client, date, warehouse), a strip of headline numbers,
  * STOCK BY SKU: opening, in, out and closing, each as "cartons + loose pieces"
    (the closing figure also gives the exact pieces underneath), closing weight,
    and the split by warehouse when stock sits in more than one,
  * BATCH DETAIL: the same per batch, with the batch text exactly as the WMS wrote it.

Notes printed under a figure when something needs saying (in words, not symbols: the Thai font has no
dagger glyph and reportlab silently drops characters a font lacks):
    "carton size not on file"   the SKU has no carton size in the master data, so it is shown in pieces only
    "does not add up"           the WMS closing balance isn't opening + in - out; the WMS figure is shown

Refusing to build: if none of the file's SKUs belong to the chosen client the WMS file would be filed
under the wrong client's name, so generate_stock_summary() raises ClientMismatchError and writes nothing.
Everything else that is merely suspicious comes back in result.warnings for a person to review first.
"""

import os
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Optional
from xml.sax.saxutils import escape as xml_escape

from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import CondPageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import engine
import stock_ledger as sl
import stock_master as sm
from grouped_report import safe_filename
from schemas import ColumnConfig, DocumentConfig, StockReportConfig


class StockReportError(Exception):
    """The report can't be built (message is user-facing)."""


class ClientMismatchError(StockReportError):
    """The file's SKUs don't belong to the selected client. Carries the check for the UI to show."""

    def __init__(self, check: "sm.ClientCheck"):
        super().__init__(" ".join(check.messages))
        self.check = check


DEFAULT_CONFIG_PATH = Path(__file__).parent / "stock" / "stock_summary.json"
DASH = "–"
MUTED = "#6B7280"
RED = "#A32D2D"
THAI_TZ = timezone(timedelta(hours=7))


def load_stock_config(path=None) -> StockReportConfig:
    import json
    path = Path(path) if path else DEFAULT_CONFIG_PATH
    try:
        return StockReportConfig(**json.loads(path.read_text(encoding="utf-8")))
    except Exception as e:
        raise StockReportError(f"The stock report configuration ({path.name}) is invalid: {e}")


@dataclass
class StockSummaryResult:
    path: str
    filename: str
    client: str
    report_date: date
    warnings: list = field(default_factory=list)
    check: Optional["sm.ClientCheck"] = None
    sku_count: int = 0
    batch_count: int = 0
    closing_qty: float = 0.0
    in_qty: float = 0.0
    out_qty: float = 0.0
    closing_wt: float = 0.0
    all_reconciled: bool = True


# ─────────────────────────────────────────────
#  TEXT FOR THE CELLS
# ─────────────────────────────────────────────

def _d(value: date) -> str:
    return value.strftime("%d-%m-%Y")


def _fill(template: str, client: str, report_date: date) -> str:
    return template.format(client=client, date=_d(report_date), date_iso=report_date.isoformat())


def stock_filename(cfg: StockReportConfig, client: str, report_date: date) -> str:
    """e.g. 'Eurosia - Daily Stock Movement - 2026-10-06.pdf', safe on any operating system."""
    return safe_filename(_fill(cfg.filename_template, client, report_date), set())


def _qty(q: "sm.CartonQty", labels: tuple) -> str:
    """Opening / in / out cell: a dash for nothing, otherwise cartons + pieces."""
    return DASH if q.total == 0 else xml_escape(sm.format_cartons(q, *labels))


def _bold(text: str, color: str) -> str:
    """Emphasised text inside a table cell: the bold face, in the dark text colour. (The b tag does nothing
    here since the Thai font family has no registered bold mapping, so the face is named explicitly. The
    colour keeps the emphasis visible even where the bold face isn't being embedded.)"""
    return f'<font name="{engine.FONT_BOLD}" color="{color}">{text}</font>'


def _closing(q: "sm.CartonQty", labels: tuple, reconciled: bool, dark: str) -> str:
    """Closing cell: bold cartons + pieces, then small notes underneath: the exact pieces (when there is at
    least one whole carton), and, only when true, why the figure needs care."""
    text = _bold(xml_escape(sm.format_cartons(q, *labels)), dark)
    small = []
    if q.cartons:                                   # otherwise the line would just repeat the pieces
        small.append(f'<font size="6.5" color="{MUTED}">{q.total:,} pcs</font>')
    if q.cartons is None:
        small.append(f'<font size="6.5" color="{MUTED}">carton size not on file</font>')
    if not reconciled:
        small.append(f'<font size="6.5" color="{RED}">does not add up</font>')
    return text + "".join("<br/>" + line for line in small)


def _warehouse_text(by_warehouse: dict, per_carton: Optional[int], labels: tuple) -> str:
    lines = []
    for name, pieces in sorted(by_warehouse.items()):
        lines.append(f"{xml_escape(name)}: " + (DASH if abs(pieces) < 0.5 else xml_escape(
            sm.format_cartons(sm.to_cartons(pieces, per_carton), *labels))))
    return "<br/>".join(lines)


def _kpi_strip(items: list, usable_w: float, c_primary, c_mid, c_light_bg) -> Table:
    value_style = ParagraphStyle("KpiValue", fontName=engine.FONT_BOLD, fontSize=15, leading=19,
                                 textColor=c_primary, alignment=TA_CENTER)
    label_style = ParagraphStyle("KpiLabel", fontName=engine.FONT_REGULAR, fontSize=7.5, leading=10,
                                 textColor=c_mid, alignment=TA_CENTER)
    cells = [[Paragraph(value, value_style), Paragraph(label, label_style)] for value, label in items]
    t = Table([cells], colWidths=[usable_w / len(items)] * len(items))
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), c_light_bg),
        ("BOX", (0, 0), (-1, -1), 0.5, HexColor("#CCCCCC")),
        ("LINEAFTER", (0, 0), (-2, -1), 0.5, HexColor("#CCCCCC")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t


# ─────────────────────────────────────────────
#  BUILDING THE PDF
# ─────────────────────────────────────────────

def _build_pdf(ledger, master, client, report_date, cfg, out_path, assets_dir):
    engine._ensure_fonts_registered()
    brand = cfg.brand
    c_primary, c_accent = HexColor(brand.primary), HexColor(brand.accent)
    c_light_bg, c_dark, c_mid = HexColor(brand.light_bg), HexColor(brand.text_dark), HexColor(brand.text_mid)
    styles = engine.make_styles(c_dark, c_mid, colors.white, c_primary)
    dark = brand.text_dark
    usable_w = engine.page_size(cfg)[0] - 2 * engine.MARGIN + 7 * mm
    # headings sit flush with the table's left edge (tables overhang the frame by 3.5mm, and the frame pads 6pt)
    heading = ParagraphStyle("StockHeading", fontName=engine.FONT_BOLD, fontSize=12, textColor=c_primary,
                             spaceBefore=10, spaceAfter=6, leftIndent=-(6 + 3.5 * mm))

    note = ParagraphStyle("StockNote", parent=styles["intro"], fontSize=8.5, spaceAfter=2)
    summaries = ledger.sku_summaries()
    warehouses = sorted({w for s in summaries for w in s.by_warehouse})
    multi_wh = len(warehouses) > 1
    used_nocarton = used_mismatch = False

    def table(columns, rows):
        section_cfg = DocumentConfig(id="stock", name="stock", document_title=cfg.document_title,
                                     columns=columns, brand=cfg.brand, orientation=cfg.orientation)
        return engine.build_table(rows, section_cfg, styles, c_primary, c_accent, c_light_bg, c_mid, c_dark,
                                  colors.white, auto_height=True)

    qty_cols = [
        ColumnConfig(key="opening", label="Opening", width_mm=34, align="center"),
        ColumnConfig(key="in", label="In", width_mm=26, align="center"),
        ColumnConfig(key="out", label="Out", width_mm=34, align="center"),
        ColumnConfig(key="closing", label="Closing", width_mm=40, align="center"),
    ]
    wh_col = [ColumnConfig(key="wh", label="Closing by warehouse", width_mm=44, align="left")] if multi_wh else []

    # ---- stock by SKU
    sku_rows, total_open = [], 0.0
    for s in summaries:
        product = master.lookup(s.sku)
        per = product.pieces_per_carton if product else None
        labels, view = sm.unit_labels(product), sm.carton_view(s, product)
        closing = _closing(view["closing"], labels, s.reconciled, dark)
        used_nocarton |= view["closing"].cartons is None
        used_mismatch |= not s.reconciled
        row = {
            "sku": xml_escape(s.sku),
            "name": xml_escape(product.name) if product else f'<font color="{MUTED}">(not in product master data)</font>',
            "opening": _qty(view["opening"], labels), "in": _qty(view["in"], labels),
            "out": _qty(view["out"], labels), "closing": closing, "weight": f"{s.closing_wt:,.1f}",
        }
        if multi_wh:
            row["wh"] = _warehouse_text(s.by_warehouse, per, labels)
        sku_rows.append(row)
    t_open, t_in = sum(s.opening_qty for s in summaries), sum(s.in_qty for s in summaries)
    t_out, t_close = sum(s.out_qty for s in summaries), sum(s.closing_qty for s in summaries)
    t_wt = sum(s.closing_wt for s in summaries)
    pcs = lambda n: _bold(f"{n:,.0f} pcs", dark)
    sku_rows.append({"sku": _bold("Total", dark), "name": _bold(f"{len(summaries)} SKU{'s' if len(summaries) != 1 else ''}", dark),
                     "opening": pcs(t_open), "in": pcs(t_in), "out": pcs(t_out), "closing": pcs(t_close),
                     "weight": _bold(f"{t_wt:,.1f}", dark), "wh": ""})
    sku_cols = ([ColumnConfig(key="sku", label="SKU", width_mm=24, align="center", emphasis=True),
                 ColumnConfig(key="name", label="Product", width_mode="flex", flex_weight=1.0, align="left")]
                + qty_cols + [ColumnConfig(key="weight", label="Closing weight (kg)", width_mm=24, align="center")] + wh_col)

    # ---- batch detail
    batch_rows = []
    for b in ledger.batches:
        product = master.lookup(b.sku)
        per = product.pieces_per_carton if product else None
        labels = sm.unit_labels(product)
        q = {k: sm.to_cartons(v, per) for k, v in
             (("opening", b.opening_qty), ("in", b.in_qty), ("out", b.out_qty), ("closing", b.closing_qty))}
        used_nocarton |= q["closing"].cartons is None
        used_mismatch |= not b.reconciled
        row = {"sku": xml_escape(b.sku), "batch": xml_escape(b.batch) or DASH,
               "opening": _qty(q["opening"], labels), "in": _qty(q["in"], labels), "out": _qty(q["out"], labels),
               "closing": _closing(q["closing"], labels, b.reconciled, dark)}
        if multi_wh:
            row["wh"] = _warehouse_text(b.warehouse_closing(), per, labels)
        batch_rows.append(row)
    batch_cols = ([ColumnConfig(key="sku", label="SKU", width_mm=24, align="center", emphasis=True),
                   ColumnConfig(key="batch", label="Batch", width_mode="flex", flex_weight=1.0, align="left")]
                  + qty_cols + wh_col)

    # ---- the page
    has_signature = any(b.type == "signature" for b in cfg.footer_blocks)
    bottom = (20 + engine.SIGNATURE_AREA_HEIGHT / mm) * mm if has_signature else 20 * mm
    doc = SimpleDocTemplate(out_path, pagesize=engine.page_size(cfg), leftMargin=engine.MARGIN,
                            rightMargin=engine.MARGIN, topMargin=30 * mm, bottomMargin=bottom,
                            title=f"{cfg.document_title} - {client} - {_d(report_date)}",
                            author=brand.company_name, compress=1)

    where = (f"Warehouse: {xml_escape(warehouses[0])}" if len(warehouses) == 1
             else f"Warehouses: {xml_escape(', '.join(warehouses))}")
    kpis = [(f"{len(summaries)}", "SKUs"), (f"{t_close:,.0f}", "Closing stock (pieces)"),
            (f"{t_wt:,.1f}", "Closing weight (kg)"), (f"{t_in:,.0f}", "Received (pieces)"),
            (f"{t_out:,.0f}", "Shipped (pieces)")]
    story = [
        Spacer(1, 4 * mm),
        Paragraph(xml_escape(cfg.document_title), styles["cat_title"]),
        Paragraph(xml_escape(_fill(cfg.intro_text_template, client, report_date)), styles["intro"]),
        Paragraph(f"{where}. Quantities are shown as cartons + loose pieces; the WMS counts everything in pieces.", note),
        Spacer(1, 3 * mm),
        _kpi_strip(kpis, usable_w, c_primary, c_mid, c_light_bg),
        Paragraph("Stock by SKU", heading),
        table(sku_cols, sku_rows),
        CondPageBreak(45 * mm),
        Paragraph("Batch detail", heading),
        table(batch_cols, batch_rows),
        Spacer(1, 4 * mm),
    ]
    notes = []
    if used_nocarton:
        notes.append('Figures marked "carton size not on file" have no carton size in the product master data, '
                     "so they are shown in pieces only.")
    if used_mismatch:
        notes.append('Figures marked "does not add up": the WMS closing balance differs from opening + in − out. '
                     "The WMS figure is shown.")
    if master.updated_at:
        try:
            stamp = datetime.fromisoformat(master.updated_at).astimezone(THAI_TZ)
            notes.append(f"Product names and carton sizes from the product master data of {_d(stamp.date())}.")
        except ValueError:
            pass
    for n in notes:
        story.append(Paragraph(xml_escape(n), styles["footer_notes"]))
    story.extend(engine.build_footer_blocks(cfg, styles, c_mid, c_dark))
    doc.build(story, canvasmaker=engine.make_branded_canvas(cfg, assets_dir))
    return summaries


# ─────────────────────────────────────────────
#  PUBLIC ENTRY POINT
# ─────────────────────────────────────────────

def generate_stock_summary(xlsx_path: str, client: str, report_date: date, master: "sm.ProductMaster",
                           cfg: StockReportConfig, out_dir: str, assets_dir: str) -> StockSummaryResult:
    """Read the WMS stock ledger and write the client's summary PDF into out_dir.

    Raises sl.LedgerError (file can't be read), ClientMismatchError (wrong client) or StockReportError
    (nothing to report); nothing is written in those cases. Everything else is in result.warnings."""
    ledger = sl.read_stock_ledger(xlsx_path)
    if not ledger.batches:
        raise StockReportError("The file has no stock rows, so there is nothing to report.")

    check = sm.check_client(master, client, [b.sku for b in ledger.batches])
    if check.blocking:
        raise ClientMismatchError(check)

    warnings = list(check.messages)
    date_warning = ledger.date_warning(report_date)
    if date_warning:
        warnings.append(date_warning)
    warnings.extend(ledger.warnings)

    os.makedirs(out_dir, exist_ok=True)
    filename = stock_filename(cfg, client, report_date)
    path = os.path.join(out_dir, filename)
    summaries = _build_pdf(ledger, master, client, report_date, cfg, path, assets_dir)
    return StockSummaryResult(
        path=path, filename=filename, client=client, report_date=report_date, warnings=warnings, check=check,
        sku_count=len(summaries), batch_count=len(ledger.batches),
        closing_qty=sum(s.closing_qty for s in summaries), in_qty=sum(s.in_qty for s in summaries),
        out_qty=sum(s.out_qty for s in summaries), closing_wt=sum(s.closing_wt for s in summaries),
        all_reconciled=ledger.all_reconciled,
    )
