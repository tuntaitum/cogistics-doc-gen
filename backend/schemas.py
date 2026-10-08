"""
schemas.py — the config schema every CoDocuments template conforms to.

A "template" (a.k.a. preset) is one of these, saved as JSON. The engine
(engine.py) never has type-specific logic — it only ever reads a
DocumentConfig. Client catalog, BU catalog, and quotation sheet are three
JSON files that use this same schema, not three code paths.
"""

from __future__ import annotations
from typing import Literal, Optional
from pydantic import BaseModel, Field, model_validator


class BrandConfig(BaseModel):
    company_name: str = "COGISTICS"
    tagline: str = "The One Stop Cold Chain Solution"
    primary: str = "#1A3C5E"      # header bar / table header
    accent: str = "#7C7C7C"       # stripes, borders
    light_bg: str = "#F5F7FA"     # alternating row background
    text_dark: str = "#1C1C1C"
    text_mid: str = "#555555"
    banner_left: Optional[str] = "cogistics header banner left.png"
    banner_right: Optional[str] = "cogistics header banner right.png"


class ColumnConfig(BaseModel):
    """One column in the output PDF table."""

    key: str
    # Label shown in the PDF header row
    label: str
    # "image" | "text"
    type: Literal["image", "text"] = "text"
    # "excel" (default): value comes from a mapped spreadsheet column.
    # "manual": value is typed in per-row by the user in the web UI at
    # generation time — for data that doesn't exist in the source file at
    # all (e.g. an employee manually assigning a Quantity per line item).
    # Manual columns are always added ad-hoc per-generation by the user,
    # never part of a saved template's default column list.
    # "computed": value is calculated from other columns in the same row
    # (e.g. a Subtotal = Unit Price x Quantity) — see compute_operation /
    # compute_operands below.
    source: Literal["excel", "manual", "computed"] = "excel"
    # Which Excel header this pulls from. Only meaningful when source="excel".
    # Left as None in the *template* until the user maps it during generation —
    # the template ships with a suggested/default value the user can override.
    source_header: Optional[str] = None
    # Only meaningful when source="computed". "multiply"/"add"/"subtract"
    # applied left-to-right across compute_operands (column keys, evaluated
    # after excel/manual values are all in place). Non-numeric or missing
    # operand values make the result blank for that row rather than erroring
    # — a quotation row missing a Quantity just has no Subtotal, not a crash.
    compute_operation: Literal["multiply", "add", "subtract"] = "multiply"
    compute_operands: list[str] = Field(default_factory=list)
    # Fixed width in mm. Ignored if width_mode="flex".
    width_mm: Optional[float] = None
    # "fixed" uses width_mm. "flex" splits remaining space proportionally
    # across all flex columns, weighted by flex_weight.
    width_mode: Literal["fixed", "flex"] = "fixed"
    flex_weight: float = 1.0
    # If true, the whole column is hidden when every product's value for it
    # is empty (mirrors today's show_price / show_advantage behavior).
    optional: bool = False
    align: Literal["left", "center"] = "left"
    # Bold styling for the "headline" column of a row (e.g. product/item name)
    emphasis: bool = False


class FooterBlock(BaseModel):
    # "signature": a fixed block pinned to the bottom of the last page only.
    # "text": free-form notes/terms, flows normally after the table (e.g.
    # numbered quotation conditions) — available on every document type,
    # not just quotation sheets.
    type: Literal["signature", "text"]
    # For type="text": the text to render. Plain text with real line breaks —
    # the engine converts them to proper PDF line breaks and escapes any
    # special characters, so no markup knowledge is needed to use this.
    text: Optional[str] = None
    # For type="signature": labels for each signature slot.
    labels: list[str] = Field(default_factory=lambda: ["Prepared by", "Approved by"])
    # For type="signature": if true, also draw a shorter "Date" line beneath
    # each signature line for that party.
    include_date: bool = True


class DocumentConfig(BaseModel):
    id: str
    name: str                      # shown in the preset picker, e.g. "Bakery BU Catalog"
    document_title: str            # printed at the top of the PDF
    intro_text_template: str = "This document lists {count} item(s)."
    # Page orientation. Defaults to portrait so every existing preset is
    # unchanged; wide tables (e.g. the Food Pipeline delivery report) opt in.
    orientation: Literal["portrait", "landscape"] = "portrait"

    # --- Excel reading ---
    header_row: int = 2            # row containing column headers
    data_start_row: int = 3        # first row of actual data
    select_column: str = "Select"  # header of the yes/no "include this row" column
    select_value: str = "yes"      # value (case-insensitive) that means "include"
    # 0-indexed Excel columns to skip when matching floating images to rows
    # (mirrors "skip col F — internal photos" in the original script).
    image_skip_columns: list[int] = Field(default_factory=list)

    columns: list[ColumnConfig]
    footer_blocks: list[FooterBlock] = Field(default_factory=list)
    brand: BrandConfig = Field(default_factory=BrandConfig)

    # If set, the table gets an extra summary row at the very bottom summing
    # this column (by key) across all rows — e.g. a quotation's grand Total
    # under a Subtotal column. Non-numeric values are skipped when summing
    # rather than erroring.
    totals_column: Optional[str] = None
    totals_label: str = "Total"

    class Config:
        json_schema_extra = {
            "example": {
                "id": "client_catalog",
                "name": "Product Suggestions Catalog",
                "document_title": "Product Suggestions Catalog",
            }
        }


# ═════════════════════════════════════════════
#  GROUPED, MULTI-SHEET REPORTS
#  (e.g. Food Pipeline's daily delivery report: one spreadsheet in, one
#  PDF per client out.)
#
#  Deliberately a SEPARATE model from DocumentConfig. A DocumentConfig is
#  "one sheet -> one table -> one PDF"; this is "several sheets -> rows
#  grouped by a column -> one PDF per group, each with a table per
#  section". Forcing both into one schema would make every DocumentConfig
#  carry fields that only make sense for reports. Both share the same
#  building blocks (ColumnConfig, FooterBlock, BrandConfig) so the table
#  and page rendering code can be reused.
#
#  Report JSON files live in backend/reports/ — NOT backend/presets/, whose
#  files are all assumed to be DocumentConfigs by the veggie preset picker.
# ═════════════════════════════════════════════

class ReportColumn(ColumnConfig):
    """A ColumnConfig for grouped reports.

    Header matching: `source_header` is compared to the sheet's header
    after collapsing all whitespace (including line breaks) to single
    spaces, so write it on one line: the sheet's "สินค้าทั้งหมด \\n( รายการ )"
    is matched by "สินค้าทั้งหมด ( รายการ )".
    """

    # If set, the cell value is split on this separator and each piece is
    # shown on its own line inside the cell. For TMS cells that pack several
    # items into one string (e.g. the product list joined with " , ").
    split_on: Optional[str] = None


class ReportSection(BaseModel):
    """One table in each client's report, fed by one sheet of the workbook."""

    key: str                       # stable id, e.g. "delivered"
    title: str                     # heading printed above this table
    # Sheet to read, WITHOUT its numeric prefix: the TMS names sheets
    # "1.กำลังจัดส่ง", "2.จัดส่งสำเร็จ"; write "กำลังจัดส่ง", "จัดส่งสำเร็จ".
    # The reader strips a leading "<digits>." before comparing.
    sheet_name: str
    # Each section has its OWN columns because the TMS sheets share most
    # headers but not all (e.g. the date column is called "delivered date" on
    # one sheet and "confirmed-failed date" on another).
    columns: list[ReportColumn]
    # Show this section for a client even with zero rows (e.g. "Delivered"
    # should always appear, "Failed" should only appear when there is a failure).
    show_when_empty: bool = False
    # Optional sanity check: if set, rows on this sheet whose status cell
    # (see GroupedReportConfig.status_header) differs are FLAGGED (never
    # silently dropped or relabelled).
    expected_status: Optional[str] = None
    # Heading colour override (hex), e.g. red for the failed section.
    title_color: Optional[str] = None


class GroupedReportConfig(BaseModel):
    id: str
    name: str
    document_title: str
    # "{group}" = the group value (e.g. the client), "{total}" = rows across sections
    intro_text_template: str = "{group} --- {total} order(s)"
    # Output file name (without .pdf); "{group}" is replaced, then sanitised.
    filename_template: str = "{group}"
    orientation: Literal["portrait", "landscape"] = "portrait"

    # --- Excel reading ---
    header_row: int = 2
    data_start_row: int = 3

    # Header of the column whose unique values each get their own PDF.
    group_by_header: str
    # Rows where the group cell is blank are NOT dropped: they are put in a
    # report under this label so nothing silently disappears.
    blank_group_label: str = "(No client specified)"
    # Header of the per-row delivery-status column, used by expected_status.
    status_header: Optional[str] = None

    sections: list[ReportSection] = Field(min_length=1)
    footer_blocks: list[FooterBlock] = Field(default_factory=list)
    brand: BrandConfig = Field(default_factory=BrandConfig)

    @model_validator(mode="after")
    def _check_structure(self):
        keys = [sec.key for sec in self.sections]
        if len(set(keys)) != len(keys):
            raise ValueError(f"section keys must be unique, got {keys}")
        sheets = [sec.sheet_name for sec in self.sections]
        if len(set(sheets)) != len(sheets):
            raise ValueError(f"each section needs its own sheet, got {sheets}")
        for sec in self.sections:
            ckeys = [c.key for c in sec.columns]
            if len(set(ckeys)) != len(ckeys):
                raise ValueError(f"section '{sec.key}': column keys must be unique, got {ckeys}")
            for c in sec.columns:
                if c.source != "excel" or c.type != "text" or not c.source_header:
                    raise ValueError(
                        f"section '{sec.key}', column '{c.key}': reports only support "
                        "text columns read from the spreadsheet (type='text', source='excel', "
                        "with a source_header). Manual, computed and image columns are not "
                        "available here."
                    )
        if any(sec.expected_status for sec in self.sections) and not self.status_header:
            raise ValueError("expected_status is set on a section but status_header is missing")
        return self
