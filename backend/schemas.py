"""
schemas.py — the config schema every CoDocuments template conforms to.

A "template" (a.k.a. preset) is one of these, saved as JSON. The engine
(engine.py) never has type-specific logic — it only ever reads a
DocumentConfig. Client catalog, BU catalog, and quotation sheet are three
JSON files that use this same schema, not three code paths.
"""

from __future__ import annotations
from typing import Literal, Optional
from pydantic import BaseModel, Field


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
    # Which Excel header this pulls from. Not required for type="image"
    # (images are matched by row position, same as today's floating-image logic).
    # Left as None in the *template* until the user maps it during generation —
    # the template ships with a suggested/default value the user can override.
    source_header: Optional[str] = None
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
    # "signature" is the first supported block (for quotation sheets).
    # More block types (terms text, totals, etc.) can be added later
    # without touching the engine's core loop.
    type: Literal["signature", "text"]
    # For type="text": the text to render.
    text: Optional[str] = None
    # For type="signature": labels for each signature slot.
    labels: list[str] = Field(default_factory=lambda: ["Prepared by", "Approved by"])


class DocumentConfig(BaseModel):
    id: str
    name: str                      # shown in the preset picker, e.g. "Bakery BU Catalog"
    document_title: str            # printed at the top of the PDF
    intro_text_template: str = "This document lists {count} item(s)."

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

    class Config:
        json_schema_extra = {
            "example": {
                "id": "client_catalog",
                "name": "Client Product Catalog",
                "document_title": "Product Suggestions Catalog",
            }
        }
