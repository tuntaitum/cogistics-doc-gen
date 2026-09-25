"""Standalone guarantee letter renderer. Inputs are deliberately limited to date/products."""

from __future__ import annotations

import calendar
from datetime import date
from io import BytesIO
from pathlib import Path

from reportlab.lib.utils import ImageReader, simpleSplit
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

PAGE_W, PAGE_H = 612, 792  # US Letter, matching NON.docx
ASSET = Path(__file__).parent / "assets" / "non-gmo-letterhead.png"
FONT = "DejaVuSerif"
pdfmetrics.registerFont(TTFont(FONT, str(Path(__file__).parent / "fonts" / "DejaVuSerif.ttf")))


def render_non_gmo(issue_date: date, products: list[str]) -> bytes:
    if not 1 <= len(products) <= 6:
        raise ValueError("Choose 1 to 6 products for this one-page letter")
    names = [p.strip() for p in products]
    if any(not p or len(p) > 90 or "\n" in p or "\r" in p for p in names):
        raise ValueError("Each product name must be one line, 1 to 90 characters")

    out = BytesIO()
    pdf = canvas.Canvas(out, pagesize=(PAGE_W, PAGE_H))
    pdf.setTitle("NON-GMO GUARANTEE LETTER")
    pdf.drawImage(ImageReader(str(ASSET)), 36, 707, width=540, height=80, mask="auto")

    pdf.setFont(FONT, 10)
    pdf.drawCentredString(306, 678, "NON-GMO GUARANTEE LETTER")
    day = f"{issue_date.day} {calendar.month_name[issue_date.month]} {issue_date.year}"
    pdf.drawRightString(524, 651, day)
    pdf.drawString(56, 624, "To Whom It May Concern,")

    def paragraph(value: str, y: float) -> float:
        lines = simpleSplit(value, FONT, 10, 500)
        for line in lines:
            pdf.drawString(83, y, line)
            y -= 19
        return y - 12

    y = paragraph(
        "We, Cogistics Co., Ltd., hereby guarantee that the following products "
        "supplied by our company are non-genetically modified (Non-GMO)", 596
    )
    for i, product in enumerate(names, 1):
        lines = simpleSplit(product, FONT, 10, 435)
        for n, line in enumerate(lines):
            pdf.drawString(70 if n == 0 else 83, y, f"{i}.   {line}" if n == 0 else line)
            y -= 19
        y -= 6
    y -= 10
    y = paragraph(
        "We confirm that the above-mentioned products do not contain, consist of, "
        "or are produced from genetically modified organisms. No genetically "
        "modified raw materials or ingredients are used in these products.", y
    )
    y = paragraph(
        "This guarantee shall remain valid unless there is any change to the raw "
        "materials, ingredients, product specifications, or manufacturing processes "
        "that may affect the Non-GMO status of the products.", y
    )
    if y < 240:
        raise ValueError("The product list is too long for the one-page letter")

    pdf.drawString(441, 179, "Yours sincerely,")
    # The supplied Word document has no signature image. Leave space to sign
    # rather than copying a signature from a previously issued PDF.
    pdf.drawCentredString(487, 123, "(Ms.Thiaranya Maharn)")
    pdf.drawCentredString(487, 103, "QA Officer")
    pdf.save()
    return out.getvalue()
