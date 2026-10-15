"""
TEMPLATE — copy to backend/<your_tool>_routes.py and rename. See docs/ADDING_A_TOOL.md.

Pattern used by report_routes.py: the feature lives in its OWN module and exposes a
router. main.py only imports it and includes it; you don't add routes to main.py.

This example is "stateless": it builds a PDF in memory and returns it, storing nothing
on the server. That is the cheapest, safest pattern (no files to clean up on the
Railway volume), and the one the NON-GMO letter uses. If your tool must save files,
read "Files" in the guide first.
"""

import io

from fastapi import APIRouter
from fastapi.responses import Response
from pydantic import BaseModel, Field
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from memory import releases_memory   # hands memory back to the OS when the request ends


class ExampleRequest(BaseModel):
    # Validate on the SERVER. The browser's maxlength is a convenience, not a safeguard.
    title: str = Field(min_length=1, max_length=90)


def make_router() -> APIRouter:
    router = APIRouter(prefix="/api/example-tool", tags=["example-tool"])

    @router.post("/make")
    @releases_memory                # put this on anything that handles images, workbooks or big PDFs
    def make(request: ExampleRequest):
        """Return a one-line PDF. Replace this with your tool's real rendering code."""
        buffer = io.BytesIO()
        pdf = canvas.Canvas(buffer, pagesize=A4)
        # NOTE: this standard font cannot draw Thai. For Thai text, call
        # engine._ensure_fonts_registered() and use engine.FONT_REGULAR instead.
        pdf.setFont("Helvetica", 16)
        pdf.drawString(72, 770, request.title)
        pdf.save()
        return Response(
            content=buffer.getvalue(),
            media_type="application/pdf",
            headers={"Content-Disposition": 'inline; filename="example.pdf"'},
        )

    # To return a PLAIN-LANGUAGE error the user will see on the page:
    #     from fastapi import HTTPException
    #     raise HTTPException(422, "The issue date can't be in the future.")

    return router
