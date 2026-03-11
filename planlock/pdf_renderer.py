from __future__ import annotations

import base64
from dataclasses import dataclass

import fitz


@dataclass(frozen=True)
class RenderedPage:
    page_number: int
    image_bytes: bytes
    native_text: str


def render_pdf_pages(pdf_bytes: bytes, max_pages: int) -> list[RenderedPage]:
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    rendered: list[RenderedPage] = []
    page_limit = min(document.page_count, max_pages)
    for page_index in range(page_limit):
        page = document.load_page(page_index)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        rendered.append(
            RenderedPage(
                page_number=page_index + 1,
                image_bytes=pixmap.tobytes("png"),
                native_text=page.get_text("text"),
            )
        )
    return rendered


def render_pdf_previews(pdf_bytes: bytes, max_pages: int, scale: float = 0.32) -> dict[int, str]:
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    previews: dict[int, str] = {}
    page_limit = min(document.page_count, max_pages)
    matrix = fitz.Matrix(scale, scale)
    for page_index in range(page_limit):
        page = document.load_page(page_index)
        pixmap = page.get_pixmap(matrix=matrix, alpha=False)
        preview_bytes = pixmap.tobytes("png")
        previews[page_index + 1] = (
            "data:image/png;base64," + base64.b64encode(preview_bytes).decode("ascii")
        )
    return previews
