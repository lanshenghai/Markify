from __future__ import annotations

from pathlib import Path

import fitz  # PyMuPDF


def pdf_to_markdown(path: str | Path) -> str:
    """Extract PDF content as Markdown (PyMuPDF markdown mode when available)."""
    path = Path(path)
    doc = fitz.open(path)
    try:
        parts: list[str] = []
        for page in doc:
            md = ""
            try:
                md = page.get_text("markdown") or ""
            except (AttributeError, ValueError, RuntimeError):
                pass
            if not md.strip():
                md = page.get_text("text") or ""
            if md.strip():
                parts.append(md.strip())
        return "\n\n".join(parts) if parts else ""
    finally:
        doc.close()
