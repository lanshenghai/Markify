from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def _paragraph_md(p: Paragraph) -> str:
    text = p.text.strip()
    if not text:
        return ""
    style = p.style.name if p.style else ""
    if style.startswith("Heading"):
        try:
            level = int(style.replace("Heading", "").strip() or "1")
            level = max(1, min(level, 6))
        except ValueError:
            level = 1
        return f"{'#' * level} {text}"
    return text


def _table_md(table: Table) -> str:
    rows: list[list[str]] = []
    for row in table.rows:
        cells = [cell.text.replace("\n", " ").strip() for cell in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    lines = []
    header = rows[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join("---" for _ in header) + " |")
    for r in rows[1:]:
        lines.append("| " + " | ".join(r) + " |")
    return "\n".join(lines)


def docx_to_markdown(path: str | Path) -> str:
    path = Path(path)
    doc = Document(path)
    blocks: list[str] = []

    for child in doc.element.body:
        if child.tag == qn("w:p"):
            para = Paragraph(child, doc)
            b = _paragraph_md(para)
            if b:
                blocks.append(b)
        elif child.tag == qn("w:tbl"):
            table = Table(child, doc)
            b = _table_md(table)
            if b:
                blocks.append(b)

    return "\n\n".join(blocks)
