from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from markify.converters.docx import docx_to_markdown
from markify.converters.pdf import pdf_to_markdown
from markify.converters.pptx import pptx_to_markdown
from markify.converters.xlsx import xlsx_to_markdown


@dataclass(frozen=True)
class ConvertOptions:
    """Options shared by CLI and GUI."""

    attachments_dir: Path | None = None
    no_ocr: bool = False
    ocr_lang: str = "chi_sim+eng"
    max_rows: int = 5000


def convert_to_markdown(path: Path | str, options: ConvertOptions | None = None) -> str:
    """Convert a supported document to Markdown text."""
    if options is None:
        options = ConvertOptions()
    path = Path(path).expanduser().resolve()
    suf = path.suffix.lower()

    if suf == ".pdf":
        return pdf_to_markdown(path)
    if suf in (".docx", ".doc"):
        if suf == ".doc":
            raise ValueError(".doc (legacy binary) is not supported; save as .docx")
        return docx_to_markdown(path)
    if suf in (".xlsx", ".xlsm"):
        return xlsx_to_markdown(path, max_rows=options.max_rows)
    if suf in (".pptx", ".ppt"):
        if suf == ".ppt":
            raise ValueError(".ppt (legacy binary) is not supported; save as .pptx")
        att = Path(options.attachments_dir).resolve() if options.attachments_dir else None
        return pptx_to_markdown(
            path,
            attachments_dir=att,
            ocr_images=not options.no_ocr,
            ocr_lang=options.ocr_lang,
        )
    raise ValueError(f"Unsupported file type: {suf}")
