from __future__ import annotations

import argparse
import sys
from pathlib import Path

from markify import __version__
from markify.converters.docx import docx_to_markdown
from markify.converters.pdf import pdf_to_markdown
from markify.converters.pptx import pptx_to_markdown
from markify.converters.xlsx import xlsx_to_markdown


def _convert(path: Path, args: argparse.Namespace) -> str:
    suf = path.suffix.lower()
    if suf == ".pdf":
        return pdf_to_markdown(path)
    if suf in (".docx", ".doc"):
        if suf == ".doc":
            raise ValueError(".doc (legacy binary) is not supported; save as .docx")
        return docx_to_markdown(path)
    if suf in (".xlsx", ".xlsm"):
        return xlsx_to_markdown(path, max_rows=args.max_rows)
    if suf in (".pptx", ".ppt"):
        if suf == ".ppt":
            raise ValueError(".ppt (legacy binary) is not supported; save as .pptx")
        att = Path(args.attachments_dir).resolve() if args.attachments_dir else None
        return pptx_to_markdown(
            path,
            attachments_dir=att,
            ocr_images=not args.no_ocr,
            ocr_lang=args.ocr_lang,
        )
    raise ValueError(f"Unsupported file type: {suf}")


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="markify", description="Convert documents to Markdown")
    p.add_argument("input", type=Path, help="Input file (.pdf, .docx, .xlsx, .pptx)")
    p.add_argument("-o", "--output", type=Path, help="Output .md file (default: stdout)")
    p.add_argument(
        "--attachments-dir",
        type=Path,
        help="For .pptx: extract embedded files (ppt/embeddings + OLE) into this folder",
    )
    p.add_argument(
        "--no-ocr",
        action="store_true",
        help="For .pptx: disable image OCR (Tesseract)",
    )
    p.add_argument(
        "--ocr-lang",
        default="chi_sim+eng",
        help="Tesseract language(s), e.g. chi_sim+eng or eng (default: chi_sim+eng)",
    )
    p.add_argument(
        "--max-rows",
        type=int,
        default=5000,
        help="For .xlsx: max rows per sheet (default: 5000)",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = p.parse_args(argv)

    path = args.input.expanduser().resolve()
    if not path.is_file():
        print(f"Not a file: {path}", file=sys.stderr)
        return 1

    try:
        md = _convert(path, args)
    except ValueError as e:
        print(str(e), file=sys.stderr)
        return 1
    except Exception as e:  # noqa: BLE001
        print(f"Conversion failed: {e}", file=sys.stderr)
        return 1

    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(md, encoding="utf-8")
    else:
        sys.stdout.buffer.write((md + "\n").encode("utf-8"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
