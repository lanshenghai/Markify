from __future__ import annotations

import argparse
import sys
from pathlib import Path

from markify import __version__
from markify.core import ConvertOptions, convert_to_markdown


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

    opts = ConvertOptions(
        attachments_dir=Path(args.attachments_dir).resolve() if args.attachments_dir else None,
        no_ocr=args.no_ocr,
        ocr_lang=args.ocr_lang,
        max_rows=args.max_rows,
    )

    try:
        md = convert_to_markdown(path, opts)
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
