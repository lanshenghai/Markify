# Markify

Convert **PDF**, **Word** (`.docx`), **Excel** (`.xlsx` / `.xlsm`), and **PowerPoint** (`.pptx`) into **Markdown**.

## Requirements

- **Python** 3.10+
- **Tesseract OCR** — required only if you use PowerPoint **image OCR**. Install [Tesseract](https://github.com/tesseract-ocr/tesseract) and the language packs you pass to `--ocr-lang` (for example `eng`, or `chi_sim+eng` for Chinese + English).

## Installation

From the repository root:

```bash
pip install -e .
```

This installs the `markify` command-line tool.

## Usage

```bash
markify INPUT [-o OUTPUT.md] [options]
```

If `-o` / `--output` is omitted, Markdown is written to **stdout**.

### Examples

```bash
# PDF → Markdown
markify report.pdf -o report.md

# Word → Markdown
markify notes.docx -o notes.md

# Excel → Markdown (each sheet becomes a section with a pipe table)
markify data.xlsx -o data.md

# PowerPoint → Markdown, extract embedded files, OCR images (Chinese + English)
markify deck.pptx -o deck.md --attachments-dir ./ppt_assets --ocr-lang chi_sim+eng

# PowerPoint without OCR (text and tables only)
markify deck.pptx -o deck.md --no-ocr
```

### CLI options

| Option | Description |
|--------|-------------|
| `-o`, `--output` | Output `.md` file path |
| `--attachments-dir` | **`.pptx` only:** extract embedded parts under `ppt/embeddings/` and OLE blobs into this directory; paths are listed in the Markdown |
| `--no-ocr` | **`.pptx` only:** disable image OCR |
| `--ocr-lang` | **`.pptx` only:** Tesseract language string (default: `chi_sim+eng`). Examples: `eng`, `deu`, `chi_sim+eng` |
| `--max-rows` | **`.xlsx` only:** max rows per sheet (default: `5000`) |
| `--version` | Print version |

You can also run the package as a module:

```bash
python -m markify report.pdf -o report.md
```

## Format behavior

| Format | Notes |
|--------|--------|
| **PDF** | Uses PyMuPDF: prefers markdown-like extraction, falls back to plain text per page |
| **DOCX** | Paragraphs (including heading styles) and tables in document order |
| **XLSX / XLSM** | One section per worksheet; large sheets may be truncated via `--max-rows` |
| **PPTX** | See below |

### PowerPoint (`.pptx`)

- **Text:** Recurses into **groups**; reads text frames and **tables** as Markdown tables.
- **Attachments:** With `--attachments-dir`, saves ZIP parts from `ppt/embeddings/` and OLE object payloads from shapes when available; lists saved paths in the output.
- **Images → text:** Optional OCR over images in `ppt/media/` and **picture** shapes; duplicate image bytes are skipped using **SHA-256**. If Tesseract is missing or misconfigured, conversion still runs; warnings appear under **OCR warnings** in the Markdown.

## Limitations

- Legacy **`.doc`** and **`.ppt`** (binary Office formats) are **not** supported. Save as **`.docx`** / **`.pptx`** first.
- Chart titles, SmartArt, and speaker notes are not specifically extracted beyond what `python-pptx` and the ZIP expose as shapes/media.

## License

See the repository for license information.
