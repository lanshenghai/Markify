from __future__ import annotations

import queue
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from markify import __version__
from markify.core import ConvertOptions, convert_to_markdown

_PREVIEW_CHARS = 12_000


def _default_output_path(input_path: Path) -> Path:
    return input_path.with_suffix(".md")


def _append_log(widget: scrolledtext.ScrolledText, line: str) -> None:
    widget.configure(state="normal")
    widget.insert(tk.END, line + "\n")
    widget.see(tk.END)
    widget.configure(state="disabled")


def _set_preview(widget: scrolledtext.ScrolledText, text: str, truncated: bool) -> None:
    widget.configure(state="normal")
    widget.delete("1.0", tk.END)
    widget.insert(tk.END, text)
    if truncated:
        widget.insert(tk.END, "\n\n---\n_(Preview truncated; full text was saved to the output file.)_\n")
    widget.configure(state="disabled")


def main() -> None:
    root = tk.Tk()
    root.title(f"Markify {__version__}")
    root.minsize(640, 520)

    events: queue.Queue = queue.Queue()
    converting = {"active": False}

    frm = ttk.Frame(root, padding=10)
    frm.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    frm.columnconfigure(1, weight=1)

    # Input
    ttk.Label(frm, text="Input file").grid(row=0, column=0, sticky="w", pady=(0, 4))
    input_var = tk.StringVar()
    input_entry = ttk.Entry(frm, textvariable=input_var, width=56)
    input_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 8))

    def browse_input() -> None:
        path = filedialog.askopenfilename(
            title="Select document",
            filetypes=[
                ("Supported", "*.pdf *.docx *.xlsx *.xlsm *.pptx"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx"),
                ("Excel", "*.xlsx *.xlsm"),
                ("PowerPoint", "*.pptx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            input_var.set(path)
            p = Path(path)
            if p.is_file():
                out_var.set(str(_default_output_path(p)))

    ttk.Button(frm, text="Browse…", command=browse_input).grid(row=1, column=2, padx=(8, 0), pady=(0, 8))

    # Output
    ttk.Label(frm, text="Output Markdown").grid(row=2, column=0, sticky="w", pady=(0, 4))
    out_var = tk.StringVar()
    ttk.Entry(frm, textvariable=out_var, width=56).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 8))

    def browse_output() -> None:
        initial = out_var.get().strip() or None
        path = filedialog.asksaveasfilename(
            title="Save Markdown as",
            defaultextension=".md",
            filetypes=[("Markdown", "*.md"), ("Text", "*.txt"), ("All files", "*.*")],
            initialfile=Path(initial).name if initial else None,
        )
        if path:
            out_var.set(path)

    ttk.Button(frm, text="Browse…", command=browse_output).grid(row=3, column=2, padx=(8, 0), pady=(0, 8))

    # Options frame
    opt_frame = ttk.LabelFrame(frm, text="Options", padding=8)
    opt_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 8))
    opt_frame.columnconfigure(1, weight=1)

    pptx_ocr = tk.BooleanVar(value=True)
    ttk.Checkbutton(opt_frame, text="PowerPoint: OCR images (Tesseract)", variable=pptx_ocr).grid(
        row=0, column=0, columnspan=2, sticky="w"
    )

    ttk.Label(opt_frame, text="OCR language").grid(row=1, column=0, sticky="w", pady=(6, 0))
    ocr_lang_var = tk.StringVar(value="chi_sim+eng")
    ttk.Entry(opt_frame, textvariable=ocr_lang_var, width=24).grid(row=1, column=1, sticky="w", pady=(6, 0))

    ttk.Label(opt_frame, text="PPTX attachments folder (optional)").grid(row=2, column=0, sticky="w", pady=(6, 0))
    att_var = tk.StringVar()
    ttk.Entry(opt_frame, textvariable=att_var, width=40).grid(row=2, column=1, sticky="ew", pady=(6, 0))

    def browse_att() -> None:
        path = filedialog.askdirectory(title="Folder for extracted PPTX attachments")
        if path:
            att_var.set(path)

    ttk.Button(opt_frame, text="Browse…", command=browse_att).grid(row=2, column=2, padx=(8, 0), pady=(6, 0))

    ttk.Label(opt_frame, text="Excel max rows / sheet").grid(row=3, column=0, sticky="w", pady=(6, 0))
    max_rows_var = tk.StringVar(value="5000")
    ttk.Entry(opt_frame, textvariable=max_rows_var, width=12).grid(row=3, column=1, sticky="w", pady=(6, 0))

    # Log
    ttk.Label(frm, text="Progress").grid(row=5, column=0, sticky="w", pady=(4, 4))
    log_box = scrolledtext.ScrolledText(frm, height=8, wrap=tk.WORD, state="disabled")
    log_box.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(0, 8))
    frm.rowconfigure(6, weight=0)

    # Preview
    ttk.Label(frm, text="Preview (start of result)").grid(row=7, column=0, sticky="w", pady=(4, 4))
    preview = scrolledtext.ScrolledText(frm, height=14, wrap=tk.WORD, state="disabled")
    preview.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=(0, 8))
    frm.rowconfigure(8, weight=1)

    convert_btn = ttk.Button(frm, text="Convert")

    def poll_events() -> None:
        try:
            while True:
                kind, payload = events.get_nowait()
                if kind == "log":
                    _append_log(log_box, payload)
                elif kind == "done":
                    md: str = payload[0]
                    out_path: Path = payload[1]
                    elapsed: float = payload[2]
                    truncated = len(md) > _PREVIEW_CHARS
                    prev = md[:_PREVIEW_CHARS] if truncated else md
                    _set_preview(preview, prev, truncated)
                    _append_log(
                        log_box,
                        f"Finished in {elapsed:.2f}s — {len(md)} characters — saved to:\n{out_path}",
                    )
                    converting["active"] = False
                    convert_btn.configure(state="normal")
                elif kind == "fail":
                    _append_log(log_box, f"Error: {payload}")
                    converting["active"] = False
                    convert_btn.configure(state="normal")
        except queue.Empty:
            pass
        root.after(80, poll_events)

    def do_convert() -> None:
        if converting["active"]:
            return
        in_path = input_var.get().strip()
        if not in_path:
            messagebox.showwarning("Markify", "Please choose an input file.")
            return
        p = Path(in_path)
        if not p.is_file():
            messagebox.showerror("Markify", f"Not a file:\n{p}")
            return
        out_str = out_var.get().strip()
        out_path = Path(out_str) if out_str else _default_output_path(p)
        try:
            max_rows = int(max_rows_var.get().strip())
            if max_rows < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Markify", "Excel max rows must be a positive integer.")
            return

        att_raw = att_var.get().strip()
        opts = ConvertOptions(
            attachments_dir=Path(att_raw).resolve() if att_raw else None,
            no_ocr=not pptx_ocr.get(),
            ocr_lang=ocr_lang_var.get().strip() or "chi_sim+eng",
            max_rows=max_rows,
        )

        log_box.configure(state="normal")
        log_box.delete("1.0", tk.END)
        log_box.configure(state="disabled")
        preview.configure(state="normal")
        preview.delete("1.0", tk.END)
        preview.configure(state="disabled")

        converting["active"] = True
        convert_btn.configure(state="disabled")
        events.put(("log", f"Input: {p}"))
        events.put(("log", f"Output: {out_path}"))
        events.put(("log", "Converting…"))

        def task() -> None:
            t0 = time.perf_counter()
            try:
                md = convert_to_markdown(p, opts)
                out_path.parent.mkdir(parents=True, exist_ok=True)
                out_path.write_text(md, encoding="utf-8")
                elapsed = time.perf_counter() - t0
                events.put(("done", (md, out_path.resolve(), elapsed)))
            except Exception as e:  # noqa: BLE001
                events.put(("fail", str(e)))

        threading.Thread(target=task, daemon=True).start()

    convert_btn.configure(command=do_convert)
    convert_btn.grid(row=9, column=0, sticky="w", pady=(4, 0))

    ttk.Label(frm, text=f"v{__version__}", foreground="gray").grid(row=9, column=2, sticky="e", pady=(4, 0))

    poll_events()
    root.mainloop()


if __name__ == "__main__":
    main()
