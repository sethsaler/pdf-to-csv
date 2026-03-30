#!/usr/bin/env python3
"""
Desktop GUI for selecting PDF files or folders and exporting tagged structure to CSV/Excel.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from extract_tagged_pdf import export_pdfs


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    seen: set[Path] = set()
    out: list[Path] = []
    for p in paths:
        try:
            r = p.resolve()
        except OSError:
            r = p
        if r not in seen:
            seen.add(r)
            out.append(p)
    return out


class PdfExportApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Tagged PDF → CSV / Excel")
        self.root.minsize(560, 380)

        self._paths: list[Path] = []

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(1, weight=1)

        ttk.Label(main, text="PDF files to export").grid(row=0, column=0, sticky="w")

        list_frame = ttk.Frame(main)
        list_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(4, 8))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        scroll = ttk.Scrollbar(list_frame)
        scroll.grid(row=0, column=1, sticky="ns")
        self._list = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scroll.set,
            font=("TkFixedFont", 11),
        )
        self._list.grid(row=0, column=0, sticky="nsew")
        scroll.config(command=self._list.yview)

        btn_row = ttk.Frame(main)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(btn_row, text="Add files…", command=self._add_files).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Add folder…", command=self._add_folder).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Remove selected", command=self._remove_selected).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Clear all", command=self._clear).pack(side=tk.LEFT)

        out_row = ttk.Frame(main)
        out_row.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        out_row.columnconfigure(1, weight=1)
        ttk.Label(out_row, text="Output file").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._output_var = tk.StringVar(value=str(Path.home() / "tagged_export.xlsx"))
        out_entry = ttk.Entry(out_row, textvariable=self._output_var)
        out_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(out_row, text="Browse…", command=self._browse_output).grid(
            row=0, column=2, padx=(8, 0)
        )

        fmt_row = ttk.Frame(main)
        fmt_row.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(fmt_row, text="Format").pack(side=tk.LEFT, padx=(0, 8))
        self._fmt_var = tk.StringVar(value="auto")
        for val, label in (
            ("auto", "Auto (from extension)"),
            ("csv", "CSV"),
            ("xlsx", "Excel (.xlsx)"),
        ):
            ttk.Radiobutton(fmt_row, text=label, variable=self._fmt_var, value=val).pack(
                side=tk.LEFT, padx=(0, 12)
            )

        ttk.Button(main, text="Export", command=self._export).grid(row=5, column=0, sticky="w")
        self._status = ttk.Label(main, text="Add one or more PDFs, choose an output path, then Export.")
        self._status.grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 0))

    def _sync_listbox(self) -> None:
        self._list.delete(0, tk.END)
        for p in self._paths:
            self._list.insert(tk.END, str(p))

    def _add_files(self) -> None:
        names = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*")],
        )
        if not names:
            return
        self._paths = _dedupe_paths(self._paths + [Path(n) for n in names])
        self._sync_listbox()
        self._set_status(f"{len(self._paths)} file(s) selected.")

    def _add_folder(self) -> None:
        d = filedialog.askdirectory(title="Folder containing PDFs")
        if not d:
            return
        folder = Path(d)
        found = sorted(
            p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
        )
        if not found:
            messagebox.showinfo("No PDFs", f"No .pdf files found in:\n{folder}")
            return
        self._paths = _dedupe_paths(self._paths + found)
        self._sync_listbox()
        self._set_status(f"{len(self._paths)} file(s) selected (added {len(found)} from folder).")

    def _remove_selected(self) -> None:
        sel = list(self._list.curselection())
        if not sel:
            return
        for i in reversed(sel):
            del self._paths[i]
        self._sync_listbox()
        self._set_status(f"{len(self._paths)} file(s) selected.")

    def _clear(self) -> None:
        self._paths.clear()
        self._sync_listbox()
        self._set_status("List cleared.")

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save export as",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel workbook", "*.xlsx"),
                ("CSV", "*.csv"),
                ("All files", "*"),
            ],
        )
        if path:
            self._output_var.set(path)

    def _set_status(self, text: str) -> None:
        self._status.config(text=text)

    def _export(self) -> None:
        if not self._paths:
            messagebox.showwarning("No PDFs", "Add at least one PDF file or folder.")
            return
        out = Path(self._output_var.get().strip()).expanduser()
        if not str(out):
            messagebox.showwarning("Output", "Choose an output file path.")
            return
        fmt = self._fmt_var.get()
        try:
            n, written = export_pdfs(self._paths, out, fmt)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc))
            self._set_status("Export failed.")
            return
        self._set_status(f"Wrote {n} row(s) to {written}")
        messagebox.showinfo("Done", f"Wrote {n} row(s) to:\n{written}")

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    PdfExportApp().run()


if __name__ == "__main__":
    main()
