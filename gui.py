#!/usr/bin/env python3
"""
Desktop GUI: pick a folder whose PDFs to export (folder can change each time), optional extras.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from extract_tagged_pdf import export_pdfs


def _pdfs_in_folder(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted(
        p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"
    )


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
        self._folder_var = tk.StringVar(value="")

        main = ttk.Frame(self.root, padding=12)
        main.grid(row=0, column=0, sticky="nsew")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main.columnconfigure(0, weight=1)
        main.rowconfigure(3, weight=1)

        folder_row = ttk.Frame(main)
        folder_row.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        folder_row.columnconfigure(1, weight=1)
        ttk.Label(folder_row, text="Folder with PDFs").grid(row=0, column=0, sticky="w", padx=(0, 8))
        folder_entry = ttk.Entry(folder_row, textvariable=self._folder_var)
        folder_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(folder_row, text="Choose folder…", command=self._choose_pdf_folder).grid(
            row=0, column=2, padx=(8, 0)
        )
        ttk.Label(
            main,
            text="All .pdf files in that folder are loaded. Pick another folder anytime.",
            font=("TkDefaultFont", 10),
            foreground="gray30",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(main, text="PDFs to export").grid(row=2, column=0, sticky="w")

        list_frame = ttk.Frame(main)
        list_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(4, 8))
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
        btn_row.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        ttk.Button(btn_row, text="Add files…", command=self._add_files).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Add PDFs from another folder…", command=self._append_folder).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Remove selected", command=self._remove_selected).pack(
            side=tk.LEFT, padx=(0, 6)
        )
        ttk.Button(btn_row, text="Clear all", command=self._clear).pack(side=tk.LEFT)

        out_row = ttk.Frame(main)
        out_row.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, 8))
        out_row.columnconfigure(1, weight=1)
        ttk.Label(out_row, text="Output file").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self._output_var = tk.StringVar(value=str(Path.home() / "tagged_export.xlsx"))
        out_entry = ttk.Entry(out_row, textvariable=self._output_var)
        out_entry.grid(row=0, column=1, sticky="ew")
        ttk.Button(out_row, text="Browse…", command=self._browse_output).grid(
            row=0, column=2, padx=(8, 0)
        )

        fmt_row = ttk.Frame(main)
        fmt_row.grid(row=6, column=0, columnspan=2, sticky="w", pady=(0, 8))
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

        ttk.Button(main, text="Export", command=self._export).grid(row=7, column=0, sticky="w")
        self._status = ttk.Label(
            main,
            text="Choose the folder that contains your PDFs, set the output file, then Export.",
        )
        self._status.grid(row=8, column=0, columnspan=2, sticky="w", pady=(12, 0))

        folder_entry.bind("<Return>", lambda e: self._load_folder_from_entry())

    def _sync_listbox(self) -> None:
        self._list.delete(0, tk.END)
        for p in self._paths:
            self._list.insert(tk.END, str(p))

    def _choose_pdf_folder(self) -> None:
        initial = self._folder_var.get().strip()
        kwargs: dict = {"title": "Select folder containing PDFs"}
        if initial:
            p = Path(initial).expanduser()
            if p.is_dir():
                kwargs["initialdir"] = str(p)
        d = filedialog.askdirectory(**kwargs)
        if not d:
            return
        self._apply_folder(Path(d))

    def _load_folder_from_entry(self) -> None:
        raw = self._folder_var.get().strip()
        if not raw:
            return
        p = Path(raw).expanduser()
        if not p.is_dir():
            messagebox.showwarning("Not a folder", f"Folder not found:\n{p}")
            return
        self._apply_folder(p)

    def _apply_folder(self, folder: Path) -> None:
        try:
            folder = folder.expanduser().resolve()
        except OSError as exc:
            messagebox.showerror("Folder", str(exc))
            return
        self._folder_var.set(str(folder))
        found = _pdfs_in_folder(folder)
        if not found:
            messagebox.showinfo("No PDFs", f"No .pdf files in:\n{folder}")
            self._paths.clear()
            self._sync_listbox()
            self._set_status("That folder has no PDF files.")
            return
        self._paths = found
        self._sync_listbox()
        self._set_status(f"Loaded {len(found)} PDF(s) from folder.")

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

    def _append_folder(self) -> None:
        d = filedialog.askdirectory(title="Add PDFs from another folder")
        if not d:
            return
        folder = Path(d)
        found = _pdfs_in_folder(folder)
        if not found:
            messagebox.showinfo("No PDFs", f"No .pdf files found in:\n{folder}")
            return
        self._paths = _dedupe_paths(self._paths + found)
        self._sync_listbox()
        self._set_status(f"{len(self._paths)} file(s) total (added {len(found)} from second folder).")

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
        self._folder_var.set("")
        self._sync_listbox()
        self._set_status("Cleared folder and file list.")

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
            messagebox.showwarning(
                "No PDFs",
                "Choose a folder that contains PDFs (or add files / a second folder).",
            )
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
