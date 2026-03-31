#!/usr/bin/env python3
"""
Extract PDF structure tags (tagged PDF / structure tree) into tabular form.

Uses PyMuPDF with structure collection enabled and parses the semantic XHTML
per page into rows: tag hierarchy, local tag name, text, and attributes.

Text comes from the PDF’s embedded character data only (no OCR). Scanned
pages that are image-only have no extractable text unless you OCR them elsewhere.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections.abc import Iterator
from html.parser import HTMLParser
from pathlib import Path

import fitz


STRUCTURE_FLAGS = fitz.TEXTFLAGS_XHTML | fitz.TEXT_COLLECT_STRUCTURE

SKIP_TAGS = frozenset({"script", "style"})

# Block-level structure tags: a new export "paragraph" / logical block starts here.
# Nested spans inside the same block share one id (see _TagTextParser).
PARAGRAPH_BLOCK_TAGS = frozenset(
    {
        "p",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "th",
        "td",
        "li",
        "lbody",
        "caption",
        "blockquote",
        "note",
        "toci",
    }
)

# Standard PDF 1.7 structure tags (ISO 32000-1) - content must be within these
PDF_STRUCTURE_TAGS = frozenset(
    {
        # Root element
        "document",
        # Container elements
        "part",
        "sect",
        "art",
        "div",
        # Paragraph and heading elements
        "p",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        # List elements
        "l",
        "li",
        "lbl",
        "lbody",
        # Table elements
        "table",
        "thead",
        "tbody",
        "tfoot",
        "tr",
        "th",
        "td",
        # Special inline elements
        "span",
        "quote",
        "code",
        "link",
        "annot",
        "form",
        "ruby",
        "rb",
        "rt",
        "rp",
        # Special block elements
        "figure",
        "formula",
        "index",
        "toc",
        "toci",
        "caption",
        "blockquote",
        # Reference and note elements
        "bibentry",
        "reference",
        "note",
    }
)


class _TagTextParser(HTMLParser):
    """Walk XHTML fragments and emit rows for text runs."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        # (tag_name, attrs, block_uid) — empty tag name means skip (e.g. script/style)
        self._stack: list[tuple[str, dict[str, str], int]] = []
        self._next_block_uid = 1
        self.rows: list[tuple[str, str, dict[str, str], str, int]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in SKIP_TAGS:
            self._stack.append(("", {}, 0))
            return
        ad: dict[str, str] = {k: v if v is not None else "" for k, v in attrs}
        parent_uid = self._stack[-1][2] if self._stack else 0
        if t in PARAGRAPH_BLOCK_TAGS:
            uid = self._next_block_uid
            self._next_block_uid += 1
        else:
            uid = parent_uid
        self._stack.append((t, ad, uid))

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        if self._stack:
            self._stack.pop()

    def handle_data(self, data: str) -> None:
        if not self._stack:
            return
        leaf, attrs, block_uid = self._stack[-1]
        if not leaf:
            return
        text = data.strip()
        if not text:
            return
        # Build path and check if content is part of semantic structure
        path_tags = [name for name, _, _ in self._stack if name]
        path = "/".join(path_tags)
        # Only include text if it's part of the PDF structure tree
        # (i.e., at least one tag in the stack is a standard PDF structure tag)
        if not any(tag in PDF_STRUCTURE_TAGS for tag in path_tags):
            return
        self.rows.append((path, leaf, dict(attrs), text, block_uid))

    def error(self, message: str) -> None:
        raise RuntimeError(message)


def _parse_xhtml_fragment(xhtml: str) -> list[tuple[str, str, dict[str, str], str, int]]:
    parser = _TagTextParser()
    parser.feed(xhtml)
    parser.close()
    return parser.rows


def _document_has_struct_tree(doc: fitz.Document) -> bool:
    """True if the PDF catalog has /StructTreeRoot (real tagging, e.g. from Acrobat)."""
    if not doc.is_pdf or doc.is_closed:
        return False
    try:
        cat = doc.pdf_catalog()
        t, val = doc.xref_get_key(cat, "StructTreeRoot")
    except Exception:
        return False
    return t != "null" and val != "null"


def _merge_paragraph_rows(
    rows: list[dict[str, str | int]],
) -> list[dict[str, str | int]]:
    """Join consecutive text runs that belong to the same logical block (same paragraph_index)."""
    if not rows:
        return rows
    keyfn = lambda r: (r.get("page", 0), r.get("paragraph_index", 0))
    merged: list[dict[str, str | int]] = []
    group_key = keyfn(rows[0])
    chunk: list[dict[str, str | int]] = [rows[0]]
    for r in rows[1:]:
        k = keyfn(r)
        if k == group_key and r.get("error") == "" and chunk[-1].get("error") == "":
            chunk.append(r)
        else:
            merged.append(_combine_chunk(chunk))
            chunk = [r]
            group_key = k
    merged.append(_combine_chunk(chunk))
    return merged


def _combine_chunk(chunk: list[dict[str, str | int]]) -> dict[str, str | int]:
    if len(chunk) == 1:
        base = dict(chunk[0])
        base.pop("paragraph_index", None)
        return base
    texts: list[str] = []
    attrs_parts: list[dict[str, str]] = []
    for r in chunk:
        t = str(r.get("text", "")).strip()
        if t:
            texts.append(t)
        aj = r.get("attributes_json", "")
        if aj:
            try:
                attrs_parts.append(json.loads(str(aj)))
            except json.JSONDecodeError:
                pass
    merged_attrs: dict[str, str] = {}
    for d in attrs_parts:
        merged_attrs.update(d)
    base = dict(chunk[0])
    base["text"] = " ".join(texts)
    base["attributes_json"] = (
        json.dumps(merged_attrs, ensure_ascii=False) if merged_attrs else ""
    )
    base.pop("paragraph_index", None)
    return base


def _iter_pdf_paths(paths: list[Path], from_dir: Path | None) -> Iterator[Path]:
    seen: set[Path] = set()
    if from_dir is not None:
        for p in sorted(from_dir.glob("*.pdf")):
            rp = p.resolve()
            if rp not in seen:
                seen.add(rp)
                yield p
    for raw in paths:
        p = raw.expanduser()
        if p.is_dir():
            for child in sorted(p.glob("*.pdf")):
                rc = child.resolve()
                if rc not in seen:
                    seen.add(rc)
                    yield child
        elif p.is_file() and p.suffix.lower() == ".pdf":
            rp = p.resolve()
            if rp not in seen:
                seen.add(rp)
                yield p


def extract_rows(
    pdf_path: Path,
    *,
    include_layout_paragraphs: bool = False,
) -> list[dict[str, str | int]]:
    out: list[dict[str, str | int]] = []
    try:
        doc = fitz.open(pdf_path)
    except Exception as exc:
        out.append(
            {
                "source_pdf": str(pdf_path),
                "page": 0,
                "tag_path": "",
                "tag": "",
                "text": "",
                "attributes_json": "",
                "error": str(exc),
            }
        )
        return out

    try:
        has_struct = _document_has_struct_tree(doc)
        if not has_struct and not include_layout_paragraphs:
            out.append(
                {
                    "source_pdf": str(pdf_path.resolve()),
                    "page": 0,
                    "tag_path": "",
                    "tag": "",
                    "text": "",
                    "attributes_json": "",
                    "error": (
                        "No PDF structure tree (/StructTreeRoot): this file is not tagged "
                        "in Acrobat (or tags were removed). Nothing exported. "
                        "Re-save with tags from Acrobat, or pass "
                        "--include-layout-paragraphs to export MuPDF layout blocks "
                        "(not the same as Acrobat metadata)."
                    ),
                }
            )
            return out

        for i in range(doc.page_count):
            page = doc[i]
            try:
                xhtml = page.get_text("xhtml", flags=STRUCTURE_FLAGS)
            except Exception as exc:
                out.append(
                    {
                        "source_pdf": str(pdf_path),
                        "page": i + 1,
                        "tag_path": "",
                        "tag": "",
                        "text": "",
                        "attributes_json": "",
                        "error": str(exc),
                    }
                )
                continue
            for path, tag, attrs, text, block_uid in _parse_xhtml_fragment(xhtml):
                out.append(
                    {
                        "source_pdf": str(pdf_path.resolve()),
                        "page": i + 1,
                        "tag_path": path,
                        "tag": tag,
                        "text": text,
                        "attributes_json": json.dumps(attrs, ensure_ascii=False)
                        if attrs
                        else "",
                        "paragraph_index": block_uid,
                        "error": "",
                    }
                )
    finally:
        doc.close()

    return out


def _write_csv(rows: list[dict[str, str | int]], path: Path) -> None:
    fieldnames = [
        "source_pdf",
        "page",
        "tag_path",
        "tag",
        "text",
        "attributes_json",
        "error",
    ]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})


def _write_xlsx(rows: list[dict[str, str | int]], path: Path) -> None:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "tagged_content"
    headers = [
        "source_pdf",
        "page",
        "tag_path",
        "tag",
        "text",
        "attributes_json",
        "error",
    ]
    ws.append(headers)
    for r in rows:
        ws.append([r.get(h, "") for h in headers])
    wb.save(path)


def _resolve_output_format(output: Path, fmt: str) -> tuple[Path, str]:
    """Normalize output path suffix and return (path, 'csv' | 'xlsx')."""
    out = output.expanduser()
    if fmt == "auto":
        suf = out.suffix.lower()
        if suf == ".csv":
            resolved_fmt = "csv"
        elif suf in (".xlsx", ".xlsm"):
            resolved_fmt = "xlsx"
        else:
            resolved_fmt = "xlsx"
            out = out.with_suffix(".xlsx")
    elif fmt == "csv":
        resolved_fmt = "csv"
        if out.suffix.lower() != ".csv":
            out = out.with_suffix(".csv")
    else:
        resolved_fmt = "xlsx"
        if out.suffix.lower() not in (".xlsx", ".xlsm"):
            out = out.with_suffix(".xlsx")
    return out, resolved_fmt


def export_pdfs(
    pdf_paths: list[Path],
    output: Path,
    fmt: str = "auto",
    *,
    include_layout_paragraphs: bool = False,
    paragraph_rows: bool = False,
) -> tuple[int, Path]:
    """
    Extract tagged structure from PDFs and write CSV or Excel.

    By default only PDFs with a real /StructTreeRoot (Acrobat-style tagging) are
    exported. Set include_layout_paragraphs=True to also export MuPDF's synthetic
    layout blocks for untagged files.

    Returns (row_count, resolved_output_path).
    """
    out, kind = _resolve_output_format(output, fmt)
    out.parent.mkdir(parents=True, exist_ok=True)

    all_rows: list[dict[str, str | int]] = []
    for pdf in pdf_paths:
        all_rows.extend(
            extract_rows(pdf, include_layout_paragraphs=include_layout_paragraphs)
        )

    if paragraph_rows:
        all_rows = _merge_paragraph_rows(all_rows)
    else:
        for r in all_rows:
            r.pop("paragraph_index", None)

    if kind == "csv":
        _write_csv(all_rows, out)
    else:
        _write_xlsx(all_rows, out)

    return len(all_rows), out


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Export tagged PDF structure (tags + text) to CSV or Excel."
    )
    p.add_argument(
        "--gui",
        action="store_true",
        help="Open a window to pick PDFs/folders and an output file.",
    )
    p.add_argument(
        "pdfs",
        nargs="*",
        type=Path,
        help="PDF files or directories containing PDFs",
    )
    p.add_argument(
        "--from-dir",
        type=Path,
        metavar="DIR",
        help="Shorthand: include every DIR/*.pdf",
    )
    p.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output file (.csv or .xlsx). Not required with --gui.",
    )
    p.add_argument(
        "--format",
        choices=("auto", "csv", "xlsx"),
        default="auto",
        help="Output format (default: from file extension)",
    )
    p.add_argument(
        "--include-layout-paragraphs",
        action="store_true",
        help=(
            "Also export untagged PDFs using MuPDF layout <p> blocks (not Acrobat "
            "metadata). Default: only export when /StructTreeRoot exists."
        ),
    )
    p.add_argument(
        "--paragraph-rows",
        action="store_true",
        help="Merge text runs that belong to the same logical block into one row.",
    )
    args = p.parse_args(argv)

    if args.gui:
        from gui import main as gui_main

        gui_main()
        return 0

    if args.output is None:
        p.error("the following arguments are required: -o/--output")

    pdf_list = list(_iter_pdf_paths(list(args.pdfs), args.from_dir))
    if not pdf_list:
        print("No PDF files found.", file=sys.stderr)
        return 2

    n, out = export_pdfs(
        pdf_list,
        args.output,
        args.format,
        include_layout_paragraphs=args.include_layout_paragraphs,
        paragraph_rows=args.paragraph_rows,
    )
    print(f"Wrote {n} row(s) to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
