#!/usr/bin/env python3
"""
Extract PDF structure tags (tagged PDF / structure tree) into tabular form.

Uses PyMuPDF with structure collection enabled and parses the semantic XHTML
per page into rows: tag hierarchy, local tag name, text, and attributes.
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


class _TagTextParser(HTMLParser):
    """Walk XHTML fragments and emit (tag_path, tag, attrs, text) for text runs."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        # (tag_name, attrs) — empty tag name means skip (e.g. script/style body)
        self._stack: list[tuple[str, dict[str, str]]] = []
        self.rows: list[tuple[str, str, dict[str, str], str]] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in SKIP_TAGS:
            self._stack.append(("", {}))
            return
        ad: dict[str, str] = {k: v if v is not None else "" for k, v in attrs}
        self._stack.append((t, ad))

    def handle_startendtag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.handle_starttag(tag, attrs)
        self.handle_endtag(tag)

    def handle_endtag(self, tag: str) -> None:
        if self._stack:
            self._stack.pop()

    def handle_data(self, data: str) -> None:
        if not self._stack:
            return
        leaf, attrs = self._stack[-1]
        if not leaf:
            return
        text = data.strip()
        if not text:
            return
        path = "/".join(name for name, _ in self._stack if name)
        self.rows.append((path, leaf, dict(attrs), text))

    def error(self, message: str) -> None:
        raise RuntimeError(message)


def _parse_xhtml_fragment(xhtml: str) -> list[tuple[str, str, dict[str, str], str]]:
    parser = _TagTextParser()
    parser.feed(xhtml)
    parser.close()
    return parser.rows


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


def extract_rows(pdf_path: Path) -> list[dict[str, str | int]]:
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
            for path, tag, attrs, text in _parse_xhtml_fragment(xhtml):
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


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="Export tagged PDF structure (tags + text) to CSV or Excel."
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
        required=True,
        help="Output file (.csv or .xlsx)",
    )
    p.add_argument(
        "--format",
        choices=("auto", "csv", "xlsx"),
        default="auto",
        help="Output format (default: from file extension)",
    )
    args = p.parse_args(argv)

    pdf_list = list(_iter_pdf_paths(list(args.pdfs), args.from_dir))
    if not pdf_list:
        print("No PDF files found.", file=sys.stderr)
        return 2

    all_rows: list[dict[str, str | int]] = []
    for pdf in pdf_list:
        all_rows.extend(extract_rows(pdf))

    out = args.output.expanduser()
    fmt = args.format
    if fmt == "auto":
        suf = out.suffix.lower()
        if suf == ".csv":
            fmt = "csv"
        elif suf in (".xlsx", ".xlsm"):
            fmt = "xlsx"
        else:
            fmt = "xlsx"
            out = out.with_suffix(".xlsx")

    out.parent.mkdir(parents=True, exist_ok=True)
    if fmt == "csv":
        if out.suffix.lower() != ".csv":
            out = out.with_suffix(".csv")
        _write_csv(all_rows, out)
    else:
        if out.suffix.lower() not in (".xlsx", ".xlsm"):
            out = out.with_suffix(".xlsx")
        _write_xlsx(all_rows, out)

    print(f"Wrote {len(all_rows)} row(s) to {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
