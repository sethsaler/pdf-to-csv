#!/bin/bash
# Double-click this file in Finder to install (if needed) and run the extractor.
# Place PDFs in a folder named "pdfs" next to this script, or edit PDF_DIR below.

set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

PDF_DIR="${PDF_DIR:-$ROOT/pdfs}"
OUT="${OUT:-$ROOT/tagged_export.xlsx}"

if [[ ! -d .venv ]]; then
  bash "$ROOT/scripts/install.sh"
fi
PY=python3
if [[ -f "$ROOT/.venv/bin/python" ]]; then
  PY="$ROOT/.venv/bin/python"
fi

mkdir -p "$PDF_DIR"
"$PY" "$ROOT/extract_tagged_pdf.py" --from-dir "$PDF_DIR" -o "$OUT" --format xlsx
echo "Output: $OUT"
echo "Press Enter to close."
read -r _
