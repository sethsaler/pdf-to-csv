#!/bin/bash
# Double-click in Finder: install dependencies if needed, then open the PDF folder GUI.
#
# Unattended batch (old behavior): export every PDF in ./pdfs to ./tagged_export.xlsx
#   BATCH=1 "/path/to/Run PDF tag extraction.command"

set -euo pipefail
ROOT="$(cd "$(dirname "$0")" && pwd)"
cd "$ROOT"

if [[ ! -d .venv ]]; then
  bash "$ROOT/scripts/install.sh"
fi
PY=python3
if [[ -f "$ROOT/.venv/bin/python" ]]; then
  PY="$ROOT/.venv/bin/python"
fi

if [[ "${BATCH:-}" == "1" ]]; then
  PDF_DIR="${PDF_DIR:-$ROOT/pdfs}"
  OUT="${OUT:-$ROOT/tagged_export.xlsx}"
  mkdir -p "$PDF_DIR"
  set +e
  # Keep batch exports working for folders that mix tagged and plain PDFs.
  "$PY" "$ROOT/extract_tagged_pdf.py" --from-dir "$PDF_DIR" -o "$OUT" --format xlsx \
    --include-layout-paragraphs
  ec=$?
  set -e
  if [[ $ec -eq 0 ]]; then
    echo "Output: $OUT"
  else
    echo "extract_tagged_pdf.py exited with status $ec." >&2
  fi
else
  set +e
  "$PY" "$ROOT/extract_tagged_pdf.py" --gui
  ec=$?
  set -e
  if [[ $ec -ne 0 ]]; then
    echo "Application exited with status $ec." >&2
  fi
fi

echo ""
echo "Press Enter to close."
read -r _
