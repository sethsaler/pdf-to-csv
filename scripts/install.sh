#!/usr/bin/env bash
# Install dependencies into a local virtual environment.
#
# From a clone of this repository:
#   bash scripts/install.sh
#
# One-liner (after replacing ORG, REPO, and BRANCH with your Git remote):
#   curl -fsSL "https://raw.githubusercontent.com/ORG/REPO/BRANCH/scripts/install.sh" | bash
#
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

if ! command -v python3 >/dev/null 2>&1; then
  echo "python3 is required but not found in PATH." >&2
  exit 1
fi

USE_VENV=1
if [[ ! -d .venv ]]; then
  if python3 -m venv .venv 2>/dev/null; then
    :
  else
    rm -rf .venv
    USE_VENV=0
    echo "Note: could not create a venv (install python3-venv if you want one)." >&2
    echo "      Installing packages with: python3 -m pip install --user -r requirements.txt" >&2
    python3 -m pip install --user -r requirements.txt
  fi
fi

if [[ "$USE_VENV" -eq 1 ]]; then
  # shellcheck source=/dev/null
  source .venv/bin/activate
  python -m pip install --upgrade pip
  python -m pip install -r requirements.txt
  echo "Done. Activate the environment with:"
  echo "  source \"${ROOT}/.venv/bin/activate\""
  echo "Then run:"
  echo "  python extract_tagged_pdf.py --from-dir ./pdfs -o ./out.xlsx"
else
  echo "Done. Run with:"
  echo "  python3 extract_tagged_pdf.py --from-dir ./pdfs -o ./out.xlsx"
fi
