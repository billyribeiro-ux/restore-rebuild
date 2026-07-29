#!/usr/bin/env bash
# Build the manuscript: Markdown -> DOCX -> PDF, plus a concatenated Markdown.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

python3 -c "import docx" 2>/dev/null || pip3 install --quiet python-docx

python3 build/build.py

# Georgia -> Bitstream Charter substitution so the PDF matches the DOCX design.
export FONTCONFIG_FILE="$ROOT/build/fonts.conf"

soffice --headless --norestore \
        --convert-to pdf:writer_pdf_Export \
        --outdir dist dist/restore-and-rebuild.docx >/dev/null

echo "  wrote dist/restore-and-rebuild.pdf"
echo
ls -1sh dist/
