#!/usr/bin/env bash
# verify.sh — the build → render → rasterize loop for a docx.
# Usage: ./verify.sh path/to/file.docx [dpi]
# Produces an ASCII-named copy, a PDF (LibreOffice ground truth), and pg-NN.jpg page images.
set -euo pipefail

SRC="${1:?usage: verify.sh file.docx [dpi]}"
DPI="${2:-90}"
SOFFICE="${SOFFICE:-/Applications/LibreOffice.app/Contents/MacOS/soffice}"
WORK="$(mktemp -d)"

# 1. ASCII copy — em-dashes / spaces in filenames break shell globbing.
cp "$SRC" "$WORK/doc.docx"

# 2. Render to PDF. If a stale soffice lock makes this fail, just re-run once.
"$SOFFICE" --headless --convert-to pdf --outdir "$WORK" "$WORK/doc.docx" >/dev/null 2>&1 \
  || "$SOFFICE" --headless --convert-to pdf --outdir "$WORK" "$WORK/doc.docx" >/dev/null 2>&1

PAGES="$(pdfinfo "$WORK/doc.pdf" | awk '/Pages/{print $2}')"
echo "Pages: $PAGES"

# 3. Rasterize every page next to the source.
OUT="$(dirname "$SRC")/_verify"
mkdir -p "$OUT"
pdftoppm -jpeg -r "$DPI" "$WORK/doc.pdf" "$OUT/pg" >/dev/null 2>&1
cp "$WORK/doc.pdf" "$OUT/doc.pdf"
echo "Wrote: $OUT/doc.pdf and $OUT/pg-*.jpg  (read the images to inspect layout)"
rm -rf "$WORK"
