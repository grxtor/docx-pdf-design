#!/usr/bin/env bash
# check-tabs.sh — detect the "tab stops collapse in Word" bug.
# A correct header/footer with left/center/right layout must use <w:tab/> ELEMENTS.
# A literal TAB character inside <w:t> renders fine in LibreOffice but COLLAPSES in Word.
# Usage: ./check-tabs.sh file.docx
# (Uses python3 for matching so it works with BSD/macOS grep, which lacks -P.)
set -euo pipefail
SRC="${1:?usage: check-tabs.sh file.docx}"

for part in word/header1.xml word/header2.xml word/header3.xml word/footer1.xml word/footer2.xml word/footer3.xml; do
  xml="$(unzip -p "$SRC" "$part" 2>/dev/null || true)"
  [ -z "$xml" ] && continue
  PART="$part" XMLDATA="$xml" python3 <<'PY'
import os, re
xml = os.environ["XMLDATA"]; part = os.environ["PART"]
tab_elems = len(re.findall(r"<w:tab\s*/>", xml))
literal   = len(re.findall(r"<w:t[^>]*>[^<]*\t[^<]*</w:t>", xml))
print(f"{part} : <w:tab/>={tab_elems}  literal-tab-in-<w:t>={literal}")
if literal:
    print('  WARN literal tab found — will COLLAPSE in Word. Use new Tab() in children, not "\\t" in text.')
PY
done
