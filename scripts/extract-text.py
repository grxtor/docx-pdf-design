#!/usr/bin/env python3
"""extract-text.py — dump a .docx body as plain text (one line per paragraph).

For CONTENT checks: cross-references, typos, stale section numbers.
Note: list auto-numbers (e.g. "(a)(b)(c)") live in <w:numPr>, NOT in text runs,
so they will NOT appear here. Verify lettering/numbering in a rendered image instead.

Usage: python3 extract-text.py file.docx
"""
import sys, zipfile, re, html

def main(path):
    with zipfile.ZipFile(path) as z:
        names = [n for n in z.namelist() if re.match(r"word/(document|header\d+|footer\d+)\.xml$", n)]
        for n in sorted(names):
            xml = z.read(n).decode("utf-8")
            txt = re.sub(r"</w:p>", "\n", xml)
            txt = html.unescape(re.sub(r"<[^>]+>", "", txt))
            lines = [l.strip() for l in txt.split("\n") if l.strip()]
            if lines:
                print(f"\n===== {n} =====")
                print("\n".join(lines))

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("usage: extract-text.py file.docx")
    main(sys.argv[1])
