#!/usr/bin/env python3
"""check-refs.py — surface every internal cross-reference in a .docx so you can
verify each one against the ACTUAL structure after renumbering.

Catches "cross-reference rot": "see Section 6.1" left behind when royalties moved
to Section 8, "Schedule C" referenced but never attached, etc.

Usage: python3 check-refs.py file.docx
"""
import sys, zipfile, re, html
from collections import Counter

def body(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    return html.unescape(re.sub(r"<[^>]+>", "", xml))

def main(path):
    t = body(path)

    secs = Counter(re.findall(r"Section[s]?\s+\d+(?:\.\d+)?", t))
    print("Section references (verify each points to the right place):")
    for s, c in sorted(secs.items(), key=lambda kv: [int(x) for x in re.findall(r"\d+", kv[0])]):
        print(f"  {c:>2}x  {s}")

    # Schedules referenced vs. schedules that actually exist as headings.
    referenced = set(re.findall(r"Schedule [A-Z]", t))
    headings = set(re.findall(r"SCHEDULE [A-Z]", t))
    print("\nSchedules referenced:", ", ".join(sorted(referenced)) or "(none)")
    print("Schedules present   :", ", ".join(sorted(h.title() for h in headings)) or "(none)")
    dangling = {r for r in referenced} - {h.title() for h in headings}
    if dangling:
        print("  ⚠ DANGLING (referenced but no heading found):", ", ".join(sorted(dangling)))

    # Clause-letter references like "clause (e)" — make sure the list actually emits that letter.
    clause = sorted(set(re.findall(r"clause \(([a-z])\)", t)))
    if clause:
        print("\nClause-letter references:", ", ".join(f"({c})" for c in clause),
              "\n  (auto-numbers are invisible to text extraction — confirm the lettered list in a render)")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("usage: check-refs.py file.docx")
    main(sys.argv[1])
