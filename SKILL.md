---
name: docx-pdf-design
description: >-
  Build polished, EDITABLE Word (.docx) documents — contracts, templates, reports — with
  docx-js that render correctly in Microsoft Word, Apple Pages, AND LibreOffice. Captures
  battle-tested fixes for real pitfalls: full-bleed cover image sizing, header/footer tab
  stops collapsing in Word, cross-reference rot after renumbering sections, orphan/widow
  lines, duplicated auto-numbers, image handling (PIL regeneration, transparency, units,
  SVG), and the build→render→verify loop. Use when creating or editing a styled .docx with
  a cover page, headers/footers, tables, fill-in fields, numbered clauses, images/logos, or
  schedules — especially long or legal documents.
origin: learned-from-session
---

# DOCX / PDF Design — Pitfalls & Fixes

Hard-won lessons from building a multi-page branded contract template with `docx-js`. The
base `docx` skill covers API syntax; THIS skill covers what actually breaks across renderers
and how to catch it. Read the base `docx` skill for syntax; read this for the traps.

Each pitfall below has a real **BEFORE / AFTER** screenshot in `assets/` (see `README.md`).
Reusable scripts live in `scripts/`; a runnable demo generator is in `examples/make_demos.js`.

## Golden rule: render with LibreOffice, never trust the build

`docx-js` producing a file ≠ the file looking right. The same `.docx` renders **differently**
in Word, Pages, and LibreOffice. Apple Pages is the worst offender (mangles headers/footers
and tables). Some bugs are renderer-specific: a literal tab looks fine in LibreOffice and
collapses in Word. So you verify with THREE different lenses, each catching a different class:

| Lens | Tool | Catches |
|------|------|---------|
| **Layout** | LibreOffice → PDF → `pdftoppm` → read .jpg | white borders, orphans, overflow, spacing |
| **Content** | `scripts/extract-text.py` | typos, stale "see Section X", missing schedules |
| **Structure** | `unzip -p ... \| grep` (`scripts/check-tabs.sh`, `check-refs.py`) | literal tabs, dangling refs |

### The build → render → verify loop

```bash
node build.js                          # 1. build the .docx
scripts/verify.sh "Nice — Name.docx"   # 2. ASCII-copy + LibreOffice PDF + pg-NN.jpg pages
#    then READ the pg-*.jpg images for layout, and:
scripts/extract-text.py doc.docx       # 3. content check
scripts/check-refs.py  doc.docx        # 4. cross-reference check
scripts/check-tabs.sh  doc.docx        # 5. header/footer tab check
```

A document is "done" only after (a) you've read the rasterized pages AND (b) the text/ref/tab
checks are clean. **Always deliver a PDF next to the .docx** — if the user only has Pages, the
PDF bypasses Pages' broken rendering and shows your true intent.

---

## Pitfall catalog

### 1. Full-bleed cover image leaves a white border  · `assets/combo_cover.jpg`
**Symptom:** A "full-page" cover floats with a uniform white margin around it.
**Cause:** `ImageRun` `transformation: { width, height }` is in **pixels @ 96 DPI**, NOT points.
`612×792` (US-Letter *points*) renders at 6.375"×8.25" — smaller than the page.
**Fix:** Use pixels. US Letter full-bleed = `816×1056` (8.5"×11" × 96). Section margins `0`.
```js
transformation: { width: 816, height: 1056 }   // NOT 612×792
```
Don't add an explicit `PageBreak` AND a section break — the section break already starts a new
page; both = a blank page. Use `SectionType.NEXT_PAGE` on the following (body) section.
→ `scripts/helpers.js` `fullBleedCover()`.

### 2. Prefer a NATIVE cover over a baked image
A cover composited in PIL/Pillow becomes one flat raster: not editable, heavy (MBs), fields
can't be typed into. For templates/forms, build the cover from real Word elements (centered
`ImageRun` logo + `TextRun` title + fill-in lines). Editable, lighter, fields work.
**Trade-off:** baked PIL image = pixel-perfect art but frozen; native = editable but limited to
Word's layout. For a fillable template, native wins. → `helpers.js` `coverField()`.

### 3. Header/footer tab stops COLLAPSE in Word  · `assets/combo_footer.jpg`
**Symptom:** Footer reads `ACME CORPDocument SubtitlePage 2 of 16` — jammed, no
spacing. **Looks fine in LibreOffice, broken in Word/Pages.**
**Cause:** Putting `"\t"` inside a `TextRun`'s **text string** emits a literal TAB char inside
`<w:t>`. LibreOffice honors it; Word collapses it to nothing.
**Fix:** Use the `Tab` element in `children` → emits a proper `<w:tab/>`:
```js
const { Tab } = require("docx");
// ❌ new TextRun({ text: "\tConfidential" })
// ✅ new TextRun({ children: [new Tab(), "Confidential — Master Template"] })
```
Define stops on the paragraph: `tabStops: [{type: CENTER, position: 4680}, {type: RIGHT, position: 9360}]`.
**Verify:** `scripts/check-tabs.sh doc.docx` — must show `<w:tab/>` elements, zero literal tabs.
**Alt approach:** a borderless 3-cell table is even more robust across renderers, but empty
cells can render as boxes and the base `docx` skill warns against tables in footers — only
reach for it if Tab() still misbehaves. → `helpers.js` `brandFooter()`.

### 4. Cross-reference rot after renumbering  · most dangerous in long docs
**Symptom:** "see Section 6.1" when royalties are now Section 8; "see Section 9" when accounting
moved to 11; `Schedule C` referenced but never attached.
**Cause:** "see Section X", "Section X.1", and schedule references are plain strings that don't
auto-update when you restructure.
**Fix:** After the structure is final, `scripts/check-refs.py doc.docx` lists every reference and
flags dangling schedules. Map each to the real target — don't trust, verify every hit.

### 5. Auto-numbering is INVISIBLE in text extraction — don't "fix" it  · `assets/combo_letters.jpg`
**Symptom:** Lettered clauses `(a)(b)(c)…` look missing in extracted text, so you add manual
letters — and the render now shows `(a) (a) Title.` (doubled).
**Cause:** List auto-numbers live in `<w:numPr>`, NOT in text runs, so extraction never shows
them. They ARE rendered.
**Fix:** Trust the `numbering` config; verify lettering by **reading the image**, not the text.
"subject to clause (e)" is already valid if the alpha list emits a–g in order.

### 6. Orphan / widow — a short item alone on its own page
**Symptom:** One 3-line clause spills onto an otherwise-empty trailing page.
**Fixes (pick one):**
- Tighten that block's `spacing.before/after/line` to reclaim ~1–2 lines (cheapest).
- `keepLines: true` / `keepNext: true` on the paragraph to stop it splitting.
- `pageBreakBefore` on the NEXT heading so the spill starts clean instead of dangling.
Re-render and re-check the page count after.

### 7. Justified inline blanks look broken  · `assets/combo_intro.jpg`
**Symptom:** A justified paragraph with inline `____` blanks stretches words across the line
("under   the   laws   of") into ugly rivers.
**Fix:** Pull fill-ins OUT of justified prose. Use a left-aligned labeled block with a
`TabStopType.LEFT` stop (e.g. a "THE SECOND PARTY: Legal name ____" block). → `helpers.js`.

### 8. Non-ASCII / em-dash filenames break the shell
`Acme — Report.docx` (em-dash, spaces) → `no matches found` / quoting failures in zsh. Work
on an ASCII copy (`doc.docx`) for all tooling; use the pretty name only for the final delivery.
`scripts/verify.sh` does this copy automatically.

---

## Images & logos — you REGENERATE, you don't pixel-edit

Honest limitation: image fixes here are **not** interactive pixel editing. You fix images two
ways: **(a) regenerate with code (Pillow/PIL)**, or **(b) swap the file inside the docx zip**.
Plan around that.

### Units are different in every layer — the #1 source of image bugs
| Layer | Unit | 1 inch |
|-------|------|--------|
| docx-js `transformation` | pixels @ 96 dpi | 96 |
| raw OOXML `<wp:extent cx/cy>` | EMU | 914400 |
| points (PDF/print) | pt | 72 |

Pitfall #1 (the white-border cover) was exactly this: points used where pixels were expected.
When in doubt, compute from inches: `px = inches*96`, `emu = inches*914400`.

### Pillow recipes (the actual "image code" that fixes things)
```python
from PIL import Image
im = Image.open(src).convert("RGBA")
im = im.crop(im.getbbox())                 # trim transparent padding around a logo
w = 640; im = im.resize((w, round(im.height*w/im.width)), Image.LANCZOS)  # crisp downscale (~2× display)
im.save("logo_trim.png")
```
- **Render at ~2× the display size** for retina crispness, then let docx scale it down.
- **Transparency matters by background:** a white-on-transparent logo is INVISIBLE on a white
  page (use it only on a dark cover); a dark/chrome logo works on white. Keep both variants and
  pick per background — this bit us when the "Normal" white logo vanished on the white body.

### Word vs docx-js image embedding
- **New doc (docx-js):** `new ImageRun({ type: "png"|"jpg"|"svg", data, transformation })` —
  `type` is REQUIRED; `transformation` is px@96dpi (see table).
- **Editing an existing .docx's image:** unzip, replace `word/media/imageN.png` with a same-named
  file. If dimensions change, also fix the drawing `<wp:extent cx cy>` (EMU) in `document.xml`,
  add a `<Default Extension=.../>` to `[Content_Types].xml` and a `<Relationship .../>` if it's
  a new file. (See the base `docx` skill "Images" XML section.)
- **SVG:** docx-js accepts `type: "svg"` but requires a raster fallback, and Word/Pages SVG
  support is uneven. Safest: **rasterize the SVG to PNG first** (`rsvg-convert`, Inkscape, or
  `cairosvg`) at 2× and embed the PNG. Don't ship SVG-only into a contract.

### Generating before/after proof images (for docs/READMEs)
You can't screenshot Word, but you CAN: build a tiny "before" docx and a "fixed" one, render
both via LibreOffice, `pdftoppm`, crop the region with Pillow, and stack them with a label
strip. That's how `assets/combo_*.jpg` were made — see `examples/make_demos.js`.
Caveat: a renderer-specific bug (like §3) won't reproduce in LibreOffice, so build the "before"
to mimic the **target renderer's** output (e.g. concatenate the footer text to show Word's jam)
and say so in the caption.

---

## Cross-renderer table & layout rules (reinforced)
- **Dual widths always:** `columnWidths` on the `Table` AND `width` on every `TableCell`, matched.
  Use `WidthType.DXA` — percentages break in Google Docs/Pages.
- **`ShadingType.CLEAR`**, never `SOLID` (SOLID renders black).
- Fill-in rows: header row + empty rows with cell `margins:{top:90,bottom:90,...}` gives writable
  height without explicit row heights. → `helpers.js` `formTable()`.
- Never use a table as a divider/rule — use a paragraph `border.bottom` instead.

## Reusable build harness shape
Keep one `build.js` that pushes paragraphs into a `children[]` array using small helpers
(`P`, `b`, `r`, `it`, `fill`, `bullet`, `SEC`, `CL`, `formTable`). Re-run end-to-end on every
change — never hand-edit the .docx. `scripts/helpers.js` ships the core set ready to import.
