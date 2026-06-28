// helpers.js — reusable docx-js building blocks that render correctly across
// Word, Pages, and LibreOffice. Import these into a build.js. Every pattern here
// encodes a fix from SKILL.md (full-bleed cover, Tab() footers, dual-width tables).
//
//   const H = require("./helpers");
//   const doc = new (require("docx").Document)({ sections: [...] });
//
// US Letter, 1" margins → content width 9360 DXA. Adjust for other page sizes.

const {
  Paragraph, TextRun, Table, TableRow, TableCell, Header, Footer, AlignmentType,
  ImageRun, BorderStyle, WidthType, ShadingType, PageNumber, TabStopType, Tab,
  SectionType, LevelFormat,
} = require("docx");

// ---- palette (tweak per brand) ----
const INK = "16181C", BODY = "20232A", SILVER = "8A8F99", RULE = "B8BCC4";
const PAGE_US = { width: 12240, height: 15840 };       // US Letter, DXA
const CONTENT_W = 9360;                                // 12240 - 2*1440

// ---- run shorthands ----
const r  = (t, o = {}) => new TextRun({ text: t, font: "Arial", size: 20, color: BODY, ...o });
const b  = (t) => new TextRun({ text: t, font: "Arial", size: 20, color: INK, bold: true });
const it = (t) => new TextRun({ text: t, font: "Arial", size: 20, color: BODY, italics: true });
// A fill-in blank that does NOT stretch justified lines — keep these in left-aligned blocks.
const fill = (w = "____________________") => new TextRun({ text: w, font: "Arial", size: 20, color: "9A9EA6" });

// ---- paragraph ----
function P(body, opts = {}) {
  const kids = Array.isArray(body) ? body : [r(body)];
  return new Paragraph({
    spacing: { before: 40, after: 110, line: 264 },
    alignment: opts.align || AlignmentType.JUSTIFIED,
    indent: opts.indent || { left: 360 },
    children: kids,
  });
}

// ---- FULL-BLEED COVER IMAGE ----
// transformation is PIXELS @ 96dpi, NOT points. US Letter full bleed = 816 x 1056.
// Put this in its OWN section with margin:0 and DO NOT add a PageBreak (the next
// section's SectionType.NEXT_PAGE already starts a new page — a PageBreak too = blank page).
function fullBleedCover(pngBuffer) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
    children: [new ImageRun({ type: "png", data: pngBuffer,
      transformation: { width: 816, height: 1056 },
      altText: { title: "Cover", description: "Cover", name: "cover" } })],
  });
}

// ---- NATIVE (editable) COVER FIELD — prefer over a baked PIL image for templates ----
function coverField(label, stop = 2160) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 140 },
    children: [
      b(label + "   "),
      new TextRun({ text: "______________________________", font: "Arial", size: 20, color: "9A9EA6" }),
    ],
  });
}

// ---- HEADER with logo + right-aligned title. Uses Tab() (renders in Word). ----
function brandHeader(logoBuffer, titleText) {
  return new Header({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
    spacing: { after: 40 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
    children: [
      new ImageRun({ type: "png", data: logoBuffer, transformation: { width: 74, height: 38 },
        altText: { title: "logo", description: "logo", name: "logo" } }),
      new TextRun({ children: [new Tab(), titleText], font: "Arial", size: 14, color: SILVER, characterSpacing: 20 }),
    ],
  })] });
}

// ---- FOOTER left / center / right. CRITICAL: new Tab(), never "\t" in text. ----
function brandFooter(leftText, centerText) {
  return new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.CENTER, position: CONTENT_W / 2 }, { type: TabStopType.RIGHT, position: CONTENT_W }],
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } },
    children: [
      new TextRun({ text: leftText, font: "Arial", size: 14, color: SILVER, characterSpacing: 14 }),
      new TextRun({ children: [new Tab(), centerText], font: "Arial", size: 14, color: SILVER }),
      new TextRun({ children: [new Tab(), "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: "Arial", size: 14, color: SILVER }),
    ],
  })] });
}

// ---- fill-in form table (Schedule rosters etc.). Dual widths + ShadingType.CLEAR. ----
function formTable(headers, widths, emptyRows) {
  const bd = { style: BorderStyle.SINGLE, size: 2, color: RULE };
  const cell = (content, w, head) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: { top: bd, bottom: bd, left: bd, right: bd },
    shading: head ? { fill: "EEEFF2", type: ShadingType.CLEAR } : undefined,
    margins: { top: 90, bottom: 90, left: 110, right: 110 }, // gives writable row height
    children: [new Paragraph({ children: [new TextRun({ text: content || "", bold: !!head, font: "Arial", size: 17, color: head ? INK : BODY })] })],
  });
  const rows = [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) })];
  for (let n = 0; n < emptyRows; n++) rows.push(new TableRow({ children: widths.map((w) => cell("", w, false)) }));
  return new Table({ width: { size: widths.reduce((a, c) => a + c, 0), type: WidthType.DXA }, columnWidths: widths, rows });
}

module.exports = {
  INK, BODY, SILVER, RULE, PAGE_US, CONTENT_W,
  r, b, it, fill, P, fullBleedCover, coverField, brandHeader, brandFooter, formTable,
};
