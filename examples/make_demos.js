// Generates BEFORE/AFTER mini-docx files for each pitfall in the docx-pdf-design skill.
// All content is GENERIC sample data — no real document content.
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType, ImageRun,
  BorderStyle, PageNumber, TabStopType, SectionType, Tab, LevelFormat,
} = require("docx");

const INK = "16181C", BODY = "20232A", SILVER = "8A8F99", RULE = "B8BCC4";
const cover = fs.readFileSync("cover_demo.png");
const US = { width: 12240, height: 15840 };
const save = (name, doc) => Packer.toBuffer(doc).then((b) => { fs.writeFileSync(name, b); console.log("WROTE", name); });

// ---------- 1. COVER full-bleed ----------
const coverDoc = (px) => new Document({ sections: [{
  properties: { page: { size: US, margin: { top: 0, right: 0, bottom: 0, left: 0 } } },
  children: [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
    children: [new ImageRun({ type: "png", data: cover, transformation: px,
      altText: { title: "c", description: "c", name: "c" } })] })],
}] });
const coverBefore = coverDoc({ width: 612, height: 792 });   // points-as-pixels bug -> white border
const coverAfter  = coverDoc({ width: 816, height: 1056 });  // true 96dpi full bleed

// ---------- 2. FOOTER tab stops ----------
function footerDoc(useTabElement) {
  const runs = useTabElement ? [
    new TextRun({ text: "ACME CORP", font: "Arial", size: 16, color: SILVER }),
    new TextRun({ children: [new Tab(), "Document Subtitle"], font: "Arial", size: 16, color: SILVER }),
    new TextRun({ children: [new Tab(), "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: SILVER }),
  ] : [
    new TextRun({ text: "ACME CORP", font: "Arial", size: 16, color: SILVER }),
    new TextRun({ text: "\tDocument Subtitle", font: "Arial", size: 16, color: SILVER }),
    new TextRun({ children: ["\t", "Page ", PageNumber.CURRENT, " of ", PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: SILVER }),
  ];
  const f = new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.CENTER, position: 4680 }, { type: TabStopType.RIGHT, position: 9360 }],
    border: { top: { style: BorderStyle.SINGLE, size: 6, color: RULE, space: 6 } }, children: runs,
  })] });
  return new Document({ sections: [{
    properties: { page: { size: US, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    footers: { default: f },
    children: [new Paragraph({ children: [new TextRun({ text: "Body text. The footer below uses a three-part left / center / right layout.", font: "Arial", size: 22, color: BODY })] })],
  }] });
}

// ---------- 3. INTRO inline blanks vs clean block ----------
const fill = (w) => new TextRun({ text: w, font: "Arial", size: 22, color: "9A9EA6" });
const r = (t, o = {}) => new TextRun({ text: t, font: "Arial", size: 22, color: BODY, ...o });
const bld = (t) => new TextRun({ text: t, font: "Arial", size: 22, color: INK, bold: true });
function introDoc(after) {
  const kids = [];
  if (!after) {
    kids.push(new Paragraph({ alignment: AlignmentType.JUSTIFIED, spacing: { line: 300 }, children: [
      bld("This DOCUMENT"), r(" is made between "), bld("Acme Corp"),
      r(", a company of "), fill("________________"), r(" (the "), bld("First Party"), r("), and "), fill("________________________________"),
      r(", a "), fill("____________________"), r(" of "), fill("__________________"),
      r(", with address at "), fill("________________________________________"),
      r(" (the "), bld("Second Party"), r("), effective on the date signed below."),
    ] }));
  } else {
    kids.push(new Paragraph({ alignment: AlignmentType.LEFT, spacing: { line: 300, after: 160 }, children: [
      bld("This DOCUMENT"), r(" is made between "), bld("Acme Corp"),
      r(", a company of Country (the "), bld("First Party"), r("), and the party identified immediately below (the "),
      bld("Second Party"), r("), effective on the date signed below."),
    ] }));
    kids.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: "THE SECOND PARTY", bold: true, font: "Arial", size: 18, color: SILVER, characterSpacing: 30 })] }));
    for (const lab of ["Legal name", "Entity type", "Jurisdiction", "Address"]) {
      kids.push(new Paragraph({ spacing: { after: 70 }, tabStops: [{ type: TabStopType.LEFT, position: 2160 }], children: [
        new TextRun({ text: lab, bold: true, font: "Arial", size: 22, color: INK }),
        new TextRun({ children: [new Tab()] }),
        new TextRun({ text: "______________________________________________", font: "Arial", size: 22, color: "9A9EA6" }),
      ] }));
    }
  }
  return new Document({ sections: [{ properties: { page: { size: US, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: kids }] });
}

// ---------- 4. Double letter (a)(a) ----------
function lettersDoc(manual) {
  const num = { config: [{ reference: "al", levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: "(%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] };
  const items = ["First provision", "Second provision", "Third provision"];
  const kids = items.map((t, i) => new Paragraph({ numbering: { reference: "al", level: 0 }, spacing: { after: 80 }, children: [
    bld((manual ? `(${String.fromCharCode(97 + i)}) ` : "") + t + ". "), r("Body text referenced elsewhere as “clause (b)”."),
  ] }));
  return new Document({ numbering: num, sections: [{ properties: { page: { size: US, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: kids }] });
}

Promise.all([
  save("cover_before.docx", coverBefore), save("cover_after.docx", coverAfter),
  save("footer_before.docx", footerDoc(false)), save("footer_after.docx", footerDoc(true)),
  save("intro_before.docx", introDoc(false)), save("intro_after.docx", introDoc(true)),
  save("letters_before.docx", lettersDoc(true)), save("letters_after.docx", lettersDoc(false)),
]).then(() => console.log("ALL DONE"));
