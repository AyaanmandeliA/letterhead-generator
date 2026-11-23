#!/usr/bin/env node
/**
 * Generate a letterhead .docx from a registered brand config + a content JSON file.
 *
 * Usage:
 *   node generate_letterhead.js --config aanka-group --data ./data.json --out "../output/example-letter.docx"
 *
 * --data JSON shape (all fields optional except body):
 * {
 *   "entityName": "Khau Galli",            // required if the config's requiresEntityField is true
 *   "date": "19 August 2026",              // omit to use "[Date]"
 *   "recipient": { "name": "...", "title": "...", "org": "...", "city": "..." }, // omit for an unaddressed announcement
 *   "salutation": "Dear ...,",             // omit if no recipient
 *   "subject": "Re: ...",                  // omit for no subject line
 *   "body": ["Paragraph one.", "Paragraph two."],
 *   "signOff": "With appreciation,",       // omit to use "With appreciation,"
 *   "signatory": { "name": "...", "title": "..." }, // omit to use the config's defaultSignatory
 *   "suppressSignOff": false,              // true skips the auto signOff+signatory block entirely —
 *                                           // use for documents (contracts, forms) that already end
 *                                           // with their own signature block inside `body`
 *   "highlightPlaceholders": false          // true highlights every [bracketed] fill-in field in
 *                                           // yellow (Word's real highlighter), so a human filling
 *                                           // in a template can't miss one. Off by default so it
 *                                           // doesn't change output for letters that don't ask for it.
 * }
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle, Header, Footer, ImageRun,
  Table, TableRow, TableCell, WidthType, VerticalAlign, TableLayoutType, ShadingType,
} = require("docx");

function parseArgs(argv) {
  const out = {};
  for (let i = 0; i < argv.length; i += 2) {
    const key = argv[i].replace(/^--/, "");
    out[key] = argv[i + 1];
  }
  return out;
}

const args = parseArgs(process.argv.slice(2));
if (!args.config || !args.data || !args.out) {
  console.error("Usage: node generate_letterhead.js --config <key> --data <path.json> --out <path.docx>");
  process.exit(1);
}

const registryPath = path.join(__dirname, "letterheads.json");
const registry = JSON.parse(fs.readFileSync(registryPath, "utf8"));
const cfg = registry[args.config];
if (!cfg) {
  console.error(`Unknown letterhead key "${args.config}". Known keys: ${Object.keys(registry).join(", ")}`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(args.data, "utf8"));

if (cfg.requiresEntityField && !data.entityName) {
  console.error(`"${cfg.name}" requires data.entityName (the legal entity this letter is issued on behalf of). None was given.`);
  process.exit(1);
}
if (!data.body || !data.body.length) {
  console.error("data.body must be a non-empty array of paragraph strings.");
  process.exit(1);
}

const INK = cfg.colors.ink;
const MUTED = cfg.colors.muted;
const ACCENT = cfg.colors.accent;
const DISPLAY_FONT = cfg.fonts.display;
const BODY_FONT = cfg.fonts.body;

const rule = (size) => new Paragraph({
  border: { bottom: { color: ACCENT, space: 4, style: BorderStyle.SINGLE, size: size || 4 } },
  spacing: { after: 0 },
});

const noTableBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function logoImageRun(logoPath, logoAspect, heightPx) {
  return new ImageRun({
    data: fs.readFileSync(path.resolve(__dirname, logoPath)),
    type: "png",
    transformation: { width: Math.round(heightPx * (logoAspect || 1)), height: heightPx },
  });
}

// "framed" wraps the mark in a thin bordered box (for logos that still carry their own
// baked-in background color — the frame stops a loose colored rectangle sitting on the
// page). "dark-chip" gives it a solid dark backing instead (for light-on-dark artwork
// that would vanish on white paper, e.g. DecoVibes). "plain" is just the image.
// Framed/dark-chip use a tight 1x1 nested table (AUTOFIT) rather than a paragraph border —
// a paragraph border draws across the full column width in Word, not just around the
// inline image, which left a stray line running past the logo.
function treatedLogoBlock(logoPath, logoAspect, heightPx, treatment) {
  const img = logoImageRun(logoPath, logoAspect, heightPx);
  if (treatment === "framed" || treatment === "dark-chip") {
    return new Table({
      width: { size: 0, type: WidthType.AUTO },
      layout: TableLayoutType.AUTOFIT,
      borders: treatment === "framed" ? {
        top: { style: BorderStyle.SINGLE, size: 4, color: "B8B8B0" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "B8B8B0" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "B8B8B0" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "B8B8B0" },
      } : noTableBorders,
      rows: [new TableRow({ children: [new TableCell({
        ...(treatment === "dark-chip" ? { shading: { type: ShadingType.CLEAR, fill: "141414" } } : {}),
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [img] })],
      })] })],
    });
  }
  return new Paragraph({ alignment: AlignmentType.RIGHT, children: [img] });
}

let headerChildren;

if (cfg.headerStyle === "brand-masthead") {
  // Brand's own letterhead: name + "An AANKA Group Company" on the left, that brand's
  // real logo on the right — a single thin table row so the two sides sit level.
  const logoCellChildren = cfg.logo
    ? [treatedLogoBlock(cfg.logo, cfg.logoAspect, 62, cfg.logoTreatment || "plain")] // Table or Paragraph, both valid TableCell children
    : [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: cfg.name, font: DISPLAY_FONT, size: 34, color: INK })] })];

  const mastheadTable = new Table({
    // Page content width with this template's margins (Letter, 1600 twips each side) is
    // 9040 twips — percentage widths on a FIXED-layout table render unreliably across
    // Word/LibreOffice/Quick Look, so the two columns are sized explicitly in twips instead.
    width: { size: 9040, type: WidthType.DXA },
    columnWidths: [5800, 3240],
    layout: TableLayoutType.FIXED,
    borders: noTableBorders,
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 5800, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        borders: noTableBorders,
        // cfg.hideNameInMasthead skips the big name line — for a brand whose logo already
        // spells its own name out (e.g. House of Zaika), repeating it in type is redundant.
        children: cfg.hideNameInMasthead
          ? [new Paragraph({ children: [new TextRun({ text: cfg.masthead || "An AANKA Group Company", font: BODY_FONT, bold: true, size: 15, color: INK })] })]
          : [
              new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: cfg.name, font: DISPLAY_FONT, bold: true, allCaps: true, size: 34, color: INK, characterSpacing: 6 })] }),
              new Paragraph({ children: [new TextRun({ text: cfg.masthead || "An AANKA Group Company", font: BODY_FONT, bold: true, size: 15, color: INK })] }),
            ],
      }),
      new TableCell({
        width: { size: 3240, type: WidthType.DXA },
        verticalAlign: VerticalAlign.CENTER,
        borders: noTableBorders,
        children: logoCellChildren,
      }),
    ] })],
  });

  headerChildren = [mastheadTable, rule(4)];
} else {
  // Group-level letterhead: centered AANKA wordmark, masthead line, optional entity field.
  const headerMark = cfg.logo
    ? logoImageRun(cfg.logo, cfg.logoAspect || 4.6167, 41)
    : new TextRun({ text: cfg.name, font: DISPLAY_FONT, size: 44, color: INK, characterSpacing: 30 });

  headerChildren = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 120 },
      children: [headerMark],
    }),
  ];
  if (cfg.masthead) {
    headerChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: cfg.requiresEntityField ? 200 : 0 },
      children: [new TextRun({ text: cfg.masthead, font: BODY_FONT, size: 15, color: MUTED, allCaps: true, characterSpacing: 20 })],
    }));
  }
  if (cfg.requiresEntityField) {
    headerChildren.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 40 },
        children: [new TextRun({ text: cfg.entityFieldLabel || "Issued on behalf of", font: BODY_FONT, size: 13, color: ACCENT, bold: true, allCaps: true, characterSpacing: 24 })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 160 },
        children: [new TextRun({ text: data.entityName, font: DISPLAY_FONT, italics: true, size: 27, color: INK })],
      }),
    );
  }
  headerChildren.push(rule(4));
}

const header = new Header({ children: headerChildren });

let footerChildren;

if (cfg.footerBrands) {
  // Group-level footer: every brand's name in a single centered line, each with a tiny
  // logo next to it where a logo file exists (dot-separated); falls back to plain text
  // for a brand with no logo on file yet (e.g. AANKA Contracting).
  const runs = [];
  cfg.footerBrands.forEach((b, i) => {
    if (i > 0) runs.push(new TextRun({ text: "   ·   ", font: BODY_FONT, size: 13, color: MUTED }));
    if (b.logo) {
      runs.push(logoImageRun(b.logo, b.logoAspect || 1, 11));
      runs.push(new TextRun({ text: "  " }));
    }
    runs.push(new TextRun({ text: b.name, font: BODY_FONT, size: 13, color: MUTED, allCaps: true, characterSpacing: 8 }));
  });
  footerChildren = [
    rule(4),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 40 },
      children: [new TextRun({ text: cfg.name, font: BODY_FONT, size: 13, color: MUTED, allCaps: true, characterSpacing: 14 })],
    }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: runs }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: cfg.footerLine2, font: BODY_FONT, size: 14, color: MUTED })],
    }),
  ];
} else {
  footerChildren = [
    rule(4),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 160, after: 60 },
      children: [new TextRun({ text: cfg.footerLine1, font: BODY_FONT, size: 13, color: MUTED, allCaps: true, characterSpacing: 14 })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: cfg.footerLine2, font: BODY_FONT, size: 14, color: MUTED })],
    }),
  ];
}

const footer = new Footer({ children: footerChildren });

// Splits "Basic Salary: [Amount] per month" into plain/bracket segments so the bracket
// portion alone can be highlighted — never breaks on signature blanks (bare underscores
// carry no brackets and are meant to stay empty for a wet-ink signature, not be typed over).
const BRACKET_RE = /(\[[^\]]*\])/g;
function splitBrackets(text) {
  return text.split(BRACKET_RE).filter((s) => s.length > 0);
}

function line(text, opts = {}) {
  const runOpts = { font: opts.font ?? BODY_FONT, size: opts.size ?? 20, color: opts.color ?? INK, bold: opts.bold, italics: opts.italics };
  const children = data.highlightPlaceholders
    ? splitBrackets(text).map((seg) => new TextRun({ ...runOpts, text: seg, ...(seg.startsWith("[") && seg.endsWith("]") ? { highlight: "yellow" } : {}) }))
    : [new TextRun({ ...runOpts, text })];
  return new Paragraph({
    spacing: { after: opts.after ?? 160, line: opts.line ?? 300 },
    alignment: opts.align ?? AlignmentType.LEFT,
    children,
  });
}

const body = [];

body.push(line(data.date || "[Date]", { align: AlignmentType.RIGHT, size: 18, color: MUTED, after: 320 }));

if (data.recipient) {
  const r = data.recipient;
  const recipientLines = [r.name, [r.title, r.org].filter(Boolean).join(", "), r.city].filter(Boolean);
  recipientLines.forEach((l, i) => body.push(line(l, { size: 18, after: i === recipientLines.length - 1 ? 260 : 20 })));
}

if (data.salutation) {
  body.push(line(data.salutation, { size: 18, after: 200 }));
}

if (data.subject) {
  body.push(new Paragraph({
    spacing: { after: 240 },
    children: [new TextRun({ text: data.subject, font: DISPLAY_FONT, italics: true, size: 24, color: INK })],
  }));
}

// A short ALL-CAPS line (e.g. "1. POSITION AND DUTIES", "FOR THE COMPANY") reads as a
// section header in contracts/forms — set apart with bold + spacing rather than sitting
// flush against surrounding prose.
const isHeaderLine = (text) => /^[A-Z0-9][A-Z0-9 ,.&/()'-]*$/.test(text.trim()) && /[A-Z]/.test(text);

data.body.forEach((p, i) => {
  const last = i === data.body.length - 1;
  if (isHeaderLine(p)) {
    body.push(line(p, { bold: true, after: 100, ...(i > 0 ? { line: 300 } : {}) }));
  } else {
    body.push(line(p, { after: last ? 400 : 160 }));
  }
});

if (!data.suppressSignOff) {
  body.push(line(data.signOff || "With appreciation,", { after: 500 }));
  const signatory = data.signatory || cfg.defaultSignatory;
  if (signatory) {
    body.push(line(signatory.name, { bold: true, after: 20 }));
    body.push(line(signatory.title, { size: 17, color: MUTED }));
  }
}

const doc = new Document({
  sections: [{
    properties: { page: { margin: { top: 1600, bottom: 1500, left: 1600, right: 1600, header: 620, footer: 620 } } },
    headers: { default: header },
    footers: { default: footer },
    children: body,
  }],
});

const outPath = path.resolve(process.cwd(), args.out);
fs.mkdirSync(path.dirname(outPath), { recursive: true });
Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("written:", outPath, `(${buf.length} bytes)`);
});
