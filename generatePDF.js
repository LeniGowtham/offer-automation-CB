import PDFDocument from 'pdfkit';
import fs from 'fs';

const OUT = 'Offer Automation Workflow.pdf';

// ─── COLOURS ──────────────────────────────────────────────────────────────────
const C = {
  green:  '#27AE60',
  blue:   '#2980B9',
  purple: '#8E44AD',
  teal:   '#16A085',
  orange: '#E67E22',
  red:    '#C0392B',
  gray:   '#7F8C8D',
  light:  '#ECF0F1',
  dark:   '#2C3E50',
  white:  '#FFFFFF',
  line:   '#BDC3C7',
  bg:     '#F5F6FA',
};

const doc = new PDFDocument({ size: 'A4', margin: 0, autoFirstPage: false });
doc.pipe(fs.createWriteStream(OUT));

// ─── HELPERS ──────────────────────────────────────────────────────────────────
const PW = 595.28;   // A4 width  (points)
const PH = 841.89;   // A4 height (points)
const ML = 50;       // left margin
const MR = 50;       // right margin
const UW = PW - ML - MR;  // usable width

function hex(h) { return h; }

// Filled + stroked rounded rectangle
function roundBox(x, y, w, h, fill, stroke = null, r = 5) {
  doc.roundedRect(x, y, w, h, r);
  if (stroke) {
    doc.fillAndStroke(fill, stroke);
  } else {
    doc.fill(fill);
  }
}

// Centred single-line text inside a box
function centredText(text, x, y, w, h, color, size = 9, bold = false) {
  doc.font(bold ? 'Helvetica-Bold' : 'Helvetica')
     .fontSize(size)
     .fillColor(color)
     .text(text, x, y + (h - size) / 2, { width: w, align: 'center', lineBreak: false });
}

// Horizontal rule
function hRule(y, color = C.line, lx = ML, rw = UW) {
  doc.moveTo(lx, y).lineTo(lx + rw, y).lineWidth(0.5).strokeColor(color).stroke();
}

// Arrow (downward)
function arrow(cx, y1, y2, color = C.gray) {
  const ah = 6; // arrowhead size
  doc.moveTo(cx, y1).lineTo(cx, y2 - ah)
     .lineWidth(1).strokeColor(color).stroke();
  doc.polygon([cx, y2], [cx - 4, y2 - ah], [cx + 4, y2 - ah])
     .fill(color);
}

// ─── PAGE 1: COVER ────────────────────────────────────────────────────────────
doc.addPage();

// Background header band
doc.rect(0, 0, PW, 200).fill(C.dark);

// Title
doc.font('Helvetica-Bold').fontSize(24).fillColor(C.white)
   .text('Offer Automation', ML, 60, { width: UW, align: 'center' });
doc.font('Helvetica').fontSize(14).fillColor('#95A5A6')
   .text('Workflow Reference', ML, 92, { width: UW, align: 'center' });

// Divider line
doc.moveTo(ML + 80, 120).lineTo(PW - MR - 80, 120)
   .lineWidth(1).strokeColor('#3D566E').stroke();

// Sub info
doc.fontSize(10).fillColor('#BDC3C7')
   .text('Capillary Stage  ·  Brand 27  ·  north-america.api.capillarytech.com', ML, 132, { width: UW, align: 'center' });

// Coloured strip
const stripY = 200;
const stripW = UW / 4;
[C.blue, C.purple, C.teal, C.green].forEach((c, i) => {
  doc.rect(ML + i * stripW, stripY, stripW, 6).fill(c);
});

// Summary boxes
const boxY = 230;
const bw = (UW - 30) / 3;
const summaryItems = [
  { label: 'Scripts', value: '3', sub: 'automate · createOffer\nupdateAttributes', color: C.blue },
  { label: 'Locales managed', value: '5', sub: 'en · en-qc · en-roc\nfr-qc · fr-roc', color: C.purple },
  { label: 'Excel sheets', value: '3', sub: 'Offer ID · Attributes\nContent', color: C.teal },
];
summaryItems.forEach((item, i) => {
  const bx = ML + i * (bw + 15);
  doc.roundedRect(bx, boxY, bw, 110, 6).fillAndStroke('#F8F9FA', C.line);
  // colour top bar
  doc.rect(bx, boxY, bw, 4).fill(item.color);
  doc.font('Helvetica-Bold').fontSize(32).fillColor(item.color)
     .text(item.value, bx, boxY + 20, { width: bw, align: 'center' });
  doc.font('Helvetica-Bold').fontSize(9).fillColor(C.dark)
     .text(item.label.toUpperCase(), bx, boxY + 62, { width: bw, align: 'center' });
  doc.font('Helvetica').fontSize(8).fillColor(C.gray)
     .text(item.sub, bx, boxY + 76, { width: bw, align: 'center' });
});

// Run order section
const roY = 370;
doc.font('Helvetica-Bold').fontSize(11).fillColor(C.dark)
   .text('Run Order', ML, roY);
hRule(roY + 16);

const steps = [
  { n: '1', label: 'Fill in Offer details.xlsx', sub: 'All three sheets: Offer ID, Attributes to be updated, Content', color: C.blue },
  { n: '2', label: 'node automate.js', sub: 'Runs Phase 1 (locale management) then Phase 2 (attributes & content) for every offer ID', color: C.purple },
  { n: '3', label: 'Review the console summary', sub: 'P1 and P2 status printed per offer — check for any FAILED rows', color: C.teal },
];

steps.forEach((s, i) => {
  const sy = roY + 26 + i * 80;
  // number circle
  doc.circle(ML + 14, sy + 14, 14).fill(s.color);
  doc.font('Helvetica-Bold').fontSize(13).fillColor(C.white)
     .text(s.n, ML + 5, sy + 7, { width: 20, align: 'center' });
  // content
  doc.font('Helvetica-Bold').fontSize(11).fillColor(C.dark)
     .text(s.label, ML + 36, sy + 2);
  doc.font('Helvetica').fontSize(9).fillColor(C.gray)
     .text(s.sub, ML + 36, sy + 17, { width: UW - 36 });
  if (i < steps.length - 1) {
    doc.moveTo(ML + 14, sy + 30).lineTo(ML + 14, sy + 68)
       .lineWidth(1).strokeColor(C.line).dash(3, { space: 3 }).stroke();
    doc.undash();
  }
});

// Scripts reference
const srY = 630;
doc.font('Helvetica-Bold').fontSize(11).fillColor(C.dark)
   .text('Available Scripts', ML, srY);
hRule(srY + 16);

const scripts = [
  { cmd: 'node automate.js',         desc: 'Combined — Phase 1 + Phase 2 for all offer IDs', color: C.green },
  { cmd: 'node createOffer.js',       desc: 'Phase 1 only — locale management for all offer IDs', color: C.purple },
  { cmd: 'node updateAttributes.js',  desc: 'Phase 2 only — attributes & content for first offer ID', color: C.teal },
];

scripts.forEach((s, i) => {
  const sy = srY + 26 + i * 46;
  doc.roundedRect(ML, sy, UW, 38, 4).fillAndStroke('#F8F9FA', C.line);
  doc.rect(ML, sy, 4, 38).fill(s.color);
  doc.font('Courier-Bold').fontSize(10).fillColor(C.dark)
     .text(s.cmd, ML + 14, sy + 8);
  doc.font('Helvetica').fontSize(8.5).fillColor(C.gray)
     .text(s.desc, ML + 14, sy + 23);
});

// Footer
doc.font('Helvetica').fontSize(8).fillColor(C.line)
   .text(`Generated ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`, ML, PH - 40, { width: UW, align: 'center' });
doc.text('Page 1 of 3', ML, PH - 28, { width: UW, align: 'center' });

// ─── PAGE 2: FLOWCHART DIAGRAM ────────────────────────────────────────────────
doc.addPage();

// page header
doc.rect(0, 0, PW, 44).fill(C.dark);
doc.font('Helvetica-Bold').fontSize(13).fillColor(C.white)
   .text('Flowchart Diagram', ML, 15, { width: UW });

// diagram constants
const cx = PW / 2;          // centre X
const bw2 = 290;            // box width
const bh  = 28;             // box height
const bx  = cx - bw2 / 2;  // box left X
let   dy  = 60;             // current Y cursor

// helper: draw a process box and advance dy
function procBox(label, fill, textColor = C.white, extraH = 0) {
  const h = bh + extraH;
  roundBox(bx, dy, bw2, h, fill, null, 5);
  doc.font('Helvetica').fontSize(8.5).fillColor(textColor)
     .text(label, bx + 6, dy + (h - 8.5) / 2, { width: bw2 - 12, align: 'center', lineBreak: false });
  dy += h;
}

// helper: draw a section header bar
function phaseHeader(label, fill) {
  doc.rect(bx - 8, dy, bw2 + 16, 22).fill(fill);
  doc.font('Helvetica-Bold').fontSize(9).fillColor(C.white)
     .text(label, bx - 8, dy + 6, { width: bw2 + 16, align: 'center' });
  dy += 22;
}

// helper: draw separator label
function separator(label) {
  doc.font('Helvetica').fontSize(7.5).fillColor(C.gray)
     .text(label, bx, dy + 4, { width: bw2, align: 'center' });
  dy += 14;
  hRule(dy, C.line, bx, bw2);
  dy += 6;
}

// helper: draw arrow between last box and next box
function gap(color = C.gray) {
  arrow(cx, dy, dy + 14, color);
  dy += 14;
}

// ── FLOW ──
// START
roundBox(cx - 60, dy, 120, 26, C.green, null, 13);
doc.font('Helvetica-Bold').fontSize(10).fillColor(C.white)
   .text('START', cx - 60, dy + 8, { width: 120, align: 'center' });
dy += 26; gap(C.blue);

procBox('Read Excel File  (Offer IDs · Attributes · Content)', C.blue);
dy += 4; gap(C.blue);
procBox('PREFLIGHT: Check & register org language metadata', C.blue);
dy += 4; gap(C.blue);

separator('─── FOR EACH OFFER ID ───');

// Phase 1 group
const p1top = dy;
phaseHeader('PHASE 1 — Locale Management', C.purple);
dy += 4;
procBox('GET /reward/{id}/brand/27', C.purple);
dy += 2;
// decision note
doc.roundedRect(bx + 4, dy, bw2 - 8, 22, 3).fill('#F3E5F5');
doc.font('Helvetica').fontSize(8).fillColor(C.purple)
   .text('If all 5 locales present → skip PUT and go to Phase 2', bx + 8, dy + 7, { width: bw2 - 16, align: 'center' });
dy += 22; gap(C.purple);
procBox('Clone missing locales  (en-roc → en-qc → en)', C.purple);
dy += 2; gap(C.purple);
procBox('PUT /reward/{id}/brand/27  (all 5 locales)', C.purple);
dy += 2; gap(C.purple);
procBox('VERIFY: GET again — confirm all 5 locales present', C.purple);
dy += 2;
// Phase 1 outcome note
doc.roundedRect(bx + 4, dy, bw2 - 8, 22, 3).fill('#FDEDEC');
doc.font('Helvetica').fontSize(8).fillColor(C.red)
   .text('Phase 1 FAILED → mark offer as FAILED, skip Phase 2', bx + 8, dy + 7, { width: bw2 - 16, align: 'center' });
dy += 22;

// draw group box around Phase 1
doc.roundedRect(bx - 8, p1top, bw2 + 16, dy - p1top, 6)
   .lineWidth(1).strokeColor(C.purple).stroke();

dy += 6; gap(C.teal);

// Phase 2 group
const p2top = dy;
phaseHeader('PHASE 2 — Attributes & Content', C.teal);
dy += 4;
procBox('GET fresh reward state  (post-Phase-1)', C.teal);
dy += 2; gap(C.teal);
procBox('Extract vendorName + cashbackValue from locale customFields', C.teal, C.white, 4);
dy += 2; gap(C.teal);
procBox('Update images (MERCHANT · SMALL · MEDIUM · LARGE) — all 5 locales', C.teal, C.white, 4);
dy += 2; gap(C.teal);
procBox('Apply text replacements  (en-qc/en-roc → EN, fr-qc/fr-roc → FR)', C.teal);
dy += 2; gap(C.teal);
procBox('PUT /reward/{id}/brand/27  (all locales in one payload)', C.teal);
dy += 2;

doc.roundedRect(bx - 8, p2top, bw2 + 16, dy - p2top, 6)
   .lineWidth(1).strokeColor(C.teal).stroke();

dy += 6; gap(C.gray);

// loop back note
doc.roundedRect(bx + 4, dy, bw2 - 8, 22, 3).fill(C.light);
doc.font('Helvetica').fontSize(8).fillColor(C.gray)
   .text('Repeat for each remaining Offer ID', bx + 8, dy + 7, { width: bw2 - 16, align: 'center' });
dy += 22; gap(C.blue);

separator('─── END OF LOOP ───');

procBox('Print Summary  (P1 and P2 status per offer)', C.blue);
dy += 4; gap(C.green);

roundBox(cx - 60, dy, 120, 26, C.green, null, 13);
doc.font('Helvetica-Bold').fontSize(10).fillColor(C.white)
   .text('END', cx - 60, dy + 8, { width: 120, align: 'center' });

// page footer
doc.font('Helvetica').fontSize(8).fillColor(C.line)
   .text('Page 2 of 3', ML, PH - 28, { width: UW, align: 'center' });

// ─── PAGE 3: REFERENCE TABLES ─────────────────────────────────────────────────
doc.addPage();
doc.rect(0, 0, PW, 44).fill(C.dark);
doc.font('Helvetica-Bold').fontSize(13).fillColor(C.white)
   .text('Reference Tables', ML, 15, { width: UW });

let ry = 60;

// ── section helper ──
function sectionTitle(label) {
  doc.font('Helvetica-Bold').fontSize(10).fillColor(C.dark).text(label, ML, ry);
  ry += 16;
  hRule(ry, C.line);
  ry += 8;
}

function tableHeader(cols, widths) {
  let tx = ML;
  doc.rect(ML, ry, UW, 20).fill(C.dark);
  cols.forEach((c, i) => {
    doc.font('Helvetica-Bold').fontSize(8).fillColor(C.white)
       .text(c, tx + 4, ry + 6, { width: widths[i] - 8, lineBreak: false });
    tx += widths[i];
  });
  ry += 20;
}

function tableRow(cols, widths, shade = false) {
  if (shade) doc.rect(ML, ry, UW, 22).fill('#F8F9FA');
  let tx = ML;
  cols.forEach((c, i) => {
    doc.font(i === 0 ? 'Helvetica-Bold' : 'Helvetica').fontSize(8).fillColor(C.dark)
       .text(c, tx + 4, ry + 7, { width: widths[i] - 8, lineBreak: false });
    tx += widths[i];
  });
  hRule(ry + 22, '#ECF0F1');
  ry += 22;
}

// ── Excel sheets ──
sectionTitle('Excel Input — Offer details.xlsx');
tableHeader(['Sheet', 'Column A', 'Column B', 'Column C'], [120, 120, 120, 135]);
[
  ['Offer ID',                  'Offer IDs (row 2+)', '—', '—'],
  ['Attributes to be updated',  'Attribute name',     'URL or asset ID', '—'],
  ['Content',                   'Field name',         'English value', 'French value'],
].forEach((r, i) => tableRow(r, [120, 120, 120, 135], i % 2 === 0));

ry += 16;

// ── Attribute name map ──
sectionTitle('Image Attribute Names');
tableHeader(['Excel name', 'API field name', 'Notes'], [140, 160, 195]);
[
  ['merchantimage', 'MERCHANT_IMAGE',    'If value starts with http → isExternal:true'],
  ['small image',   'OFFER_IMAGE_SMALL', 'Otherwise treated as internal asset ID'],
  ['medium image',  'OFFER_IMAGE_MEDIUM',''],
  ['large image',   'OFFER_IMAGE_LARGE', ''],
].forEach((r, i) => tableRow(r, [140, 160, 195], i % 2 === 0));

ry += 16;

// ── Locale routing ──
sectionTitle('Locale Routing');
tableHeader(['Locale', 'Image update', 'Text source', 'vendorName & % replaced'], [80, 90, 180, 145]);
[
  ['en',     'Yes', '— (no text replacement)', 'No'],
  ['en-qc',  'Yes', 'English column (Col B)',   'Yes'],
  ['en-roc', 'Yes', 'English column (Col B)',   'Yes'],
  ['fr-qc',  'Yes', 'French column (Col C)',    'Yes'],
  ['fr-roc', 'Yes', 'French column (Col C)',    'Yes'],
].forEach((r, i) => tableRow(r, [80, 90, 180, 145], i % 2 === 0));

ry += 16;

// ── Text replacement rules ──
sectionTitle('Text Replacement Rules');
tableHeader(['Field', 'Vendor name replaced', '% value replaced', 'Pattern'], [100, 110, 110, 175]);
[
  ['name',              'Yes', 'Yes', 'X% → cashbackValue%'],
  ['shortTitle',        'Yes', 'Yes', 'X% → cashbackValue%'],
  ['description',       'Yes', 'Yes', 'X% → cashbackValue%'],
  ['termsAndConditions','Yes', 'No',  '% values are intentionally preserved'],
].forEach((r, i) => tableRow(r, [100, 110, 110, 175], i % 2 === 0));

ry += 16;

// ── Clone priority ──
sectionTitle('Clone Source Priority (Phase 1)');
tableHeader(['Priority', 'Source locale', 'Used when'], [80, 130, 285]);
[
  ['1st', 'en-roc', 'Always tried first when a locale is missing'],
  ['2nd', 'en-qc',  'Used if en-roc is not present on the reward'],
  ['3rd', 'en',     'Last resort — falls back if neither en-roc nor en-qc exist'],
].forEach((r, i) => tableRow(r, [80, 130, 285], i % 2 === 0));

// page footer
doc.font('Helvetica').fontSize(8).fillColor(C.line)
   .text(`Generated ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`, ML, PH - 40, { width: UW, align: 'center' });
doc.text('Page 3 of 3', ML, PH - 28, { width: UW, align: 'center' });

// ─── DONE ─────────────────────────────────────────────────────────────────────
doc.end();
console.log(`✓ Generated: ${OUT}`);
