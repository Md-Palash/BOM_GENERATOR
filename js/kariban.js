/* ============================================================
   KARIBAN ENGINE
   PDF extraction (Placements section: Fabrics / Trims / Label /
   Packaging), Supply Chain sheet, Cost Sheet merge, and
   Kariban-specific UI wiring.
   Depends on: shared.js (homeView, requestUnlock)

   Self-contained by design (like decathlon.js): the generic Excel
   row/formula/merge utilities are duplicated here rather than relying
   on another engine's script having already loaded first.
   ============================================================ */

/* ============================================================
   KARIBAN — PDF EXTRACTION ENGINE

   Kariban's tech pack is a PLM-generated report (not a hand-laid-out
   document like Haddad/Decathlon's), so its BOM data already lives in
   real ruled tables split across several distinct report types:
     - "Fabrics & Trims"            -> Fabrics (N) / Trims (N) sub-blocks
     - "WK Labelling - PPE Category II" -> label items
     - "WK Hangtag & User manual"   -> hangtag/drawcord/sticker items
     - "WK Packing"                 -> packaging items
   All four share the same underlying row-detection technique: every
   real data row's Product cell begins with a stable "Code - Name"
   token (FMat1739, TZip0088, LMai0037, PBox0003 ...), which is a far
   more reliable "this line starts a new row" signal than line/rect
   detection would be, and mirrors the Decathlon engine's own approach
   of anchoring on the most stable, explicit marker available.
   ============================================================ */

const KARIBAN_PRODUCT_CODE_RE = /^[A-Z]{2}[a-z]{2}\d{4}/;

const KARIBAN_REPORT_TYPES = {
  fabrics_trims: /^Fabrics\s*&\s*Trims$/i,
  labelling: /^WK Labelling/i,
  hangtag: /^WK Hangtag/i,
  packing: /^WK Packing/i,
};

function classifyKaribanReportType(rt) {
  if (!rt) return null;
  for (const key of Object.keys(KARIBAN_REPORT_TYPES)) {
    if (KARIBAN_REPORT_TYPES[key].test(rt)) return key;
  }
  return null;
}

function normKariban(s) { return (s || '').replace(/\s+/g, ' ').trim(); }

// Reads the fixed header band (KARIBAN | Style | ReportType | Draft+Date)
// that appears at the same y-position on every page, and pulls out just
// the report-type text (e.g. "Fabrics & Trims", "WK Packing"). Far more
// stable than inferring report type from which columns are present.
function readKaribanHeaderBand(items) {
  const band = items.filter(it => it.y >= 24 && it.y <= 40 && it.str.trim());
  band.sort((a, b) => a.x - b.x);
  const out = [];
  let started = false;
  for (const it of band) {
    const t = it.str.trim();
    if (!t) continue;
    if (t === 'KARIBAN') continue;
    if (!started && /^[A-Z]{2,4}\d/.test(t)) continue; // style code repeat
    if (/^Draft/.test(t) || /^\d{2}\/\d{2}\/\d{4}/.test(t) || /^\d{1,2}:\d{2}$/.test(t)) break;
    started = true;
    out.push(t);
  }
  return out.join(' ').trim();
}

// The "Placements" sub-heading sits at a fixed y on every placement-table
// page, directly above the table's own header row.
function findKaribanPlacementsTitle(items) {
  return items.find(it => normKariban(it.str) === 'Placements' && it.y >= 65 && it.y <= 85);
}

// Column header labels per report type. Order matters: it's used to map
// x-anchors positionally, since some labels wrap onto a second line
// ("Common" / "Size") and can't always be text-matched directly.
const KARIBAN_REPORT_COLUMNS = {
  fabrics_trims: ['Placement', 'Product', 'Image', 'Common Size', 'Common Qty', 'Comments', 'Remarks'],
  labelling: ['Group', 'Placement', 'Image', 'Product', 'Remarks'],
  hangtag: ['Product', 'Image', 'Group', 'Quality', 'Type of finishing', 'Comments', 'Remarks'],
  packing: ['Product', 'Placement', 'Image', 'Remarks', 'Comments'],
};

function buildKaribanAnchors(items, titleY, reportType) {
  const wanted = KARIBAN_REPORT_COLUMNS[reportType];
  // only the table's FIRST header line, and only non-blank tokens - blank
  // (space) items would otherwise shift the positional index mapping
  const band = items.filter(it => it.y >= titleY + 20 && it.y <= titleY + 36 && it.str.trim());
  band.sort((a, b) => a.x - b.x);
  const anchors = [];
  for (let i = 0; i < Math.min(band.length, wanted.length); i++) {
    anchors.push({ key: wanted[i], x: band[i].x });
  }
  // "Common Qty" is numeric and renders right-aligned within its cell,
  // unlike every other (left-aligned) column here - its header label's x
  // sits well to the LEFT of where the data actually lands (confirmed
  // against real samples: header x~277, data values land at x~306), so
  // classification needs the data's real x or values spill into Comments.
  const qtyAnchor = anchors.find(a => a.key === 'Common Qty');
  if (qtyAnchor) qtyAnchor.x += 29;
  // Cap the last wanted column so colorway Color/No pairs (which sit
  // further right and aren't part of any wanted column) don't spill into
  // it via an open-ended classification range.
  if (band.length > wanted.length) {
    anchors.push({ key: '_ignore', x: band[wanted.length].x });
  }
  return anchors;
}

function classifyKaribanColumn(x, anchors) {
  // boundary classification: each column owns everything up to the
  // midpoint between it and the next anchor (same approach as the
  // Decathlon engine), not pure nearest-anchor, which misclassifies
  // values sitting near a boundary toward whichever neighbor is closer.
  for (let i = 0; i < anchors.length; i++) {
    const next = anchors[i + 1];
    const boundary = next ? (anchors[i].x + next.x) / 2 : Infinity;
    if (x < boundary) return anchors[i].key;
  }
  return anchors[anchors.length - 1].key;
}

function isKaribanSubHeaderRow(colTexts) {
  const nonEmpty = Object.values(colTexts).filter(v => v);
  return nonEmpty.length > 0 && nonEmpty.every(v => v === 'Color' || v === 'No');
}

function isKaribanDividerLine(lineText) {
  return /\(\d+\)\s*$/.test(lineText);
}

// Parses the WHOLE tech pack in a single pass across every placement-table
// page, returning every row found (tagged with report type + section
// divider) plus the style code/designation read once from page 1.
async function extractKaribanPlacementRows(file, onProgress) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  const results = [];
  let currentDivider = null;
  let currentReportType = null;
  let styleCode = '';
  let styleDesignation = '';
  let styleBrand = '';

  for (let p = 1; p <= pdf.numPages; p++) {
    onProgress && onProgress(p, pdf.numPages);
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    const items = content.items
      .map(it => {
        const tx = pdfjsLib.Util.transform(viewport.transform, it.transform);
        return { str: it.str, x: tx[4], y: tx[5] };
      })
      // NOTE: deliberately keep whitespace-only items (not it.str.trim()).
      // This font renders some ligatures (fi/fl) as their own tight glyph,
      // relying on real embedded space-character items on either side for
      // correct word reconstruction ("polar" + "fl" + "eece" -> "polar
      // fleece"); dropping "empty after trim" items here silently eats
      // those spaces and mangles reconstructed text.
      .filter(it => it.str.length > 0);

    // Style code + designation, read once from page 1's Properties table -
    // used for Supply Chain's Style No.
    if (p === 1 && !styleCode) {
      const styleLabel = items.find(it => normKariban(it.str) === 'Style');
      if (styleLabel) {
        const valTok = items.filter(it => Math.abs(it.y - styleLabel.y) <= 3 && it.x > styleLabel.x + 20)
          .sort((a, b) => a.x - b.x);
        styleCode = valTok.map(t => t.str.trim()).join(' ').trim();
      }
      const desigLabel = items.find(it => normKariban(it.str) === 'Designation (EN)');
      if (desigLabel) {
        const valTok = items.filter(it => Math.abs(it.y - desigLabel.y) <= 3 && it.x > desigLabel.x + 20)
          .sort((a, b) => a.x - b.x);
        styleDesignation = valTok.map(t => t.str.trim()).join(' ').trim();
      }
      // Brand - read the same way as Style/Designation above, from
      // whichever label token on page 1's Properties table matches "Brand".
      // Unlike Designation, the brand value is a short code (e.g. "WK"),
      // not a multi-word phrase - take only the single nearest token so a
      // closely-spaced adjacent Properties row can never bleed into this
      // field. Left blank (same fallback as the two fields above) if this
      // exact label text isn't found on this tech pack's page 1.
      const brandLabel = items.find(it => normKariban(it.str) === 'Brand');
      if (brandLabel) {
        const valTok = items.filter(it => Math.abs(it.y - brandLabel.y) <= 3 && it.x > brandLabel.x + 20)
          .sort((a, b) => a.x - b.x);
        styleBrand = valTok.length ? valTok[0].str.trim() : '';
      }
    }

    const headerBandText = readKaribanHeaderBand(items);
    const reportType = classifyKaribanReportType(headerBandText);
    if (!reportType) continue; // Properties / Image Data Sheet / Size Chart pages - not BOM data
    if (reportType !== currentReportType) {
      currentReportType = reportType;
      // divider persists across pages of the SAME report type's table
      // only - a brand-new report type always starts without one
      currentDivider = null;
    }

    const titleTok = findKaribanPlacementsTitle(items);
    if (!titleTok) continue;
    const anchors = buildKaribanAnchors(items, titleTok.y, reportType);
    if (!anchors.length) continue;

    // content zone: below the table's own (possibly two-line-wrapped)
    // header row, above the page footer
    const contentTop = titleTok.y + 45;
    const contentItems = items.filter(it => it.y > contentTop && it.y < 800);
    contentItems.sort((a, b) => a.y - b.y || a.x - b.x);
    const lines = [];
    for (const it of contentItems) {
      let line = lines.find(l => Math.abs(l.y - it.y) <= 2.5);
      if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
      line.items.push(it);
    }
    lines.sort((a, b) => a.y - b.y);
    for (const l of lines) l.items.sort((a, b) => a.x - b.x);

    let active = null;
    for (const line of lines) {
      const lineText = line.items.map(it => it.str).join('').trim().replace(/\s+/g, ' ');
      if (!lineText) continue;
      if (/^Displaying\s/.test(lineText)) continue;
      if (/^Page\s+\d+\s+of\s+\d+/i.test(lineText)) continue;

      const cols = {};
      for (const it of line.items) {
        const key = classifyKaribanColumn(it.x, anchors);
        // Concatenate raw text with NO inserted separator: this PDF
        // embeds real space glyphs as their own text items, and inserting
        // a synthetic space between every item both double-spaces normal
        // word gaps and breaks tight ligature glyphs.
        cols[key] = (cols[key] || '') + it.str;
      }
      for (const k of Object.keys(cols)) cols[k] = normKariban(cols[k]);

      if (isKaribanSubHeaderRow(cols)) continue; // the repeated "Color / No" sub-header row

      if (isKaribanDividerLine(lineText) && Object.keys(cols).length <= 2) {
        currentDivider = lineText;
        active = null;
        continue;
      }

      const productText = cols['Product'] || '';
      const isRowStart = KARIBAN_PRODUCT_CODE_RE.test(productText);

      if (isRowStart) {
        active = { reportType, divider: currentDivider, cols: {}, page: p };
        for (const k of Object.keys(cols)) active.cols[k] = cols[k];
        results.push(active);
      } else if (active) {
        // cross-line wrap: always insert a space between wrapped lines
        // (line-wrapping happens at word boundaries, unlike the tight
        // same-line glyph adjacency handled above)
        for (const k of Object.keys(cols)) {
          if (!cols[k]) continue;
          active.cols[k] = active.cols[k] ? active.cols[k] + ' ' + cols[k] : cols[k];
        }
      }
    }
  }

  return { rows: results, styleCode, styleDesignation, brand: styleBrand };
}

/* ============================================================
   KARIBAN — BUSINESS RULES
   Turns raw extracted rows into Fabric / Trim / Label / Packaging line
   items per the buyer's specified rules.
   ============================================================ */

const KARIBAN_YDS_KEYWORDS = /elastic|velcro|tape|cord/i;
const KARIBAN_EXCLUDE_KEYWORDS = /sewing|use guide|method/i;
// Interlining/interfacing/fusible is fabric material for costing purposes,
// even when the tech pack itself files it under Trims (Trimmings &
// Accessories) - it always belongs in the Fabric bucket regardless of
// which divider it was found under.
const KARIBAN_INTERLINING_KEYWORDS = /interlining|interfacing|fusible/i;

function karibanCodeOf(product) { return (product || '').slice(0, 8); }

// Weight (Fabric only) = last purely-numeric token in the Product string,
// in gsm (e.g. "...PU coating - 200" -> 200; "...Birdseye - 140 Anti-UV"
// -> 140, skipping the non-numeric trailing "Anti-UV").
function karibanWeightOf(product) {
  const toks = (product || '').split(/\s+/);
  for (let i = toks.length - 1; i >= 0; i--) {
    const t = toks[i].replace(/^-+|-+$/g, '');
    if (/^\d+(\.\d+)?$/.test(t)) return t;
  }
  return '';
}

function isKaribanFabricDivider(d) { return !!d && d.startsWith('Fabrics'); }
function isKaribanTrimDivider(d) { return !!d && d.startsWith('Trims'); }

// Zipper merge (Trims only): Kariban zippers are always a consecutive run
// of Trim rows sharing the exact same Placement text, starting with the
// Slider and ending with whatever the assembly's final hardware piece is
// (a literal "bottom stop" on a one-way zip, or a "pin and box" piece on a
// two-way zip) - the defining signal is "consecutive same-Placement
// TZip-code rows", not the specific name of the last part. Consumption on
// the merged row is the Slider's own consumption, per business rule.
function mergeKaribanZippers(trimItems) {
  const out = [];
  let i = 0;
  while (i < trimItems.length) {
    const r = trimItems[i];
    const isZip = /^TZip/i.test(r.code);
    if (!isZip) { out.push(r); i++; continue; }
    const group = [r];
    let j = i + 1;
    while (j < trimItems.length && /^TZip/i.test(trimItems[j].code) && trimItems[j].position === r.position) {
      group.push(trimItems[j]);
      j++;
    }
    if (group.length > 1) {
      out.push({
        bucket: 'Trim',
        item: '',
        code: group.map(g => g.code).join('+'),
        position: r.position,
        description: group.map(g => g.description).join(' + '),
        consumption_qty: group[0].consumption_qty, // slider's own consumption
        consumption_unit: group[0].consumption_unit,
        weight_gsm: '',
        size_mm: '',
      });
    } else {
      out.push(r);
    }
    i = j;
  }
  return out;
}

// Applies every field/exclusion/unit rule and returns a flat list of items
// tagged by bucket (Fabric / Trim / Label / Packaging), zippers merged.
function applyKaribanBusinessRules(rows) {
  const items = [];
  for (const d of rows) {
    const cls = d.reportType;
    const cols = d.cols;
    const product = cols['Product'] || '';
    const placement = cols['Placement'] || '';
    const comments = cols['Comments'] || '';
    const remarks = cols['Remarks'] || '';
    const group = cols['Group'] || '';
    const commonSize = cols['Common Size'] || '';
    const commonQty = cols['Common Qty'] || '';

    if (cls === 'fabrics_trims') {
      let bucket;
      if (isKaribanFabricDivider(d.divider)) bucket = 'Fabric';
      else if (isKaribanTrimDivider(d.divider)) bucket = 'Trim';
      else continue;
      if (bucket === 'Trim' && /\bthread\b/i.test(product)) continue; // Trims: exclude thread only
      // Interlining/interfacing/fusible is fabric material regardless of
      // which divider the tech pack filed it under - redirect it into the
      // Fabric bucket before weight/size/unit are computed below, so it
      // gets Fabric's own field treatment (weight_gsm, no size_mm).
      if (bucket === 'Trim' && KARIBAN_INTERLINING_KEYWORDS.test(product)) bucket = 'Fabric';
      // Source Remarks/Comments text sometimes uses "/" as its own
      // separator (e.g. "Windproof / Breathable / Waterproof") - the
      // buyer wants those joined with a plain space instead.
      const desc = [product, comments, remarks].filter(Boolean).join(' ')
        .replace(/\s*\/\s*/g, ' ').replace(/\s+/g, ' ').trim();
      const unit = KARIBAN_YDS_KEYWORDS.test(product) ? 'Yds' : 'Pcs';
      items.push({
        bucket,
        // "Item" (column A) has no defined source for Fabric/Trim rows -
        // Kariban's real example sheet uses short human-written nicknames
        // here ("Shell", "Reflective Tape") that can't be reliably
        // reproduced from the tech pack, so it's left blank rather than
        // guessed. Flagging this clearly rather than inventing content.
        item: '',
        code: karibanCodeOf(product), position: placement, description: desc,
        consumption_qty: commonQty, consumption_unit: unit,
        weight_gsm: bucket === 'Fabric' ? karibanWeightOf(product) : '',
        size_mm: bucket === 'Trim' ? commonSize : '',
      });
    } else if (cls === 'labelling' || cls === 'hangtag') {
      if (KARIBAN_EXCLUDE_KEYWORDS.test(group)) continue; // Label: exclude by Group keyword
      items.push({
        bucket: 'Label',
        // "Item" for Label rows matches the tech pack's own Group value
        // almost exactly (confirmed against the real Format sheet, e.g.
        // "Adress Label" - typo and all).
        item: group,
        code: karibanCodeOf(product), position: placement,
        description: product, consumption_qty: 1, consumption_unit: 'Pcs',
      });
    } else if (cls === 'packing') {
      if (KARIBAN_EXCLUDE_KEYWORDS.test(product)) continue; // Packaging: exclude by Product-name keyword (no Group column exists here)
      items.push({
        bucket: 'Packaging', item: '', code: karibanCodeOf(product), position: placement,
        description: product, consumption_qty: 1, consumption_unit: 'Pcs',
      });
    }
  }

  const trims = mergeKaribanZippers(items.filter(r => r.bucket === 'Trim'));
  return [
    ...items.filter(r => r.bucket === 'Fabric'),
    ...trims,
    ...items.filter(r => r.bucket === 'Label'),
    ...items.filter(r => r.bucket === 'Packaging'),
  ];
}

function karibanItemsByBucket(items) {
  const by = { Fabric: [], Trim: [], Label: [], Packaging: [] };
  for (const it of items) (by[it.bucket] = by[it.bucket] || []).push(it);
  return by;
}

// Dedupe key for "identical item" comparisons: letters/digits only, case
// and punctuation-insensitive (commas, hyphens/dashes, semicolons, dots,
// and whitespace differences don't count toward two items being distinct).
function karibanDedupeKey(text) {
  return (text || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

/* ============================================================
   KARIBAN COST SHEET TEMPLATE (Kariban_Format.xlsx, base64)
   Images stripped from the source template before encoding -
   they're removed at generation time anyway (per business
   rule: no images in output), so embedding them here would
   only bloat this file for no benefit.
   ============================================================ */
const KARIBAN_TEMPLATE_B64 = "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";

/* ============================================================
   GENERIC EXCEL UTILITIES
   (self-contained copy - same approach used by decathlon.js)
   ============================================================ */
const CELL_REF_RE_KARIBAN = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;

function shiftFormulaRowsKariban(formula, insertAt, amount) {
  return formula.replace(CELL_REF_RE_KARIBAN, (m, d1, col, d2, row) => {
    row = parseInt(row, 10);
    if (row >= insertAt) row += amount;
    return `${d1}${col}${d2}${row}`;
  });
}
function colLetterToNumKariban(letter) {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToColLetterKariban(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function getFormulaKariban(cell) {
  const v = cell.value;
  if (v && typeof v === 'object' && v.formula) return v.formula;
  return null;
}
function setFormulaKariban(cell, formula) { cell.value = { formula }; }
function shiftWholeSheetFormulasKariban(ws, insertAt, amount, maxRow, maxCol = 20) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      const f = getFormulaKariban(cell);
      if (f) setFormulaKariban(cell, shiftFormulaRowsKariban(f, insertAt, amount));
    }
  }
}
function snapshotAndUnmergeAllKariban(ws) {
  const merges = [];
  const model = ws.model.merges || [];
  for (const rangeStr of [...model]) {
    const [a, b] = rangeStr.split(':');
    const parseCell = (s) => { const m = s.match(/^([A-Z]+)(\d+)$/); return { col: colLetterToNumKariban(m[1]), row: parseInt(m[2], 10) }; };
    const pa = parseCell(a), pb = parseCell(b || a);
    merges.push({ minRow: Math.min(pa.row, pb.row), maxRow: Math.max(pa.row, pb.row), minCol: Math.min(pa.col, pb.col), maxCol: Math.max(pa.col, pb.col) });
    ws.unMergeCells(rangeStr);
  }
  return merges;
}
function remergeShiftedKariban(ws, merges, insertAt, amount) {
  for (let { minRow, maxRow, minCol, maxCol } of merges) {
    if (minRow >= insertAt) { minRow += amount; maxRow += amount; }
    const rng = `${numToColLetterKariban(minCol)}${minRow}:${numToColLetterKariban(maxCol)}${maxRow}`;
    ws.mergeCells(rng);
  }
}
function copyRowStyleKariban(ws, srcRow, dstRow, maxCol = 20) {
  for (let c = 1; c <= maxCol; c++) {
    const src = ws.getCell(srcRow, c);
    const dst = ws.getCell(dstRow, c);
    dst.style = JSON.parse(JSON.stringify(src.style));
  }
  const srcRowObj = ws.getRow(srcRow);
  ws.getRow(dstRow).height = srcRowObj.height;
}
function copyRowFormulasKariban(ws, srcRow, dstRow, formulaCols) {
  for (const col of formulaCols) {
    const srcCell = ws.getCell(`${col}${srcRow}`);
    const f = getFormulaKariban(srcCell);
    if (f) {
      const shifted = f.replace(CELL_REF_RE_KARIBAN, (m, d1, c2, d2, row) => {
        row = parseInt(row, 10);
        if (row === srcRow) row = dstRow;
        return `${d1}${c2}${d2}${row}`;
      });
      setFormulaKariban(ws.getCell(`${col}${dstRow}`), shifted);
    } else if (srcCell.value !== null && srcCell.value !== undefined) {
      ws.getCell(`${col}${dstRow}`).value = srcCell.value;
    }
  }
}
function normalizeFormulasKariban(ws, maxRow, maxCol = 20) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      const v = cell.value;
      if (v && typeof v === 'object' && (v.formula !== undefined || v.sharedFormula !== undefined)) {
        const f = cell.formula;
        if (f) cell.value = { formula: f };
      }
    }
  }
}
function mergeSectionRangeKariban(ws, row, minCol, width) {
  const rng = `${numToColLetterKariban(minCol)}${row}:${numToColLetterKariban(minCol + width - 1)}${row}`;
  try { ws.mergeCells(rng); } catch (e) {}
}
function base64ToArrayBufferKariban(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

/* ============================================================
   KARIBAN — COST SHEET MERGE ENGINE
   Fills FABRICS / BRANDING LABELS & HANGTAG / TRIMMINGS AND ACCESSORIES /
   Packaging in the Kariban_Format.xlsx template, preserving every existing
   formula. EMBROIDERY & PRINT has no mapped extraction category today and
   is left untouched. Per business decision, every section's existing rows
   are cleared and rebuilt purely from the uploaded tech pack's extraction
   (not preserved as static content) - Kariban's own reference sheet holds
   real values for an unrelated style, not universal starter content.
   ============================================================ */

const KARIBAN_FORMULA_COLS = ['M']; // Total Cost = (K+K*L)*J
const KARIBAN_MAX_COL = 19; // through column S (Fabrics' extra FOB/CNF-by-yard columns)

// firstBlank/lastBlank/totalRow reflect the real template: FABRICS
// uniquely has a section-title-only row (12, "FABRICS" in col A alone)
// between its column-header row (11) and its first real item (13): every
// other section folds the section title into column A of the header row
// itself, so their first real item follows immediately after.
//
// computedFormulas: this section's structural formulas, written fresh on
// EVERY row (original and newly inserted alike) rather than trusted from
// whatever the template currently contains. This matters because the
// live template turned out to be inconsistent/stale in several ways: some
// Fabric rows never had a CNF (R/S) formula at all, the FOB-by-Yds (P)
// formula only existed on some rows, and Label's "CNF" column held
// leftover hardcoded fractions (e.g. "=13/1000") from a previous style
// rather than a real FOB*1.1 relationship. Explicitly (re)writing every
// row's formulas here removes that dependency entirely.
//   M  (Total Cost)            = (K+K*L)*J                      - every section
//   P  (FOB Price by Yds)      = Fob Price (M) * 0.9144            - Fabric only
//   R  (CNF price (M))         = Fob Price (M) + Including transport Cost   - Fabric only
//   S  (CNF price by Yds)      = FOB Price by Yds + Including transport Cost - Fabric only
//   O  (CNF, Label/Trim)       = FOB * 1.1                        - Label & Trim only
// Fabric's O/Q (Fob Price (M), Including transport Cost) are
// manual-entry-only per business rule and are deliberately left out of
// computedFormulas - the row-clear below wipes them to a blank, unstyled
// cell like any other plain data column.
const KARIBAN_TOTAL_COST_FORMULA = r => `(K${r}+K${r}*L${r})*J${r}`;
// Sewing thread is intentionally excluded from extraction (per business
// rule, its per-style consumption isn't reliably readable from the tech
// pack), and instead represented by this single fixed line, always
// appended after whatever real Trim items were extracted - taken verbatim
// from the template's own row 35 (including its "Sewig Thread" spelling).
const KARIBAN_TRIM_STATIC_ROWS = [{
  item: 'Sewig Thread', code: '', position: '', supplier: 'Coats',
  priceTerms: 'LOCAL', paymentTerm: 'LC 120 DAYS',
  description: 'Sewig Thread--40/3 (40Tex)',
  price: 1.24, priceFmt: '"$"#,##0.00',
  consumption: 350 / 3000, consumptionFmt: '0.00" mtr"',
  wastage: 0.15,
}];
const KARIBAN_SECTIONS = {
  Fabric: {
    headerRow: 11, firstBlank: 13, lastBlank: 17, totalRow: 18, weightCol: 7, sizeWidthCol: null,
    greenCols: [15, 16, 17, 18, 19], // O-S: the FOB/CNF block, kept visually distinct with a green fill
    computedFormulas: [
      { col: 13, formula: KARIBAN_TOTAL_COST_FORMULA },     // M
      { col: 16, formula: r => `O${r}*0.9144` },            // P - FOB Price by Yds = Fob Price (M) * 0.9144
      { col: 18, formula: r => `O${r}+Q${r}` },             // R - CNF price (M)
      { col: 19, formula: r => `P${r}+Q${r}` },             // S - CNF price by Yds
    ],
  },
  Label: {
    headerRow: 19, firstBlank: 20, lastBlank: 30, totalRow: 31, weightCol: null, sizeWidthCol: null,
    computedFormulas: [
      { col: 13, formula: KARIBAN_TOTAL_COST_FORMULA },     // M
      { col: 15, formula: r => `N${r}*1.1` },               // O - CNF = FOB * 1.1
    ],
  },
  Trim: {
    headerRow: 32, firstBlank: 33, lastBlank: 35, totalRow: 36, weightCol: null, sizeWidthCol: 8,
    staticRows: KARIBAN_TRIM_STATIC_ROWS,
    computedFormulas: [
      { col: 13, formula: KARIBAN_TOTAL_COST_FORMULA },     // M
      { col: 15, formula: r => `N${r}*1.1` },               // O - CNF = FOB * 1.1
    ],
  },
  Packaging: {
    headerRow: 40, firstBlank: 41, lastBlank: 52, totalRow: 53, weightCol: null, sizeWidthCol: null,
    computedFormulas: [
      { col: 13, formula: KARIBAN_TOTAL_COST_FORMULA },     // M only - no FOB/CNF block here
    ],
  },
};
const KARIBAN_COL = { item: 1, itemCode: 2, position: 3, description: 9, consumption: 11 };
const KARIBAN_PRICE_COL = 10; // J - Price - always left blank; no section writes a price value
const KARIBAN_CONSUMPTION_COL = 11; // K
const KARIBAN_TOTAL_COST_COL = 13; // M - "(K+K*L)*J", now written via computedFormulas above
const KARIBAN_WASTAGE_COL = 12; // L - 5% default, 15% thread, 0% discount (see karibanWastageFor)

function karibanWastageFor(item) {
  const text = `${item.item || ''} ${item.description || ''} ${item.code || ''}`;
  if (/\bthread\b/i.test(text)) return 0.15;
  if (/\bdiscount\b/i.test(text)) return 0;
  return 0.05;
}

function fillKaribanSection(ws, cfg, items, maxSheetRow) {
  const first = cfg.firstBlank;
  let last = cfg.lastBlank;
  const available = last - first + 1;
  const staticRows = cfg.staticRows || [];
  const needed = items.length + staticRows.length;
  const amount = Math.max(0, needed - available);

  const oldMerges = snapshotAndUnmergeAllKariban(ws);
  const rowTemplateMerges = oldMerges.filter(m => m.minRow === first).map(m => ({ minCol: m.minCol, maxCol: m.maxCol }));

  let insertAt = null;
  if (amount > 0) {
    insertAt = last;
    const blankRows = Array.from({ length: amount }, () => []);
    ws.spliceRows(insertAt, 0, ...blankRows);
    shiftWholeSheetFormulasKariban(ws, insertAt, amount, maxSheetRow + amount, KARIBAN_MAX_COL);
    for (let i = 0; i < amount; i++) {
      const newRow = insertAt + i;
      copyRowStyleKariban(ws, insertAt + amount, newRow, KARIBAN_MAX_COL);
    }
    maxSheetRow += amount;
    last += amount;
    cfg.totalRow += amount;
  }

  // The template's own example rows carry real values (and, we've found,
  // sometimes stale/incorrect formulas - e.g. Label's "CNF" column held
  // leftover hardcoded fractions from a previous style rather than a real
  // FOB*1.1 relationship, and Fabric's CNF columns didn't always have a
  // formula at all) for a different, unrelated style. Every cell in the
  // section is wiped unconditionally - value, number format, AND fill -
  // so no leftover Accounting/Currency format or highlight color (e.g.
  // the two orange "Super Dry"/"Carton" example rows in Packaging) can
  // bleed onto a newly-written row. cfg.computedFormulas below then
  // writes this sheet's actual structural formulas back in fresh,
  // immediately after, so nothing is left uncalculated.
  const NO_FILL = { type: 'pattern', pattern: 'none' };
  const CENTER_MIDDLE = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const GREEN_DATA_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5E0B4' } };
  for (let r = first; r <= last; r++) {
    for (let c = 1; c <= KARIBAN_MAX_COL; c++) {
      const cell = ws.getCell(r, c);
      cell.value = null;
      cell.numFmt = 'General';
      // FOB/CNF block (Fabric only) keeps its green highlight instead of
      // going white like every other cleared cell - matches the
      // template's own visual marker for this manual/formula input block.
      cell.fill = (cfg.greenCols && cfg.greenCols.includes(c)) ? GREEN_DATA_FILL : NO_FILL;
      // Item Code (Fabric Code / Item Code) and Position read cleaner
      // centered, both horizontally and vertically, than left/top-aligned
      // like the rest of the row - set across the whole column so it's
      // consistent even before any item lands in a given row.
      if (c === KARIBAN_COL.itemCode || c === KARIBAN_COL.position) {
        cell.alignment = CENTER_MIDDLE;
      } else if (c === KARIBAN_COL.description) {
        // Applied blanket-wide (not just on populated rows) so a section's
        // unused capacity rows (e.g. Packaging, which rarely fills all 12
        // slots) don't keep the template's own inherited alignment.
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      } else {
        // Every other data cell still gets wrap text (per business rule:
        // every cell below the sheet's row-10 intro block wraps), just
        // without forcing center alignment.
        cell.alignment = Object.assign({}, cell.alignment, { wrapText: true });
      }
    }
  }

  // Header row (e.g. "Item Code", "Position", "Description" ...) reads
  // cleaner centered - and, for Fabric, its green FOB/CNF header cells
  // (O-S) keep the template's bright-green marker rather than the paler
  // shade used on the data rows below them.
  const GREEN_HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF92D050' } };
  for (let c = 1; c <= KARIBAN_MAX_COL; c++) {
    const headerCell = ws.getCell(cfg.headerRow, c);
    headerCell.alignment = CENTER_MIDDLE;
    if (cfg.greenCols && cfg.greenCols.includes(c)) headerCell.fill = GREEN_HEADER_FILL;
  }

  // Re-apply this section's own structural formulas fresh on every row
  // (original and newly-inserted alike) - see KARIBAN_SECTIONS above for
  // exactly which columns/relationships each section defines.
  for (const { col, formula } of (cfg.computedFormulas || [])) {
    for (let r = first; r <= last; r++) {
      setFormulaKariban(ws.getCell(r, col), formula(r));
    }
  }

  items.forEach((item, i) => {
    const r = first + i;
    ws.getCell(r, KARIBAN_COL.item).value = item.item || null;
    ws.getCell(r, KARIBAN_COL.itemCode).value = item.code || null;
    ws.getCell(r, KARIBAN_COL.position).value = item.position || null;
    const descCell = ws.getCell(r, KARIBAN_COL.description);
    descCell.value = item.description || null;
    // Left-aligned but vertically centered, with wrap text so a long
    // description doesn't overflow into neighboring cells.
    descCell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
    // Wastage (%): 5% default, 15% for thread, 0% for discount - matched
    // against the item's own text so it applies regardless of section.
    const wastageCell = ws.getCell(r, KARIBAN_WASTAGE_COL);
    wastageCell.value = karibanWastageFor(item);
    wastageCell.numFmt = '0.00%';
    const consCell = ws.getCell(r, KARIBAN_COL.consumption);
    const hasQty = item.consumption_qty !== undefined && item.consumption_qty !== null && item.consumption_qty !== '';
    consCell.value = hasQty ? Number(item.consumption_qty) : null;
    // Custom format embeds the unit directly into the cell's own display
    // (e.g. "4 Pcs", "2.35 Yds") instead of writing it as separate text -
    // the stored value stays a plain number, so the Total Cost formula's
    // K reference keeps working exactly as before.
    if (hasQty) {
      const unit = item.consumption_unit === 'Yds' ? 'Yds' : 'Pcs';
      consCell.numFmt = unit === 'Yds' ? '0.00" Yds"' : '0" Pcs"';
    }
    if (cfg.weightCol && item.weight_gsm) ws.getCell(r, cfg.weightCol).value = `${item.weight_gsm} gsm`;
    if (cfg.sizeWidthCol && item.size_mm) ws.getCell(r, cfg.sizeWidthCol).value = `${item.size_mm} mm`;
  });

  // Fixed rows (currently just Trim's "Sewig Thread" line) go right after
  // the real extracted items - always present regardless of what was
  // extracted, since they represent business-constant values rather than
  // BOM data. Their own Description cell already got left-middle+wrap
  // alignment from the blanket per-row pass above, same as every other row.
  staticRows.forEach((sr, i) => {
    const r = first + items.length + i;
    ws.getCell(r, KARIBAN_COL.item).value = sr.item || null;
    ws.getCell(r, KARIBAN_COL.itemCode).value = sr.code || null;
    ws.getCell(r, KARIBAN_COL.position).value = sr.position || null;
    if (sr.supplier) ws.getCell(r, 4).value = sr.supplier;           // D - Supplier
    if (sr.priceTerms) ws.getCell(r, 6).value = sr.priceTerms;       // F - Price Terms
    if (sr.paymentTerm) ws.getCell(r, 7).value = sr.paymentTerm;     // G - Payment Term
    ws.getCell(r, KARIBAN_COL.description).value = sr.description || null;
    if (sr.price !== undefined) {
      const priceCell = ws.getCell(r, KARIBAN_PRICE_COL);
      priceCell.value = sr.price;
      if (sr.priceFmt) priceCell.numFmt = sr.priceFmt;
    }
    if (sr.consumption !== undefined) {
      const consCell = ws.getCell(r, KARIBAN_COL.consumption);
      consCell.value = sr.consumption;
      if (sr.consumptionFmt) consCell.numFmt = sr.consumptionFmt;
    }
    if (sr.wastage !== undefined) {
      const wastageCell = ws.getCell(r, KARIBAN_WASTAGE_COL);
      wastageCell.value = sr.wastage;
      wastageCell.numFmt = '0.00%';
    }
  });

  if (amount > 0) {
    remergeShiftedKariban(ws, oldMerges, insertAt, amount);
    for (let i = 0; i < amount; i++) {
      const newRow = insertAt + i;
      for (const pat of rowTemplateMerges) mergeSectionRangeKariban(ws, newRow, pat.minCol, pat.maxCol - pat.minCol + 1);
    }
  } else {
    remergeShiftedKariban(ws, oldMerges, 0, 0);
  }

  return { maxSheetRow, totalRow: cfg.totalRow };
}

async function buildKaribanCostSheet(templateArrayBuffer, itemsByBucket, styleInfo) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateArrayBuffer.slice(0));
  const ws = wb.worksheets[0];
  ws._media = []; // strip every embedded image from the template - not wanted in generated sheets
  const originalRowCount = ws.rowCount;
  let maxRow = originalRowCount;
  normalizeFormulasKariban(ws, maxRow, KARIBAN_MAX_COL);

  // Date: always the date this Cost Sheet is generated (today), not
  // something read off the tech pack. The cell's own d-mmm-yyyy display
  // format already on the template is left exactly as-is - only the value
  // changes.
  ws.getCell('D4').value = new Date();

  if (styleInfo) {
    if (styleInfo.brand) ws.getCell('D6').value = styleInfo.brand;
    // "Designation (EN)" from the tech pack's page-1 Properties table
    // (e.g. "Unisex high-visibility reversible jacket") is what belongs in
    // the "Item Description" cell - not a separate itemDescription field,
    // which nothing in extraction ever populated.
    if (styleInfo.styleDesignation) ws.getCell('D7').value = styleInfo.styleDesignation;
    if (styleInfo.styleCode) ws.getCell('D8').value = styleInfo.styleCode;
  }

  // bottom-to-top so inserting rows in a lower section never shifts a
  // section still queued for processing above it
  const order = ['Packaging', 'Trim', 'Label', 'Fabric'];
  const counts = {};
  for (const bucket of order) {
    const cfg = { ...KARIBAN_SECTIONS[bucket] };
    const items = itemsByBucket[bucket] || [];
    counts[bucket] = items.length;
    ({ maxSheetRow: maxRow } = fillKaribanSection(ws, cfg, items, maxRow));
  }

  // Every cell below the sheet's row-10 intro/header block gets wrap text -
  // including rows fillKaribanSection never touches directly, like the
  // untouched Embroidery & Print block and the summary/commission rows at
  // the bottom. Merged (not overwritten) onto whatever alignment a cell
  // already has, so the center-middle/left-middle alignment already set
  // above for headers, Item Code/Position, and Description survives.
  for (let r = 11; r <= maxRow; r++) {
    for (let c = 1; c <= KARIBAN_MAX_COL; c++) {
      const cell = ws.getCell(r, c);
      cell.alignment = Object.assign({}, cell.alignment, { wrapText: true });
    }
  }

  // Buyer (D5, e.g. "KARIBAN BRANDS") - forced left-aligned explicitly,
  // done here (after every section's own merge unmerge/remerge cycle
  // above, which touches every merge in the whole sheet including this
  // one) rather than earlier, where it would just get reset.
  ws.getCell('D5').alignment = Object.assign({}, ws.getCell('D5').alignment, { horizontal: 'left' });

  if (styleInfo && styleInfo.styleCode) {
    ws.name = styleInfo.styleCode.replace(/[\\/?*\[\]:]/g, '-').slice(0, 31) || 'Sheet';
  }

  const buffer = await wb.xlsx.writeBuffer();
  return { buffer, counts };
}

/* ============================================================
   KARIBAN — SUPPLY CHAIN SHEET
   Multiple tech packs at once -> one flat sheet: Style No | Item Code |
   Description | Picture | Supplier. Blank row after every style's block.
   Duplicate items (same Item Code) are kept only once, globally, across
   every style processed in the same run. Zipper merge follows the same
   5-part rule as the Cost Sheet.
   ============================================================ */

async function buildKaribanSupplyChainWorkbook(sessions) {
  // sessions: [{ styleCode, items: [...] }]
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Supply Chain');
  ws.columns = [
    { header: 'Style No', key: 'style', width: 20 },
    { header: 'Item Code', key: 'code', width: 26 },
    { header: 'Description', key: 'desc', width: 85 },
    { header: 'Picture', key: 'pic', width: 16 },
    { header: 'Supplier', key: 'supplier', width: 22 },
  ];
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const seenGlobal = new Set(); // dedupe by normalized Description across every style in this run
  for (const session of sessions) {
    let wroteAny = false;
    for (const item of session.items) {
      const key = karibanDedupeKey(item.description);
      if (!key || seenGlobal.has(key)) continue;
      seenGlobal.add(key);
      const row = ws.addRow({ style: session.styleCode, code: item.code, desc: item.description, pic: '', supplier: '' });
      // Wrap long descriptions instead of letting them overflow past the
      // column width, so the sheet stays readable at a glance.
      row.getCell('desc').alignment = { wrapText: true, vertical: 'top' };
      wroteAny = true;
    }
    if (wroteAny) ws.addRow({}); // blank separator after this style's block
  }

  return await wb.xlsx.writeBuffer();
}

/* ============================================================
   KARIBAN — UI WIRING
   ============================================================ */
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

document.getElementById('brandKariban').addEventListener('click', () => {
  requestUnlock('kariban', () => {
    homeView.hidden = true;
    karibanHubView.hidden = false;
  });
});
document.getElementById('karibanHubBackBtn').addEventListener('click', () => {
  karibanHubView.hidden = true;
  homeView.hidden = false;
});
document.getElementById('karibanHubBomCard').addEventListener('click', () => {
  karibanHubView.hidden = true;
  karibanView.hidden = false;
});
document.getElementById('karibanHubSupplyCard').addEventListener('click', () => {
  karibanHubView.hidden = true;
  karibanSupplyView.hidden = false;
});
document.getElementById('backBtnKariban').addEventListener('click', () => {
  karibanView.hidden = true;
  karibanHubView.hidden = false;
});
document.getElementById('backBtnKaribanSupply').addEventListener('click', () => {
  karibanSupplyView.hidden = true;
  karibanHubView.hidden = false;
});

/* ---------------- BOM Generator view ---------------- */
const dropzoneKariban = document.getElementById('dropzoneKariban');
const fileInputKariban = document.getElementById('fileInputKariban');
const filebarKariban = document.getElementById('filebarKariban');
const fileNameKaribanEl = document.getElementById('fileNameKariban');
const clearFileKariban = document.getElementById('clearFileKariban');
const processBtnKariban = document.getElementById('processBtnKariban');
const processLabelKariban = document.getElementById('processLabelKariban');
const statusKaribanEl = document.getElementById('statusKariban');
const resultsKariban = document.getElementById('resultsKariban');
const downloadBtnKariban = document.getElementById('downloadBtnKariban');
const cbdBtnKariban = document.getElementById('cbdBtnKariban');
const cbdLabelKariban = document.getElementById('cbdLabelKariban');
const cbdStatusKariban = document.getElementById('cbdStatusKariban');

let currentFileKariban = null;
let currentKaribanItems = null;   // flat business-rule-applied item list
let currentKaribanStyle = null;   // { styleCode, styleDesignation, brand }

function setStatusKariban(msg, cls) {
  statusKaribanEl.textContent = msg;
  statusKaribanEl.className = 'status' + (cls ? ' ' + cls : '');
}

function setFileKariban(file) {
  if (!file || file.type !== 'application/pdf') {
    setStatusKariban('Please choose a PDF file.', 'err');
    return;
  }
  currentFileKariban = file;
  fileNameKaribanEl.textContent = file.name;
  filebarKariban.classList.add('show');
  processBtnKariban.disabled = false;
  setStatusKariban('');
  resultsKariban.classList.remove('show');
  currentKaribanItems = null;
}

dropzoneKariban.addEventListener('dragover', e => { e.preventDefault(); dropzoneKariban.classList.add('drag'); });
dropzoneKariban.addEventListener('dragleave', () => dropzoneKariban.classList.remove('drag'));
dropzoneKariban.addEventListener('drop', e => {
  e.preventDefault(); dropzoneKariban.classList.remove('drag');
  if (e.dataTransfer.files.length) setFileKariban(e.dataTransfer.files[0]);
});
fileInputKariban.addEventListener('change', e => {
  if (e.target.files.length) setFileKariban(e.target.files[0]);
});
clearFileKariban.addEventListener('click', e => {
  e.stopPropagation();
  currentFileKariban = null; fileInputKariban.value = '';
  currentKaribanItems = null; currentKaribanStyle = null;
  filebarKariban.classList.remove('show');
  processBtnKariban.disabled = true;
  resultsKariban.classList.remove('show');
  setStatusKariban('');
});

processBtnKariban.addEventListener('click', async () => {
  if (!currentFileKariban) return;
  processBtnKariban.disabled = true;
  processBtnKariban.classList.add('loading');
  resultsKariban.classList.remove('show');
  try {
    const { rows, styleCode, styleDesignation, brand } = await extractKaribanPlacementRows(currentFileKariban, (p, total) => {
      processLabelKariban.textContent = `Scanning page ${p} / ${total}...`;
      setStatusKariban(`Scanning page ${p} of ${total}...`);
    });
    const items = applyKaribanBusinessRules(rows);
    currentKaribanItems = items;
    currentKaribanStyle = { styleCode, styleDesignation, brand };
    if (items.length === 0) {
      setStatusKariban('No Fabrics / Trims / Label / Packaging items were found in this PDF.', 'err');
    } else {
      const by = karibanItemsByBucket(items);
      const summary = `Fabric: ${by.Fabric.length}, Trim: ${by.Trim.length}, Label: ${by.Label.length}, Packaging: ${by.Packaging.length}`;
      setStatusKariban(`Done — Style ${styleCode || '(not found)'}: ${items.length} items extracted (${summary}).`, 'ok');
      resultsKariban.classList.add('show');
    }
  } catch (err) {
    console.error(err);
    setStatusKariban('Something went wrong reading this PDF: ' + err.message, 'err');
  } finally {
    processBtnKariban.disabled = false;
    processBtnKariban.classList.remove('loading');
    processLabelKariban.textContent = 'Extract to Excel';
  }
});

downloadBtnKariban.addEventListener('click', () => {
  if (!currentKaribanItems) return;
  const wsData = [['Bucket', 'Item', 'Item Code', 'Position', 'Description', 'Consumption', 'Unit']];
  currentKaribanItems.forEach(r => wsData.push([r.bucket, r.item || '', r.code || '', r.position || '', r.description || '', r.consumption_qty, r.consumption_unit]));
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{ wch: 12 }, { wch: 20 }, { wch: 16 }, { wch: 26 }, { wch: 90 }, { wch: 12 }, { wch: 8 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Extracted Data');
  const baseName = (currentKaribanStyle && currentKaribanStyle.styleCode) || (currentFileKariban.name || 'tech_pack').replace(/\.pdf$/i, '');
  XLSX.writeFile(wb, `${baseName}_extracted.xlsx`);
});

cbdBtnKariban.addEventListener('click', async () => {
  if (!currentKaribanItems) return;
  cbdBtnKariban.disabled = true;
  cbdBtnKariban.classList.add('loading');
  cbdLabelKariban.textContent = 'Building Cost Sheet...';
  cbdStatusKariban.textContent = '';
  cbdStatusKariban.className = 'status';
  try {
    const templateBuffer = base64ToArrayBufferKariban(KARIBAN_TEMPLATE_B64);
    const itemsByBucket = karibanItemsByBucket(currentKaribanItems);
    const { buffer, counts } = await buildKaribanCostSheet(templateBuffer, itemsByBucket, currentKaribanStyle);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const baseName = (currentKaribanStyle && currentKaribanStyle.styleCode) || 'kariban_style';
    const a = document.createElement('a');
    a.href = url;
    a.download = `${baseName}_Cost_Sheet.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    cbdStatusKariban.textContent = `Done — Fabric: ${counts.Fabric}, Trim: ${counts.Trim}, Label: ${counts.Label}, Packaging: ${counts.Packaging}.`;
    cbdStatusKariban.className = 'status ok';
  } catch (err) {
    console.error(err);
    cbdStatusKariban.textContent = 'Something went wrong building the Cost Sheet: ' + err.message;
    cbdStatusKariban.className = 'status err';
  } finally {
    cbdBtnKariban.disabled = false;
    cbdBtnKariban.classList.remove('loading');
    cbdLabelKariban.textContent = 'Generate Cost Sheet Excel';
  }
});

/* ---------------- Supply Chain view ---------------- */
const dropzoneKaribanSupply = document.getElementById('dropzoneKaribanSupply');
const fileInputKaribanSupply = document.getElementById('fileInputKaribanSupply');
const fileListKaribanSupply = document.getElementById('fileListKaribanSupply');
const processBtnKaribanSupply = document.getElementById('processBtnKaribanSupply');
const processLabelKaribanSupply = document.getElementById('processLabelKaribanSupply');
const statusKaribanSupplyEl = document.getElementById('statusKaribanSupply');
const resultsKaribanSupply = document.getElementById('resultsKaribanSupply');
const countsKaribanSupply = document.getElementById('countsKaribanSupply');
const downloadBtnKaribanSupply = document.getElementById('downloadBtnKaribanSupply');

let currentKaribanSupplyFiles = [];
let currentKaribanSupplyBuffer = null;

function setStatusKaribanSupply(msg, cls) {
  statusKaribanSupplyEl.textContent = msg;
  statusKaribanSupplyEl.className = 'status' + (cls ? ' ' + cls : '');
}

function escapeHtmlKariban(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function renderKaribanSupplyFileList() {
  fileListKaribanSupply.innerHTML = '';
  currentKaribanSupplyFiles.forEach((file, idx) => {
    const row = document.createElement('div');
    row.className = 'filebar show';
    row.style.marginBottom = '8px';
    row.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      <span class="name">${escapeHtmlKariban(file.name)}</span>
      <span class="clear">✕ remove</span>`;
    row.querySelector('.clear').addEventListener('click', (e) => {
      e.stopPropagation();
      currentKaribanSupplyFiles.splice(idx, 1);
      renderKaribanSupplyFileList();
    });
    fileListKaribanSupply.appendChild(row);
  });
  processBtnKaribanSupply.disabled = currentKaribanSupplyFiles.length === 0;
  resultsKaribanSupply.classList.remove('show');
  currentKaribanSupplyBuffer = null;
}

function addKaribanSupplyFiles(fileList) {
  const pdfs = Array.from(fileList).filter(f => f.type === 'application/pdf');
  if (pdfs.length === 0) {
    setStatusKaribanSupply('Please choose PDF file(s).', 'err');
    return;
  }
  currentKaribanSupplyFiles = currentKaribanSupplyFiles.concat(pdfs);
  setStatusKaribanSupply('');
  renderKaribanSupplyFileList();
}

dropzoneKaribanSupply.addEventListener('dragover', e => { e.preventDefault(); dropzoneKaribanSupply.classList.add('drag'); });
dropzoneKaribanSupply.addEventListener('dragleave', () => dropzoneKaribanSupply.classList.remove('drag'));
dropzoneKaribanSupply.addEventListener('drop', e => {
  e.preventDefault(); dropzoneKaribanSupply.classList.remove('drag');
  if (e.dataTransfer.files.length) addKaribanSupplyFiles(e.dataTransfer.files);
});
fileInputKaribanSupply.addEventListener('change', e => {
  if (e.target.files.length) addKaribanSupplyFiles(e.target.files);
  fileInputKaribanSupply.value = '';
});

processBtnKaribanSupply.addEventListener('click', async () => {
  if (currentKaribanSupplyFiles.length === 0) return;
  processBtnKaribanSupply.disabled = true;
  processBtnKaribanSupply.classList.add('loading');
  resultsKaribanSupply.classList.remove('show');
  try {
    const sessions = [];
    let totalItems = 0;
    for (let i = 0; i < currentKaribanSupplyFiles.length; i++) {
      const file = currentKaribanSupplyFiles[i];
      processLabelKaribanSupply.textContent = `Extracting ${file.name} (${i + 1}/${currentKaribanSupplyFiles.length})...`;
      const { rows, styleCode } = await extractKaribanPlacementRows(file, (p, total) => {
        setStatusKaribanSupply(`${file.name}: scanning page ${p} of ${total}...`);
      });
      const items = applyKaribanBusinessRules(rows);
      totalItems += items.length;
      sessions.push({ styleCode: styleCode || file.name.replace(/\.pdf$/i, ''), items });
    }
    if (totalItems === 0) {
      setStatusKaribanSupply('No Fabrics / Trims / Label / Packaging items were found in these PDFs.', 'err');
    } else {
      currentKaribanSupplyBuffer = await buildKaribanSupplyChainWorkbook(sessions);
      const styleList = sessions.map(s => s.styleCode).join(', ');
      countsKaribanSupply.textContent = `${sessions.length} tech pack(s) processed — Style${sessions.length > 1 ? 's' : ''}: ${styleList}. Duplicate items (same Item Code) are kept once, across every style.`;
      resultsKaribanSupply.classList.add('show');
      setStatusKaribanSupply(`Done — processed ${sessions.length} tech pack(s).`, 'ok');
    }
  } catch (err) {
    console.error(err);
    setStatusKaribanSupply('Something went wrong reading these PDFs: ' + err.message, 'err');
  } finally {
    processBtnKaribanSupply.disabled = false;
    processBtnKaribanSupply.classList.remove('loading');
    processLabelKaribanSupply.textContent = 'Extract & Generate Supply Chain Sheet';
  }
});

downloadBtnKaribanSupply.addEventListener('click', () => {
  if (!currentKaribanSupplyBuffer) return;
  const blob = new Blob([currentKaribanSupplyBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Kariban_Supply_Chain_Sheet.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});
