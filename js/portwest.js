/* ============================================================
   PORTWEST ENGINE
   PDF "Placements" (BOM) table extraction + Cost Sheet merge.
   Depends on: shared.js (homeView)
   ============================================================ */

// Decodes a base64 string (the embedded Portwest_Format.xlsx template,
// see PORTWEST_COST_SHEET_TEMPLATE_B64 near the bottom of this file) into
// an ArrayBuffer, exactly the shape ExcelJS's workbook.xlsx.load expects.
// Embedded rather than fetched as a separate asset so the tool has no
// external file dependency at all — matches how the other engines in
// this project ship their own templates.
function portwestBase64ToArrayBuffer(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

/* ============================================================
   PORTWEST — PDF EXTRACTION ENGINE
   ============================================================ */

// Fallback column boundaries, expressed as FRACTIONS of page width (not
// fixed points), since we don't have a calibrated real-world sample to
// anchor absolute point values against the way the other engines do.
// These are only used if a page's own header row can't be parsed (see
// derivePortwestColumnBoundaries) — the header-anchor path is primary.
const PORTWEST_COLUMNS_FALLBACK_FRACTIONS = [
  { key: 'placement',      max: 0.082 },
  { key: 'mainMaterial',   max: 0.135 },
  { key: 'code',           max: 0.211 },
  { key: 'image',          max: 0.287 },
  { key: 'product',        max: 0.372 },
  { key: 'description',    max: 0.476 },
  { key: 'composition',    max: 0.530 },
  { key: 'materialWeight', max: 0.575 },
  { key: 'commonSize',     max: 0.641 },
  { key: 'qtyDefault',     max: 0.687 },
  { key: 'uom',            max: 0.726 },
  { key: 'position',       max: 0.818 },
  { key: 'comment',        max: Infinity }, // swallows Comment + any colorway swatch columns — none of that is needed downstream
];

// Columns the merge engine actually reads (buildPortwestCostSheet never
// touches mainMaterial, image, product, composition, comment, or any
// colorway columns). Anchoring only needs to be reliable for these.
const PORTWEST_REQUIRED_COLUMNS = [
  'placement', 'code', 'description', 'materialWeight', 'commonSize', 'qtyDefault', 'uom', 'position',
];

function portwestColumnsFromFractions(pageWidth) {
  return PORTWEST_COLUMNS_FALLBACK_FRACTIONS.map(c => ({ key: c.key, max: c.max === Infinity ? Infinity : c.max * pageWidth }));
}

// Derives this page's column boundaries from its own header row (same
// technique as Malacca's engine) rather than assuming a fixed layout —
// the number of colorway columns (Orange/Yellow/etc.) varies by style, and
// that shifts everything to the right of "Position"/"Comment". Anchoring
// on each required column's own label makes extraction immune to that.
function derivePortwestColumnBoundaries(headerItems) {
  function anchorX(prefix, occurrence) {
    const matches = headerItems.filter(it => it.text.startsWith(prefix)).sort((a, b) => a.x - b.x);
    return matches[occurrence] !== undefined ? matches[occurrence].x : null;
  }
  const candidates = [
    ['placement',      anchorX('Place', 0)],
    ['mainMaterial',   anchorX('Main', 0)],
    ['code',           anchorX('Code', 0)],
    ['image',          anchorX('Imag', 0)],
    ['product',        anchorX('Produ', 0)],
    ['description',    anchorX('Descr', 0)],
    ['composition',    anchorX('Compo', 0)],
    ['materialWeight', anchorX('Mater', 0)],
    ['commonSize',     anchorX('Common', 0)],
    ['qtyDefault',     anchorX('Qty', 0)],
    ['uom',            anchorX('UOM', 0)],
    ['position',       anchorX('Posi', 0)],
    ['comment',        anchorX('Comment', 0)],
  ];

  const present = candidates.filter(([, x]) => x !== null);
  const foundKeys = new Set(present.map(([key]) => key));
  for (const key of PORTWEST_REQUIRED_COLUMNS) {
    if (!foundKeys.has(key)) return null; // couldn't parse this header — caller falls back
  }

  present.sort((a, b) => a[1] - b[1]);
  return present.map(([key, x], i) => ({
    key,
    max: i + 1 < present.length ? x + (present[i + 1][1] - x) * 0.6 : Infinity,
  }));
}

function classifyPortwestColumn(x, columns) {
  for (const c of columns) if (x < c.max) return c.key;
  return columns.length ? columns[columns.length - 1].key : 'comment';
}

function portwestClusterLines(items, tol) {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [];
  for (const it of sorted) {
    const last = lines[lines.length - 1];
    let line = last && Math.abs(last.y - it.y) < tol ? last : null;
    if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
    line.items.push(it);
  }
  for (const l of lines) l.items.sort((a, b) => a.x - b.x);
  return lines;
}

// A category divider row, e.g. "Fabric (6)", "Stitching (8)", "Trims (16)",
// "Labels (7)", "Packaging (13)" — sits directly above that category's
// first row (whether that's at the top of a page or mid-page). Restricted
// to short lines (few items) to avoid false-matching a Position/Comment
// free-text line that happens to end in "(something)".
const PORTWEST_CATEGORY_TAG_RE = /^([A-Za-z][A-Za-z \/&]{1,30})\s*\((\d{1,3})\)$/;

function matchPortwestCategoryTag(line) {
  if (line.items.length > 3) return null;
  const text = line.items.map(i => i.text).join(' ').trim().replace(/\s+/g, ' ');
  const m = text.match(PORTWEST_CATEGORY_TAG_RE);
  return m ? m[1].trim() : null;
}

// Items redirected to Fabric regardless of which tech-pack category they
// were tagged under (per buyer's costing convention).
const PORTWEST_FABRIC_OVERRIDE_RE = /interlin|interfac|fusible|seam\s*tape/i;

function portwestIsFabricOverride(row) {
  return PORTWEST_FABRIC_OVERRIDE_RE.test(row.placement || '') ||
    PORTWEST_FABRIC_OVERRIDE_RE.test(row.product || '') ||
    PORTWEST_FABRIC_OVERRIDE_RE.test(row.description || '');
}

// Cover page (page 1) "Properties" table is a label/value grid — each
// visual row can hold two label:value pairs (a left pair and a right
// pair) side by side. We only need Style Code and Description, both on
// the left-hand pair, so we anchor on the known label vocabulary rather
// than assuming fixed coordinates.
const PORTWEST_COVER_LABELS = [
  'Product Manager', 'Style', 'Style Code', 'Description', 'Collection', 'Range',
  'Size Range', 'Sizes', 'Default Size', 'Created', 'Created By', 'Modified', 'Modified By',
  'Authority BOM', 'Marketing Main Material', 'Composition', 'Material Weight',
  'Product Type', 'Standards', 'Climate', 'Gender', 'Authority Size Chart',
];

function extractPortwestCoverInfo(items) {
  // Zone below the "Properties" title, above "Carton Information"/"Colourways".
  const zoneItems = items.filter(it => it.y > 40 && it.y < 450);
  if (!zoneItems.length) return null;

  const lines = portwestClusterLines(zoneItems, 3);
  const values = {};
  // Longest label first so a regex alternation match prefers "Style Code"
  // over "Style" when both could start matching at the same position.
  const labelsByLengthDesc = [...PORTWEST_COVER_LABELS].sort((a, b) => b.length - a.length);
  const labelPattern = new RegExp(labelsByLengthDesc.map(l => l.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'), 'g');

  for (const line of lines) {
    // Reconstruct this line's full text by joining items with a single
    // space. Deliberately NOT item-index-based: a label like "Style Code"
    // may arrive from pdf.js as one combined text run OR as two separate
    // word-level runs depending on how the source PDF was generated —
    // matching against the flattened string handles either shape
    // identically, whereas matching label token-count against a fixed
    // number of consecutive items (the previous approach) silently failed
    // whenever a label happened to be a single run instead of split words
    // (confirmed against a real generated PDF: every multi-word label
    // failed to match, causing Style Code to come back empty and
    // Description to swallow all the way to the end of the line).
    const lineText = line.items.map(it => it.text).join(' ');
    const hits = [...lineText.matchAll(labelPattern)];
    if (!hits.length) continue;

    for (let h = 0; h < hits.length; h++) {
      const label = hits[h][0];
      const startIdx = hits[h].index + label.length;
      const endIdx = h + 1 < hits.length ? hits[h + 1].index : lineText.length;
      const text = lineText.slice(startIdx, endIdx).trim();
      if (text) values[label] = (values[label] ? values[label] + ' ' : '') + text;
    }
  }

  return {
    styleCode: values['Style Code'] || '',
    description: values['Description'] || '',
  };
}

// Boilerplate lines that appear on every Placements page outside the
// table itself (page title, "Displaying X-Y of N results" counter, and
// the "Tech Pack - <code>" / "Page N of M" footer). Matched by content
// rather than relying solely on a y-coordinate cutoff, since we don't
// have a real calibrated PDF sample to pin exact point values against —
// this is a defense-in-depth check so none of this text can ever be
// mistaken for a table row even if the y-band guess is slightly off.
const PORTWEST_BOILERPLATE_RE = /^(Placements|Displaying\s+\d+\s*-\s*\d+\s+of\s+\d+\s+results|Tech Pack\b.*|Page\s+\d+\s+of\s+\d+)$/i;

function isPortwestBoilerplateLine(line) {
  const text = line.items.map(i => i.text).join(' ').trim().replace(/\s+/g, ' ');
  return PORTWEST_BOILERPLATE_RE.test(text);
}

// Some tech packs paginate a very wide Placements table (more colorway
// columns than fit on one page width) into a PAIR of pages: the first has
// the full row content (Placement...Comment plus as many colorway columns
// as fit), and the immediately-following page repeats the SAME rows with
// only Placement/Level/MainMaterial/Code plus the single overflow colorway
// column that didn't fit — confirmed directly against a real tech pack
// (e.g. page N shows "...Comment | Orange/Black | ... | Red/Black" and
// page N+1 shows only "Placement | Level | MainMaterial | Code |
// Yellow/Grey" for the identical set of rows, both pages sharing the same
// "Displaying X - Y of N results" footer). That continuation page carries
// no data this tool needs — everything on it duplicates the page before —
// but if it isn't recognised, its rows (which DO have non-empty Placement
// and Code, since those columns are still present) get treated as
// brand-new rows: either bogus near-empty duplicate entries, or worse,
// their stray text fragments merge into the FOLLOWING real row as a false
// "wrapped continuation", corrupting genuine data (confirmed to produce
// smashed-together codes and mislabeled placements). Detected by the
// simple, reliable signal that a genuine full data page's header always
// includes a "Description" and a "UOM" column, and a colorway-overflow
// continuation page never does.
function portwestHeaderLooksLikeContinuation(headerItems) {
  const hasDescription = headerItems.some(it => it.text.startsWith('Descr'));
  const hasUom = headerItems.some(it => it.text.startsWith('UOM'));
  return !hasDescription && !hasUom;
}

function portwestBlankRow() {
  return { placement: '', mainMaterial: '', code: '', image: '', product: '', description: '',
    composition: '', materialWeight: '', commonSize: '', qtyDefault: '', uom: '', position: '', comment: '' };
}

async function extractPortwestPdf(file, onProgress) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  const categories = {}; // 'Fabric' -> [{...row}], 'Trims' -> [...], etc.
  let currentCategory = null;
  let activeRow = null;
  let coverInfo = null;
  let columnBoundaries = null;
  let headerCollecting = false;
  let headerItemsBuf = [];
  let headerLinesCollected = 0;
  // Holds any rows encountered before the very first category tag has
  // been seen. The tag ("Fabric (6)") is expected to sort, by real
  // coordinate position, just above its category's first data row — but
  // that's inferred from a rendered screenshot rather than a verified
  // coordinate dump, so this buffer is a cheap safety net: if the tag
  // instead turns out to land after its rows in coordinate order too,
  // those rows are retroactively assigned once the tag finally appears,
  // rather than being silently dropped.
  let pendingRows = [];

  function flushRow() {
    if (!activeRow) return;
    if (!(activeRow.placement || activeRow.code || activeRow.description)) { activeRow = null; return; }
    if (currentCategory) {
      categories[currentCategory] = categories[currentCategory] || [];
      categories[currentCategory].push(activeRow);
    } else {
      pendingRows.push(activeRow);
    }
    activeRow = null;
  }

  for (let p = 1; p <= pdf.numPages; p++) {
    onProgress && onProgress(p, pdf.numPages);
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();

    const items = content.items
      .map(it => {
        const tx = pdfjsLib.Util.transform(viewport.transform, it.transform);
        return { text: it.str, x: tx[4], y: tx[5] };
      })
      .filter(it => it.text.trim().length > 0);

    if (p === 1) {
      coverInfo = extractPortwestCoverInfo(items);
    }

    // NOTE: earlier versions gated page processing on an exact text match
    // for a "Placements" title at an assumed y-coordinate (y < 60). That
    // was an unverified guess (never checked against a real rendered PDF)
    // and turned out to be wrong, causing every page to be silently
    // skipped and zero items ever extracted. There is no separate title
    // gate any more — a page is simply whatever pages contain the table's
    // own "Placement" column-header row (detected below); pages without
    // it naturally contribute nothing, with no fragile pre-check needed.
    const contentItems = items.filter(it => it.y > 0);
    const lines = portwestClusterLines(contentItems, 2.5);

    // The page banner ("<Style Name> <Code> <Code> BOM <status>, <time>")
    // repeats above the table on every single page, same as the column
    // header row. Since currentCategory/columnBoundaries persist across
    // pages, that banner text could otherwise be mistaken for a stray data
    // row on continuation pages. Rather than pattern-match the (variable,
    // style-specific) banner text, we simply refuse to treat ANYTHING as
    // table content on a given page until that page's own header row has
    // been seen — the header is confirmed to repeat on every page without
    // exception, so this cleanly discards all pre-header boilerplate.
    let sawHeaderOnThisPage = false;
    // Set once this page's own header is identified as a colorway-overflow
    // continuation page (see portwestHeaderLooksLikeContinuation) — once
    // true, every remaining line on this page is discarded.
    let pageIsContinuation = false;

    for (const line of lines) {
      if (isPortwestBoilerplateLine(line)) continue;

      // Header row repeats at the top of every page and always spans
      // exactly 2 physical lines: the main row (Placement...Comment,
      // Orange, Yellow, all on one baseline) followed by "Weight" alone
      // (the wrap of "Material Weight", since "Material" is wider than
      // the "Weight" sub-label beneath it). Verified directly against the
      // sample PDF's text coordinates — there is no 3rd header line. It is
      // NOT reliably followed by a category tag — on continuation pages
      // it's followed straight by data rows of the already-current
      // category — so we consume a fixed line count rather than scanning
      // for a delimiter (mirrors the Malacca engine's header handling).
      // IMPORTANT: consuming one line too many here (e.g. 3 instead of 2)
      // silently eats either the category tag or the page's first data
      // row into the discarded header buffer — confirmed to cause exactly
      // that data loss when this was set to 3.
      //
      // The trigger match is deliberately forgiving: real PDF text
      // extraction can split a word like "Placement" across multiple
      // text runs at kerning pairs (e.g. "Place" + "ment"), so instead of
      // requiring the line's first item to be the exact string
      // "Placement", the leading item(s) are concatenated and matched
      // case-insensitively, with no coordinate/position requirement at
      // all — "Placement" as a whole word is confirmed (via the tech
      // pack's own text content) to only ever appear as this one column
      // header, never as data, so matching on text alone is safe and
      // avoids yet another unverified coordinate assumption.
      if (!headerCollecting && !sawHeaderOnThisPage) {
        const leadText = line.items.slice(0, 3).map(i => i.text).join('').trim().toLowerCase();
        if (leadText.startsWith('placement') && line.items[0]) {
          headerCollecting = true;
          headerItemsBuf = [...line.items];
          headerLinesCollected = 1;
          continue;
        }
      }
      if (headerCollecting) {
        if (headerLinesCollected < 2) {
          headerItemsBuf.push(...line.items);
          headerLinesCollected++;
          continue;
        }
        headerCollecting = false;
        sawHeaderOnThisPage = true;

        // A colorway-overflow continuation page repeats the same rows as
        // the page before it, with none of the columns this tool actually
        // needs — skip the rest of this page entirely rather than risk
        // misreading its rows as new data or letting stray fragments merge
        // into the next real row. Checked BEFORE attempting derivation, so
        // a continuation page never touches columnBoundaries at all.
        if (portwestHeaderLooksLikeContinuation(headerItemsBuf)) {
          pageIsContinuation = true;
        } else if (!columnBoundaries) {
          // The header layout is structurally identical on every full data
          // page (it's literally the same repeated table header), so once
          // derived successfully there's no need to re-parse and re-anchor
          // it on each of the (possibly 20+) subsequent pages — a real
          // saving on a long tech pack. Only re-derive if we don't already
          // have it.
          columnBoundaries = derivePortwestColumnBoundaries(headerItemsBuf) ||
            portwestColumnsFromFractions(viewport.width);
        }
        // Fall through — this line itself still needs normal processing
        // below (it may be a category tag or a plain data row), unless
        // this turned out to be a continuation page, handled just below.
      }

      if (pageIsContinuation) continue; // discard every remaining line on a colorway-overflow continuation page

      if (!sawHeaderOnThisPage) continue; // pre-header boilerplate (title, banner, "Displaying..." counter)

      const tag = matchPortwestCategoryTag(line);
      if (tag !== null) {
        flushRow();
        if (!currentCategory && pendingRows.length) {
          categories[tag] = (categories[tag] || []).concat(pendingRows);
          pendingRows = [];
        }
        currentCategory = tag;
        continue;
      }

      if (!columnBoundaries) continue;

      const cols = portwestBlankRow();
      for (const it of line.items) {
        const key = classifyPortwestColumn(it.x, columnBoundaries);
        cols[key] = (cols[key] ? cols[key] + ' ' : '') + it.text;
      }

      // New logical row starts when this line has content in the Code
      // column. Code alone (not Code AND Placement) is the right signal:
      // Code is short and never wraps, so it's only ever present on a
      // row's true first line — a long Placement label (e.g. "Reflective
      // Tape", "Hook & Loop", "Oeko-Tex Label") can itself wrap onto a
      // second physical line, and that wrapped remainder ("Tape", "Loop",
      // "Label") lands in the Placement column with an empty Code, so it's
      // still correctly treated as a continuation under a Code-only check.
      // Requiring Placement too (an earlier version of this check) missed
      // genuine rows whose Placement column ships legitimately blank in
      // the source tech pack (confirmed: a Carton line item with Level/
      // Code/weight/UOM/position all present but no Placement text) —
      // those got silently merged into the previous row, smashing both
      // rows' Code/weight/UOM together into one corrupted entry. Any other
      // wrapped field (Description/Position most commonly run onto extra
      // lines) still has an empty Code, so it's correctly treated as a
      // continuation either way.
      if (cols.code.trim()) {
        flushRow();
        activeRow = portwestBlankRow();
        for (const k of Object.keys(cols)) activeRow[k] = cols[k].trim();
      } else if (activeRow) {
        for (const k of Object.keys(cols)) {
          if (cols[k].trim()) activeRow[k] = (activeRow[k] ? activeRow[k] + ' ' : '') + cols[k].trim();
        }
      }
    }
  }
  flushRow();
  // Anything never claimed by a category tag (shouldn't happen in a
  // well-formed export) is surfaced rather than silently dropped, so a
  // parsing gap is visible instead of invisible.
  if (pendingRows.length) {
    categories['Uncategorized'] = (categories['Uncategorized'] || []).concat(pendingRows);
  }

  return { coverInfo, categories };
}

/* ============================================================
   PORTWEST — FIELD FORMATTING HELPERS
   ============================================================ */

// "190g" -> "190 gsm" (plain text, per buyer's convention). Falls back to
// the raw value untouched if it doesn't match the simple "<number>g" shape.
function portwestFormatWeight(raw) {
  const m = (raw || '').trim().match(/^(\d+(?:\.\d+)?)\s*g$/i);
  return m ? `${m[1]} gsm` : (raw || '').trim();
}

const PORTWEST_METERS_TO_YARDS = 1.09361;

// Classifies a Trims/Labels/Packaging row's consumption cell based on the
// tech pack's own UOM field: piece-counted items ("1 Pcs", "2 Pcs" — no
// decimals) vs. length items measured in meters, converted to yards
// ("2.45 yds" — 2 decimals). Any other/unrecognised UOM falls back to a
// plain "<value> <UOM>" pieces-style label so nothing silently disappears.
// FIX (formatting bug): a column-boundary artifact in the source PDF can
// occasionally push the row's leading quantity digit into the UOM field
// itself instead of qtyDefault (observed as e.g. uom = "1 Pieces" while
// qtyDefault comes through blank/zero). The old exact-match check
// (`u === 'PIECES'`) failed on that shape, fell through to the
// "unrecognised UOM" branch, and built a numFmt of `0" 1 PIECES"` — Excel
// then rendered the cell's numeric value (0) followed by that literal
// suffix text, producing the "0 1 PIECES" display bug. This version
// strips any leading numeric token off the UOM text first, recovers it as
// the quantity if qtyDefault itself came through empty, and only then
// classifies the cleaned unit text.
function portwestClassifyConsumption(uom, qtyDefault) {
  const rawUom = (uom || '').trim();
  const leadingQtyMatch = rawUom.match(/^(\d+(?:\.\d+)?)\s*(.*)$/);
  const cleanedUom = (leadingQtyMatch ? leadingQtyMatch[2] : rawUom).trim().toUpperCase();
  const recoveredQty = leadingQtyMatch ? parseFloat(leadingQtyMatch[1]) : null;
  const qty = parseFloat(qtyDefault) || recoveredQty || 0;

  if (cleanedUom === 'PIECES' || cleanedUom === 'PCS' || cleanedUom === 'PC') {
    return { value: qty, numFmt: '0" Pcs"' };
  }
  if (cleanedUom === 'M' || cleanedUom === 'MTR' || cleanedUom === 'MTRS' || cleanedUom === 'METER' || cleanedUom === 'METERS') {
    return { value: qty * PORTWEST_METERS_TO_YARDS, numFmt: '0.00" yds"' };
  }
  // Unrecognised UOM — surface the tech pack's own (now-cleaned) unit
  // rather than guessing, but never leak stray digits into the numFmt text.
  return { value: qty, numFmt: `0" ${cleanedUom.replace(/"/g, '')}"` };
}

/* ============================================================
   PORTWEST — COST SHEET MERGE (row-insertion, formula-preserving)
   ============================================================ */

// Section layout as it exists in the original Portwest_Format.xlsx
// template. "fixedTailRows" are pre-filled anchor rows (e.g. the standard
// "Sewig Thread" line, or "Test cost"/"Discount") that must always remain
// the LAST rows of that section, directly above the subtotal — new
// extracted items are inserted above them, never overwriting them.
const PORTWEST_SECTIONS = [
  { key: 'FABRICS',      dataStart: 13, dataEnd: 17, fixedTailRows: 0, subtotalRow: 18 },
  { key: 'BRANDING',     dataStart: 20, dataEnd: 30, fixedTailRows: 0, subtotalRow: 31 },
  { key: 'TRIMS',        dataStart: 33, dataEnd: 35, fixedTailRows: 1, subtotalRow: 36 }, // row35 = "Sewig Thread"
  { key: 'EMBROIDERY',   dataStart: 38, dataEnd: 38, fixedTailRows: 0, subtotalRow: 39 },
  { key: 'PACKAGING',    dataStart: 41, dataEnd: 52, fixedTailRows: 2, subtotalRow: 53 }, // rows51-52 = "Test cost"/"Discount"
];
// O(1) lookup instead of repeated Array.find() scans over the (small but
// non-trivial, since it's called from several loops) sections list.
const PORTWEST_SECTIONS_BY_KEY = Object.fromEntries(PORTWEST_SECTIONS.map(s => [s.key, s]));

// Tech-pack category name -> cost sheet section key. "Stitching" is
// intentionally absent (skipped entirely — thread items already live in
// the template's fixed Trims row).
const PORTWEST_CATEGORY_TO_SECTION = {
  'Fabric': 'FABRICS',
  'Trims': 'TRIMS',
  'Labels': 'BRANDING',
  'Packaging': 'PACKAGING',
  'Embroidery': 'EMBROIDERY',
  'Print': 'EMBROIDERY',
};

function portwestShiftFormulaRows(formula, delta) {
  return formula.replace(/([A-Z]{1,3})(\d+)/g, (whole, col, rowStr) => col + (parseInt(rowStr, 10) + delta));
}

function portwestCopyCellStyle(srcCell, dstCell) {
  if (srcCell.style) dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
}

// IMPORTANT: cells across unrelated rows in this template frequently share
// the SAME underlying style object by reference (confirmed empirically —
// a normal consequence of how Excel/ExcelJS pool identical style records
// for storage efficiency). Setting `cell.numFmt = x` mutates that style
// object in place, which silently corrupts the number format of every
// OTHER cell that happens to share the same reference — including cells
// in completely unrelated sections. Always replace the cell's style with
// a fresh clone instead of mutating it, to guarantee the change is
// isolated to this one cell.
function portwestSetNumFmt(cell, numFmt) {
  cell.style = Object.assign({}, cell.style, { numFmt });
}

// Sets wrapText:true on a cell while preserving whatever other alignment
// properties (horizontal/vertical/indent) it already had from the template
// or from earlier styling — mirrors the "clone, don't mutate" pattern used
// by portwestSetNumFmt above, since style objects can be shared by
// reference across unrelated cells in this template.
function portwestSetWrapText(cell) {
  const currentAlignment = (cell.style && cell.style.alignment) || {};
  cell.style = Object.assign({}, cell.style, {
    alignment: Object.assign({}, currentAlignment, { wrapText: true }),
  });
}

// Excel worksheet names can't contain \ / ? * [ ] : and are capped at 31
// characters — sanitize whatever we build from the tech pack's own style
// code/description before assigning it as the tab name.
function portwestSanitizeSheetName(name) {
  return (name || '').replace(/[\\/?*[\]:]/g, '-').trim().slice(0, 31) || 'Sheet1';
}

// Unmerging via ws.unMergeCells(row, fromCol, row, toCol) turns out to be
// unreliable after this file's multiple cascaded spliceRows calls: ExcelJS
// tracks merges in an internal ws._merges map keyed by the merge's ORIGINAL
// top-left address (e.g. "A18"), but after a row shifts down (say to row
// 23), that map's key is never renamed even though the Range object it
// points to gets its own row numbers updated in place. ws.unMergeCells
// looks up merges via the CELL's current live address (now "A23"), which
// no longer matches the stale "A18" key — so the lookup silently misses,
// nothing gets unmerged, and a later ws.mergeCells() at that same live
// address throws "Cannot merge already merged cells" because the merge
// object (still tracked, just under its stale key) genuinely still covers
// that row. Confirmed directly against a real multi-section tech pack
// extraction (66 BOM rows) — this is not a hypothetical edge case.
//
// This works around it by ignoring keys entirely: scan every merge
// currently tracked, and if its own (live, correctly-updated) row/column
// span overlaps our target range, unmerge every cell in THAT recorded
// span directly and drop it from the registry — regardless of what key
// it's currently filed under.
function portwestForceClearIntersectingMerges(ws, row, fromCol, toCol) {
  const merges = ws._merges;
  if (!merges) return;
  for (const key of Object.keys(merges)) {
    const merge = merges[key];
    if (!merge) continue;
    const overlapsRow = row >= merge.top && row <= merge.bottom;
    const overlapsCol = fromCol <= merge.right && toCol >= merge.left;
    if (!overlapsRow || !overlapsCol) continue;
    for (let r = merge.top; r <= merge.bottom; r++) {
      for (let c = merge.left; c <= merge.right; c++) {
        const cell = ws.getCell(r, c);
        if (cell && typeof cell.unmerge === 'function') cell.unmerge();
      }
    }
    delete merges[key];
  }
}

// Merges a FIXED, explicitly-given column range in a row into one cell and
// applies the given alignment. The range is always passed explicitly by
// the caller (never auto-detected by scanning the row) — an earlier
// version tried to find the "value column" by scanning every cell in the
// row from column 1 onward, but several of these summary rows share their
// physical row with an unrelated left-hand data block that has its own
// early numeric cells (e.g. row 58's column C holds a plain "0" that has
// nothing to do with the right-hand "Overall working efficiency" label),
// so that scan stopped at the wrong column and produced badly wrong merge
// boundaries. Fixed, verified column numbers avoid that entirely.
//
// Also IMPORTANT and confirmed by direct testing against this template:
// ExcelJS's spliceRows does not reliably preserve merged-cell ranges once
// multiple splice operations cascade across a sheet (both insertions and
// deletions) — even merges that ship pre-baked in the template (e.g. the
// section subtotal rows' A:L merge) can end up unmerged after the section
// insert/remove passes run. The cell VALUES survive correctly, just not
// the merge state. So this function is written to be safe to call
// unconditionally regardless of whether the range's merge already existed,
// already broke, or never existed: it reads whichever cell in the range
// currently holds text (if any), clears the range, safely unmerges
// whatever may or may not be currently merged there, then re-merges fresh
// at the row's final (already-shifted) position.
function portwestMergeAndAlign(ws, row, fromCol, toCol, alignment) {
  if (toCol <= fromCol) return;
  let label = '';
  for (let c = fromCol; c <= toCol; c++) {
    const cell = ws.getCell(row, c);
    if (!label && typeof cell.value === 'string' && cell.value.trim()) label = cell.value.trim();
    cell.value = null;
  }
  portwestForceClearIntersectingMerges(ws, row, fromCol, toCol);
  try {
    ws.mergeCells(row, fromCol, row, toCol);
  } catch (e) {
    // Last-resort safety net: if some other still-unaccounted-for merge
    // slipped through the force-clear above, don't let one row's styling
    // failure abort the entire Cost Sheet generation — leave that row's
    // cells unmerged (still readable, just not centered as one cell) and
    // continue with the rest of the file.
    console.error(`portwestMergeAndAlign: could not merge row ${row} cols ${fromCol}-${toCol}: ${e.message}`);
  }
  const merged = ws.getCell(row, fromCol);
  if (label) merged.value = label;
  merged.style = Object.assign({}, merged.style, { alignment });
}

// The template stores many formula columns (M, O, P, etc.) as Excel
// "shared formula" groups — one master cell holding the real formula text
// plus a `ref` range, with every other cell in that range just pointing
// back at the master. ExcelJS's spliceRows does not keep that master/
// slave bookkeeping consistent when rows are inserted or removed (it throws
// "Shared Formula master must exist above and or left of clone"), even for
// rows nowhere near the master. Flattening every formula cell to its own
// independent (non-shared) formula — before any row is spliced — sidesteps
// this entirely; `cell.formula` already resolves a slave cell's real
// formula text, so this is a lossless rewrite.
function portwestFlattenSharedFormulas(ws, maxRow, maxCol) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      if (cell.type === ExcelJS.ValueType.Formula && cell.formula) {
        cell.value = { formula: cell.formula };
      }
    }
  }
}

function portwestCopyRowStyle(ws, fromRow, toRow, maxCol) {
  for (let c = 1; c <= maxCol; c++) {
    portwestCopyCellStyle(ws.getCell(fromRow, c), ws.getCell(toRow, c));
  }
  const srcRow = ws.getRow(fromRow);
  if (srcRow.height) ws.getRow(toRow).height = srcRow.height;
}

// Clones a template row's literal INPUT values (J price default, L wastage
// default, N/O/Q fabric-pricing defaults) onto a newly-inserted row, so a
// brand new row starts with the same sensible defaults as its neighbours.
// Formula cells (M/O/P/R/S) are deliberately NOT cloned here — they are
// always self-referencing (e.g. "=(K17+K17*L17)*J17"), so they're instead
// regenerated fresh, directly at each row's own final position, by
// portwestWriteRowFormulas below. That sidesteps relying on spliceRows to
// keep formula text in sync with a row's shifted position, which it does
// not do (confirmed: a surviving row shifted downward by an insertion
// elsewhere keeps its OLD self-referencing formula text, silently
// producing #VALUE! once it points at what is now a label/header row).
function portwestCloneRowDefaults(ws, templateRow, newRow, maxCol) {
  for (let c = 1; c <= maxCol; c++) {
    const srcCell = ws.getCell(templateRow, c);
    if (typeof srcCell.value === 'number') {
      ws.getCell(newRow, c).value = srcCell.value;
    }
  }
}

// Writes the correct self-referencing formulas for a single data/fixed row
// at its OWN final row number — always correct regardless of how many
// rows were inserted/removed anywhere else, since nothing here depends on
// an "original" row number at all.
function portwestWriteRowFormulas(ws, r, sectionKey) {
  ws.getCell(r, 13).value = { formula: `(K${r}+K${r}*L${r})*J${r}` }; // M - Total Cost
  if (sectionKey === 'FABRICS') {
    // Green "Fob Price (M) / FOB Price by Yds / ... / CNF price by Yds"
    // column group (O-S) — FOB Price by Yds = Fob Price (M) / 1.09361,
    // exactly matching the real template's own convention (verified
    // directly against Portwest_Format.xlsx rows 15-17). Mathematically
    // this is the same yard<->meter conversion as multiplying by 0.9144,
    // but not bit-identical to it (1/1.09361 ≈ 0.914402, not 0.9144), so
    // the exact divisor is used here to avoid rounding drift from the
    // template's own values.
    ws.getCell(r, 16).value = { formula: `O${r}/1.09361` }; // P
    ws.getCell(r, 18).value = { formula: `O${r}+Q${r}` };   // R
    ws.getCell(r, 19).value = { formula: `P${r}+Q${r}` };   // S
  } else if (sectionKey === 'BRANDING' || sectionKey === 'TRIMS') {
    ws.getCell(r, 15).value = { formula: `N${r}*1.1` }; // O
  }
}

async function buildPortwestCostSheet(templateArrayBuffer, extracted) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer.slice(0));
  const ws = workbook.worksheets[0];
  const MAXCOL = 45; // through column ~AS, comfortably covers A-S plus FOB/CNF columns

  // Must run before any row is spliced — see portwestFlattenSharedFormulas.
  portwestFlattenSharedFormulas(ws, ws.rowCount, MAXCOL);

  // --- Header fields ---
  const cover = extracted.coverInfo || {};
  if (cover.styleCode) ws.getCell('D8').value = cover.styleCode;       // Style Name
  if (cover.description) ws.getCell('D7').value = cover.description;  // Item Description
  ws.getCell('D4').value = new Date();                                 // Date (keeps template's existing date format)

  // --- Excel tab name: "<Style Name/Description> - <Style Code>" ---
  // Falls back gracefully if either piece is missing (e.g. cover-page
  // parsing came back partial), and is sanitized for Excel's forbidden
  // tab-name characters (\ / ? * [ ] :) and 31-character limit.
  const tabNameRaw = [cover.description, cover.styleCode].filter(Boolean).join(' - ');
  if (tabNameRaw) ws.name = portwestSanitizeSheetName(tabNameRaw);

  // --- Bucket extracted rows into their target sections ---
  const bucketed = { FABRICS: [], BRANDING: [], TRIMS: [], EMBROIDERY: [], PACKAGING: [] };
  const counts = {};
  for (const [categoryName, rows] of Object.entries(extracted.categories || {})) {
    if (categoryName === 'Stitching') continue; // thread items — not needed, already in template
    const sectionKey = PORTWEST_CATEGORY_TO_SECTION[categoryName];
    if (!sectionKey) continue; // unrecognised category — skip rather than guess
    for (const row of rows) {
      const dest = portwestIsFabricOverride(row) ? 'FABRICS' : sectionKey;
      bucketed[dest].push(row);
      counts[categoryName] = (counts[categoryName] || 0) + 1;
    }
  }

  // --- Compute each section's row delta BEFORE touching the sheet ---
  const deltas = {};
  for (const sec of PORTWEST_SECTIONS) {
    const capacity = (sec.dataEnd - sec.dataStart + 1) - sec.fixedTailRows;
    const needed = Math.max(bucketed[sec.key].length, 1); // never shrink a section to zero rows
    deltas[sec.key] = needed - capacity;
  }
  const totalDelta = Object.values(deltas).reduce((a, b) => a + b, 0);

  // --- Compute final (post-shift) row numbers for every section's data
  //     start and subtotal. A section's data-start only shifts because of
  //     insertions/removals in sections ABOVE it (its own insertion happens
  //     below its own start), whereas its subtotal shifts by sections above
  //     it AND its own delta (the subtotal sits below its own data). ---
  const order = ['FABRICS', 'BRANDING', 'TRIMS', 'EMBROIDERY', 'PACKAGING'];
  let running = 0;
  const finalDataStart = {};
  const finalSubtotalRow = {};
  for (const key of order) {
    const sec = PORTWEST_SECTIONS_BY_KEY[key];
    finalDataStart[key] = sec.dataStart + running;
    running += deltas[key];
    finalSubtotalRow[key] = sec.subtotalRow + running;
  }
  const cmPcRow = 54 + totalDelta;
  const totalCostPcRow = 55 + totalDelta;

  // --- Physically splice rows, bottom-to-top, so each section's ORIGINAL
  //     row numbers stay valid until it's this section's turn ---
  const reverseOrder = [...order].reverse();
  for (const key of reverseOrder) {
    const sec = PORTWEST_SECTIONS_BY_KEY[key];
    const delta = deltas[key];
    const templateRow = sec.dataEnd - sec.fixedTailRows; // last row of the GROWABLE capacity zone (never the fixed anchor row itself)
    if (delta > 0) {
      const insertAt = templateRow + 1; // right where the fixed tail rows (or subtotal, if none) currently begin
      ws.spliceRows(insertAt, 0, ...Array.from({ length: delta }, () => []));
      for (let i = 0; i < delta; i++) {
        const newRow = insertAt + i;
        portwestCopyRowStyle(ws, templateRow, newRow, MAXCOL);
        portwestCloneRowDefaults(ws, templateRow, newRow, MAXCOL);
      }
    } else if (delta < 0) {
      const removeCount = -delta;
      const removeAt = sec.dataStart + bucketed[key].length;
      ws.spliceRows(removeAt, removeCount);
    }
  }

  // --- Write extracted data into each section's data rows ---
  for (const key of order) {
    const items = bucketed[key];
    items.forEach((row, i) => {
      const r = finalDataStart[key] + i;
      ws.getCell(r, 1).value = row.placement || '';   // A - Item
      ws.getCell(r, 2).value = row.code || '';          // B - Item Code
      ws.getCell(r, 3).value = row.position || '';      // C - Position
      ws.getCell(r, 9).value = row.description || '';   // I - Description

      if (key === 'FABRICS') {
        ws.getCell(r, 7).value = portwestFormatWeight(row.materialWeight); // G - Weight
        const kCell = ws.getCell(r, 11); // K - Consumption
        kCell.value = 0;
        portwestSetNumFmt(kCell, '0.00" yds"');
      } else {
        ws.getCell(r, 8).value = row.commonSize || ''; // H - Size/Width
        const { value, numFmt } = portwestClassifyConsumption(row.uom, row.qtyDefault);
        const kCell = ws.getCell(r, 11); // K - Consumption
        kCell.value = value;
        portwestSetNumFmt(kCell, numFmt);
      }
    });
  }

  // --- Wrap text across every cell in Fabrics / Trims / Labels(Branding) /
  //     Packaging's data rows. Packaging's fixed "Test cost"/"Discount"
  //     tail rows are excluded here — they get their own left-aligned
  //     merge-and-wrap treatment below instead, since they're a note
  //     rather than an item row. ---
  const WRAP_SECTIONS = ['FABRICS', 'TRIMS', 'BRANDING', 'PACKAGING'];
  for (const key of WRAP_SECTIONS) {
    const sec = PORTWEST_SECTIONS_BY_KEY[key];
    const excludeTailRows = key === 'PACKAGING' ? sec.fixedTailRows : 0;
    const endRow = finalSubtotalRow[key] - 1 - excludeTailRows;
    for (let r = finalDataStart[key]; r <= endRow; r++) {
      for (let c = 1; c <= MAXCOL; c++) {
        portwestSetWrapText(ws.getCell(r, c));
      }
    }
  }

  // --- Regenerate every row-local formula (M, and O/P/R/S where
  //     applicable) at each row's OWN final position, for every row in
  //     the section's final layout — data rows AND any fixed tail rows
  //     (e.g. "Sewig Thread", "Test cost", "Discount"). This is done
  //     unconditionally, regardless of whether a given row is newly
  //     inserted or a shifted survivor, since spliceRows does not keep a
  //     surviving row's formula text in sync with its new position (see
  //     portwestCloneRowDefaults for why relocate-in-place isn't used).
  for (const key of order) {
    for (let r = finalDataStart[key]; r < finalSubtotalRow[key]; r++) {
      portwestWriteRowFormulas(ws, r, key);
    }
  }

  // --- Rewrite each section's SUM(...) subtotal to match its new row range ---
  for (const key of order) {
    const dataEndFinal = finalSubtotalRow[key] - 1;
    ws.getCell(finalSubtotalRow[key], 13).value = { formula: `SUM(M${finalDataStart[key]}:M${dataEndFinal})` }; // M
  }

  // --- Repair the summary block (rows 54-62 in the original template),
  //     which sits below every section and shifts uniformly by totalDelta,
  //     except M55 which cross-references the 5 (individually-shifted)
  //     section subtotals + the CM/Pc row ---
  if (totalDelta !== 0) {
    // Generic same-block formulas: every reference inside them points to
    // another cell within this same uniformly-shifting block, so a flat
    // +totalDelta on every embedded row number is correct.
    for (const origRow of [58, 59, 60, 61, 62]) {
      const newRow = origRow + totalDelta;
      for (const col of ['E', 'F', 'M']) {
        const cell = ws.getCell(`${col}${newRow}`);
        if (cell.formula) cell.value = { formula: portwestShiftFormulaRows(cell.formula, totalDelta) };
      }
    }
  }
  // M55 (Total Cost/Pc) — special-cased: its 6 references each shifted by
  // a different amount depending on which section they belong to.
  ws.getCell(totalCostPcRow, 13).value = {
    formula: `M${finalSubtotalRow.FABRICS}+M${finalSubtotalRow.BRANDING}+M${finalSubtotalRow.TRIMS}+M${finalSubtotalRow.EMBROIDERY}+M${finalSubtotalRow.PACKAGING}+M${cmPcRow}`,
  };
  // K7 sits in the fixed header block (never itself shifts) but points at
  // M55's new location.
  ws.getCell('K7').value = { formula: `M${totalCostPcRow}` };

  // --- Re-merge every summary/subtotal label row at its final (already
  //     row-shifted) position, using exact column ranges verified against
  //     the real Portwest_Format.xlsx template — not auto-detected, for
  //     the reasons explained in portwestMergeAndAlign's own comment.
  //     This is done unconditionally for every row here, whether or not
  //     that row's merge happened to survive the section insert/remove
  //     passes above, since that survival isn't reliable across multiple
  //     cascaded splices. ---
  const CENTER_MIDDLE_WRAP = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const LEFT_MIDDLE_WRAP = { horizontal: 'left', vertical: 'middle', wrapText: true };

  // Section subtotal rows ("Sub Total:") — label spans columns A:L (1-12),
  // value sits in M (13). Applied to every section, including Embroidery
  // (whose original template row happens not to ship pre-merged), so all
  // five subtotal rows are visually consistent.
  for (const key of order) {
    portwestMergeAndAlign(ws, finalSubtotalRow[key], 1, 12, CENTER_MIDDLE_WRAP);
  }

  // CM/Pc and Total Cost/Pc — label spans I:L (9-12), value in M (13).
  portwestMergeAndAlign(ws, cmPcRow, 9, 12, CENTER_MIDDLE_WRAP);
  portwestMergeAndAlign(ws, totalCostPcRow, 9, 12, CENTER_MIDDLE_WRAP);

  // "Price per min (supplier profit included)", "SAM (Sewing) in min", and
  // "Overall working efficiency (OWE) based on SAM (Sewing)" sit on the
  // three rows directly below Total Cost/Pc, shifting by the same
  // totalDelta as the rest of this fixed summary block. Same I:L / M
  // geometry as CM/Pc and Total Cost/Pc above, for visual consistency.
  const pricePerMinRow = 56 + totalDelta;
  const samRow = 57 + totalDelta;
  const oweRow = 58 + totalDelta;
  portwestMergeAndAlign(ws, pricePerMinRow, 9, 12, CENTER_MIDDLE_WRAP);
  portwestMergeAndAlign(ws, samRow, 9, 12, CENTER_MIDDLE_WRAP);
  portwestMergeAndAlign(ws, oweRow, 9, 12, CENTER_MIDDLE_WRAP);

  // Test cost / Discount — merged from Item Code (B, col 2) through
  // Description (I, col 9), left + middle aligned. The template's own
  // "Test cost"/"Discount" label lives in column A (the item-name column,
  // matching every other data row's own convention) and is left as-is;
  // this merge only unifies the otherwise-separate B:I cells into one
  // blank, properly-aligned block.
  const testCostRow = finalSubtotalRow.PACKAGING - 2;
  const discountRow = finalSubtotalRow.PACKAGING - 1;
  portwestMergeAndAlign(ws, testCostRow, 2, 9, LEFT_MIDDLE_WRAP);
  portwestMergeAndAlign(ws, discountRow, 2, 9, LEFT_MIDDLE_WRAP);

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, counts, bucketedCounts: Object.fromEntries(order.map(k => [k, bucketed[k].length])) };
}

/* ============================================================
   PORTWEST — SUPPLY CHAIN SHEET (multi-PDF, 2-tab, deduplicated)
   ============================================================ */

// Strips punctuation the buyer doesn't want counted as a real difference
// (commas, semicolons, periods, dashes/hyphens/en-dash/em-dash,
// underscores, quotes, parentheses, slashes), lowercases, and collapses
// whitespace — so "V226OX190 - 300D, Oxford" and "v226ox190 300d oxford"
// compare as identical.
function portwestNormalizeForDedup(text) {
  return (text || '')
    .toLowerCase()
    .replace(/[,;.\-–—_'"()/\\]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function portwestDedupKey(code, description) {
  return portwestNormalizeForDedup(code) + '||' + portwestNormalizeForDedup(description);
}

async function extractPortwestSupplyChain(files, onProgress) {
  const fabricRows = []; // { styleNo, category, code, description }
  const trimsRows = [];  // { styleNo, itemName, code, description }
  const seenFabric = new Set();
  const seenTrims = new Set();

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    onProgress && onProgress(i + 1, files.length, file.name);
    const extracted = await extractPortwestPdf(file);
    const styleNo = (extracted.coverInfo && extracted.coverInfo.styleCode) || file.name.replace(/\.pdf$/i, '');

    for (const [categoryName, rows] of Object.entries(extracted.categories || {})) {
      // Thread items ('Stitching') are skipped for the BOM/cost sheet
      // because the template already has a standard thread line, but for
      // a sourcing/supply-chain list they're still real materials — no
      // category is excluded here, everything non-Fabric just lands in Trims.
      for (const row of rows) {
        const goesToFabric = categoryName === 'Fabric' || portwestIsFabricOverride(row);
        if (goesToFabric) {
          const key = portwestDedupKey(row.code, row.description);
          if (seenFabric.has(key)) continue;
          seenFabric.add(key);
          fabricRows.push({ styleNo, category: row.placement || '', code: row.code || '', description: row.description || '' });
        } else {
          const key = portwestDedupKey(row.code, row.description);
          if (seenTrims.has(key)) continue;
          seenTrims.add(key);
          trimsRows.push({ styleNo, itemName: row.placement || '', code: row.code || '', description: row.description || '' });
        }
      }
    }
  }

  return { fabricRows, trimsRows };
}

async function buildPortwestSupplyChainWorkbook(fabricRows, trimsRows) {
  const workbook = new ExcelJS.Workbook();

  function addSheet(name, headers, rows, rowMapper) {
    const ws = workbook.addWorksheet(name);
    ws.columns = headers.map(h => ({ header: h, width: h === 'Description' ? 55 : h.includes('Style') ? 14 : 22 }));
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF374151' } };
    headerRow.alignment = { vertical: 'middle' };
    for (const row of rows) ws.addRow(rowMapper(row));
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  }

  addSheet('Fabric', ['Style No', 'Fabric Category', 'Fabric Code', 'Description', 'Supplier'], fabricRows,
    r => [r.styleNo, r.category, r.code, r.description, '']);
  addSheet('Trims', ['Style No', 'Item Name', 'Item Code', 'Description', 'Supplier'], trimsRows,
    r => [r.styleNo, r.itemName, r.code, r.description, '']);

  const buffer = await workbook.xlsx.writeBuffer();
  return buffer;
}

/* ============================================================
   PORTWEST — UI WIRING
   ============================================================ */
const portwestHubView = document.getElementById('portwestHubView');
const portwestView = document.getElementById('portwestView');
const portwestSupplyView = document.getElementById('portwestSupplyView');

document.getElementById('brandPortwest').addEventListener('click', () => {
  requestUnlock('portwest', () => {
    homeView.hidden = true;
    portwestHubView.hidden = false;
  });
});
document.getElementById('portwestHubBackBtn').addEventListener('click', () => {
  portwestHubView.hidden = true;
  homeView.hidden = false;
});
document.getElementById('portwestHubBomCard').addEventListener('click', () => {
  portwestHubView.hidden = true;
  portwestView.hidden = false;
});
document.getElementById('portwestHubSupplyCard').addEventListener('click', () => {
  portwestHubView.hidden = true;
  portwestSupplyView.hidden = false;
});
document.getElementById('backBtnPortwest').addEventListener('click', () => {
  portwestView.hidden = true;
  portwestHubView.hidden = false;
});
document.getElementById('backBtnPwSupply').addEventListener('click', () => {
  portwestSupplyView.hidden = true;
  portwestHubView.hidden = false;
});

const dropzonePw = document.getElementById('dropzonePw');
const fileInputPw = document.getElementById('fileInputPw');
const filebarPw = document.getElementById('filebarPw');
const fileNamePwEl = document.getElementById('fileNamePw');
const clearFilePw = document.getElementById('clearFilePw');
const processBtnPw = document.getElementById('processBtnPw');
const processLabelPw = document.getElementById('processLabelPw');
const statusPw = document.getElementById('statusPw');
const resultsPw = document.getElementById('resultsPw');
const cbdBtnPw = document.getElementById('cbdBtnPw');
const cbdLabelPw = document.getElementById('cbdLabelPw');
const cbdStatusPw = document.getElementById('cbdStatusPw');

let currentFilePw = null;
let currentPwExtracted = null;

function setStatusPw(msg, cls) {
  statusPw.textContent = msg;
  statusPw.className = 'status' + (cls ? ' ' + cls : '');
}

function setFilePw(file) {
  if (!file || file.type !== 'application/pdf') {
    setStatusPw('Please choose a PDF file.', 'err');
    return;
  }
  currentFilePw = file;
  fileNamePwEl.textContent = file.name;
  filebarPw.classList.add('show');
  setStatusPw('');
  resultsPw.classList.remove('show');
  currentPwExtracted = null;
  processBtnPw.disabled = false;
}

dropzonePw.addEventListener('dragover', e => { e.preventDefault(); dropzonePw.classList.add('drag'); });
dropzonePw.addEventListener('dragleave', () => dropzonePw.classList.remove('drag'));
dropzonePw.addEventListener('drop', e => {
  e.preventDefault(); dropzonePw.classList.remove('drag');
  if (e.dataTransfer.files.length) setFilePw(e.dataTransfer.files[0]);
});
fileInputPw.addEventListener('change', e => {
  if (e.target.files.length) setFilePw(e.target.files[0]);
});
clearFilePw.addEventListener('click', e => {
  e.stopPropagation();
  currentFilePw = null; fileInputPw.value = '';
  filebarPw.classList.remove('show');
  resultsPw.classList.remove('show');
  currentPwExtracted = null;
  setStatusPw('');
  processBtnPw.disabled = true;
});

function portwestItemCount(categories) {
  return Object.values(categories).reduce((s, rows) => s + rows.length, 0);
}

processBtnPw.addEventListener('click', async () => {
  if (!currentFilePw) return;
  processBtnPw.disabled = true;
  processBtnPw.classList.add('loading');
  resultsPw.classList.remove('show');
  try {
    const extracted = await extractPortwestPdf(currentFilePw, (p, total) => {
      processLabelPw.textContent = `Scanning page ${p} / ${total}...`;
      setStatusPw(`Scanning page ${p} of ${total}...`);
    });
    currentPwExtracted = extracted;
    const total = portwestItemCount(extracted.categories);
    if (total === 0) {
      setStatusPw('No Placements (BOM) line items were found in this tech pack.', 'err');
    } else {
      const summary = Object.entries(extracted.categories).map(([k, v]) => `${k}: ${v.length}`).join(', ');
      setStatusPw(`Done — ${total} items extracted. ${summary}.`, 'ok');
      resultsPw.classList.add('show');
    }
  } catch (err) {
    console.error(err);
    setStatusPw('Something went wrong reading this PDF: ' + err.message, 'err');
    currentPwExtracted = null;
  } finally {
    processBtnPw.disabled = false;
    processBtnPw.classList.remove('loading');
    processLabelPw.textContent = 'Extract Tech Pack';
  }
});

cbdBtnPw.addEventListener('click', async () => {
  if (!currentPwExtracted) return;
  cbdBtnPw.disabled = true;
  cbdBtnPw.classList.add('loading');
  cbdLabelPw.textContent = 'Building Cost Sheet...';
  cbdStatusPw.textContent = '';
  cbdStatusPw.className = 'status';
  try {
    const templateBuffer = portwestBase64ToArrayBuffer(PORTWEST_COST_SHEET_TEMPLATE_B64);
    const { buffer, bucketedCounts } = await buildPortwestCostSheet(templateBuffer, currentPwExtracted);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const styleCode = (currentPwExtracted.coverInfo && currentPwExtracted.coverInfo.styleCode) || 'portwest';
    const a = document.createElement('a');
    a.href = url;
    a.download = `${styleCode}_CostSheet.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    const parts = Object.entries(bucketedCounts).map(([k, v]) => `${k}: ${v}`).join(' | ');
    cbdStatusPw.textContent = `Done — ${parts}`;
    cbdStatusPw.className = 'status ok';
  } catch (err) {
    console.error(err);
    cbdStatusPw.textContent = 'Something went wrong building the Cost Sheet: ' + err.message;
    cbdStatusPw.className = 'status err';
  } finally {
    cbdBtnPw.disabled = false;
    cbdBtnPw.classList.remove('loading');
    cbdLabelPw.textContent = 'Generate Cost Sheet';
  }
});

/* ============================================================
   PORTWEST — SUPPLY CHAIN UI WIRING (multi-file)
   ============================================================ */
const dropzonePwSupply = document.getElementById('dropzonePwSupply');
const fileInputPwSupply = document.getElementById('fileInputPwSupply');
const fileListPwSupply = document.getElementById('fileListPwSupply');
const processBtnPwSupply = document.getElementById('processBtnPwSupply');
const processLabelPwSupply = document.getElementById('processLabelPwSupply');
const statusPwSupply = document.getElementById('statusPwSupply');
const resultsPwSupply = document.getElementById('resultsPwSupply');
const countsPwSupply = document.getElementById('countsPwSupply');
const downloadBtnPwSupply = document.getElementById('downloadBtnPwSupply');

let currentFilesPwSupply = [];
let currentSupplyChainBuffer = null;

function setStatusPwSupply(msg, cls) {
  statusPwSupply.textContent = msg;
  statusPwSupply.className = 'status' + (cls ? ' ' + cls : '');
}

function renderFileListPwSupply() {
  fileListPwSupply.innerHTML = '';
  currentFilesPwSupply.forEach((file, idx) => {
    const bar = document.createElement('div');
    bar.className = 'filebar show';
    bar.style.marginBottom = '10px';
    bar.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      <span class="name"></span>
      <span class="clear">✕ remove</span>`;
    bar.querySelector('.name').textContent = file.name;
    bar.querySelector('.clear').addEventListener('click', () => {
      currentFilesPwSupply.splice(idx, 1);
      renderFileListPwSupply();
    });
    fileListPwSupply.appendChild(bar);
  });
  processBtnPwSupply.disabled = currentFilesPwSupply.length === 0;
}

function addFilesPwSupply(fileList) {
  const pdfs = Array.from(fileList).filter(f => f.type === 'application/pdf');
  if (pdfs.length < fileList.length) {
    setStatusPwSupply('Some selected files were skipped (not PDFs).', 'err');
  } else {
    setStatusPwSupply('');
  }
  currentFilesPwSupply = currentFilesPwSupply.concat(pdfs);
  renderFileListPwSupply();
  resultsPwSupply.classList.remove('show');
  currentSupplyChainBuffer = null;
}

dropzonePwSupply.addEventListener('dragover', e => { e.preventDefault(); dropzonePwSupply.classList.add('drag'); });
dropzonePwSupply.addEventListener('dragleave', () => dropzonePwSupply.classList.remove('drag'));
dropzonePwSupply.addEventListener('drop', e => {
  e.preventDefault(); dropzonePwSupply.classList.remove('drag');
  if (e.dataTransfer.files.length) addFilesPwSupply(e.dataTransfer.files);
});
fileInputPwSupply.addEventListener('change', e => {
  if (e.target.files.length) addFilesPwSupply(e.target.files);
  fileInputPwSupply.value = '';
});

processBtnPwSupply.addEventListener('click', async () => {
  if (!currentFilesPwSupply.length) return;
  processBtnPwSupply.disabled = true;
  processBtnPwSupply.classList.add('loading');
  resultsPwSupply.classList.remove('show');
  try {
    const { fabricRows, trimsRows } = await extractPortwestSupplyChain(currentFilesPwSupply, (fileIdx, totalFiles, name) => {
      processLabelPwSupply.textContent = `Scanning ${fileIdx} / ${totalFiles}: ${name}...`;
      setStatusPwSupply(`Scanning file ${fileIdx} of ${totalFiles} (${name})...`);
    });
    currentSupplyChainBuffer = await buildPortwestSupplyChainWorkbook(fabricRows, trimsRows);
    countsPwSupply.textContent = `Fabric: ${fabricRows.length} unique items | Trims: ${trimsRows.length} unique items (across ${currentFilesPwSupply.length} style${currentFilesPwSupply.length === 1 ? '' : 's'}).`;
    setStatusPwSupply(`Done — Fabric: ${fabricRows.length}, Trims: ${trimsRows.length}.`, 'ok');
    resultsPwSupply.classList.add('show');
  } catch (err) {
    console.error(err);
    setStatusPwSupply('Something went wrong reading these PDFs: ' + err.message, 'err');
    currentSupplyChainBuffer = null;
  } finally {
    processBtnPwSupply.disabled = false;
    processBtnPwSupply.classList.remove('loading');
    processLabelPwSupply.textContent = 'Extract & Generate Supply Chain Sheet';
  }
});

downloadBtnPwSupply.addEventListener('click', () => {
  if (!currentSupplyChainBuffer) return;
  const blob = new Blob([currentSupplyChainBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Portwest_SupplyChain.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});

/* ============================================================
   PORTWEST — EMBEDDED COST SHEET TEMPLATE (Portwest_Format.xlsx)
   Base64-encoded so the tool has no external asset dependency.
   Decoded via portwestBase64ToArrayBuffer() at the top of this file,
   then loaded straight into ExcelJS in buildPortwestCostSheet().
   Regenerate by base64-encoding a fresh Portwest_Format.xlsx if the
   buyer ever issues an updated cost sheet template.
   ============================================================ */
const PORTWEST_COST_SHEET_TEMPLATE_B64 =
  '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' +
  'KKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA' +
  'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACskstqxDAMRfeF/oPRvnEyfVCGcWbRUphtm36AcJQ4TGIHW33k72tS' +
  'OsnAkG6yMUjC9x6Ju9t/d634JB8aZxVkSQqCrHZlY2sF78XLzSOIwGhLbJ0lBQMF2OfXV7tXapHjp2CaPoioYoMCw9xvpQzaUIchcT3ZOKmc75Bj6WvZoz5iTXKTpg/SzzUgP9MUh1KBP5S3IIqhj87/a7uqajQ9O/3RkeULFjLw0MYFRIG+JlbwWyeREeRl+82a9hzP' +
  'QpP7WMrxzZYYsjUZvpw/BkPEE8epFeQ4WYS5XxNGY6ufDDZ2gjm1li5yt2ooDHoq39jHzM+zMW//wciz2Oc/AAAA//8DAFBLAwQUAAYACAAAACEAgX0u/P4vAAD/GgEAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnhtbKydW3PcRral3yfi/AcFn3rsEMW6kRTD0gmg' +
  '7vf79Y2WKYthSfSQtN09E/PfzwaQu7AzP4AiLXccn3Z/tZBAYWWuSiQ2wJ/++99fPr/68+b+4fbu67uTyunZyaubrx/ufrn9+uu7k/Wq8/ry5NXD4/XXX64/3329eXfyn5uHk/9+/1//66e/7u5/e/h0c/P4Slr4+vDu5NPj4+9Xb948fPh08+X64fTu95uv8snHu/sv' +
  '14/yP+9/ffPw+/3N9S/pRl8+v6menZ2/+XJ9+/Uka+Hq/jlt3H38ePvhpnX34Y8vN18fs0bubz5fP8rxP3y6/f1BW/vy4TnNfbm+/+2P319/uPvyuzTx8+3n28f/pI2evPry4ar/69e7++ufP8v3/nelfv3h1b/v5f+q8k9Nd5Ny7OnL7Yf7u4e7j4+n0vKb7Jj59d++' +
  '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';
