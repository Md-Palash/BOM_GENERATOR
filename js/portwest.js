/* ============================================================
   PORTWEST ENGINE
   PDF "Placements" (BOM) table extraction + Cost Sheet merge.
   Depends on: shared.js (homeView)
   ============================================================ */

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
    ['placement',      anchorX('Placement', 0)],
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
  // Longest label first so "Style Code" claims its tokens before the
  // shorter "Style" label gets a chance to match the same starting token.
  const labelsByLengthDesc = [...PORTWEST_COVER_LABELS].sort((a, b) => b.split(' ').length - a.split(' ').length);

  for (const line of lines) {
    // Find every label occurrence on this line (sorted by x), each
    // introduces a value that runs up to the next label (or line end).
    const claimed = new Array(line.items.length).fill(false);
    const labelHits = [];
    for (const label of labelsByLengthDesc) {
      const labelTokens = label.split(' ');
      for (let i = 0; i <= line.items.length - labelTokens.length; i++) {
        if (claimed[i]) continue;
        const slice = line.items.slice(i, i + labelTokens.length);
        if (slice.map(s => s.text).join(' ') === label) {
          labelHits.push({ label, x: slice[0].x, endIdx: i + labelTokens.length });
          for (let k = i; k < i + labelTokens.length; k++) claimed[k] = true;
          break;
        }
      }
    }
    if (!labelHits.length) continue;
    labelHits.sort((a, b) => a.x - b.x);
    for (let h = 0; h < labelHits.length; h++) {
      const startIdx = labelHits[h].endIdx;
      const nextX = h + 1 < labelHits.length ? labelHits[h + 1].x : Infinity;
      const valueItems = line.items.filter((it, idx) => idx >= startIdx && it.x < nextX);
      const text = valueItems.map(it => it.text).join(' ').trim();
      if (text) values[labelHits[h].label] = (values[labelHits[h].label] ? values[labelHits[h].label] + ' ' : '') + text;
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

    // Confirm this is a "Placements" (BOM) page: a "Placements" title
    // sits near the top of the page.
    const isPlacementsPage = items.some(it => it.text === 'Placements' && it.y < 60);
    if (!isPlacementsPage) continue;

    const contentItems = items.filter(it => it.y > 30 && it.y < viewport.height - 10);
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

    for (const line of lines) {
      if (isPortwestBoilerplateLine(line)) continue;

      // Header row repeats at the top of every page and always spans
      // exactly 3 physical lines: main row, the "Weight" wrap of
      // "Material Weight", then the remaining labels. It is NOT reliably
      // followed by a category tag — on continuation pages it's followed
      // straight by data rows of the already-current category — so we
      // consume a fixed line count rather than scanning for a delimiter
      // (mirrors the Malacca engine's header handling).
      if (line.items[0] && line.items[0].text === 'Placement' && line.items[0].x < 30) {
        headerCollecting = true;
        headerItemsBuf = [...line.items];
        headerLinesCollected = 1;
        continue;
      }
      if (headerCollecting) {
        if (headerLinesCollected < 3) {
          headerItemsBuf.push(...line.items);
          headerLinesCollected++;
          continue;
        }
        headerCollecting = false;
        sawHeaderOnThisPage = true;
        // The header layout is structurally identical on every page (it's
        // literally the same repeated table header), so once derived
        // successfully there's no need to re-parse and re-anchor it on
        // each of the (possibly 20+) subsequent pages — a real saving on
        // a long tech pack. Only re-derive if we don't already have it.
        if (!columnBoundaries) {
          columnBoundaries = derivePortwestColumnBoundaries(headerItemsBuf) ||
            portwestColumnsFromFractions(viewport.width);
        }
        // Fall through — this line itself still needs normal processing
        // below (it may be a category tag or a plain data row).
      }

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

      // New logical row starts when this line has content in the
      // Placement column (leftmost); anything else is a wrapped
      // continuation of the previous row's fields (Description/Position
      // most commonly run onto extra lines).
      if (cols.placement.trim()) {
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
function portwestClassifyConsumption(uom, qtyDefault) {
  const u = (uom || '').trim().toUpperCase();
  const qty = parseFloat(qtyDefault) || 0;
  if (u === 'PIECES' || u === 'PCS' || u === 'PC') {
    return { value: qty, numFmt: '0" Pcs"' };
  }
  if (u === 'M' || u === 'MTR' || u === 'MTRS' || u === 'METER' || u === 'METERS') {
    return { value: qty * PORTWEST_METERS_TO_YARDS, numFmt: '0.00" yds"' };
  }
  // Unrecognised UOM — surface the tech pack's own unit rather than guessing.
  return { value: qty, numFmt: `0" ${u.replace(/"/g, '')}"` };
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
    // column group (O-S) — FOB Price by Yds = Fob Price (M) * 0.9144
    // (the exact yard-to-meter factor: 1 yard = 0.9144 m, equivalent to
    // dividing by 1.09361 but expressed as the buyer's preferred formula).
    ws.getCell(r, 16).value = { formula: `O${r}*0.9144` }; // P
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
    const templateResp = await fetch('assets/portwest-format.xlsx');
    if (!templateResp.ok) throw new Error('Could not load the Portwest cost sheet template asset.');
    const templateBuffer = await templateResp.arrayBuffer();
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
