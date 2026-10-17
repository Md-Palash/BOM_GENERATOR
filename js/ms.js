/* ============================================================
   M&S ENGINE
   PDF extraction (Header on page 1 + "BOM" pages only: Fabrics / Trims),
   business rules, "M&S F.O.B. Open Costing" merge, and M&S-specific UI.
   Depends on: shared.js (requestUnlock), haddad.js (homeView)

   Self-contained by design (like kariban.js / decathlon.js): the generic
   Excel helpers are duplicated here with an "MS" suffix.
   ============================================================ */

/* ============================================================
   M&S — PDF EXTRACTION ENGINE
   The M&S tech pack is a PLM report. Only two things are used:
     (1) page 1 header block  -> Stroke / Style Number / Style Description / Season
     (2) pages whose footer says "-- BOM: BOM" -> Fabrics + Trims tables
         (Colourway pages, Labels, Packaging, Processes, Other are ignored)
   Table columns:  Name | Supplier Name | Article Number | Placement |
                   Comments | Width (Trims: Size) | Consumption | [null] | Price
   Every data row starts with a 4-7 digit item code in the Name column
   ("87187 Woven Plain Weave ...") - that is the row-start anchor, the same
   idea Kariban uses. All cells wrap independently, so each cell is rebuilt
   from the items inside the row's vertical band, by x-column.
   ============================================================ */

const MS_LINE_TOL = 2.5;     // pt: items closer than this in y are on one visual line
const MS_COL_TOL = 3;        // pt: slack when matching an item's x to a header's x
const MS_CODE_RE = /^\d{4,7}(?=\D|$)/;
const MS_SECTION_RE = /^(Fabrics?|Trims?|Labels?|Packaging|Processes|Other)\b/i;
const MS_HEADER_KEYS = [
  ['name',        /(^|\s)Name$/i],
  ['supplier',    /^Supplier/i],
  ['article',     /^Article/i],
  ['placement',   /^Placement/i],
  ['comments',    /^Comment/i],
  ['width',       /^(Width|Size)$/i],
  ['consumption', /^Consum/i],
  ['price',       /^Price/i],
];

function msNorm(s) { return (s || '').replace(/\s+/g, ' ').trim(); }

function msClusterLines(items, tol) {
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = [];
  for (const it of sorted) {
    const last = lines[lines.length - 1];
    if (last && Math.abs(last.y - it.y) <= tol) last.items.push(it);
    else lines.push({ y: it.y, items: [it] });
  }
  for (const l of lines) l.items.sort((a, b) => a.x - b.x);
  return lines;
}

// Items on one visual line -> text. Whitespace-only items are dropped
// upstream, so word gaps are recovered from the x-gap instead.
function msJoinLine(items) {
  let out = '', prevEnd = 0;
  for (const it of items) {
    const t = it.str;
    if (out === '') out = t;
    else {
      const gap = it.x - prevEnd;
      const needSpace = !/\s$/.test(out) && !/^\s/.test(t) && gap > 1.0;
      out += (needSpace ? ' ' : '') + t;
    }
    prevEnd = it.x + (it.w || 0);
  }
  return msNorm(out);
}

// Cell text: lines top-to-bottom. Wrapped lines break at word boundaries
// in this report ("Plain" / "Weave"), so wrapped lines are joined with a space.
function msCellText(items) {
  return msClusterLines(items, MS_LINE_TOL).map(l => msJoinLine(l.items)).join(' ').replace(/\s+/g, ' ').trim();
}

function parseMSHeader(items) {
  const flat = msNorm(msClusterLines(items, MS_LINE_TOL).map(l => msJoinLine(l.items)).join(' '));
  const grab = re => { const m = flat.match(re); return m ? msNorm(m[1]) : ''; };

  const styleNo = grab(/Style Number\s*:\s*(\d+)/i) || grab(/Style Number\s*:\s*([A-Za-z0-9-]+?)(?=\s|Product|$)/i);
  const stroke = grab(/Stroke\s*:\s*(\d+)/i);
  const description = grab(/Style Description\s*:\s*(.+?)\s*(?=Stroke\s*:|Block Reference|Sourcing Office|Supplier\s*:|$)/i);

  // Season: only "<year> <season>", e.g. "2025 Spring". The header line
  // reads "Season : T49 CASUAL OUTERWEAR 2025 Spring".
  const si = flat.search(/Season\s*:/i);
  const seg = si >= 0 ? flat.slice(si, si + 400) : flat;
  const SEAS = '(Spring|Summer|Fall|Autumn|Winter)';
  let year = '', season = '';
  let m = seg.match(new RegExp('\\b(20\\d{2})\\s*' + SEAS + '\\b', 'i'));
  if (m) { year = m[1]; season = m[2]; }
  else if ((m = seg.match(new RegExp('\\b' + SEAS + '\\s*(20\\d{2})\\b', 'i')))) { season = m[1]; year = m[2]; }
  season = season ? season.charAt(0).toUpperCase() + season.slice(1).toLowerCase() : '';
  return { styleNo, stroke, description, season: [year, season].filter(Boolean).join(' ') };
}

// A header line is any line where at least 4 known column labels are found
// (tolerant of wrapped/merged labels and stray whitespace).
function msFindHeaderAnchors(texts) {
  const anchors = [];
  for (const [key, re] of MS_HEADER_KEYS) {
    const hit = texts.find(i => re.test(msNorm(i.str)));
    if (hit) anchors.push({ key, x: hit.x, w: hit.w || 0 });
  }
  const keys = anchors.map(a => a.key);
  if (anchors.length < 4 || !(keys.includes('placement') || keys.includes('article'))) return null;
  anchors.sort((a, b) => a.x - b.x);
  // Column ownership does not depend on how the header label is aligned
  // (left / centred): a column starts where the previous label's text ends
  // (never later than its own label's start), and the first (Name) column
  // owns everything to its left.
  anchors.forEach((a, i) => {
    if (i === 0) { a.bound = -Infinity; return; }
    const p = anchors[i - 1];
    const prevEnd = p.w > 0 ? p.x + p.w : (p.x + a.x) / 2;
    a.bound = Math.min(prevEnd, a.x);
  });
  return anchors;
}
// Header cells are not guaranteed to share one baseline: wrapped / bold labels
// ("Consumptio" + "n") float at different heights. So the header is found as a
// small vertical BAND, and its zone extends over any wrapped label lines that
// follow it closely (but never into the first data row).
function msFindHeaders(items) {
  const lines = msClusterLines(items, MS_LINE_TOL);
  // only items that ARE column labels take part, so neighbouring data rows
  // can never split or pollute a header
  const labelItems = items.filter(i => MS_HEADER_KEYS.some(([, re]) => re.test(msNorm(i.str)))).sort((a, b) => a.y - b.y);
  const groups = [];
  for (const it of labelItems) {
    const g = groups[groups.length - 1];
    if (g && it.y - g.last <= 9) { g.items.push(it); g.last = it.y; }
    else groups.push({ items: [it], last: it.y });
  }
  const out = [];
  for (const g of groups) {
    const anchors = msFindHeaderAnchors(g.items);
    if (!anchors) continue;
    const ys = g.items.map(i => i.y);
    const top = Math.min(...ys), bottom = Math.max(...ys);
    // the section title cell ("Fabrics", "Trims" ...) sits in the same header row
    const titles = items.filter(i => MS_SECTION_RE.test(msNorm(i.str)) && i.y >= top - 9 && i.y <= bottom + 9);
    let yEnd = bottom;
    for (const l of lines) {
      if (l.y <= yEnd + 0.1) continue;
      if (l.y - yEnd <= 11 && l.y <= bottom + 20) yEnd = l.y; else break;
    }
    out.push({ y: bottom, top, yEnd, anchors, bandItems: [...g.items, ...titles] });
  }
  return out;
}
function msClassify(x, anchors) {
  let key = null;
  for (const a of anchors) if (x >= a.bound - 0.01) key = a.key;
  return key;
}

function msSectionKey(title) {
  const t = title.toLowerCase();
  return t.startsWith('fab') ? 'fabric' : t.startsWith('tri') ? 'trim' : t.startsWith('lab') ? 'label'
    : t.startsWith('pack') ? 'packaging' : t.startsWith('proc') ? 'process' : 'other';
}

// Contiguous text blocks within one column: consecutive lines closer than
// `gap` belong to the same cell.
function msBlocks(items, gap) {
  const blocks = [];
  for (const l of msClusterLines(items, MS_LINE_TOL)) {
    const last = blocks[blocks.length - 1];
    if (last && l.y - last.bottom <= gap) { last.items.push(...l.items); last.bottom = l.y; }
    else blocks.push({ items: [...l.items], top: l.y, bottom: l.y });
  }
  for (const b of blocks) b.center = (b.top + b.bottom) / 2;
  return blocks;
}

// Parses one BOM page; `state.lastHeader` carries a table's column anchors
// onto a continuation page that doesn't repeat the header row.
//
// Row assignment is alignment-proof. Cells in one row share a common top
// (top-aligned), centre (middle-aligned) or bottom (bottom-aligned), and the
// Name cell always begins with the item code. So: each row's Name block gives
// that row's top/centre/bottom; every other cell block is attached to the
// row it lines up with, in whichever of the three modes fits the whole page
// best. A fixed y-band per row would glue a tall Supplier cell onto the
// previous row whenever cells are middle-aligned.
function parseMSBomPage(items, state, pageNo, debug) {
  const lines = msClusterLines(items, MS_LINE_TOL);

  const headers = [];
  for (const hb of msFindHeaders(items)) {
    const titleItem = hb.bandItems.find(i => MS_SECTION_RE.test(msNorm(i.str)));
    const section = titleItem ? msSectionKey(msNorm(titleItem.str)) : (state.lastHeader ? state.lastHeader.section : 'other');
    headers.push({ y: hb.y, top: hb.top, yEnd: hb.yEnd, section, anchors: hb.anchors });
  }
  if (headers.length) state.lastHeader = headers[headers.length - 1];

  // Page-level boundaries: the repeated page-header band (ends at the
  // "Designer :" line) and the legal footer.
  const designer = items.filter(i => /^Designer/i.test(i.str));
  const topY = designer.length ? Math.max(...designer.map(i => i.y)) + 2 : -Infinity;
  const foot = items.filter(i => /^All orders placed|^All garments must comply|CONFIDENTIAL TO M&S/i.test(i.str));
  const botY = foot.length ? Math.min(...foot.map(i => i.y)) - 2 : Infinity;

  // One segment per table; a continuation page is a single segment that
  // reuses the previous page's header anchors.
  const segments = headers.length
    ? headers.map((h, i) => ({ h, y0: h.yEnd + 0.5, y1: Math.min(botY, i + 1 < headers.length ? headers[i + 1].top - 1 : Infinity) }))
    : (state.lastHeader ? [{ h: state.lastHeader, y0: topY, y1: botY }] : []);

  const rows = [];
  const dbgSeg = [];
  for (const seg of segments) {
    const inSeg = items.filter(it => it.y > seg.y0 && it.y < seg.y1 && it.y > topY);
    const classified = inSeg.map(it => ({ it, key: msClassify(it.x, seg.h.anchors) })).filter(c => c.key);
    const nameItems = classified.filter(c => c.key === 'name').map(c => c.it);
    const starts = nameItems.filter(it => MS_CODE_RE.test(msNorm(it.str))).sort((a, b) => a.y - b.y);
    dbgSeg.push({ section: seg.h.section, rows: starts.length });
    if (!starts.length) continue;

    // Name block per row -> top / centre / bottom of that row
    const rowInfo = starts.map((st, i) => {
      const end = i + 1 < starts.length ? starts[i + 1].y - 1 : Infinity;
      const mine = nameItems.filter(it => it.y >= st.y - 1 && it.y < end);
      const ys = mine.map(it => it.y);
      const top = Math.min(...ys), bottom = Math.max(...ys);
      return { start: st, name: mine, top, bottom, center: (top + bottom) / 2, cols: {} };
    });

    // typical wrapped-line spacing inside a cell -> where one cell ends and the next row's begins
    const diffs = [];
    for (const r of rowInfo) {
      const ly = msClusterLines(r.name, MS_LINE_TOL).map(l => l.y);
      for (let k = 1; k < ly.length; k++) diffs.push(ly[k] - ly[k - 1]);
    }
    diffs.sort((a, b) => a - b);
    const lineGap = diffs.length ? diffs[Math.floor(diffs.length / 2)] : 9;
    const blockGap = lineGap * 1.4;

    // blocks of every other column
    const colKeys = [...new Set(classified.map(c => c.key))].filter(k => k !== 'name');
    const blocks = [];
    for (const k of colKeys) {
      const colItems = classified.filter(c => c.key === k).map(c => c.it);
      for (const b of msBlocks(colItems, blockGap)) { b.key = k; blocks.push(b); }
    }

    const nearest = (b, mode) => {
      let best = 0, bd = Infinity;
      rowInfo.forEach((r, i) => { const d = Math.abs(b[mode] - r[mode]); if (d < bd) { bd = d; best = i; } });
      return { best, d: bd };
    };
    let mode = 'top', modeCost = Infinity;
    for (const m of ['top', 'center', 'bottom']) {
      const cost = blocks.reduce((sum, b) => sum + nearest(b, m).d, 0);
      if (cost < modeCost - 0.5) { modeCost = cost; mode = m; }
    }
    for (const b of blocks) rowInfo[nearest(b, mode).best].cols[b.key] = (rowInfo[nearest(b, mode).best].cols[b.key] || []).concat(b.items);
    dbgSeg[dbgSeg.length - 1].align = mode;

    for (const r of rowInfo) {
      const t = k => msCellText(r.cols[k] || []);
      rows.push({
        page: pageNo, section: seg.h.section,
        name: msCellText(r.name), supplier: t('supplier'), article: t('article'),
        placement: t('placement'), width: t('width'),
        hasPrice: (r.cols.price || []).some(it => /^\$\s*\d/.test(it.str)) || Object.values(r.cols).some(arr => arr.some(it => /^\$\s*\d/.test(it.str))),
      });
    }
  }
  debug.push({ page: pageNo, headers: headers.map(h => ({ section: h.section, y: Math.round(h.y), anchors: h.anchors.map(a => a.key + '@' + Math.round(a.x)).join(' ') })), segments: dbgSeg });
  return rows;
}

async function extractMSPdf(file, onProgress) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  let header = { styleNo: '', stroke: '', description: '', season: '' };
  const rows = [];
  const debug = [];
  const state = { lastHeader: null };
  const dump = [`pages=${pdf.numPages}`];
  const diag = { pages: pdf.numPages, bomPages: 0, headersFound: 0 };

  for (let p = 1; p <= pdf.numPages; p++) {
    onProgress && onProgress(p, pdf.numPages);
    const page = await pdf.getPage(p);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();
    // pdfjs transform x viewport transform = rotation-safe, top-down coordinates
    const items = content.items
      .map(it => {
        const tx = pdfjsLib.Util.transform(viewport.transform, it.transform);
        return { str: it.str, x: tx[4], y: tx[5], w: it.width || 0 };
      })
      .filter(it => it.str.trim().length > 0);

    // raw dump of every item so a misread layout can be diagnosed
    dump.push(`\n=== PAGE ${p}  rotation=${page.rotate}  size=${Math.round(viewport.width)}x${Math.round(viewport.height)}  items=${items.length}`);
    msClusterLines(items, MS_LINE_TOL).forEach(l => dump.push(`y=${l.y.toFixed(1)}\t` + l.items.map(i => `[${i.x.toFixed(1)}|w${i.w.toFixed(1)}] ${JSON.stringify(i.str)}`).join('  ')));

    if (p === 1) header = parseMSHeader(items);

    const flat = items.map(i => i.str).join(' ');
    const footerSaysBom = /BOM:\s*BOM\b/i.test(flat);
    const headerLines = msFindHeaders(items);
    diag.headersFound += headerLines.length;
    if (!footerSaysBom && !headerLines.length) { debug.push({ page: p, skipped: true, items: items.length }); continue; }
    diag.bomPages++;
    const before = rows.length;
    rows.push(...parseMSBomPage(items, state, p, debug));
    debug[debug.length - 1].rows = rows.length - before;
  }
  return { header, rows, debug, diag, dump: dump.join('\n') };
}

/* ============================================================
   M&S — BUSINESS RULES
   ============================================================ */
// Trims-table items that are really fabric-type materials go to the Fabric
// section and follow the Fabric rule (Reference = "Article @ Placement").
const MS_FABRIC_ROUTE_RE = /padding|filling|\bfill\b|insulat|interlining|fusing|fusible|wadding/i;
const MS_THREAD_RE = /\bthread\b/i;

function msFabricItem(r) {
  const ref = r.article ? `${r.article} @ ${r.placement}` : (r.placement ? `@ ${r.placement}` : '');
  return { supplier: r.supplier, reference: ref.trim(), quality: r.name, width: r.width };
}
function msTrimItem(r) {
  const label = r.placement ? `${r.placement} : ${r.name}` : r.name;
  return { label, placement: r.placement, supplier: r.supplier };
}

function applyMSBusinessRules(rows) {
  const fabrics = [], moved = [], trims = [];
  for (const r of rows) {
    if (r.section === 'fabric') fabrics.push(msFabricItem(r));
    else if (r.section === 'trim') {
      if (MS_FABRIC_ROUTE_RE.test(`${r.name} ${r.placement}`)) moved.push(msFabricItem(r));
      else trims.push({ ...msTrimItem(r), isThread: MS_THREAD_RE.test(`${r.name} ${r.placement}`) });
    }
  }
  // thread goes in first, the rest keep tech-pack order
  const orderedTrims = [...trims.filter(t => t.isThread), ...trims.filter(t => !t.isThread)];
  return { fabric: [...fabrics, ...moved], trim: orderedTrims, movedCount: moved.length };
}

/* ============================================================
   M&S COST SHEET TEMPLATE (MS_FORMAT.xlsx, base64) - untouched original
   ============================================================ */
const MS_TEMPLATE_B64 = "UEsDBBQABgAIAAAAIQAOtUlHjgEAABkGAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIooAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsVEtOwzAQ3SNxh8hb1LiwQAg17YLPCgES5QDGnjZWHdvyDNDenokLVYXaRlWzSeTY8z6T8RtNlo0rviChDb4Sl+VQFOB1MNbPK/E+fRzciAJJeaNc8FCJFaCYjM/PRtNVBCy42mMlaqJ4KyXqGhqFZYjgeWcWUqOIl2kuo9ILNQd5NRxeSx08gacBtRhiPLqHmfp0VDws+fNayYf1orhbn2upKqFidFYrYqHyy5t/JIMwm1kNJujPhqFLjAmUwRqAGlfGZJkxvQERG0Mhd3ImcHgc6a+rkiuzMKxtxAu2voeh3dnv6rfuhX9HsgaKV5XoWTXsXS6d/A5p8RHCojwMcmxrcovKRln/p/sAfz6MMr8uexbS+svAHTqAhyR55Z6sX6DcXvWtaBu7QxTx4IPMz9NVZJgOQqSVA+x7FjJoF3OtEpg34is1713ANnaHDq2cvqt5bntuwgb3ED/nzGsKETnKEhwv4C832upBZCBIZGGTHLtu4IaRc/Bkx+1kewNmB7fMwT7+AQAA//8DAFBLAwQUAAYACAAAACEAtVUwI/QAAABMAgAACwAIAl9yZWxzLy5yZWxzIKIEAiigAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKySTU/DMAyG70j8h8j31d2QEEJLd0FIuyFUfoBJ3A+1jaMkG92/JxwQVBqDA0d/vX78ytvdPI3qyCH24jSsixIUOyO2d62Gl/pxdQcqJnKWRnGs4cQRdtX11faZR0p5KHa9jyqruKihS8nfI0bT8USxEM8uVxoJE6UchhY9mYFaxk1Z3mL4rgHVQlPtrYawtzeg6pPPm3/XlqbpDT+IOUzs0pkVyHNiZ9mufMhsIfX5GlVTaDlpsGKecjoieV9kbMDzRJu/E/18LU6cyFIiNBL4Ms9HxyWg9X9atDTxy515xDcJw6vI8MmCix+o3gEAAP//AwBQSwMEFAAGAAgAAAAhAIHa7uKdBAAAYg4AAA8AAAB4bC93b3JrYm9vay54bWysVv9vokgU//2S+x9YQrI/URgEVFK7EZCem9qY6rZ3uV42UxgLKTDeMFSbzf7v9wbFavUKt3tGB+fbh89783lv3vmndZZKz4QVCc0HMjrTZYnkIY2S/HEgf5kHak+WCo7zCKc0JwP5hRTyp4tffzlfUfb0QOmTBAB5MZBjzpeOphVhTDJcnNElyWFmQVmGOXTZo1YsGcFRERPCs1QzdN3WMpzk8gbBYW0w6GKRhMSnYZmRnG9AGEkxB/pFnCyLGi0L28BlmD2VSzWk2RIgHpI04S8VqCxloTN+zCnDDymYvUaWtGbwteGHdGiM+k0wdfSqLAkZLeiCnwG0tiF9ZD/SNYQOXLA+9kE7JFNj5DkRZ7hjxewfZGXvsOxXMKT/NBoCaVVaccB5P4hm7bgZ8sX5IknJ7Ua6El4ur3EmTiqVpRQXfBQlnEQDuQtduiIHA6xcumWSwmxHNzu6rF3s5DxlUkQWuEz5HIRcw0Nk2HbfsMRKEMYw5YTlmBOP5hx0uLXrZzVXYXsxBYVLN+TvMmEEAgv0BbZCi0MHPxRTzGOpZOlA9pz7LwWYfz/FYHCsep5qG/c+XeUphTi73xMoPo6G/yBRHAq7NTB8Q27z/60TgCNzahlOOZPg/9i/gqOY4Wc4GDj+aBu3Y/B87+s33R1aBrJ1dTgybNX0vK7aG1pd1TeGgdt3EXK7/newgtlOSHHJ4+1hC8yBbMLJHk1N8LqeQbpTJtHr+7/p248qnm+aeu67sFSktduErIpXWYiutL5L8oiuBrKKdEiLL4fdVTV5l0Q8HshGx7AgfDZjv5HkMQbGyLC6IqaYIZgN5ANG/oZRAB9VNAeMtD1KVQIFatVTyivRz/hLSqRregbJWuRX4WBLlpgjXsPGEaoOsN5J1pV80xuyIAyyPQG0o7HXzUa1+dQmiJQkJ5EIPIDY621pebTgkCdTGuJ0VvPS5YuPO74fPyi+YlgOtJ3uubYH8W94ZhPepWJ2HAVauxUeHOT7/C6VTr8RKQRL+01IE8XsNSIF+EG4DS6y92lNFNR1FGibyU1x+IQf4UY/gfkn+ms3/UEZKsgRTaeRJmjHo2XOWQIFwQmqgLuTl+RjjuGkrxXDgcZqdoJAL5nQ5ktb8JGCAB1asw13nyxPu/gUb1fwdhWrWaDAexrjor1LxgJ6rLTyyIzgAqqctg4JBHSgNMsDSM845mVr5M8C+XML0nPKcYqalWygxjOroDotYr/ZlRUUZMem+LKao2DOkkyEfmO0XlZJDlLJWw3tpzzIoZAoQyhBxEPkcNRHutEX+RfS71XBqyfc/glcHsjUh129b6r6qGOpZq9vqD2zY6ie6RsjqzvyR64lbk9Rnjv/R5FaFSFOXfcLljFmfM4gu0BuARW5oHzgXF8X+2Rdq+fqHaBoBihQTdTXVde1TdXyg47VRb43soJXssL8xQ+WiD2t2k1A0FA+icqp6juiDbaju8HFZmB7XR2URM6NLwzZ7n5v4QysT0nLxcFty4Xe9WQ+abn2ajT/ehe0XTycuP6w/frhzc3wj/no9/oV2kmHaiBQqAtrmWq1TC7+AQAA//8DAFBLAwQUAAYACAAAACEAtqgYShcBAADZAwAAGgAIAXhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAArJNBS8QwEIXvgv8hzN2mXXUR2XQPirDgSdcfENJpG5omJRN1++8NFdsuLPXS43tD3nwwebv9qTXsCz1pZwVkSQoMrXKFtpWAj+PLzQMwCtIW0jiLAnok2OfXV7s3NDLER1TrjlhMsSSgDqF75JxUja2kxHVo46R0vpUhSl/xTqpGVsg3abrlfp4B+VkmOxQC/KG4BXbsu7j5/2xXllrhs1OfLdpwYQUPkQtjoPQVBgGD/DWzJIICv8ywWZMBTwG9leZV22ZCmbvE52qRLFuT7Nv5hmrEMGGNFvFhsgizXRNGSaOeaqntBDNaS7e6XxOCaumxeA8+1oEmkDN7CeZuVZjQm9i+8ffSoP/W87NC5j8AAAD//wMAUEsDBBQABgAIAAAAIQC0Zd7Q/SQAAO/BAAAYAAAAeGwvd29ya3NoZWV0cy9zaGVldDEueG1szF3pciO5kf6/EfsOtYrZjfEcEos3ud1yeMT7Gt1j+4+DTZVajOYhk1QfdvjPvor/7lP4UfZJNlHIBBJIUCy2KI87Zlrqr4AsIPEhkZlAVb357ef5LPqYrNbT5eLtUXycO4qSxWR5N128f3t0c936sXoUrTfjxd14tlwkb4++JOuj357++7+9+bRcfVg/JMkmAgmL9dujh83msX5ysp48JPPx+nj5mCzgyv1yNR9v4J+r9yfrx1UyvksrzWcn+VyufDIfTxdHWkJ9lUXG8v5+Okkay8nTPFlstJBVMhtvoP3rh+njmqTNJ1nEzcerD0+PP06W80cQ8W46m26+pEKPovmk3n2/WK7G72bQ789xcTyJPq/gvzz8X6DbpLi403w6WS3Xy/vNMUg+0W2W3a+d1E7GEyNJ9j+TmLh4sko+TtUAWlH5r2tSXDKy8lZY4SuFlY0wpa5V/Wl69/borzn88yP8jNVfuR9zsfqL/fnb0embuymMsOpVtEru3x79lK+fjWqlo5PTNymDbqfJpzX7PdqM310ls2SySeAu8VH0l+VyfjUZq8GrAavNP0eKkTMNKhK/Wy4/KGFdqJaD+65TIeq+48lm+jE5S2ZQ+rwM8+DPaUvgV2jEiWkF/51a1Eppf76K7pL78dNsc7ac/TK92zzAbY/jXK1QKR3Rpcvlp04yff+wgVYXjpXaFdXqd18ayXoCHIdWHRfUHSfLGYiHv6P5VM1VoOj4c/rzkxYdH5fLxVw5D7LXmy+q41Bm8rTeLOd48xjlaAlwr1QC/DQSisV4HxEwsKkI+IkiisXjUqlYrqoeZmpFEUXAT2pF5bhaLVfi7DLgXmkz4CfJKO0rAwY4lQE/UUZmRVawKvyk2xf3HAsws+nd4SeKKB/vORY1FAE/UURtz6GIYZ5oTsEvRsZ+pIoNLxmtCvv2JSZiqV9Ip/G+3SFqxZZb+7IzJmqpXyy39hyamMilfiEp1hLYefIuWW9aU2UKnp25MTEutnwJiNNGA80F0SO2/Ij31WieCKJ+oX7k9p1peaKI+sVKIbuYzXYZ4+XQbE+7kSeaqV9sS/Y0o0SzvKXZjsFQlkWbXsaqEB+es97EqZrV4o7b1khp6hfb3YC9zsjDmpmm5UJ1V+9P9AKWrpuN8WZ8+ma1/BSB1wONWT+OlQ+Zr+fVP94elUAvuu96JU0nhFoiC/ExUH7zMJ18+Gmp50lovSzBQj5Rsn/S8vLpwqeAMw3EZZhCUH8NBT+e1nJvTj7COj7BWg1bSM0hVa8poZaE2hLqSKgroZ6E+hIaSGhIkFIL9cfrzsiWUZ4LqN3oHgYwk+6DTgnpWAmBMUv9olRZZ4hUjNobAmkKpCWQtkA6AukKpCeQvkAGAhkKZIRINfX3uNaA6pkZGxePofiz+lPiQH81S1KNVGDczJjGHkV1Ee0ZpgQ1dYiyrYCUvCul7UvpaCAuwU9z64JbqWvL0K16CJWLpg99CQ0kNJTQSEPlnFC7cimzGooMalfiXLVrBNRO/WoIpCmQlkDaAukIpCuQnkD6BrFjUXTHYhAoUnKLDANFym6RkS4SULkKew6ociXOVblGHKZXPKbrIpZZTQ3EZfBkDEerns1r2UI0mG0JdSTUlVDPNJJE9f02DbCInY9Dv8hIAwElw1J0SCUrca6SNcJ5LZCmQFoCaQukI5CuQHoC6QtkIJChQEYaCehP5agOSFIlztWfRhySVj2SYiVjOZoacA1pza3UsmUMRzVkl9IOlinDYmTXAo/sXSmoh2229th0gu41CHQr9paZoS5jaT3CBkrrDLPxkKOgxLmjoJF0NU69jIYG7LLZ9Eu0fKDtAx0f6PpAzwf6PjDwgaEPjDQQoK4KuQ+ptVSeqzaE+OyXUFNCLQm1JdSRUFdCPQn1JTSQ0FBCI4RC2oRQKLM2VfT5rFumUhkeBRFyHTPPp2pgoRicITtly54T1WSljMMWwNrBW3rCOqZimiBNZ0c3gPUCWD+ADQLYMICNEAsNxT5RXQZnLcbAkDnJCDnE1qUY1JSlWhJqS6gjoa6EehLqS2ggoaGERgiFtJk1TkvTyLsjDpVW86mtIZfanrfZwHpxGXwZsxqVvVJNVspSW4tXNc1ah+XcW3rea4eEgaG065/nv3atJBORSKgvoYGEhggV7EIzQig0MCpYOaDvoVKV/sBoiK17WMgiTYG0BNIWSEcgXYH0BNIXyEAgQ4GMEAmpcJ+wrlQ5Burp1JDZNtllx3V0A7aKyHGmsrmg5gKPIGI/7sBCxdTV0UkgCbUk1JZQR0JdCfUk1JfQQEJDCY0QAptCvT4nqMJzX1Vvwl2wUlTzMoBdG4wrsewp8SZcyvOeb8OlrLvs5KxUEj37Kl9VCcPn13kZlaa3eHtU4St47Lv8wUKei9/EQnERGGfsV94PVlkpax0xpORtyHvueSfUhrznkHTDbfB8iF6gDf2geD8FESzkJyGChfw0BBbSRiKdcedUj5v/qkeyi2AhbyQubSFS8TV1uWK9ipsAdutiLhkDwTtsZ29JX2dxdHSYG+e4uSLMxmENtQcEJizOsRQgYcBbu/BiOdiQtdSSdTsBeV0jL5+myT3W9sxl24S+EVMMVRmYy7Y1Q4OVocp9Oksexqvk7kjv/MP1+lBtdK2n6Q49tOo7uM2bk3uVuvdT3UaYzT2fm3ZWQ226MFXUZXl/uF6/oPvHR6cg7rvccS0uwjwINeHSyOOGseI19JpKQTQT6MeNc/n+FIp/DxJC97v1ikJVUdClbCBf8jLKYvrBoSxhnLKIwXQiKjZTxQKNHcpiOYeyUl6H6rIp0DXytlCWRHPKkugtlKXLnLKECcpqkkL3AtSkSpya1J4t1KQqgprqPqmFEPe5NGp5nn80GFv4xy/LSXFDEwLUfB1Xt3PzeTG3JKZwdAoid/BWZVE8X/tlvNVpGdfUEsZ5K7Gm2kZX5rdivcOWwazpbhPGONoJYF1Tdwtv6Xact9SsLbyly5y3hG0zteo6MAskQovAzMIiGjazJIhzmdq4hctUZZuZxXvD3u053Pt5E0uynqc4ltpmYvnlEMWxQaCr67i2neLPi7lNRxZ0WlEUl/p0TLM6V3HIcDKV56YDEeJZEwk1JdSSUFtCHYSKdgp0CbKRSE9CfQkNJDSUdxwhFIgsVYbvcOnAVJqnS50hZGtXA0sVmVeGEF/hJNSWUIdk4fw+znlOdVfW6Umoj5AXgng+9IBKldgWFnVZNeD+9Opm+G1r/O5sud6sf/Pdt/H30MDfbLEPWDM0KHslBlXUtMdhD5kkxAMlLHsikKZAWgJpC6SDCCQ2yJnpEmQtbk9CfQkNCLKO0ZAga2BHCIWUuk9+MA9HC3eExOo8lJeGQkid67RxrB8SYyFY1Y2DJ6GWhNok3dqIDkIQOFv9YhKxwtOPeS/Q65F4KEU1+wFsEMCGAWyEGI9KqViOn9CoeMH5BZaq2iG8lNCVhK4ldCOhW4Lk+RR1Ei274UtTas9mSFJ5nunTWUknQ1IQx6hChTwtNVG4a558US1WyoSxiLlt8DMkVMgyqxu+o5c16QXu2A/e0UukDIKFvKTJMEvbR1iobBfTc2qV2nkx87Ditf3ClHL46TX0MlzKa+lVsFTVm/nXphTfp656CZ4bU8pJanmlbk0pbmsKtl2uyxRIH8d5kyT2zg8Wtx+xNycGKV1sp+1ZehQcvH1m2BoGq4UC+CZdLlor1DKYFd0mrMD2/AmDVdTaPmxWESIqezLL8wh65g42o9U30jgZCp7SB6YUH8Cyn8Yzpcgn8bOiI1PC2u1z00frW1wYjFtH7GPerh9XppyVd20wG3zdGAxSNIEsjr4ObjhEH3Dv70DulnSKERSwq4HM83auxSW1V/G8YcXDRpDCMXsTirsqsnS4RtgWruFlh2uEca4hBi4z3a5Dt3O4RnUdrnkeY49qFjnX6A4O1zxjMTD3fJ5rJMv4v56ZGxk5nGtYq8C5RhjnGknnXKNynGuEca4Rto1r5fpNQYVywDcYsYt8GfgGEy6YvqNO6A1H17aFkstbbVt+91ZHHhO/Dt8I4z4yYVv4hpcdvhHG+YaYwzfCuG2jug7fhF9HpTjfSBrnW9HzQwbUbeUxGdtZ9I86mVKGb14LRqYE5xu1wPb7wpSzKr0MYFcGsz26Jgz4Q3P0xmCCb5phoXTjramUej4ur1RWzs+kbeVVFjtGaT5uxwjjvCJsC6/wssMrwjivEHN4RRjnFdXlvCoKR49KcV6RNIdXwtOjUo4d81bWIaxmqUEHtyZNCx/nhB2jEpxXiDl2jDDLjUsjndsxksd5JeveUN3CXrwiQQFeqVTYQXlFuTXOK8I4rwjbwiu87PCKMM4rxBxeEcZ5RXUdXnmOcy9PpTivSBrnVdXjw4Bq7rBXJMvwyuPnyMjhvMJaDq8I47wi6ZxXhHFeybo3dN/9eEWCJK/StTQzrzKsg6lA5WMxXhFWsAmdhikX5hVd5rwyGOMVYZxXBmO8MnUdXvk+vinFeGWk2TB3YDBn1fOtkym1Lcs4MiUYi4yymJdlMMYiU5exyJRjXlagrvag1FPR+1gnIyjAokA6+EXeOzRMe+qcRYg5LKJyW1iElx0WEcZZhJjDIsI4i6iuwyLfe6fGc++dMJVotH6S772bUixTbDBikW+KTAFOItIVJxFhnETUSU4iKsdJJOve0H0lidKH+fV1tZcK26AXhfi7qwK4hUE/3UgKMGu/nPbOsLAg89gI1cDbtM+8eEpuUCHcgDvO+YeOsUDMtmdaBrOmo42YwzRxtLZLNZ1jdFV/GTTy7Yj2EataCzawEDtZKoyVbgRsqW3xpFAKz5iRUnjCrOop5oIKwT602rkXmrukAsHt2Su6qp47/XjqBRbXVrjx6AlKd1bEkZebQh4CSBVYKGaC+YP2ATPBTdDMFM27JXn63RTOE57O1oAOXLc+XpvB1Qe/wd8aQMjlpv+8IhXS3IyPizX3T/NHzwtqYg2XrLgDUORk1ZhDVnOY2eTPSBqQlbCeuQOnpq7qUFNAQ6pphY0QcqinK6ontO2k9XcJSDGaehkUc0k1tnARb7qFi3jVmtwbEie4iOxLo8sM7bolQQESOhsSu0i4+/w6rK6ChBpySFjxXKcG1qttN5Baiss5wjjnNOZwDjc72EYD3i5mu8E9g7FHZRFjKbiBhIYSGiHkcA71wDnn6+GC9LDd3KGULRTDq1sohlc5xRB6lmLS6t5SOwOMctL+B2CUTj3zh7ALGgJGkcFoEBQ84dLEqy55KG3PyaMxhzwaYqcyukaa1WPPYJw8KM16gAMsxvg0lNAIIYc82GUr64K6jFRxn2C8pKtbeILiNE/8E5BWsl0WscLzPHHbcEty7AbKHxGqpP1wX3igqGJCyQPwBkWw55JgJ0cZJ4c3CG3hDaaTS9a/baGQmBGyjZjDG13V4Q1Kc4wOYZw3GnOMjoCGdE/LiBH1zy6a5wRx3mCXt/AGr27hDV7dwhsj2fIGof14g5U4b1CdAd44qXgcdGBr8C0lGY55p96du6WOkMMbnVym9colfhMrxA5vMB3t8EZjDm805PCG0uvc3hDGeYPSuL0R0BDbxtg1ov5x3mD/OG8Q2sIbvLqFN3h1C2+MZMsbhPbjDVYKLEoqNXpI46JTrc6ipCF+fA82jpS9YemDJkKsVEtCbapoF7iOLNVFyHViMAdc5LzQmGNPBDREaZDZpEEYIQRv6QPINdYqJZhVn+rI2fOvqdEJRkebGnJOkBS9XZ+GCm9BvzZ9L6LqgJSS/4gPSmFD0kYI3rVCx+vU2br8D9er6XxiTtnB/emUXeW4knfjJhE1dVAoHBOwoQ8lovnMJoyPIHaUz2yErPcyJHXwEdSlAiMID88ddEak8rx35ahbqH1946VhIXYITSAtgbQF0kEEDgQYVSLkTAaD8Zff6EbxyUDSmCoJYqpEKKRKld3KPBl27wVDDsiPoRBy8jkl8eIhzLJx9QqohaLYIYY2QmwGdCTUpYrcfTEY17C+p6NhhLiGEeIa1pB2ex1zA0mFw5JVpvHSW3gPFpb8FwhgIfdEXMm3OqwU0bMVwNqIQZ6XnRvyT8lRu1jwamQxq2EwPhC6l85AIMQHAiE+EBoKHGeFsT/sQMicVXoL9aCDVUrJf9ody8Ts5GfTYNZItgzGYjzEYFM9nFDsYAHugZEch/qU8OIax4QXs9N0O65xLMU1rqGQxg+coAHHQBgXDbHdNSzE4+UmYSXbk5Ypx/WLCRh2ag6LORql1A3nMGFcoyiNaxQhrlGEuEY1FNLogRMU6gFi70g2QlyjlG2wHWliqdjRaCArgeWAsWROOgg5GsVEhS3Vo1LWjeyTLK5PTFNwfSLE9amhkD5VzHVA3xrcZaFPDXF9UtzM9YmYo08qxxmKwTTXJ4aX3MpSTc5QwjhDURrXKEJcowhxjWoopNEDh7RFHRc57zEUYWYDS/FXtEioJaG2hDoS6kqoRxDXJrbLQgNZaihljRAKaXOf2C/DEQV1ztef7zL2w1KONrGUs7R5ifCWrNemG7rRSNE84SM2nTpSSFdCPYK4+il8hNluNkZK/olkrBjzJ5JkI4f5Uh1GyjbTfwxjhHVCY8bjS/VSXKXwFyR1lI/lj5mGeHTiI02sxrJ/AmkjAof6rHX2BXVFtZ5A+gIZkGh7/yFBbHsLoYAWSzzGO4AWU3neizl1OMWzHliKM19CLQm1JdRBiG0qdmWpnoT6suJAlhpKaIRQSJs8zDuENmWYV9KQo02M4CAXQfxqYjH1CkMTYASwNmLQJUNNeYeuhHoS6ktoIKGhhEYEpYx1ojpYp63jsFOhyiJlf+Awle1RVZyTaFAp5kQgxDajW7JUGyGVE7Bm0jvI05H1ugixAe5JqC+hgYSGEhoRxPY8CAPrmZ6AOC7lK4VCpVqpVYqlGotw3aHhcd6uodmZ34N1wre+CDn5vZL/0CLV48ODoliQJ0u1EYKDQZb4WJF5eFTRQj2EvADfyzv2WSmSPyCMr4mI1YKny0Z4tWpn589GCF9+K97ye25KWQf2wmDpMuSOJY8gD2G3ZASpXkqt3tplNdlAiOWiBNISSJsEWTe5gxBbX7sS6glRfSpkWzAQhYYCGSHC9l5/pkLcb6t4k/2cCjkj5xH6ggqB528fEfTYdUZNwBlr09XuqPIodueoltX3a/YxnxiG2gl0ptwvGGNnbRJQU5ZqSagtoY6EuhLqSagvoYGEhhIaIcRySD8jVFXnWsXrW87pqtriuj/NncTm2Je7/XdBBfVOGeiebwa4o8hj55fbWRk5w7v/U9ML1s++FtJ/GhgLVfhkxYCYub6yVFtCHYK4nRXBdQ9LuXa27OU++6yUtbMU5FsTMaRO6gDperkZz+Lv1Y5N9Xu1W1P7PoVK+gdka/GsXjWu5Evs41b+yVgUG3IFeYB+CJMqA/SShlh4IpCmQFoCaQukI5AuIsx89yTUl9BAQkOCrGs6QiikRx6a7/YAdx7Phf12xXb+QkmE2LsiGxJqSqiFEHs5SFtCHQl1JdSTUF9CAwkNJTRCCA7S+JuykN/Zx5/erU3cieSLAG6o8oVeQM20Ic5S0ZJQW0IdhGIVhClztVmlphZelKInrfogYR1yjRP4hBp8s3CdrD4m8A4gNeGjCN6AsppOIjXvI++7GEwuGZJeAOsbjCXzCFPGHF73lhqX8EH0IRaFMzeqaLfVvLz8+fLbQal28s3gm3Lxh6OjbW9lwZqBKQLQQQc1ledOEYT4yi6hpoRaEmpLqEN39Aa1WaBXLunxU3vr6eB5Y8eqm7ELYH2DsbEjTI9duiLg0GXYsB9ibT2c4ig4XK4P1ZvE08O44BOY8S7nxHjXjnNxXKgUSpVqIZevxHFZnA8Y4e1CHNgr87D74XZwbn0ziZDDAZGMaMpSLQm1JdShO/oTu4hvlntuYp+PJx/G7+G7p8GpzSRbelDOhJ0eN+U4PbCcpkfqKhjnwE9WDlGAZAOOP9xdPZ9QzVfz+Vo1VyznKpVSxfcpUEpokPfLhuy03mV8IoNZb4ScQRZJkaYs1ZJQW0IduqM/0WmQT3/ePCSr6Hy1nCTrdXCuY5tZxqrHpNIA9w3GBxPraqe7csw/Wwq/V8SEG6KU50e0ehyXS2VIlZRztXxcKsWOM5++CneEgkKDuk8eJcMzHupjI56Dg5AzqLoUz8DKUi0JtSXUoTt6gwqZkt1L8tnTJoIvFEfD8YdET15/YWbS7ezFPjIS9E05PuBYji3MW94cMsTqO+au9+klrBQa1b0yKhnsMW5886mqIWdUBdRMlx/X0ZJQW0IdhIyjlca7Zf/11KyUHR1sqzM6hPHRQUyPjnpDHfhB9UG5YPeDqiU3/vJSKnpxhaxS8FWu+vXp4pGuEbY5NGo8Y0Kv89+6p1TY83OGZdwft57xGUE2Um0gBAduTTJflmpJqC2hjoS6EupJqC+hgYSGBLFAjiAZeqjPTpgzAM/p9rlzoM1Uytsj9iWKloTaBLENN4LY0UoJ9RBKSepkYdQnJw/Qeh2xO60XUCe9F/SRNxVLWaiHpeCZaT/Ig4qHaKsOlJ22Cqid3gvaarcTOgTx5mNF3nwNhZrPY9SvJ4qOOUFLZv9LfXMOVkYGdQjibcVSvK0Y5QpNwwuEX67pbirFGe8eQpKH8FTMy+/YSaW4DJNQD6FAI7gf+rUj1IGFKx0OpnsJ9RAKNIL7TV/fCDxVxxshoF4F/QhJAL7M790IYCZ+Rjf94jodkHqJnBcZWdaeF5k7JudFpojJeZFNsHJgJ+0F04fJ4dMQXwKmUvriiaLgYsbk8Jn0Ejl8MrxEDufzS+RwR+olcjifXyIn63G5XeOV9aDYLjmczy/oFzxqkO1g5Y72qKfcMx3Q3CXH5/OzD/XYdD6t0WcSakioKaGWhNoS6kioK6GehPoSGkhoKKGRhH6W0LmELiR0KaErCV1L6EZCtxL6RUK/l9AfJPRHB3LcZvWReYdXz/MBD2DYZ3bOUgHpwT6iSENCTQm1JNSWUEdCXQn1JNSX0EBCQwmNJPSzhM4ldCGhSwldSehaQjcSupXQLxL6vYT+IKE/OpDLB399eZ4PeIyE80FAjZqAmhJqSagtoY6EuhLqSagvoYGEhhIaOZCrK38NfV5XOrkAToi1pQJq1ATUlFBLQm0JdSTUlVBPQn0JDSQ0lNDIgbSuTtYPSbJpjDfj0zfzZPU+OUtms3U0WT4tYHcQPpbHYP1Jqm6hUu/p7IqpQFeqcCWNBLwr8CW8uvq4GKjWuwJfv6urT4zJK/DFvrr69p68Al/fq6vPl8krzXy13tIvUPbuA9+sqatv1Mg68K2Xei98JV+uqxf9Bu4Tl0Fa6Ap8bw6khetUoE4aTvt6g49sqe+Jhe4D/dGPEYo6oOvgFdjJrasPWgR6mi/U1Uv+Q1eK0NM0DeXrDU6Yq1d3h+qUoE7oylkeGhCoMczXRyG8U6x3Q/fuFeFwe0groPyQhtvlOnzAPqBFUGKIlZ1qHT4eLsvDR4xB6yENwreG4XNsoSvwdVioE9ItnDisq7ODAQ0WCjBfQleaBeiifim6mGMwUvr9ev6VIlzRz6T7YwgztrVlxgL39TlcIS0H0tLcmLgSw5WgDor5unoKM9BTuNILXmkUQW/6mT9xH9BO8EoDeqqeiwuMdAHYqo9tCB1AnaDeYPulrjY/pDRI4cOV0H1gsxWuhHUALdCf7fBa0CjCmOr3FIiegs0IXukXq/WBfprdl1YCjsK5b9lq2Daqq6OyoStluBKaIXAqpK5OZoR0kIOehngAe5FwJTja8Aa3nj7pJEYBGLLlCvAgOBe6MBd6wbkAX88A2xS2WsB4fZzV13UebO2WK2BRw2tHIQdtC84FuNILXmkWYHz04Wgxg2H+bLkCFjXM3jysXvBi2ZAVhhkcvNKE1aulV68T04Q1fOhhubibbqbLxXimvw+xUfv+6z+nr4SEDzvUG4VKpL7MAEXvL59mSbT58giHgZLP6jjQGioeRXef77t3sG0E8cbjarpcTTdf0nf+wJmc5Wr+NBufftP5Ji68Pfq///nfI9jERPDNiZYIb7oItmJX44AJ0LgsDYMw2jYsPbOYrV07ewxxvhWsdOR1+Jtgd3eKhTSEFZtuxrnt/cffD6dGGNsIxzmDKtXrjNkYp9+vc9v2dV1Wr99mctX5fU+Vh+1zud4GXvcL4FMUyvCzUh8UqvATDF+xHMEDgJHePFVk2TVc6u3ytu0xmOHMxN8pGhJ2XHR6EPog+obIhQtON41fkWSwrrdhTcugTAg/WcNgtTucMh37FMNTfAear+odwVyZ6jOir8he+OpvHcKUSH8NuBapl+dkUKxjq9Q58MMp1rFXcTrOBzGE6nuRnAzK2L+iYuHzQPV+IY7S52hhNdutVPXNVtbAw6lUfcjYCj7YtFdf77ZiX3nSDyEAyqBCzxKp9NNm+di9v149gZ/BjWn3Sp/vBcHqTMth3IghBMcZWumYJf0G7cP4EKAkPiSHG2rHKKVsfsWpAy+5g+g8i4FXX2tjUybAwW1e4q7FEt5b7liLQ5l31wgF7OYhXZNhuuLt6qprdiCUPpgtd30ZCLgPtEa6jgyE669qyYcQ8e6e1K7lyQUc7a+louvE5ALr1ldGBg7DcwHn6KBchOxABsfCnc8H1KJnKQ6mRddOQIr6pVy8g4T57Xg2hZ8QEK+jz79MF3fLTxASwxN+0Rf6V0XZesylQ2B/+sath5H1XTKZzsezo2g8my0//TQbLz6oZTBaPyw/dRePT5shxN3j9+namILN1Wq54mCigLdHwyd4LuddEo2jj+PZU/IfapyW80dI5DcXGzgWvkmf4lGH8KPlfaSeBln/6TMc3B7/CW6mEwAQHKpACd4muao/TSHG/ysd8v4Rfsmrv3I/5uDL8+lv6hf1529WobH6gEAup54m0stlbHSdP4XiBs9DFsBVh1DPV+vD7TbkGjar8WQT/flpvNhArsL2FgYrW0cL1FFIwXiDeKBWzseLp3to5dNK5WM203kSTRfgGcHp7vHqLsKBmk8XT5tkfWy7oB5Lz9aH4qv34ep3w3X0CEyDAbh7Ao0LeqlXVWRrLT42+jeV9Hodjf+nnnDjxSSJgKzRpzFo+31iVJu+/DRbY8vbVauzZ7PpevOCCU6EvkpmCah18rRaJYvJl+hpndxF//j7D9E3EXQAFq9ouZhZfneU/cnWg0pgJl8m92d4Jz6dd03bV+mx2pBcfVF2C9KN8KiSHSXcecnYz+qWfqbyp8l6n46+eOZ3F/B85SZaPz0+zqYwbRZjmPXqcQ6AoS3R4mn+LlmxrmaePbVXn+tqGVGWCuYmmzKZDerv/skzhtaAx4fx2rZ3AF5qtvnxU5g350rcXqR5jdmxTsZrlZfHRbyfuVdn4V5dpfJ+9W6ZhRvs8uZpbfo3zNy/xpb+pfJ+bZt2lzyOV5t5sthE9+CnRXqNoJ2YzH1shvvYSB43/1xrpp3M+6fZzKz/d8l6spo+KifZjF4PYqVsc6716kYM2jd9v1DHVxK1oCY/RNP7aPxxPJ2N382smWhnbTL6w+AQS69FT/xN8nkzSBbvNw9f4eBfbVZLeM5vlC4L0Rzd/aJa+0vR5GGsvNxkpbxD1wMuRnewZsJCo+unK8zT/f30s+6vWn0mTocbmTscigDgXcIB5189OahDggyu/0GsJE2lrOFMnA9MJTWP/qlOAcZqD0n0EcYSZg66Acr3gad7IZ5MF15rDbOuufHrBzF3yQbmzlo1NRlPHqKNeuPAdJPMTWubEGDCx4KiM/h5lq9FZ3BY4CxzwAlPU6Z/XjEq0K9KUM/V/td4/vjf6ZO1T4+p1tkKlDnqghdob22xJvmvEf6rcAdMhHpI3EaSuGWXzTjHoYBBhf7/SpE/PQffCCxDZ9DdM/Ui5ExrUWzihlcLR6mxEHXpxShdi0K2WR2ZytzwZyKBg7qi2mz5BgBeJPNBRQlqyUlfJ7MA10vZNGUU/OxAF3bn4XBb1AXb0IVtOTgmE3Vhh14djso4TCasMFmn+NS83WJ/h+9XSs3Z93F0A2rCrRfctczsTYUWt3+1CXtFkfBIRcLKSjmLx0Ad24DjGvCFwGgAJBmo79Jmm7+vn/zauvb5NFenO9PvV2XzgvX68YxL+WuQVCX55slmpTzmxWT2dKcm+f0qmb5/0C+kuHvafPG8S5u0xAMNWccuFHvnjv+VFpoxuGqQOozu9QuynlS6nidCbWgORzn66pRDNtaa+PzVVh1sMZL3hyg5fn8MScTkx09jGL/JQzL5EM3HG/jlDi5uJnYQm9AVOFaftSsmFH/trphsGvbJ6P4MGnyWvcEmrn6FBmPiDzYkZrAfYeNPluqr1BvZ2/qakTK2FTgNIXH0aXqXhq56u6gDKu1kbiauQM+YssOmyZHb6QZcMHfchua3szc/FOuqHPmvkjvW/pZneObj1QewzAnE9pMpJDRUEvnb//wNrkD+QjSA7g+yd984EDAhvD0AOFf8CCZvOF5Ban4dzZJ72HTMHUPgDcn6B/odTvmkKGzDv1tuNss5/eshGd8lsIGZO4b9ivvlEpIY+h9wJyX3KtlACPY4hlXnavoX2AgFqwPbAJA5S51JeFPhcgVbfFPYYlnV1a7lqnun33jwabn6kD7idPr/AAAA//8DAFBLAwQUAAYACAAAACEAsUUN7FwHAAACIQAAEwAAAHhsL3RoZW1lL3RoZW1lMS54bWzsWUuPGzcSvgfIfyD6LuvVrcfAsqGnJ/aMbXhkL3LkSJSaHnZTIKkZC4GBwDnlEiBAdrGXBfa2hyBIgARYI5f8GAM2stkfsUV2S02OqHhsjxdOMDPAjJr6qlisKn6sLl6/+SRh6JQISXnaCarXKgEi6YRPaTrvBA/Ho1IrQFLhdIoZT0knWBEZ3Lzx8UfX8Z6KSUIQyKdyD3eCWKnFXrksJzCM5TW+ICl8N+MiwQoexbw8FfgM9CasXKtUGuUE0zRAKU5A7b3ZjE4IqlWqdVSCf7UaGmv1wY31REMGj6mSemDCxJGehuyUNnLTk6pGy5XsM4FOMesEMP+Un43JExUghqWCLzpBxfwE5RvXy3gvF2Jqh6wlNzI/uVwuMD2pmTnF/HgzaRhGYaO70W8ATG3jhs1hY9jY6DMAPJnAqjNbXJ3NWj/MsRYo++jRPWgO6lUHb+mvb9ncjfSvgzegTH+4hR+N+uBFB29AGT7awke9dm/g6jegDN/Ywjcr3UHYdPQbUMxoerKFrkSNen+92g1kxtm+F96OwlGzlisvUJANm0zTU8x4qi6Sdwl+zMUIwFqIYUVTpFYLMsMTyPQ+ZvRYUHRA5zEk4QKnXMJwpVYZVerwV/+G5pOJLt4j2JLWNoJVcmtI24bkRNCF6gS3QWtgQV4+f/7i2U8vnv37xRdfvHj2fT63UeXI7eN0bsv99q+v//uPz9F/fvznb9/8NZv6PF7a+Fffffnq519+Tz2suHDFy7/98OqnH17+/atfv/3Go70r8LENH9OESHSXnKEHPIEFeuwnx+LNJMYxpo4EjkG3R/VQxQ7w7gozH65HXBc+EsA4PuCt5WPH1qNYLBX1zHwnThzgIeesx4XXAXf0XJaHx8t07p9cLG3cA4xPfXP3ceoEeLhcAO1Sn8p+TBwz7zOcKjwnKVFIf8dPCPGs7lNKHb8e0ongks8U+pSiHqZel4zpsZNIhdA+TSAuK5+BEGrHN4ePUI8z36oH5NRFwrbAzGP8mDDHjbfwUuHEp3KME2Y7/ACr2Gfk0UpMbNxQKoj0nDCOhlMipU/mnoD1WkG/AwzjD/shWyUuUih64tN5gDm3kQN+0o9xsvDaTNPYxn4iTyBFMbrPlQ9+yN0dop8hDjjdGe5HlDjhfj0RPARytU0qEkR/sxSeWN4i3N2PKzbDxMcyXZE47NoV1JsdveXcSe0DQhg+w1NC0MNPPBb0+MLxeWH07RhYZZ/4Eus2dnNVP6dEEmRqnG2KPKDSSdkjMuc77DlcnSOeFU4TLHZpvgtRd1IXTjkvld5jkxMbeJdCiQj54nXKPQk6rOQe7tJ6P8bO2aWfpT9fV8KJ30X2GOzLx2+6L0GGvLEMEPuFfTPGzJmgSJgxhgLDR7cg4oS/ENHnqhFbeuVm7qYtwgBFklPvJDR9bfFzruyJ/j9lj7+AuYSCx6/4XUqdXZSyf67A2YX7A5Y1A7xM7xM4SbY566qquapqgj99VbNrL1/VMle1zFUt43v7ei+1TFG+QGVTdHxM/ye5UPtnRhk7UitGDqTpAEl4u5mOYNC0qUzfctMaXMTwMW88Obi5wEYGCa7+QlV8FOMFtImqprE5l7nquUQLLqF7ZIZN65Wc0216UMvkkE+zDmi1qrudmTslVsV4JdqMQ8dKZehGs+jqbdSbPuncdGLXBmjZNzHCmsw1ou4xorkehIj8nhFmZZdiRdtjRUurX4dqHcWNK8C0TVTg9RvBS3sniMKsswyNOSjVpzpOWZN5HV0dnEuN9C5nMjsDoNxeZ0AR6ba2defy9OqyVLtApB0jrHRzjbDSMIaX4jw77Vb8Zca6XYTUMU+7Yr0bCjOarfcRa00o57iBpTZTsBSddYJGPYJbmAledIIZdI/hY7KA3JH6DQyzOVzTTJTINvzbMMtCSDXAMs4cbkgnY4OEKiIQo0kn0MvfZANLDYcY26o1IIQP1rg20MqHZhwE3Q0ymc3IRNlht0a0p7NHYPiMK7zfGvG3B2tJvoRwH8XTM3TMluIBhhSLmlXtwCmVcIlQzbw5pXBDtiGyIv/OHUw57dpXVCaHsnHMFjHOTxSbzDO4IdGNOeZp4wPrKV8zOHTbhcdzfcC+86n7+qNae84izeLMdFhFn5p+Mn1/h7xlVXGIOlZl1G3er2XBde0110Giek+J15y6FzgQLNOKyRzTtMXbNKw5Ox91TbvEgsDyRGOH3zZnhNcTb3vyg9z5rNUHxLrGNIlvrtjtm29+/BjIYwB3iUumpAkl3GULDEVfdjOZ0QZskScqrxHhE1oK2gk+q0TdsF+L+qVKKxqWwnpYKbWibr3UjaJ6dRhVK4Ne7SkcLCpOqlF2vT+C6wy2yi/5zfjWRX+yvrG5NuFJmZuL/LIx3Fz0V2sXu+hHFAjos0Zt1K63e41Su94dlcJBr1Vq9xu90qDRbw5Gg37Uao+eBujUgMNuvR82hq1So9rvl8JGRS+l1S41w1qtGza7rWHYfZqXNOCFjEpyv4CrjY03/gcAAP//AwBQSwMEFAAGAAgAAAAhANiiNX6XDgAA19kAAA0AAAB4bC9zdHlsZXMueG1s7F3rjqNGFv6/0r4DYmajZLVuwI1vPe2ezPTEUqRsNNLMSlllVi0al7tRuDiAJ+6s8mdfJX/3KfIo+yR7qgC7MBSUMXRD0SMlbmOqqK/OqVPnVofL11vHlj4jP7A8dy5rZ6osIdf0lpZ7N5f/8XExmMpSEBru0rA9F83lBxTIr6/+/KfLIHyw0Yd7hEIJunCDuXwfhusLRQnMe+QYwZm3Ri78svJ8xwjhq3+nBGsfGcsAN3JsZaiqY8UxLFeOerhwTJ5OHMP/abMemJ6zNkLr1rKt8IH0JUuOefHtnev5xq0NQ91qumFKW23sD6WtnzyEXM08x7FM3wu8VXgG/SreamWZKDvcmTJTDHPfE/RcrSdtpKjDFPatX7EnXfHRZwuTT766dDfOwgkDyfQ2bgjk1HfXpOinb5dzWddlKaLKtbeEebr58oufN1746mX08Vfpxd9evFDPVPXmq1eZ3z59mfz6KefXqIdB9PH6Neng65uvZCUZGjUMbXwwjqjVH79Hn8ljGG1HaQz5bVVG43G68XKpOI7y8PDqa8b9k/T9ydBY3cOSoecXpvIvjI5nmTtVRqcTWJgpog0oOg1efUp/vYGvh7QYvLoZfH0zyB/JRMuHeCNB1zuKS4zGwyyMT/HzN8EyGgij6XkeHWNOLGaByQH7YI5NnuqEfuFTD7iHbro2g8KmB7xDN/3nuw/7pkq8Fq8uV567X5JDeDQRQRc/ud4v7gL/BgsV1im+7eoy+FX6bNhwRcMTZnq250shCFRYp+SKazgouuPasK1b38K3rQzHsh+iy0N8gcjg+D7HAomILyrREw6eE/Hbrtc3vmXYuX2mmt/ihyRDPb2LoU7YoyZs9OCGZD7qmjW652l6yJwTl6ZvE+OKOKfOni2a2AfzyQk71UWtjIxJ0lzvLB4/bfGlVk9jDKpVI1VqcOM61+VjybascIrkqH93O5cXC1WdqNc1iL0DNj5WAJARFAn4yqK47k6prej0KUyvqLoIk8daBWOtcfnyTU7d8qJ03qo+kOzyAagUlm2n7Qh85eoSbK4Q+e4Cvkjx3x8f1qCduGAeRloGua/k7jvfeNBAFyJqCU+DwLOtJdaS7q6JThQTdwH/IpX5Nv7Bcpdoi8DMAfMC904NGKtAVZ6161InxH6MJ83w2B8F0nnDT4qVVyLtGgRE8cOs6cmLIYGZF1rYzFbPJrPZbKqNp9PpTD/XdJ3svA2ijUcAdmkyghGMYHY+nY2HMBBVn5Kl1fwIwJp54hGACZ2M4BymYDIaTUfabKjDf02v1pgKYBE+8RyAbyAawUA903TMi/pEVyf6aDg+fhKIjIQd4Nbzl+AXTHxJIKyTa1eXNlqFIKB86+4ef4beGosrLwzBeXZ1ubSMO881bCx+o164WoKfEVyKc9lBS2vjwNOiTc7YhF5s/EJv+BHlT+DoCkNIIeBoAzgrw2yse46OCZ0ImWofRZWZPoZXkv7De/AT57EERcfSe7PzwOyWInVpt7xcmayYoyjW1GTVvmyZU3nsouUhdTT6Uso8BhWZw20Zt3GMs7oQP4YSR+8QT0LmkwRbjVvZaUKj+T310N46TR4f7P2HnTexnhveYTqCqGSYpRrEY1KqZKwd0giq8N7RlHpUHYchcOqVEscgKpmuJpSThvXQJ0TUKO+dgCs2M8FqNZFtf8B25A+rnek6BsNhu6JSD0gs3Q1xMgT+E3yC8Z+RtRp9gY2X1QgHyxtsRLICouGxnyQZ67X9gOPHBET0DQa1//aWGPD7729s6851EN3gve+FyAxJ7g/x0KUQa1HSQQtGEqWsxAPBwfv82W9mThSapyIOo5hLm6qV2Evarkr5DOdVMJBGrYE9KDZIvlGETi7RhMbJBpA7EPGCdO/51q/ARTjpwATmQJApBflgoWXSV37xjfVHtMU5RdhfvN7xjXRvLZcIp44Rb/t2xV41ONUjDw2MsQhNwsZtRKdTSxUnhuTBg+tpeN9vnFvkL0heHJ63PBI+KWjbM3/CERSQjfzU5YYfiawEdkeAamOS9xNJoFKosG7JshSI0iRPK4KPs+H6xuhxsiIn/YG5cUYleu+jlbXdr3HhVz5wyV6FAtlYJNdxCHk/NW2W8pBGC7mXGdoXg+WhdYemYD8DfVz+lMGAk3tj6ZdiAJzXSTS6/A29O6TmwAosIAZWkric2dUA3l6KVVLfOrSwxyRbnVO4AcdzKzad5HhKuIm5l1FbGTdWsbYyymTr6QxQEp4yZVL8rsE3QZS3Me1OY3lURDbS8zc5YIITN7mOmO60jwbORvTaSUPjh52ggAHSXpoW7eX8zimNXvl99M5BQmqu4VpM+Y7t9qc6K9PbHshK4SybPrI+teexVJyMe0osNffZZ5Ufm+mT147iAe5QIs8qaFOAUYuOrnN65mFGRIvMUNacRmUzCOqKH5NaDL32xYtF8Er6G3ZP5HrhhfFM53upQJ/vjWeaYnNq80qr6+I4qTjEGkzI8Ykkacu9pZlBmNnzYmoZxEVeiZZiY+7JnQVHpYKwbCvBWJXDlhIMMeHaOOWBFSURDDJHXEg0IvPEfneYxYz9amNSC6zYoBCO7hTm/q3u0k0r4fTWxkBKsnLzA34pzbkXG3QqztE1xMfEuagwT6kqLaYYp1i+pzOgTfq37CnEJEc3tiH7lL9HOwN5Twu11VJmeAGyMdnOZI7nUye1E2cTrURCB1vwQTxdJHTZnG+R0GV96Z1BRzlsqIyv1LrLoGMEPkWiqCY2w2YPkgpFvBaL0mpBPMZ+r2WScTt6nKLg8H60KR5ji7Uq5yJ6dUNpYQCecFW7cPGVXugervwNEdRTuuJFrKp1lSvzA8ZA0S6mOfMfvC+VMicXTsAFlXKKQlStA8EdNOxgtJfCJnx+PbXeuIMIHSQpy58ozhEp2o1GSZMUA+OA8KnZJx05MMQ1HRlbQ9hjEynHMn82y/Hz0TG3ZDZ/J2VttgkNhzuEd3m3FiOHIxzeBXUgw9pcdYQHUMZ53HVAGX9x1wFl1lXzgKrq5dxnY7sHAd6C8OgrvyoZqGRnfFx3b7xqEOZv0+HE+gG2rGJU5YXEXSKpRbXwShwOHCoEbw2Fjjvl4bV61KIclqiBra5rNqFSuSnpn0bYUrLWsMuV4OxQ1KHLhJQq15bVJlQea9f4tzrsfMPgQCwJEgtl1XrOhghJIezE1d0RT1OqQgkOzienoACeOOf8+IPDjHy+1Ibb9Xp81ULlVNXJ/KBdftXJtrrU+lREk23PPc1h5cNoHrxTR65nMyoV1sC6OeXqu1JHbF9LiRbV+cIpQSpuVIDWQkoJf0xKQQejRBT+/NzzjrLDEUkILIkuTEEGbUK7QKigcHGCiVhFpehFzz0FXT1Gp6WsagouDoL2sgwHJn98xqbUC5ayxzpZCxq79vLhllifNaJt5k08bUDW8HuH6DdGccCFHTxHMxWMbbNRsK6K5qOUz0e3OkosqkoeAEy73rx3AidqJGCL99qOvkyKSjOnsfbU0UW5/YZUYbfTtayOmJKUNKO5obhgQ1cyaeny672nLjd3CyDXsLnct9fE0QV3Tz14KoABxX00Jc3uLY1a0CpYu83fo3IjUm4ONsZOxN2OBE6lD3QceC0BnCPmQAif3jPN53KBUMumkIi1JfXBhcmd092dTNRUIKa3S5i2mbm9JQK9f5ittaTMZvG0FoazoDT9r/uyW2Ol1WfPptYozfYRmPoSZZhASrbck/xZDQBhvVmqUYKcGjg69YVhKbckPqFTdCZHpCPPAPz5BLjnL5H/7RJeXN+fE+DUliNyoQdm1Ek8ccaGKpw8Y8YcqlW06J5AZ09AJQnWlaMVdDnkvlSt4CvVgU2FTm/k/GmpTAU1W5e2i+eHODbmLLEFyPkRuYIY5bFLpYNg6/dpSynXcOI2jShbz+CRj4bXjihbwKDNiCgfIqXXP5f9kpmz0UubD1u6uZkc4tlFbKjCKQxsqGKYgJRuxFIXWlZPp5YYPkvj3QVCHj0b2kRuiPyiI6bVjkRT+WXFebOtVXhP9UT3qcIsK5uwzcKK305lHaAVAx0lkwBowck5kY+j7JUlYSQwk2t3bNtpqHQ1Egpq2obdBcUSqMKe+OeajgzlhZ0OniNoT7jmmziDRpe9YETBH33BN4FT/IrvNPNSu7OY78rQJhhi/DJz7blygxjVSRUT2fYPq+DqEv/xIXywUSCZ3sYN5/JYpq5KruGguXztOY4h7QpmY3+Of7GxIJPj32r8bwCfI/w/df+/5LffZHAW756U9LnxfeSaD4lnHmu3G8sOLRdniOh5Tb73fMewkwZYiu4bqOwG0i5mh//gGvh5NPTcgUej2M8G3ra5OtULOn2PfGzr74eKJ5ur1xHVK6FrRM6ry+V2tSPquQ5UhQtXl+7GWTihFH2QXJzofVQws+G1twRif/Hzxgtf/fF79Pniby9eqGcqmV6F9MDoZkJepLzv5seXLwcjXX3zL/4e8OLCJI4H8uPL//3nvwP9/IgunrFkyPtMF5rZJ33mMbyPgyD23CVyAyR9Nuy5TBY22obIXVIXTM/2fMmCG7cIS2MNS1clal8oBcbkWFiZMOGQJuSUf4E0Ke9iShS1InFS3sezPHmWJ8W7Xg3ypD1LRrDVe5QEwGQGhTg0bm2U1ohBjizRytjY4cfdj3N5//ff0dLaOKBbxne9tz57IeliLu///s66uw/h9awgSEHcfheAJIZPaeNboEV/83Yye/fNYjiYqm+noPKg0WA2evsOFKjrt+/eLWbqUL3+DTRCx3aDi62mz+X7MFxfKEpg3iPHCM4cy/S9wFuFZ6bnKN5qZZlICdY+MpbBPUKhYytDVZ0pM8UxLBeUQejkIrDhLj8GGw/+w/7aXKa+RMMn2wAMmx77bDhW34w0dbA4V7WBPjamg+n4fDRYjLThu7H+9pvRYkSNfVRt7JqqaNp+8KOL0HKQbbkJrRIK0VeBSPC1AISSUEIJsA30Ac/U1f8BAAD//wMAUEsDBBQABgAIAAAAIQDl2CX1/gMAANgLAAAUAAAAeGwvc2hhcmVkU3RyaW5ncy54bWyUVm1v2zYQ/j5g/+EgbEOKppa9OFvbOSpsxV6DWbEbyQm6b4x8tolIlEZScb1fv6OctZiOwdpvEo/39tzdcxy9+1QW8IjayEpdBINePwBUebWWansRrLLZq9cBGCvUWhSVwovggCZ4F33/3cgYC6SrzEWws7Z+G4Ym32EpTK+qUZFkU+lSWPrV29DUGsXa7BBtWYQ/9/u/hKWQKoC8apQlv/3zABol/2owPp68GQTRyMhoZKMMy7oQFuEWhr3+r6PQRqPQiY7i1OrqAeG6Ke9Rd4WXaORWwQ1ueh5RrmVtKfGOyOX11tQip3wpcIP6EYMoRWEqBdC1s9wJgzwoYRvDXdZC2xKV7UriyliCHC4pz65sSiJCEuGDPcBJStcKNC+YRywKZ2GpZc5M3B4L/HV5NnVdSNQs0bgqXejdrPxotWXUB26k0Zoa7NAN/4fuwcog2J00YDB3NQLqJ5iJe0rPnEJcKauFsfQ5l4ryNkA9ClfKoiYY3AEr+FEZ0qf8ug6pR9CFxsD70IhCWhZxXEG1gYWWW9ltoOhOru3OV2MIIUGrmY+VEVt2OBcHmG42MpcOsR+79rLKioKQMN1u8lfkeP0JwlaNFecz5ukT5ovr+ccW+EzLkrXz89eXIn8QhMsWrixyRa81TyLRf+14KkZ4pvJvPn7P1Pi9UFtGEn64Vlk6HITDIZwk6avkcvh6OJi46j39nE1ecCo4mm8j6jKR38mfy3gFhhLg03Yzu7pkM7GMw7v+2fnZ+c2sKxtfwoQpTKbzMLuZw1zcY9HVIOGbAaORvataJvMHZCS1sDsKVPoq+vsqgWy8nIZ/yPK+YPWIifYYzUbnsCzYWI2LotoLzxge+9e1jq/nIy9f7IXZUT6ngLQciB/NzjHYKZRVU7gVdwqmxlzSGG2INGg/GUCbdzFZ6ipHw9r/M2HAyZ6AQRCOOHNB+TN2zpBNaXSE88n4t87xl8lwc8Mm+QhWmqTdXBKhmo3IbaNdnb+RPOLGtjSbCFq4rd9nHM8Wk1b+dbviGOy/Ov9jkjWsqvYWP0FaV9qaPQoNc7tm5D+hyS/EGg2j5UyUXZtX2bx75KLbtNsnrN2GDZNpNr3x3WrFcH+AEmkZvdxolNsdX/jXM/+qbrUY1Tv6haPlly3waVOWQrPpcXTL9g65alWfjZZpCOpmL2e8nxLPLGazq3gKccJqkdwyU+zScsmO4nG6Gs9hsSJI76ZjBqujeH88d/FifjY4H/bOuo7vqkdUkNAr06vp5+P2/aCEe28QJxSONT3PIGWa0vdujDItlHF96GWoj2sTLnNGI/SUYWeJ1b6ryU+irH9LYdZb9CY9WNBDu/VEo/wl/ZBe5tE/AAAA//8DAFBLAwQUAAYACAAAACEAO20yS8EAAABCAQAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzhI/BisIwFEX3A/5DeHuT1oUMQ1M3IrhV5wNi+toG25eQ9xT9e7McZcDl5XDP5Tab+zypG2YOkSzUugKF5GMXaLDwe9otv0GxOOrcFAktPJBh0y6+mgNOTkqJx5BYFQuxhVEk/RjDfsTZsY4JqZA+5tlJiXkwyfmLG9Csqmpt8l8HtC9Ote8s5H1Xgzo9Uln+7I59Hzxuo7/OSPLPhEk5kGA+okg5yEXt8oBiQet39p5rfQ4Epm3My/P2CQAA//8DAFBLAwQUAAYACAAAACEAD3I7qnwHAAAXIgAAIgAAAHhsL2V4dGVybmFsTGlua3MvZXh0ZXJuYWxMaW5rMS54bWycWs1S40YQvqcq76By5ZAcFluShTG1OIX4WcyC14VNch6kwZ6yrHGNJC9wzKvkmqfIo+RJ0sIL0Tc7LahcwC1/PT090/11z8gff31YZ95WmkLp/Kjj7/U6nswTnap8cdS5nZ9/OOh4RSnyVGQ6l0edR1l0fh39+MNH+VBKk4vsSuUrjwbJi6POsiw3h91ukSzlWhR7eiNz+uZem7UoSTSLbrExUqTFUspynXWDXm+/uxYq7+xGOFwn7xlkLcyq2nxI9HojSnWnMlU+Po/V8dbJ4XiRayPuMprsg9/3Hh6ywH8Znx58Z2CtEqMLfV/u0YBdfX+vEvn9PIfdIcz0edj/N1bgd2nxMlq3IujR3Eavaxlr/W0tD817FmI32VOdVGuZl7vlNDKjRdF5sVSbouOZQ5Uedcw4re08r/tErGXR+OxtRXbUCYZB78NMZyrtdN1fHn+Zcl+dCLOVrOLUqLyUqceipiJZiQVFnGP8G3kvDUWk9E5FKWpA9zsv6m9msvw27VrynjFj8rz3OiY+95nnAfM8ZJ73aVmN/urRftUrnMgsqz8fU8iVR52iNPRwOyLXpXcl7mT2sbslF2oY/SPFV+2gqR2A9piWj1Kt3lWRvTFM2BwmhGF+11uZ79S9n4tf2JnULr360YchLkS+kIbVjJqaEWjenI9PWb39pt6+w6JXqCfpbdSGHWLQHGIAQ8RnV935zdUbC3fQHIBIr7F7s6XRVcqaHjY1h6A5NyIvVOnF+oFV93sQN0TATdNfKS28uUpWsuRHwMjD0JuVtbIpeG2IPB9Db6qzRy8WC14bAs7HiPtSLqXxVCnXDvO7PK6z0pm4dSx9S6zmBGOc37znN6f2ErbniDq+9fsu2NhydikK6TkHvLSQmchzCSHxYnpimc7uRK4Eu37N5Ytx9ea9wOkbomZTP3L6hrDps2/OAS8ReSq3MtObuqK4xp0gOiZGyEQqi+W7+CRGPpn3QqeLiKLtc7uIsJ2LzgEvEXk7PXG6hqiYKqmp+MRpsl2MbDfv7Tv9QhRtnRM2RtjOL2cAX1pIxi9ExbqgiOxeSPMkF3qrcj46m8QcIzHPewOni4iirXO7iLCdi85NvkTkldpK594hLDbiSfHFtlktYqwW897Q6ReiaOuc7o8RtvPLuQKXiDxXuaKmxU0pVj2rsoUwLZzSLGUxlrK533N6hyjaNbd3CNt550ReWhW0KugskDLuIfZErO/o/MEHZbPexla99d3lAFG0eQfOGELYybItNaBsxyQ1y/bcdzLtxIKdGE1dHe+p3yzsMUlow0kJEwt28iSTpXcjN9VdphK+jkORtbqAuZt/zyzY3386V9VCnS0eNy29DBREq5+Yu8nkzIL95J4GFq6zzJuJbCtSzfezfrMVjknC9XdG/pkF++ePv9zTwdHOilK39Qk+VBuScCrueLZg59SNJsCecALxge1JQhNOWpxYsE+Szvr5Ix9mwLy+Rb2Bk5wmFuyTkbLNDaA/3+K/gElMhF3oPK2MaGmagYR8i4UCZxMysWBkZOF9pj/sagXNo0FMEuxI4KyXEwt2UeVULfgdeb6AeGle4/qqpHH8mAfuwLJg47yNrwPs360GPnTvOSk1J0IW6PqpaOHKAIiDJPAjdO+IBRsXRrQd0YEOAosOQicdTCzYuBRZy2ZAlgeQ5fUV2WGxEQldbNEtWiHpzqUzmocHnpNfUHl0qU0qcj7SIPcDyH3GcN9nDFsNm9B8GgVAByQ1No0zGzBmkUuuRLltixYgiQCynzMcMoaROq5Uuaxaj3wBcAdJ73B6yNhG4rmm4KA8aTtuAqOEFqNE7ly0YNciE49tqRhC20LS2w5GTDCh8uhaVEaVquVQFgLZkPQO20xEofLoWj6oRLMZFOJJHviHiaio795VGqpJXtc6S/W25QoBaCkEWuIsR4xl5LRrbXTS5jPQVfgeuooGjGXsaK4fRU7X/N24on6CX3NgrfA9rBUxdInKo4msr67oliflySsE8iIJKk7k7pcs2FSsVP16hXcRiCq0upl9JmHtcxq9ItlsFBVQ3g6QUmg1NPvOY9XEgtFlIS0Ya6IP3EMSrNe+uzGzYFNtymoh+JN9H7iHJDTiPjJZsBtNPWwLifaBY0hCG87j/sSC3VRFG4X2gU1IepvG9pkagcqjmTR3bb4BmfStHmfgjgMLNqM7a7HRhj9q9IE4SIIVHDChYN2eZUSKqzZXgB361nlmwIQCwmZkRLbGApBA3yKBARML1mWSrsolnQTo1SifPMADfYsHBu7e04LNNvSqlbcADNC3GGDgPghYsJlR9JolX/F+REACJMHOH7jjy4LNHttuvyJgAJLQgju2bJhQX1tIOQICIAlNuM8aNmwpVCtdRkABJKERd2TZsIpuFVtSJIJsJwltuKPKhlVmJfkjTYTX5FaqH7jDipSaM7nN6R1S6n0mXkn1mo3gCBKeJPBmyGwLwm5XhpKEZ64I8p2kt4l5yLQ8qDy63ZvtHfO+NfN/EmH+/6ZkmQvHwjTfsf33+fldef0bhOcfcNQ/OqB30c3fc4z+BQAA//8DAFBLAwQUAAYACAAAACEAN33tpMQDAABEDgAAJwAAAHhsL3ByaW50ZXJTZXR0aW5ncy9wcmludGVyU2V0dGluZ3MxLmJpbuxWy25cRRCtOt3Tt+c9Y48dO/FjPMHGCcSZGBsM5DFmnIeJcSQIEkIssHQDtmQ5UmLWWEjsEYss8hXwC/wEG36BFWKFxHD63rFjJySMgy2UiGqVbt/urqrT1VXVfV/W5B65LvOyLev8fiGbcpftnvRCaqcHfpH1ovlZRSUrD/NzPmYvkk8QRoSssihzPWk73KKgHYmF9Pu4dHtt6+4WBzer3RkKBBnjRH60Kxn5POt/M23ufov73eLeN+RDmaUnZqXJv4/lGv+X2eqyIrefCi5g6EQiv+PXYCFWAzXGIgMXq7JftpWMk2j/jzdZzWkeBRRRQtlILszCIBFDBI9sUBRGNEyCetRqJkedESITcaGv03L+CUGpxVqE50bJlpypu3pUl2KisGtApbSrHxl1dieyLW9bUg6jvm583R4wRMOlSKWSCgUk6jRS7/uykCoxSF+sBNQfoNqG3InpiEcW4MDFyCIHbtoUtaRlU0FV+7RfaxjAIE5gyAzjJE5hBKMYwzjqmEADp/EKJjGFVzGNMziL1/R1Paczel6bekFn9Q2d03l9U9/Cgr6Nd/AuLuISLuMKWljU97QNQQDiEufaBpEE79qcJRLbKpqSKWtFq6YP/ahpFwoSKHbEjtoxO+4IRRt6WrtQdFrP6FklFJyLZ+LzcRMHoOgCEijmol7Sy3pFW0ooaOsSrrpruI4bWMb79iZW8AFWccs3xKRx88izmkSIFjRxVkCYOssNuEF3wg1Fw9FJPaUjOqpjOm7iPYham9TalNaEJ+gn4KuS6Z6aMHjAmPIV6ytpQBpGqsmYxCkqPo1R5ZCfcC5y3mVdzuVdwRVdCcKYpHyRnvTDTLrkhENe3FyQuooTy36WbY5R9zAvTKSUQzSu8//Ct8+eVXzDwvGDFdWhkKpxknRBK6lMdYyztKww8LsJ+V1YuMh6tcFM3nxqlj57Iko140FP8ixtn7VXP20252eW2+3eTT6vHPf8N/agJITq84JQOKgsudN5cTAfM9Lvqb8tt2SVd82yLMlV9m6zt8gbp1f6t/I3uoaOJJF2dnb2cP8UDlz4FgjNknshxvOS3OEzZE2+Yjpv8/L9iP/bbBu8pL+U+0dyJCxIBabPEJ8pBeY+604AKRzu/JG0jj7QPzsMVSTAmbq904G1PQgeRvc/o4gI+UmNyYhPJ0IRPjx1y/A+QVbjY6BQ0R7nXdu74+m78muuDBxo91o4DkT/63xeD2TS90BIfxXDY/X7njXl8KDSSZ3aC6NB6a29lAfyH9zkR5LAfwEAAP//AwBQSwMEFAAGAAgAAAAhANE4/BFIAQAAOgQAABAAAAB4bC9jYWxjQ2hhaW4ueG1sdFNNa4QwFLwX+h9C7t2sZk0/WN1DoevBQBfq3oObroJGMVLaf99QiNu8p8dMxpk38+L+8N215EuPtulNSqPNlhJtqv7SmGtKy4+3hydK7KTMRbW90Sn90ZYesvu7faXa6rVWjSFOwdiU1tM0vDBmq1p3ym76QRt389mPnZrccbwyO4xaXWyt9dS1LN5uBeucAM32FRmdG3fmTUoTSlo3CmUej59XcDfbH//GfFxmcg6ZfIcQ5xuqcYEQrz878ghyIj+t55xnhLiabrlOiCkRUkbQ8TwjoVoZwTbOMxIyjxxOWMY+KWiexyDde+JbCplyh3exwrxNFWxZoqQy9t2GXoXw78Q3XAi433wBgVlyAXeXI+U8gV0V6Ksjyn5CuzjhdMly58WCo3+roPOZGe5XihW+gA1IlEWiBuT/Jtn812e/AAAA//8DAFBLAwQUAAYACAAAACEAAdSKgGkBAACiAgAAEQAIAWRvY1Byb3BzL2NvcmUueG1sIKIEASigAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAfJJRT8IwFIXfTfwPS99H18EQmzESMTxJQsKMxremvUDj2i5tFfj3doNNUOPjds79cs5J89lBVdEnWCeNniIySFAEmhsh9XaKnstFPEGR80wLVhkNU3QEh2bF7U3Oa8qNhZU1NVgvwUWBpB3l9RTtvK8pxo7vQDE3CA4dxI2xivnwabe4ZvydbQGnSTLGCjwTzDPcAOO6J6IzUvAeWX/YqgUIjqECBdo7TAYEf3s9WOX+PGiVC6eS/liHTue4l2zBT2LvPjjZG/f7/WA/bGOE/AS/Lp/WbdVY6mYrDqjIBafcAvPGFmtTMamYjufG+TBsji/EZsiKOb8Mm28kiIdjsWLhxy6ez+NxmuPfeneyslJ7EEWapGlMkjglZZLRjNDs7q2/60whUNv/lApEFBrRU/9OeRnOH8sF6nkkLZMRzSZ0eB94P+6bhiegOif/nzhuEpKsTCZ0dHdN7ABFG/r6VRVfAAAA//8DAFBLAwQUAAYACAAAACEAYKKa+OQBAADYBAAAEAAIAWRvY1Byb3BzL2FwcC54bWwgogQBKKAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACcVFFP2zAQfp+0/5D5hSeaFDoElWM0lSEeNlatgT0b59JYOHZkH1GzXz87GWkKFWV7O999+vz5uzvTy02logask0anZDpJSARamFzqdUrusuvjcxI55DrnymhISQuOXLKPH+jSmhosSnCRp9AuJSViPY9jJ0qouJv4svaVwtiKoz/adWyKQgq4MuKpAo3xSZKcxbBB0Dnkx/VASHrGeYP/S5obEfS5+6ytvWBGv9S1koKjfyX7LoU1zhQYfd0IUDQeF6lXtwLxZCW2LKHx+EhXgitYeGJWcOWAxtsEvQEeTFtyaR2jDc4bEGhs5ORvb9uMRA/cQZCTkoZbyTV6WQHWH7pY1Q4t+2XsoysB0NHYA/pkF46x41jO2LQD+OBNYM91yyvIo59cr+FfrrjYf0XQ2L/V373rQiZRgftRLLnFPaZM/bBtXem09Z70MlfYKohuzWRsw2DI0VA++rQwDk/ehZq9C3V+ECX8jb0ff9uzV9c1fwjSdkzeC8wMctW38C2+DnZ6UFwH+3wYZmUV3vF6BrrJ9d180b+FqWquW18Yom9SP7q7OjNXHOF5K3aTdFVyC7lfpGFrhgS98QthVSBZlGEe82fM60LY4fv+o2LTs0lymvj1HOVovP2S2B8AAAD//wMAUEsDBBQABgAIAAAAIQBdDI3ANgEAAAICAAAtAAAAeGwvZXh0ZXJuYWxMaW5rcy9fcmVscy9leHRlcm5hbExpbmsxLnhtbC5yZWxzhFFda8IwFH0f7D+UwPYW03ZMN2crQycTJorTt75k6W0bbD5IMtF/v7vhWAVhD+HcJOeecz9G44Nqoz04L43OSNKLSQRamFLqOiPbzYw+kMgHrkveGg0ZOYIn4/z6arSGlgdM8o20PkIV7TPShGCHjHnRgOK+Zyxo/KmMUzzg1dXMcrHjNbA0jvvMdTVIfqYZzcuMuHmZkGhztOj8v7apKilgasSnAh0uWDA4BHCat29S71Y8NKjNXQ0hI5VsAStnr8NiccuVfXq/SeMNvRsUz9ukX6xhLz2U+LbEnhAmxgcEimdApyBoct8hORAg9z/8yhmFpD/RC3QcEFJm/MNJgcEKAYpvB9xC0THsxjR9RPskpUsR0PxUy8TYY+/QevXb2cKUOLuXU9+E5SN2trn8CwAA//8DAFBLAQItABQABgAIAAAAIQAOtUlHjgEAABkGAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAi0AFAAGAAgAAAAhALVVMCP0AAAATAIAAAsAAAAAAAAAAAAAAAAAxwMAAF9yZWxzLy5yZWxzUEsBAi0AFAAGAAgAAAAhAIHa7uKdBAAAYg4AAA8AAAAAAAAAAAAAAAAA7AYAAHhsL3dvcmtib29rLnhtbFBLAQItABQABgAIAAAAIQC2qBhKFwEAANkDAAAaAAAAAAAAAAAAAAAAALYLAAB4bC9fcmVscy93b3JrYm9vay54bWwucmVsc1BLAQItABQABgAIAAAAIQC0Zd7Q/SQAAO/BAAAYAAAAAAAAAAAAAAAAAA0OAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECLQAUAAYACAAAACEAsUUN7FwHAAACIQAAEwAAAAAAAAAAAAAAAABAMwAAeGwvdGhlbWUvdGhlbWUxLnhtbFBLAQItABQABgAIAAAAIQDYojV+lw4AANfZAAANAAAAAAAAAAAAAAAAAM06AAB4bC9zdHlsZXMueG1sUEsBAi0AFAAGAAgAAAAhAOXYJfX+AwAA2AsAABQAAAAAAAAAAAAAAAAAj0kAAHhsL3NoYXJlZFN0cmluZ3MueG1sUEsBAi0AFAAGAAgAAAAhADttMkvBAAAAQgEAACMAAAAAAAAAAAAAAAAAv00AAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzUEsBAi0AFAAGAAgAAAAhAA9yO6p8BwAAFyIAACIAAAAAAAAAAAAAAAAAwU4AAHhsL2V4dGVybmFsTGlua3MvZXh0ZXJuYWxMaW5rMS54bWxQSwECLQAUAAYACAAAACEAN33tpMQDAABEDgAAJwAAAAAAAAAAAAAAAAB9VgAAeGwvcHJpbnRlclNldHRpbmdzL3ByaW50ZXJTZXR0aW5nczEuYmluUEsBAi0AFAAGAAgAAAAhANE4/BFIAQAAOgQAABAAAAAAAAAAAAAAAAAAhloAAHhsL2NhbGNDaGFpbi54bWxQSwECLQAUAAYACAAAACEAAdSKgGkBAACiAgAAEQAAAAAAAAAAAAAAAAD8WwAAZG9jUHJvcHMvY29yZS54bWxQSwECLQAUAAYACAAAACEAYKKa+OQBAADYBAAAEAAAAAAAAAAAAAAAAACcXgAAZG9jUHJvcHMvYXBwLnhtbFBLAQItABQABgAIAAAAIQBdDI3ANgEAAAICAAAtAAAAAAAAAAAAAAAAALZhAAB4bC9leHRlcm5hbExpbmtzL19yZWxzL2V4dGVybmFsTGluazEueG1sLnJlbHNQSwUGAAAAAA8ADwAPBAAAN2MAAAAA";

/* ============================================================
   GENERIC EXCEL UTILITIES (self-contained copy, "MS" suffix)
   ============================================================ */
const CELL_REF_RE_MS = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;
function colLetterToNumMS(letter) { let n = 0; for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function numToColLetterMS(n) { let s = ''; while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function shiftRowsMS(text, insertAt, amount) {
  return text.replace(CELL_REF_RE_MS, (m, d1, col, d2, row) => {
    row = parseInt(row, 10);
    if (row >= insertAt) row += amount;
    return `${d1}${col}${d2}${row}`;
  });
}
function getFormulaMS(cell) { const v = cell.value; return (v && typeof v === 'object' && v.formula) ? v.formula : null; }
function normalizeFormulasMS(ws, maxRow, maxCol) {
  for (let r = 1; r <= maxRow; r++) for (let c = 1; c <= maxCol; c++) {
    const cell = ws.getCell(r, c), v = cell.value;
    if (v && typeof v === 'object' && (v.formula !== undefined || v.sharedFormula !== undefined)) {
      const f = cell.formula; if (f) cell.value = { formula: f };
    }
  }
}
function shiftWholeSheetFormulasMS(ws, insertAt, amount, maxRow, maxCol) {
  for (let r = 1; r <= maxRow; r++) for (let c = 1; c <= maxCol; c++) {
    const cell = ws.getCell(r, c), f = getFormulaMS(cell);
    if (f) cell.value = { formula: shiftRowsMS(f, insertAt, amount) };
  }
}
function snapshotAndUnmergeAllMS(ws) {
  const merges = [];
  const parse = s => { const m = s.match(/^([A-Z]+)(\d+)$/); return { col: colLetterToNumMS(m[1]), row: parseInt(m[2], 10) }; };
  for (const rangeStr of [...(ws.model.merges || [])]) {
    const [a, b] = rangeStr.split(':');
    const pa = parse(a), pb = parse(b || a);
    merges.push({ minRow: Math.min(pa.row, pb.row), maxRow: Math.max(pa.row, pb.row), minCol: Math.min(pa.col, pb.col), maxCol: Math.max(pa.col, pb.col) });
    ws.unMergeCells(rangeStr);
  }
  return merges;
}
function remergeShiftedMS(ws, merges, insertAt, amount) {
  for (let { minRow, maxRow, minCol, maxCol } of merges) {
    if (minRow >= insertAt) { minRow += amount; maxRow += amount; }
    ws.mergeCells(`${numToColLetterMS(minCol)}${minRow}:${numToColLetterMS(maxCol)}${maxRow}`);
  }
}
function copyRowStyleMS(ws, srcRow, dstRow, maxCol) {
  for (let c = 1; c <= maxCol; c++) ws.getCell(dstRow, c).style = JSON.parse(JSON.stringify(ws.getCell(srcRow, c).style));
  ws.getRow(dstRow).height = ws.getRow(srcRow).height;
}
// copies formulas (re-pointed to the new row) AND constants (the 0 inputs, "meter")
function copyRowContentMS(ws, srcRow, dstRow, cols) {
  for (const col of cols) {
    const src = ws.getCell(`${col}${srcRow}`), f = getFormulaMS(src);
    if (f) {
      const shifted = f.replace(CELL_REF_RE_MS, (m, d1, c2, d2, row) => `${d1}${c2}${d2}${parseInt(row, 10) === srcRow ? dstRow : row}`);
      ws.getCell(`${col}${dstRow}`).value = { formula: shifted };
    } else if (src.value !== null && src.value !== undefined) ws.getCell(`${col}${dstRow}`).value = src.value;
  }
}

// ExcelJS' spliceRows only moves cells. Everything else that points at rows
// has to be moved by hand so the sheet stays correct.
function shiftDataValidationsMS(ws, insertAt, amount, cloneFromRow) {
  const model = ws.dataValidations.model, out = {};
  for (const [addr, v] of Object.entries(model)) {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    if (!m) { out[addr] = v; continue; }
    let r = parseInt(m[2], 10);
    if (r >= insertAt) r += amount;
    out[m[1] + r] = v;
  }
  // new rows inherit the validations of the row they were cloned from
  for (const [addr, v] of Object.entries({ ...out })) {
    const m = addr.match(/^([A-Z]+)(\d+)$/);
    if (m && parseInt(m[2], 10) === cloneFromRow) for (let i = 0; i < amount; i++) out[m[1] + (insertAt + i)] = v;
  }
  ws.dataValidations.model = out;
}
function shiftConditionalFormattingsMS(ws, insertAt, amount) {
  for (const cf of (ws.conditionalFormattings || [])) cf.ref = shiftRowsMS(cf.ref, insertAt, amount);
}
function sanitizeSheetNameMS(name) {
  let n = (name || 'Sheet').replace(/[\\/?*\[\]:]/g, '-').replace(/^'+|'+$/g, '').trim();
  if (n.length > 31) n = n.slice(0, 31).trim();
  return n || 'Sheet';
}
function base64ToArrayBufferMS(base64) {
  const bin = atob(base64), bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

/* ============================================================
   M&S — COST SHEET MERGE ENGINE
   No formula text is edited. Only values are written; when Fabric needs
   more rows they are inserted exactly like Excel's "Insert row" would.
   ============================================================ */
const MS_MAXCOL_FORMULAS = 22;   // formulas exist up to column V
const MS_MAXCOL_STYLE = 92;      // template has stray cells out to column CN

const MS_FABRIC = { first: 17, last: 19, cols: { supplier: 3, reference: 4, quality: 5, width: 8 } };
// Trims name slots: left column (C) down to the Woven Main Label row, then
// the second column (E:F); C35:C37 are only a last-resort overflow.
const MS_TRIM_SLOTS = [
  ...[25, 26, 27, 28, 29, 30, 31].map(r => ({ r, c: 3 })),
  ...[25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37].map(r => ({ r, c: 5 })),
  ...[35, 36, 37].map(r => ({ r, c: 3 })),
];
// Right-hand "Trims price + Cost Summary": rows 32-34 hold the fixed label rows.
const MS_SUMMARY_ROWS = [25, 26, 27, 28, 29, 30, 31, 35, 36, 37];
const MS_SUMMARY_ITEM_COL = 16;      // P
const MS_SUMMARY_SUPPLIER_COL = 22;  // V

function msLinesNeeded(text, widthChars) {
  if (!text) return 1;
  const cpl = Math.max(3, Math.floor(widthChars * 0.8));
  return String(text).split('\n').reduce((n, para) => n + Math.max(1, Math.ceil(para.length / cpl)), 0);
}
function msColWidth(ws, from, to) {
  let w = 0;
  for (let c = from; c <= to; c++) w += ws.getColumn(c).width || 8.43;
  return w;
}
function msGrowRow(ws, row, lines, perLine) {
  const need = lines * perLine;
  const cur = ws.getRow(row).height || 15;
  if (need > cur) ws.getRow(row).height = need;
}
function msSetText(cell, value, horizontal) {
  cell.value = value || null;
  cell.alignment = Object.assign({}, cell.alignment, { horizontal: horizontal || (cell.alignment && cell.alignment.horizontal) || 'center', vertical: 'middle', wrapText: true });
}

async function buildMSCostSheet(templateArrayBuffer, items, header) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateArrayBuffer.slice(0));
  const ws = wb.worksheets[0];
  const warnings = [];

  // --- Template plumbing -------------------------------------------------
  // The template's dropdown lists live in an external workbook on someone's
  // local drive (names like [1]Reference Data). ExcelJS can't carry that
  // link, and a dangling reference would make Excel "repair" the file - so
  // those names and the dropdowns that use them are removed.
  const names = wb.definedNames.model.filter(n => !n.ranges.some(r => r.includes('[1]')));
  const deadList = /^(Ref\w+|Packaging)$/i;
  const dv = ws.dataValidations.model;
  for (const addr of Object.keys(dv)) {
    const f = dv[addr] && dv[addr].formulae && dv[addr].formulae[0];
    if (typeof f === 'string' && deadList.test(f.trim())) delete dv[addr];
  }
  normalizeFormulasMS(ws, ws.rowCount, MS_MAXCOL_FORMULAS);
  let maxRow = ws.rowCount;

  // --- Header ------------------------------------------------------------
  if (header) {
    if (header.stroke) ws.getCell('D4').value = header.stroke;          // Stroke No
    if (header.styleNo) ws.getCell('G4').value = header.styleNo;        // Design Ref
    if (header.description) ws.getCell('J4').value = header.description; // Description (J4:M4)
    if (header.season) ws.getCell('K6').value = header.season;          // Season, e.g. "2025 Spring"
  }

  // --- Trims (left + second column) and right-hand summary table ----------
  const trims = items.trim;
  const nameW = { 3: msColWidth(ws, 3, 3), 5: msColWidth(ws, 5, 6) };
  trims.forEach((t, i) => {
    const slot = MS_TRIM_SLOTS[i];
    if (!slot) { warnings.push(`Trim "${t.label.slice(0, 40)}…" had no free slot`); return; }
    msSetText(ws.getCell(slot.r, slot.c), t.label);
    msGrowRow(ws, slot.r, msLinesNeeded(t.label, nameW[slot.c]), 15);
  });
  trims.forEach((t, i) => {
    const r = MS_SUMMARY_ROWS[i];
    if (!r) { if (i === MS_SUMMARY_ROWS.length) warnings.push(`Trims summary table only has ${MS_SUMMARY_ROWS.length} free rows - extra trims are listed on the left only`); return; }
    msSetText(ws.getCell(r, MS_SUMMARY_ITEM_COL), t.placement);
    msSetText(ws.getCell(r, MS_SUMMARY_SUPPLIER_COL), t.supplier);
    msGrowRow(ws, r, Math.max(msLinesNeeded(t.placement, msColWidth(ws, 16, 16)), msLinesNeeded(t.supplier, msColWidth(ws, 22, 22))), 15);
  });
  // merge the summary table's title across to the Supplier column
  try {
    ws.mergeCells('P23:V23');
    ws.getCell('P23').alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  } catch (e) { console.warn('P23:V23 merge skipped', e); }

  // --- Fabric (insert rows between the first and last slot if needed) ----
  const fab = items.fabric;
  const capacity = MS_FABRIC.last - MS_FABRIC.first + 1;
  const amount = Math.max(0, fab.length - capacity);
  if (amount > 0) {
    const insertAt = MS_FABRIC.last;
    const oldMerges = snapshotAndUnmergeAllMS(ws);
    const rowMerges = oldMerges.filter(m => m.minRow === insertAt).map(m => ({ minCol: m.minCol, maxCol: m.maxCol }));
    ws.spliceRows(insertAt, 0, ...Array.from({ length: amount }, () => []));
    shiftWholeSheetFormulasMS(ws, insertAt, amount, maxRow + amount, MS_MAXCOL_FORMULAS);
    for (let i = 0; i < amount; i++) {
      const nr = insertAt + i;
      copyRowStyleMS(ws, insertAt + amount, nr, MS_MAXCOL_STYLE);
      copyRowContentMS(ws, insertAt + amount, nr, ['I', 'K', 'M', 'P', 'Q', 'R', 'T', 'U', 'V']);
    }
    remergeShiftedMS(ws, oldMerges, insertAt, amount);
    for (let i = 0; i < amount; i++) for (const p of rowMerges) {
      try { ws.mergeCells(`${numToColLetterMS(p.minCol)}${insertAt + i}:${numToColLetterMS(p.maxCol)}${insertAt + i}`); } catch (e) {}
    }
    shiftDataValidationsMS(ws, insertAt, amount, insertAt + amount);
    shiftConditionalFormattingsMS(ws, insertAt, amount);
    // Named ranges (FabCosts, Total1 ...) - shifting them is what makes the
    // untouched formulas (=SUM(FabCosts)...) cover the new rows.
    for (const n of names) n.ranges = n.ranges.map(rg => { const i = rg.indexOf('!'); return rg.slice(0, i + 1) + shiftRowsMS(rg.slice(i + 1), insertAt, amount); });
    maxRow += amount;
  }
  const fc = MS_FABRIC.cols;
  const widths = { supplier: msColWidth(ws, 3, 3), reference: msColWidth(ws, 4, 4), quality: msColWidth(ws, 5, 6), width: msColWidth(ws, 8, 8) };
  fab.forEach((f, i) => {
    const r = MS_FABRIC.first + i;
    msSetText(ws.getCell(r, fc.supplier), f.supplier);
    msSetText(ws.getCell(r, fc.reference), f.reference);
    msSetText(ws.getCell(r, fc.quality), f.quality);
    msSetText(ws.getCell(r, fc.width), f.width);
    msGrowRow(ws, r, Math.max(...['supplier', 'reference', 'quality', 'width'].map(k => msLinesNeeded(f[k], widths[k]))), 15);
  });

  // --- Sheet name + named ranges (rewritten with the new sheet name) -----
  const tabName = sanitizeSheetNameMS(header && header.styleNo);
  ws.name = tabName;
  const quoted = `'${tabName.replace(/'/g, "''")}'`;
  wb.definedNames.model = names.map(n => ({ name: n.name, ranges: n.ranges.map(rg => quoted + rg.slice(rg.indexOf('!'))) }));
  wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });

  let buffer = await wb.xlsx.writeBuffer();
  // ExcelJS writes a purely numeric sheet name unquoted inside named ranges
  // (10818268!$M$17), which Excel reads as invalid and the totals break.
  // Patch workbook.xml so the names use '10818268'!$M$17.
  if (/^\d/.test(tabName)) {
    if (typeof JSZip === 'undefined') warnings.push('JSZip not loaded - named ranges may not resolve when the sheet name starts with a digit');
    else {
      const zip = await JSZip.loadAsync(buffer);
      let xml = await zip.file('xl/workbook.xml').async('string');
      const esc = tabName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      xml = xml.replace(new RegExp('(<definedName[^>]*>)' + esc + '!', 'g'), `$1'${tabName.replace(/'/g, "''")}'!`);
      zip.file('xl/workbook.xml', xml);
      buffer = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
    }
  }
  return { buffer, counts: { fabric: fab.length, trim: trims.length, insertedRows: amount }, warnings };
}

/* ============================================================
   M&S — UI WIRING
   ============================================================ */
if (typeof document !== 'undefined' && document.getElementById('brandMS')) {
  if (typeof pdfjsLib !== 'undefined' && !pdfjsLib.GlobalWorkerOptions.workerSrc) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  const msHubView = document.getElementById('msHubView');
  const msView = document.getElementById('msView');

  document.getElementById('brandMS').addEventListener('click', () => {
    requestUnlock('ms', () => { homeView.hidden = true; msHubView.hidden = false; });
  });
  document.getElementById('msHubBackBtn').addEventListener('click', () => { msHubView.hidden = true; homeView.hidden = false; });
  document.getElementById('msHubBomCard').addEventListener('click', () => { msHubView.hidden = true; msView.hidden = false; });
  document.getElementById('backBtnMS').addEventListener('click', () => { msView.hidden = true; msHubView.hidden = false; });

  const dropzoneMS = document.getElementById('dropzoneMS');
  const fileInputMS = document.getElementById('fileInputMS');
  const filebarMS = document.getElementById('filebarMS');
  const fileNameMSEl = document.getElementById('fileNameMS');
  const clearFileMS = document.getElementById('clearFileMS');
  const processBtnMS = document.getElementById('processBtnMS');
  const processLabelMS = document.getElementById('processLabelMS');
  const statusMSEl = document.getElementById('statusMS');
  const resultsMS = document.getElementById('resultsMS');
  const downloadBtnMS = document.getElementById('downloadBtnMS');
  const cbdBtnMS = document.getElementById('cbdBtnMS');
  const cbdLabelMS = document.getElementById('cbdLabelMS');
  const cbdStatusMS = document.getElementById('cbdStatusMS');

  const debugBtnMS = document.getElementById('debugBtnMS');
  let lastDumpMS = '';
  let currentFileMS = null, currentMSRows = null, currentMSItems = null, currentMSHeader = null;

  function setStatusMS(msg, cls) { statusMSEl.textContent = msg; statusMSEl.className = 'status' + (cls ? ' ' + cls : ''); }

  function setFileMS(file) {
    if (!file || file.type !== 'application/pdf') { setStatusMS('Please choose a PDF file.', 'err'); return; }
    currentFileMS = file;
    fileNameMSEl.textContent = file.name;
    filebarMS.classList.add('show');
    processBtnMS.disabled = false;
    setStatusMS('');
    resultsMS.classList.remove('show');
    currentMSRows = null; currentMSItems = null;
  }

  dropzoneMS.addEventListener('dragover', e => { e.preventDefault(); dropzoneMS.classList.add('drag'); });
  dropzoneMS.addEventListener('dragleave', () => dropzoneMS.classList.remove('drag'));
  dropzoneMS.addEventListener('drop', e => {
    e.preventDefault(); dropzoneMS.classList.remove('drag');
    if (e.dataTransfer.files.length) setFileMS(e.dataTransfer.files[0]);
  });
  fileInputMS.addEventListener('change', e => { if (e.target.files.length) setFileMS(e.target.files[0]); });
  clearFileMS.addEventListener('click', e => {
    e.stopPropagation();
    currentFileMS = null; fileInputMS.value = ''; debugBtnMS.hidden = true;
    currentMSRows = null; currentMSItems = null; currentMSHeader = null;
    filebarMS.classList.remove('show');
    processBtnMS.disabled = true;
    resultsMS.classList.remove('show');
    setStatusMS('');
  });

  processBtnMS.addEventListener('click', async () => {
    if (!currentFileMS) return;
    processBtnMS.disabled = true;
    processBtnMS.classList.add('loading');
    resultsMS.classList.remove('show');
    try {
      const { header, rows, debug, diag, dump } = await extractMSPdf(currentFileMS, (p, total) => {
        processLabelMS.textContent = `Scanning page ${p} / ${total}...`;
        setStatusMS(`Scanning page ${p} of ${total}...`);
      });
      lastDumpMS = dump; debugBtnMS.hidden = false;
      console.log('[M&S] header', header);
      console.log('[M&S] table headers per page', debug);
      console.table(rows);
      currentMSRows = rows; currentMSHeader = header;
      currentMSItems = applyMSBusinessRules(rows);
      const fabricRows = rows.filter(r => r.section === 'fabric').length;
      const trimRows = rows.filter(r => r.section === 'trim').length;
      if (fabricRows + trimRows === 0) {
        setStatusMS(`No Fabrics / Trims rows were found. Pages: ${diag.pages}, BOM pages detected: ${diag.bomPages}, table header lines found: ${diag.headersFound}. Please click "Download debug info" below and send it so the layout can be checked.`, 'err');
      } else {
        const noPrice = rows.filter(r => (r.section === 'fabric' || r.section === 'trim') && !r.hasPrice).length;
        const warn = noPrice ? ` — ${noPrice} row(s) look incomplete (no price found), please check the extracted data` : '';
        const missing = ['stroke', 'styleNo', 'description', 'season'].filter(k => !header[k]);
        const hdrWarn = missing.length ? ` — header field(s) not found: ${missing.join(', ')}` : '';
        setStatusMS(`Done — Style ${header.styleNo || '(not found)'}, Stroke ${header.stroke || '(not found)'}: Fabric ${currentMSItems.fabric.length} (incl. ${currentMSItems.movedCount} moved from Trims), Trims ${currentMSItems.trim.length}${warn}${hdrWarn}.`, (noPrice || missing.length) ? 'err' : 'ok');
        resultsMS.classList.add('show');
      }
    } catch (err) {
      console.error(err);
      setStatusMS('Something went wrong reading this PDF: ' + err.message, 'err');
    } finally {
      processBtnMS.disabled = false;
      processBtnMS.classList.remove('loading');
      processLabelMS.textContent = 'Extract Tech Pack';
    }
  });

  debugBtnMS.addEventListener('click', () => {
    const blob = new Blob([lastDumpMS], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `ms_debug_${(currentFileMS && currentFileMS.name || 'techpack').replace(/\.pdf$/i, '')}.txt`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  downloadBtnMS.addEventListener('click', () => {
    if (!currentMSRows) return;
    const routed = new Set(currentMSItems.fabric.map(f => f.quality));
    const wsData = [['Page', 'Section', 'Name', 'Supplier', 'Article Number', 'Placement', 'Width / Size', 'Goes To']];
    currentMSRows.forEach(r => wsData.push([r.page, r.section, r.name, r.supplier, r.article, r.placement, r.width,
      r.section === 'fabric' ? 'Fabric' : r.section === 'trim' ? (routed.has(r.name) ? 'Fabric (moved from Trims)' : 'Trims') : 'ignored']));
    const sheet = XLSX.utils.aoa_to_sheet(wsData);
    sheet['!cols'] = [{ wch: 6 }, { wch: 10 }, { wch: 70 }, { wch: 34 }, { wch: 22 }, { wch: 26 }, { wch: 12 }, { wch: 24 }];
    const wbk = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wbk, sheet, 'Extracted Data');
    XLSX.writeFile(wbk, `${(currentMSHeader && currentMSHeader.styleNo) || 'ms_tech_pack'}_extracted.xlsx`);
  });

  cbdBtnMS.addEventListener('click', async () => {
    if (!currentMSItems) return;
    cbdBtnMS.disabled = true;
    cbdBtnMS.classList.add('loading');
    cbdLabelMS.textContent = 'Building Cost Sheet...';
    cbdStatusMS.textContent = ''; cbdStatusMS.className = 'status';
    try {
      const { buffer, counts, warnings } = await buildMSCostSheet(base64ToArrayBufferMS(MS_TEMPLATE_B64), currentMSItems, currentMSHeader);
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${(currentMSHeader && currentMSHeader.styleNo) || 'ms_style'}_Cost_Sheet.xlsx`;
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      cbdStatusMS.textContent = `Done — Fabric: ${counts.fabric}${counts.insertedRows ? ` (${counts.insertedRows} row(s) inserted)` : ''}, Trims: ${counts.trim}.` + (warnings.length ? ' Note: ' + warnings.join('; ') : '');
      cbdStatusMS.className = 'status ' + (warnings.length ? 'err' : 'ok');
    } catch (err) {
      console.error(err);
      cbdStatusMS.textContent = 'Something went wrong building the Cost Sheet: ' + err.message;
      cbdStatusMS.className = 'status err';
    } finally {
      cbdBtnMS.disabled = false;
      cbdBtnMS.classList.remove('loading');
      cbdLabelMS.textContent = 'Generate Cost Sheet';
    }
  });
}
