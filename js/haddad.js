/* ============================================================
   HADDAD ENGINE
   PDF extraction, Cost Break Down merge, bug report form,
   and Haddad-specific UI wiring.
   Depends on: shared.js (homeView, requestUnlock)
   ============================================================ */
/* ============================================================
   EXTRACTION ENGINE
   Locates each of the three "detail grid" sections (Fabric, Trim,
   Labels & Packaging) on every page, finds each blue-numbered item
   column, and merges its first 11 field values into one string.
   ============================================================ */

const SECTIONS = {
  'Fabric': {
    fields: ['Internal Code','Dev Code','Mill / Vendor Ref no','Mill / Vendor','Name','Content','Construction','Weight','Width','Finish','Coated'],
    terminator: 'COLORWAYS',
    outputCount: 11,
  },
  'Trim': {
    fields: ['Internal Code','Dev Code','Vendor Ref no','Vendor','Name','Type','Size','Quantity','Special Instructions','Location/Placement','Size scale'],
    terminator: 'COLORWAYS',
    outputCount: 7,
    typeField: 'Type',
    // Colorway-variant columns under the same physical slot (e.g. 4 shank
    // colors, all "CENTER SHANK-SINGLE PRONG", "27 LIGNE (L)", "@ CF
    // CLOSURE") share these three fields verbatim, differing only in Name/
    // Internal Code/Vendor Ref no.
    dedupeKeyFields: ['Type','Size','Location/Placement'],
  },
  'Labels & packaging': {
    fields: ['Internal Code','Dev Code','Vendor Ref no','Vendor','Name','Type','Art Page','Placement','Variable','Quantity','Special instructions','Size scale'],
    terminator: 'COLORWAYS',
    outputCount: 6,
    typeField: 'Type',
    dedupeKeyFields: ['Type','Placement'],
  },
};

function norm(s){ return s.replace(/\s+/g,' ').trim().toLowerCase(); }
// Same as norm(), but also collapses "XS - XL" / "XS-XL" spacing
// differences so a page-1 size label and an item's own Size-scale text
// compare equal regardless of how the dash was spaced in the source PDF.
function normSize(s){ return norm(s || '').replace(/\s*-\s*/g, '-'); }

function nearestAnchor(x, anchors){
  let best = 0, bestDist = Infinity;
  anchors.forEach((a,i)=>{ const d = Math.abs(x-a); if(d<bestDist){bestDist=d; best=i;} });
  return best;
}

function extractFromPageItems(items, pageNum){
  if (items.length === 0) return null;
  // The page's section/pagination title (e.g. "Fabric 5-6/6") always sits in
  // the top title row of the page. Searching the whole page's text for this
  // pattern risks false positives from body content that happens to contain
  // words like "trim" or "fabric" (e.g. "HOOD TRIM 60MM"). Restrict the
  // search to the topmost row only.
  const maxY = Math.max(...items.map(it => it.y));
  const titleRowItems = items.filter(it => it.y >= maxY - 40);
  const titleText = titleRowItems.map(i=>i.str).join(' ');
  const secMatch = titleText.match(/(Fabric|Trim|Labels\s*&\s*packaging)\s*[\d>]/i);
  if (!secMatch) return null;
  let sectionKey = secMatch[1].replace(/\s+/g,' ').trim();
  sectionKey = Object.keys(SECTIONS).find(k => norm(k) === norm(sectionKey));
  if (!sectionKey) return null;
  const { fields, terminator, outputCount } = SECTIONS[sectionKey];

  const startTok = items.find(it => norm(it.str) === norm(fields[0]));
  if (!startTok) return null;
  const labelX = startTok.x;

  const labelYs = [startTok.y];
  let searchAbove = startTok.y;
  for (let i=1;i<fields.length;i++){
    const candidates = items.filter(it => norm(it.str)===norm(fields[i]) && Math.abs(it.x-labelX)<=12 && it.y<searchAbove);
    if (candidates.length===0){ labelYs.push(null); continue; }
    candidates.sort((a,b)=>b.y-a.y);
    labelYs.push(candidates[0].y);
    searchAbove = candidates[0].y;
  }
  const lastKnownY = [...labelYs].reverse().find(y=>y!==null);
  const termCandidates = items.filter(it => norm(it.str).includes(norm(terminator)) && it.y < lastKnownY);
  termCandidates.sort((a,b)=>b.y-a.y);
  const termY = termCandidates.length ? termCandidates[0].y : lastKnownY - 100;

  const rowTol = 4;
  const valueTokens = items.filter(it => Math.abs(it.y-startTok.y)<=rowTol && it!==startTok && Math.abs(it.x-labelX)>12);
  const anchors = [...new Set(valueTokens.map(t=>Math.round(t.x)))].sort((a,b)=>a-b);
  const numCols = anchors.length;
  if (numCols===0) return null;

  for (let i=0;i<labelYs.length;i++){
    if (labelYs[i]===null){
      const prev = labelYs[i-1];
      const nextIdx = labelYs.findIndex((v,j)=> j>i && v!==null);
      const next = nextIdx!==-1 ? labelYs[nextIdx] : termY;
      labelYs[i] = prev - (prev-next)/((nextIdx!==-1?nextIdx:fields.length)-i+1);
    }
  }

  const boundaries = labelYs.map((y,i)=> i===0 ? y+6 : labelYs[i-1]);
  const lowerBounds = labelYs.map((y,i)=> i<labelYs.length-1 ? labelYs[i+1] : termY+8);

  // Fabric-only: parse the Colorways sub-table (below the COLORWAYS
  // divider at termY) to see, per column, which colorway rows have a real
  // (non "N/A") value. This is the only reliable signal available for
  // detecting colorway-variant duplicate Fabric columns that share one
  // physical slot - unlike Trim/Labels, Fabric's Internal Code/Name/
  // Content can all differ between such alternates, and the blue numbered
  // heading itself is drawn as vector graphics, not text.
  let colorwayCoverage = null;
  if (sectionKey === 'Fabric') {
    const rows = [];
    for (const it of items) {
      if (it.y >= termY - 2) continue; // above/at the divider - not a colorway row
      let row = rows.find(r => Math.abs(r.y - it.y) <= 4);
      if (!row) { row = { y: it.y, items: [] }; rows.push(row); }
      row.items.push(it);
    }
    colorwayCoverage = Array.from({length:numCols}, () => new Set());
    for (const row of rows) {
      const rowName = row.items.filter(it => Math.abs(it.x - labelX) <= 12)
        .sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' ').trim();
      if (!rowName) continue;
      const byCol = Array.from({length:numCols}, () => []);
      row.items.forEach(it => {
        if (Math.abs(it.x - labelX) <= 12) return;
        byCol[nearestAnchor(it.x, anchors)].push(it);
      });
      byCol.forEach((toks, ci) => {
        const val = toks.sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' ').trim();
        if (val && norm(val) !== 'n/a') colorwayCoverage[ci].add(rowName);
      });
    }
  }

  const cols = Array.from({length:numCols}, ()=> Array.from({length:fields.length}, ()=>[]));

  const overallUpper = boundaries[0];
  const overallLower = lowerBounds[lowerBounds.length - 1];
  for (const it of items){
    if (Math.abs(it.x-labelX)<=12 && fields.some(f=>norm(f)===norm(it.str))) continue;
    if (!(it.y <= overallUpper && it.y > overallLower)) continue;
    // assign to whichever field's label is nearest in y - this correctly
    // resolves cases where a wrapped line sits closer to (or even slightly
    // above) an adjacent field's label than its own
    let best = 0, bestDist = Infinity;
    for (let i=0;i<labelYs.length;i++){
      const d = Math.abs(it.y - labelYs[i]);
      if (d < bestDist){ bestDist = d; best = i; }
    }
    const ci = nearestAnchor(it.x, anchors);
    cols[ci][best].push(it);
  }

  // Column header/title text (e.g. "CF ZIPPER TEETH") sits directly above
  // the field-value grid. Note: the blue numbered badge itself is drawn as
  // vector graphics, not text, so it never appears here - this only
  // recovers the title text next to it. Some columns' title carries a
  // "CANADA" region qualifier that isn't part of any extracted field value
  // - capture it here so it can be filtered out later. Bounded relative to
  // the table's own position (not the page's absolute top) so the page's
  // buyer-info band ("Fall 2026 ... Player: NONE") is never mistaken for a
  // column title on pages where that band sits close to the table.
  const headerTokens = items.filter(it => it.y > boundaries[0] && it.y < boundaries[0] + 35 && Math.abs(it.x - labelX) > 12);
  const headerCols = Array.from({length:numCols}, ()=>[]);
  headerTokens.forEach(it => { headerCols[nearestAnchor(it.x, anchors)].push(it); });
  const columnHeaders = headerCols.map(tokens=>{
    const byY = {};
    tokens.forEach(t=>{ const key = Math.round(t.y/3)*3; (byY[key]=byY[key]||[]).push(t); });
    return Object.keys(byY).map(Number).sort((a,b)=>b-a)
      .map(y => byY[y].sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' '))
      .join(' ').trim();
  });

  const results = [];
  const { typeField, dedupeKeyFields } = SECTIONS[sectionKey];
  const keyIdxs = (dedupeKeyFields || []).map(f => fields.indexOf(f)).filter(i => i >= 0);
  const typeIdx = typeField ? fields.indexOf(typeField) : -1;
  const sizeScaleIdx = fields.indexOf('Size scale');
  const quantityIdx = fields.indexOf('Quantity');

  for (let c=0;c<numCols;c++){
    const fieldTexts = cols[c].map(tokens=>{
      const byY = {};
      tokens.forEach(t=>{
        const key = Math.round(t.y/3)*3;
        (byY[key]=byY[key]||[]).push(t);
      });
      const lines = Object.keys(byY).map(Number).sort((a,b)=>b-a).map(y=>
        byY[y].sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' ')
      );
      return lines.map(l => l.replace(/^[•\u2022\-]\s*/, '').trim()).join(' ').trim();
    });
    const nameIdx = fields.indexOf('Name');
    const internalCodeIdx = fields.indexOf('Internal Code');
    const vendorIdx = fields.indexOf('Vendor') >= 0 ? fields.indexOf('Vendor') : fields.indexOf('Mill / Vendor');

    // Fabric-specific: "Coated" is the last field - drop it if "NO", show
    // "COATED" instead of "YES" if coated
    if (sectionKey === 'Fabric'){
      const coatedIdx = fields.indexOf('Coated');
      if (coatedIdx >= 0){
        const v = norm(fieldTexts[coatedIdx]);
        if (v === 'no') fieldTexts[coatedIdx] = '';
        else if (v === 'yes') fieldTexts[coatedIdx] = 'COATED';
      }
    }

    const typeVal = typeIdx >= 0 ? fieldTexts[typeIdx] : '';
    let unit;
    if (sectionKey === 'Fabric') unit = 'YDS';
    else if (/elastic/i.test(typeVal)) unit = 'YDS';
    else unit = 'PCS'; // Trim, Labels & packaging, Print/Heat Transfer/Embroidery default

    // drop thread items in the Trim section entirely
    if (sectionKey === 'Trim' && norm(typeVal) === 'thread') continue;

    results.push({
      section: sectionKey,
      page: pageNum,
      itemName: fieldTexts[nameIdx] || '',
      internalCode: internalCodeIdx >= 0 ? (fieldTexts[internalCodeIdx] || '') : '',
      supplier: vendorIdx >= 0 ? (fieldTexts[vendorIdx] || '') : '',
      unit: unit,
      extractedInfo: fieldTexts.slice(0, outputCount).filter(v=>v).join(' '),
      _type: typeVal,
      _name: fieldTexts[nameIdx] || '',
      // Which size tier(s) this item applies to (e.g. "BIG GIRLS ALPHA
      // (XS - XL)" + "TODDLER" stacked) - empty for Fabric, where the field
      // doesn't exist, and for any item whose Size scale couldn't be read.
      _sizeScale: sizeScaleIdx >= 0 ? (fieldTexts[sizeScaleIdx] || '') : '',
      // Raw Quantity field text (e.g. "1", "2") - the merge step turns this
      // into the +5% consumption value written to the Accessories sheet.
      _quantity: quantityIdx >= 0 ? (fieldTexts[quantityIdx] || '') : '',
      // Identifies colorway-variant columns that belong to the same
      // physical slot (e.g. 4 shank-hardware colors, all "CENTER
      // SHANK-SINGLE PRONG" / "27 LIGNE (L)" / "@ CF CLOSURE", differing
      // only in Name/Internal Code) - used to collapse them down to one per
      // size tab. The blue numbered heading itself can't be used for this:
      // it's drawn as vector graphics on the page, not text, so it isn't
      // present in the extracted PDF text at all.
      _groupKey: keyIdxs.map(i => norm(fieldTexts[i])).join('|'),
      // Fabric-only: sorted colorway names (pipe-joined) where this
      // column's value isn't "N/A" - used to detect colorway-variant
      // duplicates of the same slot (see dedupeFabricByColorway).
      _fabricColorways: colorwayCoverage ? [...colorwayCoverage[c]].sort().join('|') : '',
      // Raw column-header text, kept only long enough to check for a
      // "CANADA" region qualifier; stripped before results are returned.
      _columnHeader: columnHeaders[c] || '',
    });
  }
  // Fabric-only: expose the column anchors so a continuation page (no
  // repeated "Internal Code" row, just more Colorways rows - see
  // parseFabricColorwayContinuation) can align its values to the same
  // columns.
  if (sectionKey === 'Fabric') results._fabricAnchors = anchors;
  return results;
}

// Drops any item whose column header calls out a Canada-only region (e.g.
// a "... - CANADA" variant column) - its global counterpart, which exists
// as its own separate column without that qualifier, is kept instead.
function filterCanadaItems(results){
  return results.filter(r => !/canada/i.test(r._columnHeader || ''));
}

// Merges separate zipper component rows (tape / teeth / puller-slider) that
// belong to the same zipper into one combined row per size grouping. Parts
// are usually identical across every size tier (one teeth column, one tape
// column, etc. shared by all sizes); when a part instead has its own
// column per size tier, that's visible the same way as everywhere else -
// via that part's own Size scale field - so this now produces one combined
// row per distinct grouping found among the parts, instead of keeping only
// the "biggest size" grouping and discarding the rest.
function mergeZipperItems(results){
  const ZIPPER_RE = /zipper|zip\b/i;
  const TEETH_RE = /teeth|chain/i;
  const SLIDER_RE = /puller|slider|pull(?:\s|-)?tab/i;
  const TAPE_RE = /tape/i;
  const PART_RE = /tape|teeth|chain|puller|slider|pull(?:\s|-)?tab/i;

  // A bucket with just one part is shared by every size grouping, exactly
  // as before. A bucket with several parts picks the one whose own Size
  // scale matches this particular grouping (falls back to the first part
  // if none match exactly).
  function pickForGroup(bucket, key){
    if (bucket.length === 0) return null;
    if (bucket.length === 1) return bucket[0];
    return bucket.find(it => normSize(it._sizeScale) === key) || bucket[0];
  }

  const out = [];
  let i = 0;
  while (i < results.length){
    const r = results[i];
    const isZipperPart = ZIPPER_RE.test(r._type) || ZIPPER_RE.test(r._name) ||
                          (PART_RE.test(r._type) && ZIPPER_RE.test(results.slice(Math.max(0,i-2), i+3).map(x=>x._name+' '+x._type).join(' ')));
    if (isZipperPart){
      const group = [r];
      let j = i + 1;
      while (j < results.length && results[j].section === r.section &&
             (ZIPPER_RE.test(results[j]._type) || ZIPPER_RE.test(results[j]._name) || PART_RE.test(results[j]._type) || PART_RE.test(results[j]._name))){
        group.push(results[j]);
        j++;
      }
      if (group.length > 1){
        const teethBucket = group.filter(g => TEETH_RE.test(g._type) || TEETH_RE.test(g._name));
        const sliderBucket = group.filter(g => SLIDER_RE.test(g._type) || SLIDER_RE.test(g._name));
        const tapeBucket = group.filter(g => TAPE_RE.test(g._type) || TAPE_RE.test(g._name));
        const leftover = group.filter(g => !teethBucket.includes(g) && !sliderBucket.includes(g) && !tapeBucket.includes(g));
        const allBuckets = [teethBucket, sliderBucket, tapeBucket, leftover];

        // Only a bucket that actually has more than one part can express a
        // real size split; a lone shared part shouldn't manufacture an
        // extra grouping just because its own Size scale happens to differ
        // in wording from another bucket's.
        const splitBuckets = allBuckets.filter(b => b.length > 1);
        const groupKeys = splitBuckets.length
          ? [...new Set(splitBuckets.flatMap(b => b.map(it => normSize(it._sizeScale))))]
          : [normSize(r._sizeScale)];

        for (const key of groupKeys){
          const chosen = allBuckets.map(bucket => pickForGroup(bucket, key)).filter(Boolean);
          if (chosen.length === 0) continue;
          out.push({
            section: r.section,
            page: r.page,
            itemName: chosen.map(g=>g.itemName).filter(Boolean).join(' + '),
            internalCode: chosen.map(g=>g.internalCode).filter(Boolean).join(' + '),
            supplier: chosen.map(g=>g.supplier).filter(Boolean).join(' + '),
            unit: 'PCS',
            extractedInfo: chosen.map(g=>g.extractedInfo).filter(Boolean).join(' + '),
            _type: 'ZIPPER', _name: r._name,
            _sizeScale: key,
            _groupKey: '', // already a single combined item - not subject to group-key dedup
            // The combined zipper is one physical unit - its consumption is
            // driven by the first chosen part's own Quantity (normally the
            // teeth, since that's what's present in virtually every group).
            _quantity: chosen[0] ? chosen[0]._quantity : '',
            // One physical slot has one printed heading (e.g. "CF ZIPPER"),
            // even though its teeth/tape/puller parts sit in separate PDF
            // columns that may each pick up their own fragment of header
            // text. Prefer the group's first-encountered part's header
            // (leftmost column, where the shared heading is most often
            // printed); fall back to the first non-blank header found
            // anywhere else in the group rather than concatenating every
            // part's header together, which would read as a run-on mess.
            _columnHeader: (r._columnHeader && r._columnHeader.trim())
              || ((chosen.find(g => g._columnHeader && g._columnHeader.trim()) || {})._columnHeader || ''),
          });
        }
        i = j;
        continue;
      }
    }
    out.push(r);
    i++;
  }
  return out;
}

function stripInternalFields(results){
  // _sizeScale, _quantity, _groupKey, and _fabricColorways are intentionally
  // kept - mergeIntoCostBreakDown uses them to route items into the
  // correct size tab(s), compute the +5% consumption value, and collapse
  // colorway-variant duplicate columns down to one per slot. _columnHeader
  // is kept too (renamed to the public field `columnHeader`) - the Supply
  // Chain Sheet's Trims tab uses it as the displayed "Item name" (the
  // heading printed above the item's own column in the tech pack, e.g.
  // "CF ZIPPER"), rather than the item's own descriptive Name field value
  // (e.g. "#5 OPEN END VISLON"). Unused by the BOM / Cost Break Down path.
  return results.map(({ _type, _name, _columnHeader, ...rest }) => ({ ...rest, columnHeader: (_columnHeader || '').trim() }));
}

// Fixed placement (in PDF points, page 1) of the product sketch image in
// this tech-pack template layout.
const PAGE1_IMAGE_RECT = { x0: 307, x1: 762, top: 47, bottom: 536.9 };

async function extractProductImage(pdfDoc){
  try {
    const page = await pdfDoc.getPage(1);
    const scale = 3;
    const viewport = page.getViewport({ scale });
    const canvas = document.createElement('canvas');
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext('2d');
    await page.render({ canvasContext: ctx, viewport }).promise;

    const cropX0 = PAGE1_IMAGE_RECT.x0 * scale;
    const cropX1 = PAGE1_IMAGE_RECT.x1 * scale;
    const cropY0 = PAGE1_IMAGE_RECT.top * scale;
    const cropY1 = PAGE1_IMAGE_RECT.bottom * scale;
    const cropW = Math.round(cropX1 - cropX0);
    const cropH = Math.round(cropY1 - cropY0);
    if (cropW <= 0 || cropH <= 0) return null;

    const cropCanvas = document.createElement('canvas');
    cropCanvas.width = cropW;
    cropCanvas.height = cropH;
    const cropCtx = cropCanvas.getContext('2d');
    cropCtx.drawImage(canvas, cropX0, cropY0, cropW, cropH, 0, 0, cropW, cropH);

    const blob = await new Promise(resolve => cropCanvas.toBlob(resolve, 'image/png'));
    if (!blob) return null;
    const buffer = await blob.arrayBuffer();
    return { buffer, width: cropW, height: cropH };
  } catch (err) {
    console.error('Product image extraction failed:', err);
    return null;
  }
}

// Turns a full page-1 size label into the short token used in a Cost Break
// Down tab name, e.g.:
//   "BIG GIRLS ALPHA (XS - XL)" -> "XS-XL"   (parenthetical range)
//   "GIRLS 4-6X"                -> "4-6X"    (trailing size code)
//   "TODDLER"                   -> "TODDLER" (single word, used as-is)
function shortenSizeLabel(label){
  const paren = label.match(/\(([^)]+)\)/);
  if (paren) return paren[1].replace(/\s*-\s*/g, '-').trim();
  const tokens = label.trim().split(/\s+/);
  if (tokens.length > 1) return tokens[tokens.length - 1];
  return label.trim();
}

function extractHeaderInfo(items){
  const maxSize = Math.max(...items.map(it => it.size || 0));
  const styleTokens = items.filter(it => Math.abs((it.size||0) - maxSize) < 0.5 && it.y > 500)
    .sort((a,b) => a.x - b.x);
  const style = styleTokens.map(t => t.str.trim()).join(' ').trim();

  const nameLabel = items.find(it => norm(it.str) === 'name');
  let itemName = '';
  if (nameLabel){
    const valTok = items.find(it => Math.abs(it.y - nameLabel.y) <= 3 && it.x > nameLabel.x + 20);
    if (valTok) itemName = valTok.str.trim();
  }

  const sizeLabel = items.find(it => norm(it.str) === 'size / sample');
  let sizeCandidates = [];
  if (sizeLabel){
    const labelX = sizeLabel.x;
    const below = items.filter(it => Math.abs(it.x - labelX) <= 3 && it.y < sizeLabel.y - 2)
      .sort((a,b) => b.y - a.y);
    const nextLabelY = below.length ? below[0].y : sizeLabel.y - 100;
    const firstVal = items.filter(it => Math.abs(it.y - sizeLabel.y) <= 3 && it.x > labelX + 20)
      .sort((a,b)=>a.x-b.x)[0];
    const colX = firstVal ? firstVal.x : labelX + 68;
    sizeCandidates = items.filter(it => Math.abs(it.x - colX) <= 5 && it.y <= sizeLabel.y + 2 && it.y > nextLabelY + 2)
      .sort((a,b) => b.y - a.y)
      .map(it => it.str.trim())
      .filter(Boolean);
  }

  // Every size candidate is kept - previously only the single "biggest"
  // tier survived and the rest were discarded. One Cost Break Down tab is
  // generated per candidate found here.
  const sizes = sizeCandidates.map(label => ({ label, token: shortenSizeLabel(label) }));

  return { buyer: 'HADDAD', style, item: itemName, sizes };
}

async function extractPdf(file, onProgress){
  const arrayBuffer = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({data: new Uint8Array(arrayBuffer)}).promise;
  let all = [];
  let headerInfo = null;
  let lastFabricAnchors = null;
  let lastFabricItems = null;
  for (let p=1; p<=doc.numPages; p++){
    onProgress(p, doc.numPages);
    const page = await doc.getPage(p);
    const textContent = await page.getTextContent();
    const items = textContent.items
      .map(it => ({str: it.str, x: it.transform[4], y: it.transform[5], size: it.transform[0]}))
      .filter(it => it.str.trim() !== '');
    if (p === 1) headerInfo = extractHeaderInfo(items);
    const res = extractFromPageItems(items, p);
    if (res) {
      all = all.concat(res);
      if (res._fabricAnchors) {
        lastFabricAnchors = res._fabricAnchors;
        lastFabricItems = res.filter(r => r.section === 'Fabric');
      } else {
        // A different section was found - Fabric's colorway table can no
        // longer be continuing here or on any later page.
        lastFabricAnchors = null;
        lastFabricItems = null;
      }
    } else if (lastFabricAnchors && lastFabricItems) {
      // No section header found - check whether this is a Colorways
      // continuation of the last Fabric page (see
      // parseFabricColorwayContinuation) before moving on.
      const extra = parseFabricColorwayContinuation(items, lastFabricAnchors);
      if (extra) {
        lastFabricItems.forEach((item, ci) => {
          const existing = new Set((item._fabricColorways || '').split('|').filter(Boolean));
          extra[ci].forEach(name => existing.add(name));
          item._fabricColorways = [...existing].sort().join('|');
        });
      }
    }
  }
  all = filterCanadaItems(all);
  all = mergeZipperItems(all);
  const productImage = await extractProductImage(doc);
  return { items: stripInternalFields(all), productImage, headerInfo };
}

// Supply Chain Sheet extraction - reuses the same per-column field parsing
// as extractPdf. Zipper-part merging IS applied here (same as the BOM
// path - teeth/tape/puller collapse into one combined row per size
// grouping, with Item code becoming e.g. "T0047 + T0356 + T0560A"), run
// per file/tech pack since the merge relies on rows still being in their
// original consecutive PDF order. Canada filtering and the BOM's
// colorway/group-key dedup are still skipped here - the Supply Chain
// Sheet's own dedup (buildSupplyChainWorkbook) operates on exact-row
// matches instead. No product image is needed here either.
async function extractPdfRaw(file, onProgress){
  const arrayBuffer = await file.arrayBuffer();
  const doc = await pdfjsLib.getDocument({data: new Uint8Array(arrayBuffer)}).promise;
  let all = [];
  let headerInfo = null;
  for (let p=1; p<=doc.numPages; p++){
    onProgress(p, doc.numPages);
    const page = await doc.getPage(p);
    const textContent = await page.getTextContent();
    const items = textContent.items
      .map(it => ({str: it.str, x: it.transform[4], y: it.transform[5], size: it.transform[0]}))
      .filter(it => it.str.trim() !== '');
    if (p === 1) headerInfo = extractHeaderInfo(items);
    const res = extractFromPageItems(items, p);
    if (res) all = all.concat(res);
  }
  all = mergeZipperItems(all);
  return { items: stripInternalFields(all), headerInfo };
}

/* ============================================================
   COST BREAK DOWN MERGE ENGINE
   Inserts extracted items into the LVB-PX Cost Break Down Excel
   template (Fabric / Print-Heat Transfer-Embroidery / Accessories),
   preserving all existing formulas, formatting, and pre-filled
   standard accessory rows. Inserts new rows if more space is needed.
   ============================================================ */
const COST_BREAK_DOWN_TEMPLATE_B64 = "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";

const SECTION_CONFIG = {
  Fabric: { headerRow: 11, firstBlank: 12, lastBlank: 14, totalRow: 15, descCol: 2, formulaCols: ['K', 'N', 'O'] },
  Print: { headerRow: 17, firstBlank: 18, lastBlank: 18, totalRow: 19, descCol: 2, formulaCols: ['N', 'O'] },
  Accessories: { headerRow: 21, firstBlank: 22, lastBlank: 32, totalRow: 47, descCol: 2, formulaCols: ['J', 'N', 'O'] },
};

const CELL_REF_RE = /(\$?)([A-Z]{1,3})(\$?)(\d+)/g;

function shiftFormulaRows(formula, insertAt, amount) {
  return formula.replace(CELL_REF_RE, (m, d1, col, d2, row) => {
    row = parseInt(row, 10);
    if (row >= insertAt) row += amount;
    return `${d1}${col}${d2}${row}`;
  });
}

function colLetterToNum(letter) {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
function numToColLetter(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function getFormula(cell) {
  const v = cell.value;
  if (v && typeof v === 'object' && v.formula) return v.formula;
  return null;
}
function setFormula(cell, formula) {
  cell.value = { formula };
}

function shiftWholeSheetFormulas(ws, insertAt, amount, maxRow, maxCol = 20) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      const f = getFormula(cell);
      if (f) setFormula(cell, shiftFormulaRows(f, insertAt, amount));
    }
  }
}

function snapshotAndUnmergeAll(ws) {
  const merges = [];
  const model = ws.model.merges || [];
  for (const rangeStr of [...model]) {
    const [a, b] = rangeStr.split(':');
    const parseCell = (s) => {
      const m = s.match(/^([A-Z]+)(\d+)$/);
      return { col: colLetterToNum(m[1]), row: parseInt(m[2], 10) };
    };
    const pa = parseCell(a), pb = parseCell(b || a);
    merges.push({ minRow: Math.min(pa.row, pb.row), maxRow: Math.max(pa.row, pb.row), minCol: Math.min(pa.col, pb.col), maxCol: Math.max(pa.col, pb.col) });
    ws.unMergeCells(rangeStr);
  }
  return merges;
}

function remergeShifted(ws, merges, insertAt, amount) {
  for (let { minRow, maxRow, minCol, maxCol } of merges) {
    if (minRow >= insertAt) { minRow += amount; maxRow += amount; }
    const rng = `${numToColLetter(minCol)}${minRow}:${numToColLetter(maxCol)}${maxRow}`;
    ws.mergeCells(rng);
  }
}

function copyRowStyle(ws, srcRow, dstRow, maxCol = 20) {
  for (let c = 1; c <= maxCol; c++) {
    const src = ws.getCell(srcRow, c);
    const dst = ws.getCell(dstRow, c);
    dst.style = JSON.parse(JSON.stringify(src.style));
  }
  const srcRowObj = ws.getRow(srcRow);
  ws.getRow(dstRow).height = srcRowObj.height;
}

function copyRowFormulas(ws, srcRow, dstRow, formulaCols) {
  for (const col of formulaCols) {
    const srcCell = ws.getCell(`${col}${srcRow}`);
    const f = getFormula(srcCell);
    if (f) {
      const shifted = f.replace(CELL_REF_RE, (m, d1, c2, d2, row) => {
        row = parseInt(row, 10);
        if (row === srcRow) row = dstRow;
        return `${d1}${c2}${d2}${row}`;
      });
      setFormula(ws.getCell(`${col}${dstRow}`), shifted);
    } else if (srcCell.value !== null && srcCell.value !== undefined) {
      ws.getCell(`${col}${dstRow}`).value = srcCell.value;
    }
  }
}

function getMergeWidth(ws, row, descCol) {
  const model = ws.model.merges || [];
  for (const rangeStr of model) {
    const [a, b] = rangeStr.split(':');
    const m = a.match(/^([A-Z]+)(\d+)$/);
    const col = colLetterToNum(m[1]), r = parseInt(m[2], 10);
    if (r === row && col === descCol) {
      const m2 = b.match(/^([A-Z]+)(\d+)$/);
      return colLetterToNum(m2[1]) - col + 1;
    }
  }
  return 1;
}

function mergeSectionRange(ws, row, descCol, width) {
  const rng = `${numToColLetter(descCol)}${row}:${numToColLetter(descCol + width - 1)}${row}`;
  try { ws.mergeCells(rng); } catch (e) {}
}

function renumberSlNo(ws, firstRow, lastRow) {
  let n = 1;
  for (let r = firstRow; r <= lastRow; r++) ws.getCell(r, 1).value = n++;
}

function fillSection(ws, cfg, items, maxSheetRow) {
  const first = cfg.firstBlank;
  let last = cfg.lastBlank;
  const available = last - first + 1;
  const needed = items.length;
  const amount = Math.max(0, needed - available);
  const descCol = cfg.descCol;

  const oldMerges = snapshotAndUnmergeAll(ws);
  // capture EVERY merge pattern that exists on the section's template row
  // (description column, but also things like the J:K "CON/DZ" merge)
  const rowTemplateMerges = oldMerges
    .filter(m => m.minRow === first)
    .map(m => ({ minCol: m.minCol, maxCol: m.maxCol }));

  let insertAt = null;
  if (amount > 0) {
    insertAt = last;
    const blankRows = Array.from({ length: amount }, () => []);
    ws.spliceRows(insertAt, 0, ...blankRows);
    shiftWholeSheetFormulas(ws, insertAt, amount, maxSheetRow + amount);
    for (let i = 0; i < amount; i++) {
      const newRow = insertAt + i;
      copyRowStyle(ws, insertAt + amount, newRow);
      copyRowFormulas(ws, insertAt + amount, newRow, cfg.formulaCols);
    }
    maxSheetRow += amount;
    last += amount;
    cfg.totalRow += amount;
  }

  const UNIT_COL = 13; // column M
  items.forEach((item, i) => {
    ws.getCell(first + i, descCol).value = item.extractedInfo;
    if (item.unit) ws.getCell(first + i, UNIT_COL).value = item.unit;
    if (cfg.extraCols){
      for (const { col, compute } of cfg.extraCols){
        const val = compute(item);
        if (val !== null && val !== undefined) ws.getCell(first + i, col).value = val;
      }
    }
  });

  renumberSlNo(ws, first, last);

  // re-apply merges: pre-existing ranges shifted (if insertion happened),
  // plus every template merge pattern for each brand-new item row
  if (amount > 0) {
    remergeShifted(ws, oldMerges, insertAt, amount);
    for (let i = 0; i < amount; i++) {
      const newRow = insertAt + i;
      for (const pat of rowTemplateMerges) {
        mergeSectionRange(ws, newRow, pat.minCol, pat.maxCol - pat.minCol + 1);
      }
    }
  } else {
    remergeShifted(ws, oldMerges, 0, 0);
  }

  return { maxSheetRow, totalRow: cfg.totalRow };
}

function normalizeFormulas(ws, maxRow, maxCol = 20) {
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      const v = cell.value;
      if (v && typeof v === 'object' && (v.formula !== undefined || v.sharedFormula !== undefined)) {
        // re-assign as a plain, non-shared formula (resolves shared-formula
        // slave cells too, via the cell.formula getter)
        const f = cell.formula;
        if (f) cell.value = { formula: f };
      }
    }
  }
}


// Browser-adapted version: loads/writes via ArrayBuffer instead of file paths
const FABRIC_PRICE_RULES = [
  { re: /\bdown\s*fill\b/i, price: 1.85 },
  { re: /2\s*\.?\s*oz.*poly\s*\-?\s*fill|2\s*oz.*polyfill/i, price: 0.45 },
  { re: /4\s*\.?\s*oz.*poly\s*\-?\s*fill|4\s*oz.*polyfill/i, price: 0.68 },
  { re: /6\s*\.?\s*oz.*poly\s*\-?\s*fill|6\s*oz.*polyfill/i, price: 0.98 },
  { re: /8\s*\.?\s*oz.*poly\s*\-?\s*fill|8\s*oz.*polyfill/i, price: 1.28 },
  { re: /scrim(mage)?/i, price: 0.10 },
  { re: /interfacing/i, price: 0.26 },
];

// Consumption = Quantity + 5%, e.g. 1 -> 1.05, 2 -> 2.1. Written as an
// actual number (not a string) so it behaves like a number in Excel -
// participates correctly in SUM formulas and any "number format" applied
// to the column, instead of being stored as text.
function computeConsumption(item){
  // Quantity should just be a short number like "1" or "2", but a rare
  // field-boundary overflow (a very long Special Instructions block
  // bleeding into the next field) can tack extra text - and extra digits -
  // onto it. Reading only the leading numeric token (where a real
  // Quantity value always sits) avoids treating stray digits later in the
  // string as part of the number.
  const m = String(item._quantity || '').trim().match(/^(\d+(?:\.\d+)?)/);
  if (!m) return null;
  const q = parseFloat(m[1]);
  if (!isFinite(q)) return null;
  return Math.round(q * 1.05 * 100) / 100;
}

// Known standard prices for common Haddad accessory codes. Price is either
// a plain number or an ExcelJS formula object (no leading "="); every
// formula here is self-contained literal arithmetic with no cell
// references, so it can never produce a formula error regardless of which
// row it ends up on.
const ACCESSORY_PRICE_RULES = {
  'T0140': { price: { formula: '16.71/1000*1.1' }, supplier: 'YKK', remark: 'FREIGHT CHARGE ADDED' },
  'T0664': { price: 0.06, supplier: 'Approximate', remark: '' },
  'T0142': { price: { formula: '2.4/144' }, supplier: 'EMSIG', remark: '' },
  'P0110': { price: { formula: '0.008*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P1029': { price: { formula: '0.028*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P1132': { price: { formula: '0.019*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P1130': { price: { formula: '0.015*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P2610': { price: 0.016, supplier: 'UNIGLORY', remark: '' },
  'P3540': { price: { formula: '0.015*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P3541': { price: { formula: '0.01*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'T0647': { price: { formula: '24.85/1000*1.1' }, supplier: 'YKK', remark: 'FREIGHT CHARGE ADDED' },
  'T0648': { price: { formula: '24.85/1000*1.1' }, supplier: 'YKK', remark: 'FREIGHT CHARGE ADDED' },
  'T0646': { price: { formula: '22.27/1000*1.1' }, supplier: 'YKK', remark: 'FREIGHT CHARGE ADDED' },
  'T0138': { price: { formula: '3.55/1000*1.1' }, supplier: 'YKK', remark: 'FREIGHT CHARGE ADDED' },
  'T0340': { price: 0.10, supplier: 'REGULAR PRICE', remark: '' },
  'P0108': { price: { formula: '0.008*1.1' }, supplier: 'MAXIM', remark: 'FREIGHT CHARGE ADDED' },
  'P1098': { price: 0.016, supplier: 'UNIGLORY', remark: '' },
  'T0330': { price: 0.04, supplier: 'REGULAR PRICE', remark: '' },
};

// A merged zipper row's internalCode is a "+"-joined concatenation (e.g.
// "T0047 + T0356 + T0560A") - check each individual code, not just an
// exact match on the whole string.
function findAccessoryRule(item){
  const codes = String(item.internalCode || '').split('+').map(s => s.trim());
  for (const c of codes) if (ACCESSORY_PRICE_RULES[c]) return ACCESSORY_PRICE_RULES[c];
  return null;
}
function matchAccessoryPrice(item){ const r = findAccessoryRule(item); return r ? r.price : null; }
function matchAccessorySupplier(item){ const r = findAccessoryRule(item); return r ? r.supplier : null; }
function matchAccessoryRemark(item){ const r = findAccessoryRule(item); return (r && r.remark) ? r.remark : null; }

function matchFabricPrice(item){
  if (item._fixedFabricPrice != null) return item._fixedFabricPrice;
  const text = (item.itemName || '') + ' ' + (item.extractedInfo || '');
  for (const rule of FABRIC_PRICE_RULES){
    if (rule.re.test(text)) return rule.price;
  }
  return null;
}

// Excel sheet names: <=31 chars, no \ / ? * [ ] : , can't start/end with '.
function sanitizeSheetName(name){
  let n = (name || 'Sheet').replace(/[\\/?*\[\]:]/g, '-').replace(/^'+|'+$/g, '').trim();
  if (n.length > 31) n = n.slice(0, 31).trim();
  return n || 'Sheet';
}

// Resolves which of the tech pack's known size labels an item's own
// Size-scale text actually names. A plain substring test isn't safe here:
// some size groups share a leading word with an entirely different, more
// specific size group on the very same tech pack - e.g. "TODDLER" and
// "TODDLER CHILD EU" are two distinct size groups (not the same group
// written two ways), but "toddler" is literally a substring of "toddler
// child eu", so a naive .includes() wrongly treats an item scoped only to
// "TODDLER CHILD EU" as also belonging to the plain "TODDLER" tab.
//
// Fix: first find every known label that appears anywhere in the item's
// Size-scale text, then drop any label that is itself fully swallowed by
// a LONGER label that also matched - the longer, more specific label is
// what the field actually names, so the shorter one it contains isn't a
// genuine separate match. Two size groups that are both present (e.g. an
// item genuinely stocked on two stacked lines, "BIG GIRLS ALPHA (XS - XL)"
// and "TODDLER") are unaffected, since neither is a substring of the
// other - both remain.
function resolveItemSizeLabels(item, allSizeLabels){
  const scaleNorm = normSize(item._sizeScale);
  let candidates = allSizeLabels.filter(lbl => scaleNorm.includes(normSize(lbl)));
  candidates = candidates.filter(lbl => {
    const lblNorm = normSize(lbl);
    return !candidates.some(other => {
      const otherNorm = normSize(other);
      return otherNorm !== lblNorm && otherNorm.includes(lblNorm);
    });
  });
  return candidates;
}

// True if this item belongs on the tab for sizeLabel. An item with no
// Size-scale text at all (extraction miss, or a section like Fabric that
// doesn't carry the field) is included on every tab rather than dropped.
// When sizeLabel itself is blank (the single-tab fallback used when page 1
// yielded no size candidates at all), every item applies too. allSizeLabels
// is the full list of size labels found on this tech pack's page 1 - see
// resolveItemSizeLabels for why every label needs to be known at once
// rather than comparing the item against just the one tab's label.
function itemAppliesToSize(item, sizeLabel, allSizeLabels){
  if (!item._sizeScale) return true;
  if (!sizeLabel) return true;
  const matches = resolveItemSizeLabels(item, allSizeLabels || []);
  return matches.some(lbl => normSize(lbl) === normSize(sizeLabel));
}

// Fabric-only fallback for the same "one physical slot, several colorway
// variants" pattern handled elsewhere by _groupKey. Fabric has no reliable
// shared field to key on - Internal Code, Name, and Content can all differ
// between alternates for the same slot - so this instead uses each
// column's colorway coverage (_fabricColorways): if a column's non-"N/A"
// colorways are entirely disjoint from the running group's, it's an
// alternate material for the same slot (each colorway maps to exactly one
// of the candidates) and gets dropped, keeping the first one encountered.
// A column whose coverage overlaps the group (or a column with no
// coverage data at all) starts a fresh group instead.
function dedupeFabricByColorway(items){
  const out = [];
  let groupCoverage = null;
  for (const it of items){
    const cov = new Set((it._fabricColorways || '').split('|').filter(Boolean));
    if (groupCoverage && cov.size && [...cov].every(c => !groupCoverage.has(c))){
      cov.forEach(c => groupCoverage.add(c));
      continue;
    }
    out.push(it);
    groupCoverage = cov.size ? cov : null;
  }
  return out;
}
// A Fabric page's Colorways sub-table can spill onto a second physical PDF
// page that repeats the column title band but not the "Internal Code" row
// - so extractFromPageItems finds no section header there and skips it
// entirely, silently dropping the rest of the colorway data. This detects
// that pattern using the previous Fabric page's own column anchors (the
// label column here is found fresh, as whatever sits well to the left of
// the first anchor) and returns the additional per-column coverage found,
// or null if this page doesn't look like a continuation at all.
function parseFabricColorwayContinuation(items, prevAnchors){
  if (!prevAnchors || !prevAnchors.length) return null;
  const minAnchor = Math.min(...prevAnchors);
  const labelCandidates = items.filter(it => it.x < minAnchor - 15);
  if (labelCandidates.length === 0) return null;
  const labelX = Math.min(...labelCandidates.map(it => it.x));

  const rows = [];
  [...items].sort((a,b)=> b.y - a.y).forEach(it => {
    let row = rows.find(r => Math.abs(r.y - it.y) <= 4);
    if (!row){ row = { y: it.y, items: [] }; rows.push(row); }
    row.items.push(it);
  });

  const numCols = prevAnchors.length;
  const coverage = Array.from({length:numCols}, () => new Set());
  let foundAny = false;
  for (const row of rows){
    const rowName = row.items.filter(it => Math.abs(it.x - labelX) <= 12)
      .sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' ').trim();
    // Colorway rows consistently follow a "CODE - NAME" pattern; requiring
    // the dash avoids picking up a wrapped continuation line (e.g. a lone
    // "YET" from "DAD - ARE WE THERE" wrapping to a second line) as if it
    // were its own row.
    if (!rowName || !rowName.includes('-')) continue;
    const byCol = Array.from({length:numCols}, () => []);
    row.items.forEach(it => {
      if (Math.abs(it.x - labelX) <= 12) return;
      byCol[nearestAnchor(it.x, prevAnchors)].push(it);
    });
    byCol.forEach((toks, ci) => {
      const val = toks.sort((a,b)=>a.x-b.x).map(t=>t.str.trim()).join(' ').trim();
      if (val){ foundAny = true; if (norm(val) !== 'n/a') coverage[ci].add(rowName); }
    });
  }
  return foundAny ? coverage : null;
}

// One physical slot = one line item. Several columns can share a slot when
// they're colorway-variants of the same component (e.g. 4 shank hardware
// colors) - identified by _groupKey (matching Type/Size/Placement, since
// the blue numbered heading itself is drawn as vector graphics, not text,
// and can't be read at all). Keeps the first one encountered per group,
// scoped by section. Items with no group key (Fabric, which has no
// reliable shared field to key on, and already-merged zipper rows) pass
// through untouched.
function dedupeByGroupKey(items){
  const seen = new Set();
  const out = [];
  for (const it of items){
    if (!it._groupKey){ out.push(it); continue; }
    const key = it.section + '|' + it._groupKey;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(it);
  }
  return out;
}

// Full manual worksheet clone (values/formulas, styles, merges, data
// validations, view settings) - ExcelJS doesn't carry any of this over
// automatically. Used to produce one tab per discovered size from the
// single master "Format" layout in the template.
function cloneWorksheetInto(workbook, sourceWs, newName, maxCol = 22){
  const maxRow = sourceWs.rowCount;
  const newWs = workbook.addWorksheet(newName);

  for (let c = 1; c <= maxCol; c++){
    const srcCol = sourceWs.getColumn(c);
    const dstCol = newWs.getColumn(c);
    if (srcCol.width) dstCol.width = srcCol.width;
    if (srcCol.hidden) dstCol.hidden = srcCol.hidden;
  }

  for (let r = 1; r <= maxRow; r++){
    const srcRow = sourceWs.getRow(r);
    const dstRow = newWs.getRow(r);
    if (srcRow.height) dstRow.height = srcRow.height;
    for (let c = 1; c <= maxCol; c++){
      const srcCell = srcRow.getCell(c);
      const dstCell = dstRow.getCell(c);
      const f = getFormula(srcCell);
      if (f) dstCell.value = { formula: f };
      else if (srcCell.value !== null && srcCell.value !== undefined) dstCell.value = srcCell.value;
      if (srcCell.style) dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
      if (srcCell.dataValidation) dstCell.dataValidation = JSON.parse(JSON.stringify(srcCell.dataValidation));
    }
  }

  const merges = sourceWs.model.merges || [];
  merges.forEach(rangeStr => { try { newWs.mergeCells(rangeStr); } catch (e) {} });

  if (sourceWs.views) newWs.views = JSON.parse(JSON.stringify(sourceWs.views));
  if (sourceWs.pageSetup) newWs.pageSetup = JSON.parse(JSON.stringify(sourceWs.pageSetup));

  return newWs;
}

// Signature used to decide whether two Supply Chain Sheet rows are "the
// same item" for dedup purposes: every displayed field except Style No,
// normalized down to just its letters/digits (lowercased, with every
// space, comma, hyphen, semicolon, period, slash, and other punctuation
// stripped out entirely) so two rows that only differ in spacing or
// punctuation - e.g. "T0047, D01" vs "T0047 - D01" vs "T0047D01" - still
// count as the exact same item. Two rows with the same signature are the
// same physical item regardless of which style(s) they were extracted
// from.
function supplyNormalizeForSignature(v){
  return (v || '').toString().toLowerCase().replace(/[^a-z0-9]+/g, '');
}
function supplyRowSignature(...fields){
  return fields.map(supplyNormalizeForSignature).join('\u0001');
}

// Builds the Supply Chain Sheet workbook: a plain two-tab file (no
// template to preserve, unlike the Cost Break Down merge) - Fabric items
// on one tab, everything else (Trim + Labels & Packaging) on the other.
// `rows` is a flat list of raw extracted items, each already tagged with
// its own tech pack's style number under `styleNo`, and already
// zipper-merged per file (see extractPdfRaw).
//
// Layout: Fabric has a blank spacer column between Style No/Fabric Code
// and another between Description/Supplier; Trims has one blank spacer
// column between Description/Supplier. These are pure visual gaps - never
// written to - added on request for manual notes/spacing.
//
// Trims "Item name" is the column HEADING text printed above the item's
// own column in the tech pack (e.g. "CF ZIPPER"), not the item's own
// descriptive Name field (e.g. "#5 OPEN END VISLON"). If a row's heading
// couldn't be read at all (extraction miss), it falls back to the Name
// field rather than leaving the cell blank in a real deliverable file.
//
// Exact-duplicate dedup: an item that appears more than once - whether
// twice within the same style or once each in several different styles -
// is written only once, under whichever style it was encountered under
// first (rows arrive in upload order, so this is simply "first style
// uploaded wins"). Every later exact match, same style or different, is
// dropped entirely rather than written as its own row. No blank separator
// row is written between styles - the sheet is just the deduped row list.
async function buildSupplyChainWorkbook(rows){
  const workbook = new ExcelJS.Workbook();
  const fabricWs = workbook.addWorksheet('Fabric');
  const trimsWs = workbook.addWorksheet('Trims');

  fabricWs.columns = [
    { header: 'Style No', key: 'style', width: 20 },
    { header: '', key: 'blank1', width: 3 },
    { header: 'Fabric Code', key: 'code', width: 18 },
    { header: 'Description', key: 'desc', width: 75 },
    { header: '', key: 'blank2', width: 3 },
    { header: 'Supplier', key: 'supplier', width: 24 },
  ];
  trimsWs.columns = [
    { header: 'Style No', key: 'style', width: 20 },
    { header: 'Item name', key: 'name', width: 28 },
    { header: 'Item code', key: 'code', width: 18 },
    { header: 'Description', key: 'desc', width: 75 },
    { header: '', key: 'blank1', width: 3 },
    { header: 'Supplier', key: 'supplier', width: 24 },
  ];
  [fabricWs, trimsWs].forEach(ws => {
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF4FF' } };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
  });

  const seenFabric = new Set();
  const seenTrims = new Set();
  let fabricWritten = 0, trimsWritten = 0;

  rows.forEach(r => {
    if (r.section === 'Fabric'){
      const key = supplyRowSignature(r.internalCode, r.extractedInfo, r.supplier);
      if (seenFabric.has(key)) return; // exact duplicate of an already-written item - skip
      seenFabric.add(key);
      fabricWs.addRow({ style: r.styleNo || '', code: r.internalCode || '', desc: r.extractedInfo || '', supplier: r.supplier || '' });
      fabricWritten++;
    } else if (r.section === 'Trim' || r.section === 'Labels & packaging'){
      const displayName = (r.columnHeader && r.columnHeader.trim()) || r.itemName || '';
      const key = supplyRowSignature(displayName, r.internalCode, r.extractedInfo, r.supplier);
      if (seenTrims.has(key)) return; // exact duplicate of an already-written item - skip
      seenTrims.add(key);
      trimsWs.addRow({ style: r.styleNo || '', name: displayName, code: r.internalCode || '', desc: r.extractedInfo || '', supplier: r.supplier || '' });
      trimsWritten++;
    }
  });

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, fabricWritten, trimsWritten };
}

async function mergeIntoCostBreakDown(templateArrayBuffer, extractedItems, productImage, headerInfo) {
  const isInterfacing = r => /interfacing/i.test(r.itemName || '');

  const trimAndLabels = [
    ...extractedItems.filter(r => r.section === 'Trim'),
    ...extractedItems.filter(r => r.section === 'Labels & packaging'),
  ];
  // Fabric (and interfacing, wherever it's routed from) stays identical on
  // every tab regardless of its own Size scale value - only the remaining
  // Trim/Labels accessory items get filtered per tab. Deduped once here
  // (rather than per tab) since the same universal list is reused on every
  // tab regardless of size. dedupeByGroupKey handles interfacing rows
  // (which carry Trim's dedupeKeyFields); dedupeFabricByColorway handles
  // true Fabric rows via colorway coverage instead, since Fabric has no
  // dedupeKeyFields of its own (see SECTIONS and dedupeFabricByColorway).
  const fabricItems = dedupeFabricByColorway(dedupeByGroupKey([
    ...extractedItems.filter(r => r.section === 'Fabric'),
    ...trimAndLabels.filter(isInterfacing),
  ]));
  // Always-present Fabric line, fixed on every tab regardless of what was
  // extracted - price is a hardcoded constant (see matchFabricPrice),
  // not matched against FABRIC_PRICE_RULES like extracted items are.
  fabricItems.push({
    section: 'Fabric', page: 0, itemName: 'INTERLINING, 100% POLYESTER', internalCode: '',
    supplier: '', unit: 'YDS', extractedInfo: 'INTERLINING, 100% POLYESTER',
    _sizeScale: '', _quantity: '', _groupKey: '', _fabricColorways: '',
    _fixedFabricPrice: 0.26,
  });
  const printItems = dedupeByGroupKey(extractedItems.filter(r => r.section.toLowerCase().startsWith('print')));
  const accessoryItemsAll = trimAndLabels.filter(r => !isInterfacing(r));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateArrayBuffer);
  const masterWs = workbook.getWorksheet('Format');

  // Fall back to a single unlabeled tab (today's old single-sheet
  // behavior) if page 1 didn't yield any size candidates at all.
  const sizes = (headerInfo && headerInfo.sizes && headerInfo.sizes.length)
    ? headerInfo.sizes
    : [{ label: '', token: '' }];
  // Every real size label found on page 1, passed to itemAppliesToSize so
  // it can tell apart size groups that share a leading word (e.g.
  // "TODDLER" vs "TODDLER CHILD EU") - see resolveItemSizeLabels.
  const allSizeLabels = sizes.map(s => s.label).filter(Boolean);

  // The product image buffer only needs to be registered with the
  // workbook once; the same image id can be placed on every tab.
  const imgId = (productImage && productImage.buffer)
    ? workbook.addImage({ buffer: productImage.buffer, extension: 'png' })
    : null;

  const usedNames = new Set();
  const counts = [];

  for (let s = 0; s < sizes.length; s++){
    const { label: sizeLabel, token } = sizes[s];
    const rawName = headerInfo && headerInfo.style
      ? `${headerInfo.style} (${token || sizeLabel || (s + 1)})`
      : (token || sizeLabel || `Sheet${s + 1}`);
    let tabName = sanitizeSheetName(rawName);
    let n = 1;
    while (usedNames.has(tabName)) tabName = sanitizeSheetName(`${rawName} ${++n}`);
    usedNames.add(tabName);

    // Every size gets a full manual clone of the pristine template layout
    // - never reuse/mutate masterWs directly for any size, even the
    // first. Writing into masterWs and then cloning it for later sizes
    // used to carry the first size's already-filled rows into every
    // subsequent clone, which the "preserve already-filled rows" logic in
    // fillSection would then mistake for pre-existing static content and
    // append past - doubling Accessories/Labels items on every tab after
    // the first, and throwing off the row math badly enough to corrupt
    // the Total Accessories Cost formula.
    const ws = cloneWorksheetInto(workbook, masterWs, tabName);

    let maxRow = ws.rowCount;
    normalizeFormulas(ws, maxRow, 22);

    if (headerInfo){
      if (headerInfo.buyer) ws.getCell('B5').value = headerInfo.buyer;
      if (headerInfo.style) ws.getCell('B6').value = headerInfo.style;
      if (sizeLabel) ws.getCell('B7').value = sizeLabel;
      if (headerInfo.item) ws.getCell('B8').value = headerInfo.item;
    }

    // Test cost, fixed at 0.6 on every tab - the value cell (N50) to the
    // right of the merged "FAB PERFORMANCE+ FAB RSL+ ACC RSL+Garment
    // PERFORMANCE + FABRIC CPSIA" label (A50:M50). This also replaces the
    // template's original N50 formula (=3000/L6*12), which was one of the
    // pre-existing #DIV/0! cells flagged earlier - L6 is a blank input
    // cell, so that formula always errored until now.
    ws.getCell('N50').value = 0.6;

    const accessoryItems = dedupeByGroupKey(accessoryItemsAll.filter(r => itemAppliesToSize(r, sizeLabel, allSizeLabels)));

    const cfgAcc = { ...SECTION_CONFIG.Accessories, extraCols: [
      { col: 9, compute: computeConsumption },
      { col: 12, compute: matchAccessoryPrice },
      { col: 16, compute: matchAccessorySupplier },
      { col: 17, compute: matchAccessoryRemark },
    ] };
    ({ maxSheetRow: maxRow } = fillSection(ws, cfgAcc, accessoryItems, maxRow));

    const cfgPrint = { ...SECTION_CONFIG.Print };
    ({ maxSheetRow: maxRow } = fillSection(ws, cfgPrint, printItems, maxRow));

    const cfgFab = { ...SECTION_CONFIG.Fabric, extraCols: [{ col: 12, compute: matchFabricPrice }] };
    ({ maxSheetRow: maxRow } = fillSection(ws, cfgFab, fabricItems, maxRow));

    if (imgId !== null) {
      const displayWidth = 320;
      const displayHeight = Math.round(displayWidth * (productImage.height / productImage.width));
      ws.addImage(imgId, {
        tl: { col: 17, row: 0 },
        ext: { width: displayWidth, height: displayHeight },
      });
    }

    counts.push({ tab: tabName, fabric: fabricItems.length, print: printItems.length, accessories: accessoryItems.length });
  }

  // masterWs was only ever a cloning source, never one of the output
  // tabs - drop it so the file doesn't ship an extra untagged "Format"
  // sheet.
  workbook.removeWorksheet(masterWs.id);

  const buffer = await workbook.xlsx.writeBuffer();
  return { buffer, counts };
}



/* ---- Haddad brand navigation ---- */
const homeView = document.getElementById('homeView');
const haddadHubView = document.getElementById('haddadHubView');
const haddadView = document.getElementById('haddadView');
const haddadSupplyView = document.getElementById('haddadSupplyView');

document.getElementById('brandHaddad').addEventListener('click', ()=>{
  requestUnlock('haddad', () => {
    homeView.hidden = true;
    haddadHubView.hidden = false;
  });
});
document.getElementById('hubBackBtn').addEventListener('click', ()=>{
  haddadHubView.hidden = true;
  homeView.hidden = false;
});
document.getElementById('hubBomCard').addEventListener('click', ()=>{
  haddadHubView.hidden = true;
  haddadView.hidden = false;
});
document.getElementById('hubSupplyCard').addEventListener('click', ()=>{
  haddadHubView.hidden = true;
  haddadSupplyView.hidden = false;
});
document.getElementById('backBtn').addEventListener('click', ()=>{
  haddadView.hidden = true;
  haddadHubView.hidden = false;
});
document.getElementById('backBtnSupply').addEventListener('click', ()=>{
  haddadSupplyView.hidden = true;
  haddadHubView.hidden = false;
});
// Decathlon card is a coming-soon placeholder — intentionally not wired up yet.

/* ============================================================
   BUG REPORT FORM
   Submits directly to palashonwork@gmail.com via FormSubmit.co —
   a formless email-relay service, no backend/API key required.
   ============================================================ */
const bugToggle = document.getElementById('bugToggle');
const bugForm = document.getElementById('bugForm');
const bugSendBtn = document.getElementById('bugSendBtn');
const bugSendLabel = document.getElementById('bugSendLabel');
const bugStatus = document.getElementById('bugStatus');

bugToggle.addEventListener('click', ()=>{
  bugForm.hidden = !bugForm.hidden;
  if (!bugForm.hidden) document.getElementById('bugName').focus();
});

bugForm.addEventListener('submit', async (e)=>{
  e.preventDefault();
  bugSendBtn.disabled = true;
  bugSendBtn.classList.add('loading');
  bugSendLabel.textContent = 'Sending...';
  bugStatus.textContent = '';
  bugStatus.className = 'status';
  try{
    const formData = new FormData(bugForm);
    formData.append('_subject', 'Tech Pack Utility — Bug Report');
    formData.append('_captcha', 'false');
    const res = await fetch('https://formsubmit.co/ajax/palashonwork@gmail.com', {
      method: 'POST',
      headers: { 'Accept': 'application/json' },
      body: formData,
    });
    if (!res.ok) throw new Error('Request failed (' + res.status + ')');
    bugStatus.textContent = 'Thanks — your report was sent.';
    bugStatus.className = 'status ok';
    bugForm.reset();
    setTimeout(()=>{ bugForm.hidden = true; bugStatus.textContent = ''; }, 2500);
  } catch(err){
    console.error(err);
    bugStatus.textContent = 'Could not send the report. Please try again or email palashonwork@gmail.com directly.';
    bugStatus.className = 'status err';
  } finally {
    bugSendBtn.disabled = false;
    bugSendBtn.classList.remove('loading');
    bugSendLabel.textContent = 'Send Report';
  }
});

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const filebar = document.getElementById('filebar');
const fileNameEl = document.getElementById('fileName');
const clearFile = document.getElementById('clearFile');
const processBtn = document.getElementById('processBtn');
const processLabel = document.getElementById('processLabel');
const statusEl = document.getElementById('status');
const resultsEl = document.getElementById('results');
const resultsBody = document.getElementById('resultsBody');
const countsEl = document.getElementById('counts');
const downloadBtn = document.getElementById('downloadBtn');

let currentFile = null;
let currentResults = null;
let currentProductImage = null;
let currentHeaderInfo = null;

function setFile(file){
  if (!file || file.type !== 'application/pdf'){
    setStatus('Please choose a PDF file.', 'err');
    return;
  }
  currentFile = file;
  fileNameEl.textContent = file.name;
  filebar.classList.add('show');
  processBtn.disabled = false;
  setStatus('');
  resultsEl.classList.remove('show');
}

dropzone.addEventListener('dragover', e=>{ e.preventDefault(); dropzone.classList.add('drag'); });
dropzone.addEventListener('dragleave', ()=> dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e=>{
  e.preventDefault(); dropzone.classList.remove('drag');
  if (e.dataTransfer.files.length) setFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e=>{
  if (e.target.files.length) setFile(e.target.files[0]);
});
clearFile.addEventListener('click', e=>{
  e.stopPropagation();
  currentFile = null; fileInput.value = '';
  currentProductImage = null;
  currentHeaderInfo = null;
  filebar.classList.remove('show');
  processBtn.disabled = true;
  resultsEl.classList.remove('show');
  setStatus('');
});

function setStatus(msg, cls){
  statusEl.textContent = msg;
  statusEl.className = 'status' + (cls ? ' '+cls : '');
}

const sectionClass = s => s === 'Fabric' ? 'fabric' : s === 'Trim' ? 'trim' : 'labels';

processBtn.addEventListener('click', async ()=>{
  if (!currentFile) return;
  processBtn.disabled = true;
  processBtn.classList.add('loading');
  resultsEl.classList.remove('show');
  try{
    const { items: results, productImage, headerInfo } = await extractPdf(currentFile, (p,total)=>{
      processLabel.textContent = `Scanning page ${p} / ${total}...`;
      setStatus(`Scanning page ${p} of ${total}...`);
    });
    currentResults = results;
    currentProductImage = productImage;
    currentHeaderInfo = headerInfo;
    if (results.length === 0){
      setStatus('No matching Fabric / Trim / Labels & Packaging grids were found in this PDF.', 'err');
    } else {
      resultsEl.classList.add('show');
      setStatus(`Done — ${results.length} items extracted. Use the buttons below to download.`, 'ok');
    }
  } catch(err){
    console.error(err);
    setStatus('Something went wrong reading this PDF: ' + err.message, 'err');
  } finally {
    processBtn.disabled = false;
    processBtn.classList.remove('loading');
    processLabel.textContent = 'Extract to Excel';
  }
});

const SWATCH = { fabric:'#2563EB', trim:'#D97706', labels:'#DC2626' };
const SECTION_ORDER = ['Fabric', 'Trim', 'Labels & packaging'];

function renderResults(results){
  resultsBody.innerHTML = '';
  const bySection = {};
  results.forEach(r=>{ (bySection[r.section] = bySection[r.section] || []).push(r); });

  const orderedSections = [
    ...SECTION_ORDER.filter(s => bySection[s]),
    ...Object.keys(bySection).filter(s => !SECTION_ORDER.includes(s)),
  ];

  orderedSections.forEach((section, cardIdx) => {
    const items = bySection[section];
    const cls = sectionClass(section);
    const card = document.createElement('div');
    card.className = 'result-card';
    card.style.animationDelay = (cardIdx * 0.08) + 's';
    card.innerHTML = `
      <div class="result-card-head">
        <div class="title"><span class="swatch" style="background:${SWATCH[cls]}"></span>${escapeHtml(section)}</div>
        <span class="n">${items.length}</span>
      </div>
      <div class="result-items"></div>`;
    const itemsWrap = card.querySelector('.result-items');
    items.forEach(r => {
      const row = document.createElement('div');
      row.className = 'result-item';
      row.innerHTML = `
        <div class="result-item-top">
          <span class="code-chip">${escapeHtml(r.internalCode || '—')}</span>
          <span class="supplier-tag">${escapeHtml(r.supplier || '')}</span>
        </div>
        <div class="result-item-info">${escapeHtml(r.extractedInfo)}</div>`;
      itemsWrap.appendChild(row);
    });
    resultsBody.appendChild(card);
  });

  countsEl.innerHTML = Object.entries(bySection).map(([k,v])=>
    `<span class="count-chip"><span class="count-badge">${v.length}</span>${escapeHtml(k)}</span>`
  ).join('') + `<span class="count-chip total"><span class="count-badge">${results.length}</span>total</span>`;
  resultsEl.classList.add('show');
}

function escapeHtml(s){
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

downloadBtn.addEventListener('click', ()=>{
  if (!currentResults) return;
  const wsData = [['Section','Internal Code','Extracted Info','Supplier']];
  currentResults.forEach(r=> wsData.push([r.section, r.internalCode || '', r.extractedInfo, r.supplier || '']));
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  ws['!cols'] = [{wch:20},{wch:16},{wch:90},{wch:20}];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Extracted Data');
  const baseName = (currentFile.name || 'tech_pack').replace(/\.pdf$/i,'');
  XLSX.writeFile(wb, `${baseName}_extracted.xlsx`);
});

function base64ToArrayBuffer(base64){
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i=0;i<binary.length;i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

const cbdBtn = document.getElementById('cbdBtn');
const cbdLabel = document.getElementById('cbdLabel');
const cbdStatus = document.getElementById('cbdStatus');

cbdBtn.addEventListener('click', async ()=>{
  if (!currentResults) return;
  cbdBtn.disabled = true;
  cbdBtn.classList.add('loading');
  cbdLabel.textContent = 'Building Cost Break Down file...';
  cbdStatus.textContent = '';
  cbdStatus.className = 'status';
  try {
    const templateBuffer = base64ToArrayBuffer(COST_BREAK_DOWN_TEMPLATE_B64);
    const { buffer, counts } = await mergeIntoCostBreakDown(templateBuffer, currentResults, currentProductImage, currentHeaderInfo);
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const baseName = (currentFile.name || 'tech_pack').replace(/\.pdf$/i,'');
    const a = document.createElement('a');
    a.href = url;
    a.download = `${baseName}_cost_break_down.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    const tabSummary = counts.map(c =>
      `${c.tab} — Fabric: ${c.fabric}, Print/Heat Transfer/Embroidery: ${c.print}, Accessories: ${c.accessories}`
    ).join(' | ');
    cbdStatus.textContent = `Done — ${counts.length} size tab${counts.length === 1 ? '' : 's'} generated. ${tabSummary}.`;
    cbdStatus.className = 'status ok';
  } catch(err){
    console.error(err);
    cbdStatus.textContent = 'Something went wrong building the Cost Break Down file: ' + err.message;
    cbdStatus.className = 'status err';
  } finally {
    cbdBtn.disabled = false;
    cbdBtn.classList.remove('loading');
    cbdLabel.textContent = 'Generate Cost Break Down Excel';
  }
});

/* ============================================================
   SUPPLY CHAIN SHEET — multi tech pack, raw Fabric/Trims export
   ============================================================ */
const dropzoneSupply = document.getElementById('dropzoneSupply');
const fileInputSupply = document.getElementById('fileInputSupply');
const fileListSupply = document.getElementById('fileListSupply');
const processBtnSupply = document.getElementById('processBtnSupply');
const processLabelSupply = document.getElementById('processLabelSupply');
const statusSupply = document.getElementById('statusSupply');
const resultsSupply = document.getElementById('resultsSupply');
const countsSupply = document.getElementById('countsSupply');
const downloadBtnSupply = document.getElementById('downloadBtnSupply');

let currentSupplyFiles = [];
let currentSupplyBuffer = null;

function setStatusEl(el, msg, cls){
  el.textContent = msg;
  el.className = 'status' + (cls ? ' '+cls : '');
}

function renderSupplyFileList(){
  fileListSupply.innerHTML = '';
  currentSupplyFiles.forEach((file, idx) => {
    const row = document.createElement('div');
    row.className = 'filebar show';
    row.style.marginBottom = '8px';
    row.innerHTML = `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6"/></svg>
      <span class="name">${escapeHtml(file.name)}</span>
      <span class="clear">✕ remove</span>`;
    row.querySelector('.clear').addEventListener('click', (e)=>{
      e.stopPropagation();
      currentSupplyFiles.splice(idx, 1);
      renderSupplyFileList();
    });
    fileListSupply.appendChild(row);
  });
  processBtnSupply.disabled = currentSupplyFiles.length === 0;
  resultsSupply.classList.remove('show');
  currentSupplyBuffer = null;
}

function addSupplyFiles(fileList){
  const pdfs = Array.from(fileList).filter(f => f.type === 'application/pdf');
  if (pdfs.length === 0){
    setStatusEl(statusSupply, 'Please choose PDF file(s).', 'err');
    return;
  }
  currentSupplyFiles = currentSupplyFiles.concat(pdfs);
  setStatusEl(statusSupply, '');
  renderSupplyFileList();
}

dropzoneSupply.addEventListener('dragover', e=>{ e.preventDefault(); dropzoneSupply.classList.add('drag'); });
dropzoneSupply.addEventListener('dragleave', ()=> dropzoneSupply.classList.remove('drag'));
dropzoneSupply.addEventListener('drop', e=>{
  e.preventDefault(); dropzoneSupply.classList.remove('drag');
  if (e.dataTransfer.files.length) addSupplyFiles(e.dataTransfer.files);
});
fileInputSupply.addEventListener('change', e=>{
  if (e.target.files.length) addSupplyFiles(e.target.files);
  fileInputSupply.value = ''; // allows re-adding a removed file
});

processBtnSupply.addEventListener('click', async ()=>{
  if (currentSupplyFiles.length === 0) return;
  processBtnSupply.disabled = true;
  processBtnSupply.classList.add('loading');
  resultsSupply.classList.remove('show');
  try {
    let allRows = [];
    for (let i = 0; i < currentSupplyFiles.length; i++){
      const file = currentSupplyFiles[i];
      processLabelSupply.textContent = `Extracting ${file.name} (${i+1}/${currentSupplyFiles.length})...`;
      const { items, headerInfo } = await extractPdfRaw(file, (p, total)=>{
        setStatusEl(statusSupply, `${file.name}: scanning page ${p} of ${total}...`);
      });
      const styleNo = (headerInfo && headerInfo.style) ? headerInfo.style : file.name.replace(/\.pdf$/i, '');
      items.forEach(it => { it.styleNo = styleNo; });
      allRows = allRows.concat(items);
    }
    if (allRows.length === 0){
      setStatusEl(statusSupply, 'No matching Fabric / Trim / Labels & Packaging grids were found in these PDFs.', 'err');
    } else {
      const { buffer, fabricWritten, trimsWritten } = await buildSupplyChainWorkbook(allRows);
      currentSupplyBuffer = buffer;
      const writtenTotal = fabricWritten + trimsWritten;
      const duplicatesSkipped = allRows.length - writtenTotal;
      countsSupply.textContent = `Fabric: ${fabricWritten} | Trims: ${trimsWritten} | Total: ${writtenTotal} unique row(s) across ${currentSupplyFiles.length} tech pack(s)` +
        (duplicatesSkipped ? ` (${duplicatesSkipped} duplicate${duplicatesSkipped === 1 ? '' : 's'} skipped).` : '.');
      resultsSupply.classList.add('show');
      setStatusEl(statusSupply, `Done — ${writtenTotal} unique item(s) written from ${currentSupplyFiles.length} tech pack(s)` +
        (duplicatesSkipped ? `, ${duplicatesSkipped} duplicate${duplicatesSkipped === 1 ? '' : 's'} skipped.` : '.'), 'ok');
    }
  } catch(err){
    console.error(err);
    setStatusEl(statusSupply, 'Something went wrong reading these PDFs: ' + err.message, 'err');
  } finally {
    processBtnSupply.disabled = false;
    processBtnSupply.classList.remove('loading');
    processLabelSupply.textContent = 'Extract & Generate Supply Chain Sheet';
  }
});

downloadBtnSupply.addEventListener('click', ()=>{
  if (!currentSupplyBuffer) return;
  const blob = new Blob([currentSupplyBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'Haddad_Supply_Chain_Sheet.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
});
