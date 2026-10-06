/* ============================================================
   DECATHLON — CLOSED BOOK ENGINE
   Takes an already-generated (multi-tab) Decathlon costing workbook and
   re-prices it from fresh Fabric / Accessories price lists, using exactly
   the same pricing rules as the Open Book engine (decathlon.js), then
   produces a separate change report (before vs after).

   Depends on: decathlon.js (pricing engine + computeDecathlonPricing,
   decathlonFlatRateFor, decathlonIsInterlining and the price-list
   parse/build helpers), haddad.js (normalizeFormulas), shared.js.

   Nothing here modifies the Open Book logic.
   ============================================================ */

/* ============================================================
   1. MINI FORMULA ENGINE
   The cost sheet's totals are formulas, and a browser has no Excel to
   recalculate them. To report "before vs after" totals for ANY uploaded
   sheet (not only ones with cached results), this evaluates the formulas
   used by the Decathlon template: + - * / ^ & % comparisons, ROUND, SUM,
   IF, IFERROR, AND, OR, NOT, LEFT, RIGHT, VLOOKUP (exact), MAX, MIN, ABS,
   ISBLANK, ISNUMBER, ISERROR, cross-sheet refs, whole-column ranges.
   Anything else throws DecUnsupported and the caller falls back safely.
   ============================================================ */
class DecXlError { constructor(code) { this.code = code; } }
class DecUnsupported extends Error {}

function dcbColNum(letters) {
  let n = 0;
  for (const ch of letters.replace(/\$/g, '')) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

const DCB_REF_RE = /^(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?(?:(\$?[A-Z]{1,3})\$?(\d+)(?::(\$?[A-Z]{1,3})\$?(\d+))?|(\$?[A-Z]{1,3}):(\$?[A-Z]{1,3}))/;

function dcbTokenize(f) {
  const toks = [];
  let i = 0;
  const n = f.length;
  while (i < n) {
    const ch = f[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (ch === '"') {
      let j = i + 1, s = '';
      while (j < n) {
        if (f[j] === '"') { if (f[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
        s += f[j++];
      }
      toks.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    const two = f.substr(i, 2);
    if (two === '<>' || two === '<=' || two === '>=') { toks.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%,()'.indexOf(ch) !== -1) { toks.push({ t: 'op', v: ch }); i++; continue; }
    const rest = f.slice(i);
    if (/[0-9.]/.test(ch)) {
      const m = rest.match(/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/);
      toks.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    const m = rest.match(DCB_REF_RE);
    if (m && !/[A-Za-z0-9_(.]/.test(rest[m[0].length] || '')) {
      const sheet = m[1] ? m[1].replace(/''/g, "'") : (m[2] || null);
      if (m[3]) {
        const c1 = dcbColNum(m[3]), r1 = parseInt(m[4], 10);
        if (m[5]) toks.push({ t: 'range', sheet, c1, r1, c2: dcbColNum(m[5]), r2: parseInt(m[6], 10) });
        else toks.push({ t: 'cell', sheet, c: c1, r: r1 });
      } else {
        toks.push({ t: 'range', sheet, c1: dcbColNum(m[7]), r1: 1, c2: dcbColNum(m[8]), r2: Infinity });
      }
      i += m[0].length;
      continue;
    }
    const idm = rest.match(/^[A-Za-z_][A-Za-z0-9_.]*/);
    if (idm) {
      const name = idm[0];
      if (rest[name.length] === '(') { toks.push({ t: 'fn', name: name.toUpperCase() }); i += name.length; continue; }
      if (/^(TRUE|FALSE)$/i.test(name)) { toks.push({ t: 'bool', v: name.toUpperCase() === 'TRUE' }); i += name.length; continue; }
      throw new DecUnsupported('name ' + name);
    }
    throw new DecUnsupported('char ' + ch);
  }
  return toks;
}

function dcbParse(f) {
  const toks = dcbTokenize(f);
  let p = 0;
  const isOp = v => toks[p] && toks[p].t === 'op' && toks[p].v === v;
  const CMP = ['=', '<>', '<', '>', '<=', '>='];
  function parseCmp() {
    let l = parseConcat();
    while (toks[p] && toks[p].t === 'op' && CMP.indexOf(toks[p].v) !== -1) { const op = toks[p++].v; l = { t: 'bin', op, a: l, b: parseConcat() }; }
    return l;
  }
  function parseConcat() { let l = parseAdd(); while (isOp('&')) { p++; l = { t: 'bin', op: '&', a: l, b: parseAdd() }; } return l; }
  function parseAdd() { let l = parseMul(); while (isOp('+') || isOp('-')) { const op = toks[p++].v; l = { t: 'bin', op, a: l, b: parseMul() }; } return l; }
  function parseMul() { let l = parsePow(); while (isOp('*') || isOp('/')) { const op = toks[p++].v; l = { t: 'bin', op, a: l, b: parsePow() }; } return l; }
  function parsePow() { let l = parseUnary(); while (isOp('^')) { p++; l = { t: 'bin', op: '^', a: l, b: parseUnary() }; } return l; }
  function parseUnary() {
    if (isOp('-')) { p++; return { t: 'un', x: parseUnary() }; }
    if (isOp('+')) { p++; return parseUnary(); }
    return parsePostfix();
  }
  function parsePostfix() { let x = parsePrimary(); while (isOp('%')) { p++; x = { t: 'pct', x }; } return x; }
  function parsePrimary() {
    const tk = toks[p++];
    if (!tk) throw new DecUnsupported('unexpected end');
    if (tk.t === 'num') return { t: 'num', v: tk.v };
    if (tk.t === 'str') return { t: 'str', v: tk.v };
    if (tk.t === 'bool') return { t: 'bool', v: tk.v };
    if (tk.t === 'cell' || tk.t === 'range') return tk;
    if (tk.t === 'op' && tk.v === '(') {
      const e = parseCmp();
      if (!isOp(')')) throw new DecUnsupported('missing )');
      p++;
      return e;
    }
    if (tk.t === 'fn') {
      if (!isOp('(')) throw new DecUnsupported('missing ( after ' + tk.name);
      p++;
      const args = [];
      if (isOp(')')) { p++; return { t: 'fn', name: tk.name, args }; }
      for (;;) {
        if (isOp(',') || isOp(')')) args.push({ t: 'blank' }); else args.push(parseCmp());
        if (isOp(',')) { p++; continue; }
        if (isOp(')')) { p++; break; }
        throw new DecUnsupported('bad argument list');
      }
      return { t: 'fn', name: tk.name, args };
    }
    throw new DecUnsupported('unexpected token');
  }
  const ast = parseCmp();
  if (p < toks.length) throw new DecUnsupported('trailing tokens');
  return ast;
}

class DecFormulaEngine {
  constructor(workbook) {
    this.wb = workbook;
    this.cache = new Map();
    this.stack = new Set();
    this.parsed = new Map();
  }

  get(sheetName, r, c) {
    const key = sheetName + '\u0001' + r + '\u0001' + c;
    if (this.cache.has(key)) return this.cache.get(key);
    if (this.stack.has(key)) throw new DecUnsupported('circular reference');
    const ws = this.wb.getWorksheet(sheetName);
    if (!ws) throw new DecUnsupported('missing sheet ' + sheetName);
    const cell = ws.findCell(r, c);
    let val = null;
    if (cell) {
      this.stack.add(key);
      try { val = this._cellValue(sheetName, cell); } finally { this.stack.delete(key); }
    }
    this.cache.set(key, val);
    return val;
  }

  _cellValue(sheetName, cell) {
    if (cell.type === ExcelJS.ValueType.Formula) {
      const f = cell.formula;
      if (!f) return this._norm(cell.result);
      let ast = this.parsed.get(f);
      if (!ast) { ast = dcbParse(f); this.parsed.set(f, ast); }
      const v = this.ev(ast, sheetName);
      if (v && v.range) throw new DecUnsupported('range result');
      return v;
    }
    return this._norm(cell.value);
  }

  _norm(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return (v.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
    if (typeof v === 'object') {
      if (v.error) return new DecXlError(v.error);
      if (v.richText) return v.richText.map(t => t.text).join('');
      if (v.text !== undefined) return String(v.text);
      if ('result' in v) return this._norm(v.result);
      return null;
    }
    return v;
  }

  scalar(v) { if (v && v.range) throw new DecUnsupported('range used as scalar'); return v; }

  num(v) {
    if (v instanceof DecXlError) return v;
    if (v === null) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const s = String(v).trim();
    if (s === '' || isNaN(Number(s))) return new DecXlError('#VALUE!');
    return Number(s);
  }

  str(v) {
    if (v === null) return '';
    if (typeof v === 'number') return String(Number(v.toPrecision(15)));
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  bool(v) {
    if (v instanceof DecXlError) return v;
    if (v === null) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    const u = String(v).toUpperCase();
    if (u === 'TRUE') return true;
    if (u === 'FALSE') return false;
    return new DecXlError('#VALUE!');
  }

  rank(v) { return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : 3; }

  compare(a, b) {
    if (a === null && b === null) { a = 0; b = 0; }
    else if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    else if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = this.rank(a), rb = this.rank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (typeof a === 'string') { a = a.toLowerCase(); b = b.toLowerCase(); }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  round(x, d) {
    const f = Math.pow(10, d);
    const sign = x < 0 ? -1 : 1;
    return sign * Math.round(Number((Math.abs(x) * f).toPrecision(15))) / f;
  }

  rangeBounds(rg) {
    const ws = this.wb.getWorksheet(rg.sheet);
    if (!ws) throw new DecUnsupported('missing sheet ' + rg.sheet);
    return { r1: rg.r1, r2: rg.r2 === Infinity ? ws.rowCount : rg.r2, c1: rg.c1, c2: rg.c2 };
  }

  // Collects numeric-relevant values from args for SUM/MAX/MIN: ranges skip
  // text/bools/blank, scalars are coerced. Returns array or DecXlError.
  collectNumbers(args, ctx) {
    const out = [];
    for (const a of args) {
      const v = this.ev(a, ctx);
      if (v && v.range) {
        const b = this.rangeBounds(v);
        for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) {
          const x = this.get(v.sheet, r, c);
          if (x instanceof DecXlError) return x;
          if (typeof x === 'number') out.push(x);
        }
      } else {
        if (v instanceof DecXlError) return v;
        if (v === null && a.t === 'blank') continue;
        const n = this.num(v);
        if (n instanceof DecXlError) return n;
        out.push(n);
      }
    }
    return out;
  }

  collectBools(args, ctx) {
    const out = [];
    for (const a of args) {
      const v = this.ev(a, ctx);
      if (v && v.range) {
        const b = this.rangeBounds(v);
        for (let r = b.r1; r <= b.r2; r++) for (let c = b.c1; c <= b.c2; c++) {
          const x = this.get(v.sheet, r, c);
          if (x instanceof DecXlError) return x;
          if (typeof x === 'boolean') out.push(x); else if (typeof x === 'number') out.push(x !== 0);
        }
      } else {
        const b = this.bool(v);
        if (b instanceof DecXlError) return b;
        out.push(b);
      }
    }
    return out;
  }

  ev(node, ctx) {
    switch (node.t) {
      case 'num': case 'str': case 'bool': return node.v;
      case 'blank': return null;
      case 'cell': return this.get(node.sheet || ctx, node.r, node.c);
      case 'range': return { range: true, sheet: node.sheet || ctx, r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2 };
      case 'un': {
        const x = this.num(this.scalar(this.ev(node.x, ctx)));
        return x instanceof DecXlError ? x : -x;
      }
      case 'pct': {
        const x = this.num(this.scalar(this.ev(node.x, ctx)));
        return x instanceof DecXlError ? x : x / 100;
      }
      case 'bin': return this.evBin(node, ctx);
      case 'fn': return this.evFn(node, ctx);
    }
    throw new DecUnsupported('node ' + node.t);
  }

  evBin(node, ctx) {
    const a = this.scalar(this.ev(node.a, ctx));
    const b = this.scalar(this.ev(node.b, ctx));
    if (a instanceof DecXlError) return a;
    if (b instanceof DecXlError) return b;
    const op = node.op;
    if (op === '&') return this.str(a) + this.str(b);
    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      const c = this.compare(a, b);
      return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
    }
    const x = this.num(a), y = this.num(b);
    if (x instanceof DecXlError) return x;
    if (y instanceof DecXlError) return y;
    if (op === '+') return x + y;
    if (op === '-') return x - y;
    if (op === '*') return x * y;
    if (op === '/') return y === 0 ? new DecXlError('#DIV/0!') : x / y;
    if (op === '^') return Math.pow(x, y);
    throw new DecUnsupported('operator ' + op);
  }

  evFn(node, ctx) {
    const A = node.args;
    switch (node.name) {
      case 'IF': {
        const c = this.bool(this.scalar(this.ev(A[0], ctx)));
        if (c instanceof DecXlError) return c;
        if (c) return A[1] ? this.scalar(this.ev(A[1], ctx)) : true;
        return A[2] ? this.scalar(this.ev(A[2], ctx)) : false;
      }
      case 'IFERROR': {
        const v = this.scalar(this.ev(A[0], ctx));
        return v instanceof DecXlError ? this.scalar(this.ev(A[1], ctx)) : v;
      }
      case 'ROUND': {
        const x = this.num(this.scalar(this.ev(A[0], ctx)));
        const d = this.num(this.scalar(this.ev(A[1], ctx)));
        if (x instanceof DecXlError) return x;
        if (d instanceof DecXlError) return d;
        return this.round(x, Math.trunc(d));
      }
      case 'SUM': { const v = this.collectNumbers(A, ctx); return v instanceof DecXlError ? v : v.reduce((s, x) => s + x, 0); }
      case 'MAX': { const v = this.collectNumbers(A, ctx); return v instanceof DecXlError ? v : (v.length ? Math.max(...v) : 0); }
      case 'MIN': { const v = this.collectNumbers(A, ctx); return v instanceof DecXlError ? v : (v.length ? Math.min(...v) : 0); }
      case 'ABS': { const x = this.num(this.scalar(this.ev(A[0], ctx))); return x instanceof DecXlError ? x : Math.abs(x); }
      case 'AND': { const v = this.collectBools(A, ctx); return v instanceof DecXlError ? v : (v.length ? v.every(Boolean) : new DecXlError('#VALUE!')); }
      case 'OR': { const v = this.collectBools(A, ctx); return v instanceof DecXlError ? v : (v.length ? v.some(Boolean) : new DecXlError('#VALUE!')); }
      case 'NOT': { const b = this.bool(this.scalar(this.ev(A[0], ctx))); return b instanceof DecXlError ? b : !b; }
      case 'TRUE': return true;
      case 'FALSE': return false;
      case 'ISBLANK': return this.scalar(this.ev(A[0], ctx)) === null;
      case 'ISNUMBER': return typeof this.scalar(this.ev(A[0], ctx)) === 'number';
      case 'ISERROR': return this.scalar(this.ev(A[0], ctx)) instanceof DecXlError;
      case 'LEFT': case 'RIGHT': {
        const t = this.scalar(this.ev(A[0], ctx));
        if (t instanceof DecXlError) return t;
        let n = 1;
        if (A[1]) { n = this.num(this.scalar(this.ev(A[1], ctx))); if (n instanceof DecXlError) return n; n = Math.trunc(n); }
        if (n < 0) return new DecXlError('#VALUE!');
        const s = this.str(t);
        return node.name === 'LEFT' ? s.slice(0, n) : (n === 0 ? '' : s.slice(-n));
      }
      case 'VLOOKUP': {
        const lk = this.scalar(this.ev(A[0], ctx));
        if (lk instanceof DecXlError) return lk;
        const tbl = this.ev(A[1], ctx);
        if (!tbl || !tbl.range) throw new DecUnsupported('VLOOKUP table');
        const ci = this.num(this.scalar(this.ev(A[2], ctx)));
        if (ci instanceof DecXlError) return ci;
        const exactArg = A[3] ? this.bool(this.scalar(this.ev(A[3], ctx))) : true;
        if (exactArg !== false) throw new DecUnsupported('approximate VLOOKUP');
        const b = this.rangeBounds(tbl);
        if (ci < 1 || b.c1 + ci - 1 > b.c2) return new DecXlError('#REF!');
        for (let r = b.r1; r <= b.r2; r++) {
          const v = this.get(tbl.sheet, r, b.c1);
          if (v === null || v instanceof DecXlError) continue;
          if (this.rank(v) === this.rank(lk) && this.compare(v, lk) === 0) return this.get(tbl.sheet, r, b.c1 + ci - 1);
        }
        return new DecXlError('#N/A');
      }
    }
    throw new DecUnsupported('function ' + node.name);
  }
}

/* ============================================================
   2. SHEET ANALYSIS (finds sections/totals by label, not fixed rows,
      so sheets that grew by row insertion are still handled)
   ============================================================ */
const DCB_SECTION_DEFS = [
  { re: /^fabrics?$/i,            name: 'Fabrics',             label: 'Fabrics' },
  { re: /^legal marking$/i,       name: 'Legal Marking',       label: 'Legal Marking' },
  { re: /^labels?$/i,             name: 'Label',               label: 'Labels' },
  { re: /^graphics?$/i,           name: 'Graphic',             label: 'Graphics' },
  { re: /^accessories$/i,         name: 'Accessories',         label: 'Accessories' },
  { re: /^sales packaging$/i,     name: 'Sales Packaging',     label: 'Sales Packaging' },
  { re: /^transport packaging$/i, name: 'Transport Packaging', label: 'Transport Packaging' },
];
// Column numbers in the Decathlon cost sheet
const DCB = { type: 1, desig: 2, dsm: 3, model: 4, supplier: 6, width: 7, widthUnit: 8, price: 9, usd: 21, localSum: 19, marker: 22, summaryUsd: 12, localTransportVal: 9 };
const DCB_MAXCOL = 27;
// Accounting number formats (unit price needs 5 decimals so flat rates like 0.00054 stay visible)
const DCB_ACCT5 = '_($* #,##0.00000_);_($* (#,##0.00000);_($* "-"?????_);_(@_)';
const DCB_ACCT4 = '_($* #,##0.0000_);_($* (#,##0.0000);_($* "-"????_);_(@_)';
const DCB_NUMERIC_TEXT_RE = /^\s*-?\d+(\.\d+)?\s*$/;

function dcbSectionRowSet(L) {
  const set = new Set();
  for (const s of L.sections) for (let r = s.firstRow; r <= s.lastRow; r++) set.add(r);
  return set;
}

// Usable width + unit are kept together in ONE cell ("140 cm"); the Width unit column stays blank.
function dcbWidthUnitHint(ws, r) {
  const g = dcbCellText(ws, r, DCB.width).trim();
  const h = dcbCellText(ws, r, DCB.widthUnit).trim();
  if (h) return h;
  const m = g.match(/[A-Za-z]+/);
  return m ? m[0] : '';
}
function dcbCombinedWidth(g, h) {
  g = (g || '').trim(); h = (h || '').trim();
  return (g && h && !/[A-Za-z]/.test(g)) ? `${g} ${h}` : g;
}
function dcbNormalizeWidth(ws, r, newWidth, unitHint) {
  const gCell = ws.getCell(r, DCB.width);
  let w = (newWidth !== undefined && newWidth !== null && String(newWidth).trim() !== '') ? String(newWidth).trim() : dcbCellText(ws, r, DCB.width).trim();
  if (w && !/[A-Za-z]/.test(w)) {
    const unit = unitHint || (newWidth !== undefined && newWidth !== null && newWidth !== '' ? 'cm' : '');
    if (unit) w = `${w} ${unit}`;
  }
  if (w) gCell.value = w;
  ws.getCell(r, DCB.widthUnit).value = null;
}

// Excel flags "inconsistent formula" (LEFT(B,7) vs LEFT(B,10) etc.) and similar benign hints. Real errors
// (#VALUE!, #DIV/0! ...) are NOT suppressed. Needs JSZip (loaded on demand).
function dcbLoadJSZip() {
  if (window.JSZip) return Promise.resolve(window.JSZip);
  return new Promise((res, rej) => {
    const sc = document.createElement('script');
    sc.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
    sc.onload = () => res(window.JSZip);
    sc.onerror = () => rej(new Error('JSZip failed to load'));
    document.head.appendChild(sc);
  });
}
async function dcbInjectIgnoredErrors(buffer, maxRow) {
  const JSZip = await dcbLoadJSZip();
  const zip = await JSZip.loadAsync(buffer);
  const tag = `<ignoredErrors><ignoredError sqref="A1:AB${maxRow}" formula="1" formulaRange="1" numberStoredAsText="1" unlockedFormula="1"/></ignoredErrors>`;
  for (const name of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) continue;
    let xml = await zip.file(name).async('string');
    if (xml.includes('<ignoredErrors')) continue;
    const m = xml.match(/<(smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)[\s>\/]/);
    xml = m ? xml.slice(0, m.index) + tag + xml.slice(m.index) : xml.replace('</worksheet>', tag + '</worksheet>');
    zip.file(name, xml);
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function dcbCellText(ws, r, c) {
  const cell = ws.findCell(r, c);
  if (!cell) return '';
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object') {
    if (v instanceof Date) return '';
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.error) return '';
    if ('result' in v) return (v.result === null || v.result === undefined || typeof v.result === 'object') ? '' : String(v.result);
    if (v.text !== undefined) return String(v.text);
    return '';
  }
  return String(v);
}

function dcbFindLabelRow(ws, label, maxCol) {
  const t = label.toLowerCase();
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= maxCol; c++) {
      if (dcbCellText(ws, r, c).trim().toLowerCase() === t) return r;
    }
  }
  return null;
}

function dcbAnalyzeSheet(ws) {
  const sections = [];
  let open = null;
  let bomRow = null;
  for (let r = 1; r <= ws.rowCount; r++) {
    const a = dcbCellText(ws, r, DCB.type).trim();
    if (!a) continue;
    if (/^total bom cost$/i.test(a)) { bomRow = r; open = null; continue; }
    const def = DCB_SECTION_DEFS.find(d => d.re.test(a));
    // A section title row carries text in column A only; an item whose Type
    // happens to read e.g. "Label" has a Designation in column B.
    // (Title rows are merged across A:B, so B either is empty or mirrors A.)
    const b = dcbCellText(ws, r, DCB.desig).trim();
    if (def && (!b || b.toLowerCase() === a.toLowerCase())) {
      open = { name: def.name, label: def.label, headerRow: r, totalRow: null };
      sections.push(open);
      continue;
    }
    if (open && /^total\s+.+\s+cost$/i.test(a)) { open.totalRow = r; open = null; }
  }
  const valid = sections.filter(s => s.totalRow);
  if (!valid.length || !bomRow) return null;
  valid.forEach(s => { s.firstRow = s.headerRow + 1; s.lastRow = s.totalRow - 1; });
  return {
    sections: valid,
    bomRow,
    cmRow: dcbFindLabelRow(ws, 'TOTAL CM PRICE', 10),
    otherRow: dcbFindLabelRow(ws, 'TOTAL OTHER COSTS', 10),
    exwRow: dcbFindLabelRow(ws, 'EXW PRICE', 10),
    vendorRow: dcbFindLabelRow(ws, 'VENDOR PRICE', 10),
    localRow: dcbFindLabelRow(ws, 'Local Transport', 10),
  };
}

/* ============================================================
   3. SNAPSHOT (evaluated totals) + ROW HELPERS
   ============================================================ */
function dcbEvalNumber(engine, ws, r, c, cachedFallback) {
  try {
    const v = engine.get(ws.name, r, c);
    return typeof v === 'number' ? v : null;
  } catch (e) {
    if (e instanceof DecUnsupported && cachedFallback) {
      const cell = ws.findCell(r, c);
      return (cell && typeof cell.result === 'number') ? cell.result : null;
    }
    return null;
  }
}

function dcbSnapshot(engine, ws, L, itemRows, cachedFallback) {
  const v = (r, c) => (r ? dcbEvalNumber(engine, ws, r, c, cachedFallback) : null);
  const snap = { sections: {}, rowCost: new Map() };
  for (const s of L.sections) snap.sections[s.name] = v(s.totalRow, DCB.usd);
  snap.bom = v(L.bomRow, DCB.usd);
  snap.cm = v(L.cmRow, DCB.summaryUsd);
  snap.other = v(L.otherRow, DCB.summaryUsd);
  snap.exw = v(L.exwRow, DCB.summaryUsd);
  snap.vendor = v(L.vendorRow, DCB.summaryUsd);
  snap.local = v(L.localRow, DCB.localTransportVal);
  for (const it of itemRows) snap.rowCost.set(it.r, v(it.r, DCB.usd));
  return snap;
}

// Price-cell colour rule (updated cost sheet, column I):
//   green = price changed (updated / newly added)
//   pink  = price same as previous
//   red   = any issue (lookup failed / list not loaded / no DSM) - previous price is KEPT, value untouched
// Fixed template rows with no DSM in the packaging sections are not an issue and are left uncoloured.
const DCB_PRICE_COLORS = { changed: 'FFB7E1A1', same: 'FFFFC0CB', issue: 'FFFF7C80' };
function dcbIsIssue(it) {
  if (it.kind === 'fail') return true;
  if (it.kind === 'skip') return !(it.status === 'Skipped – no DSM' && (it.sec.name === 'Sales Packaging' || it.sec.name === 'Transport Packaging'));
  return false;
}
function dcbPriceColorKey(it) {
  if (it.kind === 'ok') return (it.status === 'Price updated' || it.status === 'Price added') ? 'changed' : 'same';
  return dcbIsIssue(it) ? 'issue' : null;
}

function dcbMarkerString(ws, r) {
  const cell = ws.findCell(r, DCB.marker);
  if (!cell || cell.value === null || cell.value === undefined) return '';
  if (cell.type === ExcelJS.ValueType.Formula) return '=' + (cell.formula || '');
  return dcbCellText(ws, r, DCB.marker);
}

// DSM = the FIRST 10 digits of the Designation (sometimes the first 7);
// item code = the LAST 7 digits. The digits do not have to be a separate
// "word" - they are read straight off the start / end of the text. The sheet's
// own DSM column (C, =LEFT(B,10) or LEFT(B,7)) is used as a hint for the 7-vs-10
// decision when available.
function dcbParseDesignation(des, dsmHint) {
  const s = des.trim();
  let dsm = '';
  if (dsmHint && /^\d{7,10}$/.test(dsmHint) && s.startsWith(dsmHint)) {
    dsm = dsmHint;
  } else {
    const m10 = s.match(/^(\d{10})/);
    const m7 = s.match(/^(\d{7})/);
    dsm = m10 ? m10[1] : (m7 ? m7[1] : '');
  }
  let itemCode = '';
  const m = s.match(/(\d{7})\s*$/);
  if (m && m.index >= dsm.length) itemCode = m[1];
  return { dsm, itemCode };
}

function dcbApplyPricing(ws, r, pricing, unitHint, skipWidth) {
  ws.getCell(r, DCB.price).value = pricing.unitPrice;

  const supCell = ws.getCell(r, DCB.supplier);
  supCell.value = (pricing.supplierText === '' || pricing.supplierText === null || pricing.supplierText === undefined) ? null : pricing.supplierText;
  const argb = supCell.font && supCell.font.color && supCell.font.color.argb;
  if (argb === 'FFCC0000') { // previous run left a red error message here
    supCell.style = Object.assign({}, supCell.style, { font: Object.assign({}, supCell.font, DECATHLON_DEFAULT_FONT) });
  }

  // Api marker column: formula only - never text comments.
  const vCell = ws.getCell(r, DCB.marker);
  vCell.value = pricing.needsApiMarker ? { formula: `S${r}*0.15` } : null;

  // Model code column stays blank here; the (red) comment is written after template repair.
  ws.getCell(r, DCB.model).value = null;

  if (pricing.usableWidth !== undefined && !skipWidth) {
    if (!pricing.usableWidth) { ws.getCell(r, DCB.width).value = null; ws.getCell(r, DCB.widthUnit).value = null; }
    else dcbNormalizeWidth(ws, r, pricing.usableWidth, unitHint);
  }
}

/* ============================================================
   3b. TEMPLATE REPAIR
   The original Decathlon template (embedded in decathlon.js) is the single
   source of truth for how a cost sheet should look and calculate. After
   re-pricing, every cost-sheet tab is checked against it:
     - a formula the template has but the sheet lost  -> written back
     - a cell whose colours / font / alignment / borders / number format
       drifted from the template                      -> reverted
     - a dropdown (Width Unit / Per unit) that is missing -> restored
   Existing formulas are never overwritten. Rows are matched to the template
   by section (so sheets that grew by row insertion line up correctly).
   The red "price not found" text in Component supplier is intentional
   formatting and is kept.
   ============================================================ */
const DCB_TPL_KEYS = ['Fabrics', 'Legal Marking', 'Label', 'Graphic', 'Accessories', 'Sales Packaging', 'Transport Packaging'];
const DCB_RED = 'FFCC0000';
// Columns the Open Book engine writes as INPUTS (type, designation, model, supplier, width, unit, price, ...).
// A template formula sitting in one of these (e.g. Sales Packaging's  L = 1/G173)  is a starter value, not
// a calculation - it must never be forced back onto an item row.
const DCB_INPUT_COLS = new Set([1, 2, 4, 6, 7, 8, 9, 10, 12, 13, 14, 17, 18, 22]);

function dcbTemplateInfo(tplWs) {
  const secs = {};
  for (const k of DCB_TPL_KEYS) {
    const c = DECATHLON_SECTION_CONFIG[k];
    secs[k] = { header: c.headerRow, first: c.firstBlank, last: c.lastBlank, total: c.totalRow };
    const def = DCB_SECTION_DEFS.find(d => d.name === k);
    if (!def.re.test(dcbCellText(tplWs, c.headerRow, 1).trim())) return null; // template layout isn't what we expect
  }
  let bom = null, last = 0;
  for (let r = 1; r <= tplWs.rowCount; r++) {
    if (bom === null && /^total bom cost$/i.test(dcbCellText(tplWs, r, 1).trim())) bom = r;
    for (let c = 1; c <= DCB_MAXCOL; c++) {
      const x = tplWs.findCell(r, c);
      if (x && x.value !== null && x.value !== undefined && x.value !== '') { last = r; break; }
    }
  }
  if (!bom) return null;
  return { ws: tplWs, secs, bom, last };
}

function dcbMakeMapper(L, T) {
  const cur = {};
  L.sections.forEach(s => { cur[s.name] = s; });
  // template row -> this sheet's row
  const toCur = rT => {
    if (rT <= 10) return rT;
    for (const k of DCB_TPL_KEYS) {
      const t = T.secs[k], c = cur[k];
      if (!c) continue;
      if (rT === t.header) return c.headerRow;
      if (rT === t.last) return c.lastRow;
      if (rT >= t.first && rT < t.last) return Math.min(c.firstRow + (rT - t.first), c.lastRow);
      if (rT === t.total) return c.totalRow;
    }
    if (rT >= T.bom) return L.bomRow + (rT - T.bom);
    return null;
  };
  // this sheet's row -> template row (inserted rows take the last blank row's look)
  const toTpl = r => {
    if (r <= 10) return r;
    for (const k of DCB_TPL_KEYS) {
      const t = T.secs[k], c = cur[k];
      if (!c) continue;
      if (r === c.headerRow) return t.header;
      if (r === c.lastRow) return t.last;
      if (r >= c.firstRow && r < c.lastRow) return t.first + Math.min(r - c.firstRow, t.last - t.first);
      if (r === c.totalRow) return t.total;
    }
    if (r >= L.bomRow && r <= L.bomRow + (T.last - T.bom)) return T.bom + (r - L.bomRow);
    return null;
  };
  return { cur, toCur, toTpl };
}

function dcbColorKey(c) {
  if (!c) return '';
  if (c.argb) return c.argb;
  if (c.theme !== undefined) return 't' + c.theme + ':' + (c.tint || 0);
  if (c.indexed !== undefined) return 'i' + c.indexed;
  return '';
}
function dcbStyleParts(st) {
  st = st || {};
  const f = st.fill || {}, fo = st.font || {}, a = st.alignment || {}, b = st.border || {};
  const side = k => (b[k] ? (b[k].style || '') + dcbColorKey(b[k].color) : '');
  return {
    fill: [f.type || '', f.pattern || '', dcbColorKey(f.fgColor), dcbColorKey(f.bgColor)].join('|'),
    fontColor: dcbColorKey(fo.color),
    font: [fo.name || '', fo.size || '', !!fo.bold, !!fo.italic, !!fo.underline, !!fo.strike].join('|'),
    alignment: [a.horizontal || '', a.vertical || '', !!a.wrapText, a.indent || 0, a.textRotation || 0].join('|'),
    border: ['left', 'right', 'top', 'bottom'].map(side).join('|'),
    numFmt: st.numFmt || 'General',
  };
}
const DCB_PART_LABEL = { fill: 'fill colour', fontColor: 'font colour', font: 'font', alignment: 'alignment', border: 'borders', numFmt: 'number format' };

function dcbTranslateFormula(f, mapRow) {
  return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, d1, col, d2, row) => d1 + col + d2 + mapRow(parseInt(row, 10)));
}

function dcbRepairSheet(ws, L, T, items) {
  const tplWs = T.ws;
  const M = dcbMakeMapper(L, T);
  const out = { formulas: [], styles: [], dropdowns: [] };
  const dsm7 = new Set(items.filter(it => it.dsm && it.dsm.length === 7).map(it => it.r));
  const secRows = dcbSectionRowSet(L);

  const restoreFormulas = (r, tr, mapRow, isItemRow) => {
    for (let c = 1; c <= DCB_MAXCOL; c++) {
      const tc = tplWs.findCell(tr, c);
      if (!tc) continue;
      const cell = ws.getCell(r, c);
      if (tc.type === ExcelJS.ValueType.Formula && tc.formula) {
        if (cell.type === ExcelJS.ValueType.Formula) continue; // never overwrite an existing formula
        if (isItemRow && DCB_INPUT_COLS.has(c)) continue;      // input column on an item row - leave the user's value
        let f = dcbTranslateFormula(tc.formula, mapRow);
        if (isItemRow && c === DCB.dsm && dsm7.has(r)) f = f.replace(/,\s*10\)/, ',7)');
        out.formulas.push({ addr: cell.address, formula: '=' + f, cell, old: cell.value });
        cell.value = { formula: f };
      } else if (isItemRow && c === 20 && tc.value === 'USD') { // "Local currency" label that every row carries
        if (cell.value === null || cell.value === undefined || cell.value === '') { cell.value = 'USD'; out.formulas.push({ addr: cell.address, formula: 'USD', label: true }); }
      }
    }
  };

  for (const k of DCB_TPL_KEYS) {
    const c = M.cur[k], t = T.secs[k];
    if (!c) continue;
    for (let r = c.firstRow; r <= c.lastRow; r++) {
      const tr = M.toTpl(r);
      restoreFormulas(r, tr, ref => (ref === tr ? r : (M.toCur(ref) ?? ref)), true);
    }
    restoreFormulas(c.totalRow, t.total, ref => M.toCur(ref) ?? ref, false);
  }
  for (let tr = T.bom; tr <= T.last; tr++) restoreFormulas(L.bomRow + (tr - T.bom), tr, ref => M.toCur(ref) ?? ref, false);

  // Local Transport (summary block) is written by the Open Book engine, not the template:
  // = SUM of the Api marker column over the whole BOM area.
  if (L.localRow && M.cur['Fabrics'] && M.cur['Transport Packaging']) {
    const lc = ws.getCell(L.localRow, DCB.localTransportVal);
    if (lc.type !== ExcelJS.ValueType.Formula) {
      const f = `SUM(V${M.cur['Fabrics'].headerRow}:V${M.cur['Transport Packaging'].totalRow})`;
      out.formulas.push({ addr: lc.address, formula: '=' + f, cell: lc, old: lc.value });
      lc.value = { formula: f };
    }
  }

  // Safety net: a restored formula must never introduce an error. If one would (bad/missing inputs it
  // depends on), put the previous cell content back and report it instead.
  out.notRestored = [];
  for (let pass = 0; pass < 4; pass++) {
    const eng = new DecFormulaEngine(ws.workbook);
    const bad = out.formulas.filter(x => x.cell && !x.undone && (() => {
      try { return eng.get(ws.name, x.cell.row, x.cell.col) instanceof DecXlError; } catch (e) { return false; }
    })());
    if (!bad.length) break;
    for (const x of bad) { x.cell.value = x.old === undefined ? null : x.old; x.undone = true; out.notRestored.push({ addr: x.addr, formula: x.formula }); }
  }
  out.formulas = out.formulas.filter(x => !x.undone).map(x => ({ addr: x.addr, formula: x.formula, label: x.label }));

  // styles (ascending row order on purpose - see ExcelJS note in kariban.js)
  const lastRow = L.bomRow + (T.last - T.bom);
  for (let r = 1; r <= lastRow; r++) {
    const tr = M.toTpl(r);
    if (tr === null) continue;
    for (let c = 1; c <= DCB_MAXCOL; c++) {
      const tc = tplWs.findCell(tr, c);
      if (!tc) continue;
      const cell = ws.getCell(r, c);
      const desired = JSON.parse(JSON.stringify(tc.style || {}));
      const keepRed = c === DCB.supplier && cell.font && cell.font.color && cell.font.color.argb === DCB_RED;
      if (keepRed) desired.font = Object.assign({}, desired.font, { color: { argb: DCB_RED } });
      if (c === DCB.price && secRows.has(r)) desired.numFmt = DCB_ACCT5;
      const a = dcbStyleParts(cell.style), b = dcbStyleParts(desired);
      const diff = Object.keys(a).filter(k => a[k] !== b[k]);
      if (diff.length) {
        cell.style = desired;
        out.styles.push({ addr: cell.address, what: diff.map(k => DCB_PART_LABEL[k]).join(', ') });
      }
      // dropdowns (Width Unit / Per unit)
      if (tc.dataValidation && !cell.dataValidation) {
        cell.dataValidation = JSON.parse(JSON.stringify(tc.dataValidation));
        out.dropdowns.push({ addr: cell.address });
      }
    }
  }
  return out;
}

/* ============================================================
   4. PER-SHEET UPDATE
   ============================================================ */
function dcbProcessSheet(wb, ws, L, opts) {
  const name = ws.name;
  const eng1 = new DecFormulaEngine(wb);

  const items = [];
  for (const sec of L.sections) {
    for (let r = sec.firstRow; r <= sec.lastRow; r++) {
      const des = dcbCellText(ws, r, DCB.desig).trim();
      if (des) items.push({ sec, r, des });
    }
  }

  // BEFORE: evaluate totals/row costs on the untouched sheet (cached results
  // are only a fallback if some formula is outside the supported set).
  const before = dcbSnapshot(eng1, ws, L, items, true);
  for (const it of items) {
    const p = dcbEvalNumber(eng1, ws, it.r, DCB.price, true);
    try {
      const hv = eng1.get(name, it.r, DCB.dsm);
      it.dsmHint = (typeof hv === 'string' || typeof hv === 'number') ? String(hv).trim() : '';
    } catch (e) { it.dsmHint = ''; }
    it.old = {
      price: p,
      supplier: dcbCellText(ws, it.r, DCB.supplier).trim(),
      width: dcbCombinedWidth(dcbCellText(ws, it.r, DCB.width), dcbCellText(ws, it.r, DCB.widthUnit)),
      marker: dcbMarkerString(ws, it.r),
    };
  }

  // Flatten shared formulas BEFORE editing: if a shared-formula master cell
  // (e.g. the first "S*0.15" Api marker) were overwritten, every dependent
  // cell would be corrupted.
  normalizeFormulas(ws, ws.rowCount, DCB_MAXCOL);

  for (const it of items) {
    const { sec, r, des } = it;
    const part = dcbCellText(ws, r, DCB.type).trim();
    const { dsm, itemCode } = dcbParseDesignation(des, it.dsmHint);
    // Same row shape the Open Book engine prices: Type + DSM + Item code,
    // with the whole Designation kept in `component` so keyword rules
    // (e.g. interlining detection) still see the full text.
    const row = { part, dsm, model: '', component: des, itemCode, gridValue: '', items: '', qty: '', unit: '', comments: '' };
    it.part = part; it.dsm = dsm; it.itemCode = itemCode;

    // --- clean-up of every item row ---
    // numbers stored as text -> real numbers
    for (const c of [9, 10, 11, 12, 14, 17, 18]) {
      const cc = ws.getCell(r, c);
      if (typeof cc.value === 'string' && DCB_NUMERIC_TEXT_RE.test(cc.value)) cc.value = parseFloat(cc.value);
    }
    // Model code column: always blank (no formula); Api marker: remove old text comments (formulas stay)
    ws.getCell(r, DCB.model).value = null;
    const mk = ws.getCell(r, DCB.marker);
    if (mk.value !== null && mk.value !== undefined && mk.type !== ExcelJS.ValueType.Formula) mk.value = null;
    // width unit moves into the usable-width cell
    it.unitHint = dcbWidthUnitHint(ws, r);
    if (sec.name !== 'Sales Packaging') dcbNormalizeWidth(ws, r, undefined, it.unitHint);
    // DSM code column: restore the formula if it is missing
    if (dsm && ws.getCell(r, DCB.dsm).type !== ExcelJS.ValueType.Formula) {
      ws.getCell(r, DCB.dsm).value = { formula: `LEFT(B${r},${dsm.length === 7 ? 7 : 10})` };
    }

    const fabricStyle = sec.name === 'Fabrics' && !decathlonIsInterlining(row);
    const flat = fabricStyle ? null : decathlonFlatRateFor(row);

    let pricing = null;
    if (!flat && !dsm) {
      it.kind = 'skip'; it.status = 'Skipped – no DSM'; it.remark = 'No 10/7-digit DSM at the start of the Designation (starts with "' + des.slice(0, 24) + '"); row left unchanged.';
    } else if (fabricStyle && !opts.fabricIdx) {
      it.kind = 'skip'; it.status = 'Skipped – fabric list not loaded'; it.remark = 'Upload a fabric price list to update this row.';
    } else if (!fabricStyle && !flat && !opts.accIdx) {
      it.kind = 'skip'; it.status = 'Skipped – accessories list not loaded'; it.remark = 'Upload an accessories price list to update this row.';
    } else {
      pricing = computeDecathlonPricing(row, sec.name, opts.accIdx, opts.fabricIdx);
      if (!pricing) {
        it.kind = 'skip'; it.status = 'Skipped'; it.remark = 'No pricing rule applies.';
      } else if (pricing.isError) {
        it.kind = 'fail'; it.status = 'Not updated'; it.remark = pricing.supplierText + ' — previous values kept.';
      } else {
        it.kind = 'ok';
      }
    }

    if (it.kind === 'ok') {
      dcbApplyPricing(ws, r, pricing, it.unitHint, sec.name === 'Sales Packaging');
      it.pricing = pricing;
      it.remark = flat ? 'Flat rate' : (pricing.apiMarkerText ? 'Fabric price, region ' + pricing.apiMarkerText : (pricing.needsApiMarker ? 'Api marker +15% (non-Bangladesh origin)' : ''));
      if (pricing.highestPriceTaken) it.remark += (it.remark ? '; ' : '') + 'Highest price taken (Item code not matched)';
    }
  }

  // Template repair: restore missing formulas, revert drifted colours / alignments.
  let repairs = { formulas: [], styles: [], dropdowns: [] };
  if (opts.tpl) repairs = dcbRepairSheet(ws, L, opts.tpl, items);

  // --- final formatting (after template repair so it is not reverted) ---
  const secRowsF = dcbSectionRowSet(L);
  for (const r of secRowsF) {
    const pc = ws.getCell(r, DCB.price);
    pc.style = Object.assign({}, pc.style, { numFmt: DCB_ACCT5 });
  }
  for (const it of items) {
    if (it.kind === 'ok' && it.pricing && it.pricing.highestPriceTaken) {
      const dc = ws.getCell(it.r, DCB.model);
      dc.value = 'Highest price taken';
      dc.style = Object.assign({}, dc.style, { font: Object.assign({}, dc.font, DECATHLON_RED_FONT) });
    }
  }
  for (let r = 1; r <= 6; r++) {
    const bc = ws.getCell(r, 2);
    bc.style = Object.assign({}, bc.style, { alignment: Object.assign({}, bc.alignment, { horizontal: 'center', vertical: 'middle' }) });
  }
  for (const a of ['B2', 'B4']) {
    const bc = ws.getCell(a);
    if (typeof bc.value === 'string' && DCB_NUMERIC_TEXT_RE.test(bc.value)) bc.value = parseFloat(bc.value);
  }

  // AFTER
  const eng2 = new DecFormulaEngine(wb);
  const after = dcbSnapshot(eng2, ws, L, items, false);
  const counts = { total: items.length, priceUpdated: 0, priceAdded: 0, infoUpdated: 0, noChange: 0, notUpdated: 0, skipped: 0 };
  for (const it of items) {
    it.costBefore = before.rowCost.get(it.r);
    it.costAfter = after.rowCost.get(it.r);
    it.new = {
      price: dcbEvalNumber(eng2, ws, it.r, DCB.price, false),
      supplier: dcbCellText(ws, it.r, DCB.supplier).trim(),
      width: dcbCellText(ws, it.r, DCB.width).trim(),
      marker: dcbMarkerString(ws, it.r),
    };
    if (it.kind === 'ok') {
      const priceChanged = it.old.price === null || it.new.price === null || Math.abs(it.old.price - it.new.price) > 1e-9;
      const infoChanged = it.old.supplier !== it.new.supplier
        || (it.pricing.usableWidth !== undefined && it.old.width !== it.new.width)
        || it.old.marker !== it.new.marker;
      if (priceChanged) { it.status = it.old.price === null ? 'Price added' : 'Price updated'; if (it.old.price === null) counts.priceAdded++; else counts.priceUpdated++; }
      else if (infoChanged) { it.status = 'Info updated'; counts.infoUpdated++; }
      else { it.status = 'No change'; counts.noChange++; }
    } else if (it.kind === 'fail') counts.notUpdated++;
    else counts.skipped++;
  }

  // Colour the Unit price cells (whole-style assignment on purpose: assigning .fill alone can bleed
  // across cells that share a style record in ExcelJS).
  for (const it of items) {
    const key = dcbPriceColorKey(it);
    if (!key) continue;
    const cell = ws.getCell(it.r, DCB.price);
    const st = JSON.parse(JSON.stringify(cell.style || {}));
    st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DCB_PRICE_COLORS[key] } };
    cell.style = st;
    it.priceColor = key;
  }

  dcbRefreshCachedResults(wb, ws);

  const readCell = (r, c) => dcbCellText(ws, r, c).trim();
  return {
    sheetName: name,
    productName: readCell(1, 2), cc: readCell(2, 2), r3: readCell(4, 2),
    layout: L, items, before, after, counts, repairs,
  };
}

// After editing, store freshly evaluated results next to each formula so
// viewers that don't recalculate still show correct numbers. Excel itself
// recalculates on open (fullCalcOnLoad is set), so this is a convenience.
// Any formula the mini-engine can't evaluate is simply left without a cached
// result.
function dcbRefreshCachedResults(wb, ws) {
  const eng = new DecFormulaEngine(wb);
  for (let r = 1; r <= ws.rowCount; r++) {
    for (let c = 1; c <= DCB_MAXCOL; c++) {
      const cell = ws.findCell(r, c);
      if (!cell || cell.type !== ExcelJS.ValueType.Formula) continue;
      const f = cell.formula;
      if (!f) continue;
      try {
        const v = eng.get(ws.name, r, c);
        if (v instanceof DecXlError) cell.value = { formula: f, result: { error: v.code } };
        else if (v === null) cell.value = { formula: f, result: 0 };
        else cell.value = { formula: f, result: v };
      } catch (e) {
        if (!(e instanceof DecUnsupported)) console.warn('result refresh failed at', ws.name, r, c, e);
        cell.value = { formula: f };
      }
    }
  }
}

async function dcbRunUpdate(arrayBuffer, opts) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  const results = [];
  const skippedTabs = [];
  if (!opts.tpl) {
    try {
      const tplWb = new ExcelJS.Workbook();
      await tplWb.xlsx.load(base64ToArrayBuffer(DECATHLON_TEMPLATE_B64));
      opts.tpl = dcbTemplateInfo(tplWb.getWorksheet('Format') || tplWb.worksheets[0]);
    } catch (e) { console.warn('Template repair disabled:', e); opts.tpl = null; }
  }
  for (const ws of wb.worksheets) {
    ws.views = (ws.views || []).map(v => { const { state, xSplit, ySplit, topLeftCell, activePane, pane, ...rest } = v; return rest; });
  }
  for (const ws of [...wb.worksheets]) {
    const L = dcbAnalyzeSheet(ws);
    if (!L) { skippedTabs.push(ws.name); continue; }
    opts.onProgress && opts.onProgress(ws.name);
    results.push(dcbProcessSheet(wb, ws, L, opts));
  }
  if (!results.length) throw new Error('No Decathlon cost-sheet tabs were recognised in this workbook (looked for FABRICS … Total bom cost).');
  wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
  let buffer = await wb.xlsx.writeBuffer();
  try {
    const maxRow = Math.max(...wb.worksheets.map(w => w.rowCount)) + 5;
    buffer = await dcbInjectIgnoredErrors(buffer, maxRow);
  } catch (e) { console.warn('Could not add ignoredErrors (Excel hints stay visible):', e); }
  return { buffer, results, skippedTabs };
}

/* ============================================================
   5. CHANGE REPORT (separate workbook)
   ============================================================ */
// Status colours in the report mirror the price-cell colours in the cost sheet
// (grey = skipped fixed rows that are not an issue).
function dcbStatusFill(it) {
  const key = dcbPriceColorKey(it);
  return key ? DCB_PRICE_COLORS[key] : 'FFE5E7EB';
}

function dcbBuildReport(run, meta) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'BOM Generator — Decathlon Closed Book';
  const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
  const HEAD_FONT = { bold: true, color: { argb: 'FFFFFFFF' } };
  const TOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF4FF' } };
  const THIN = { style: 'thin', color: { argb: 'FF9CA3AF' } };
  const BORDER = { top: THIN, left: THIN, bottom: THIN, right: THIN };
  const CENTER = { horizontal: 'center', vertical: 'middle', wrapText: true };
  const MONEY = '#,##0.0000';
  const PCT = '0.00%';

  // Writes a value; digit-only text (DSM, item code, CC, tab name ...) becomes a REAL number so
  // Excel never shows the "number stored as text" warning. Zero-padded number formats keep
  // leading zeros visible, e.g. 0012345 stays 0012345.
  function put(ws, r, c, v) {
    const cell = ws.getCell(r, c);
    if (v === null || v === undefined || v === '') return cell;
    if (typeof v === 'string') {
      const t = v.trim();
      if (/^\d{1,15}$/.test(t)) { cell.value = Number(t); cell.numFmt = '0'.repeat(t.length); return cell; }
      if (/^-?\d+\.\d+$/.test(t) && t.length <= 16) { cell.value = Number(t); return cell; }
    }
    cell.value = v;
    return cell;
  }
  // borders + alignment on a whole rectangle (empty cells inside a table get borders too)
  function finishTable(ws, r1, c1, r2, c2, centerCols) {
    const ctr = new Set(centerCols || []);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(r, c);
        const v = cell.value;
        const isNum = typeof v === 'number' || (v && typeof v === 'object' && v.formula !== undefined);
        cell.border = BORDER;
        cell.alignment = { vertical: 'middle', horizontal: ctr.has(c) ? 'center' : (isNum ? 'right' : 'left'), wrapText: true };
      }
    }
  }
  function styleHeader(ws, r, c1, c2, plain) {
    for (let c = c1; c <= c2; c++) {
      const cell = ws.getCell(r, c);
      if (!plain) { cell.fill = HEAD_FILL; cell.font = HEAD_FONT; } else { cell.font = { bold: true }; }
      cell.alignment = CENTER; cell.border = BORDER;
    }
    ws.getRow(r).height = 30;
  }
  function mergeTabBlock(ws, r1, r2) {
    if (r2 > r1) ws.mergeCells(r1, 1, r2, 1);
    for (let r = r1; r <= r2; r++) ws.getCell(r, 1).border = BORDER;
    const c = ws.getCell(r1, 1);
    c.alignment = CENTER; c.font = { bold: true };
  }
  // before / after / diff / diff% — diff cells are live formulas (with cached results)
  function putDelta(ws, r, cBefore, cAfter, cDiff, cPct, b, a) {
    const L = n => ws.getColumn(n).letter;
    if (typeof b === 'number' && typeof a === 'number') {
      ws.getCell(r, cDiff).value = { formula: `${L(cAfter)}${r}-${L(cBefore)}${r}`, result: a - b };
      ws.getCell(r, cPct).value = { formula: `IF(${L(cBefore)}${r}=0,"",${L(cDiff)}${r}/${L(cBefore)}${r})`, result: b === 0 ? '' : (a - b) / b };
    }
    ws.getCell(r, cDiff).numFmt = MONEY;
    ws.getCell(r, cPct).numFmt = PCT;
  }

  const lineDefs = s => [
    ...s.layout.sections.map(sec => ({ label: sec.label, get: sn => sn.sections[sec.name], kind: 'section' })),
    { label: 'Total BOM cost', get: sn => sn.bom, kind: 'total' },
    { label: 'Total CM price', get: sn => sn.cm, kind: 'line' },
    { label: 'Local Transport (Api marker total)', get: sn => sn.local, kind: 'line' },
    { label: 'Total other costs', get: sn => sn.other, kind: 'line' },
    { label: 'EXW price', get: sn => sn.exw, kind: 'line' },
    { label: 'VENDOR PRICE', get: sn => sn.vendor, kind: 'total' },
  ];

  /* ---------- Summary ---------- */
  const ws1 = wb.addWorksheet('Summary');
  ws1.getCell('A1').value = 'Decathlon Price Update — Change Report';
  ws1.getCell('A1').font = { bold: true, size: 16 };
  const info = [
    ['Generated', meta.generated],
    ['Cost sheet (previous version)', meta.costSheetName],
    ['Fabric price list', meta.fabricInfo],
    ['Accessories price list', meta.accInfo],
    ['All amounts', 'USD price. Δ = After − Before.'],
    ['Price cell colours (cost sheet, Unit price)', 'Green = price changed | Pink = same as previous | Red = issue (previous price kept)'],
  ];
  info.forEach((p, i) => { put(ws1, 3 + i, 1, p[0]); put(ws1, 3 + i, 2, p[1]); ws1.getCell(3 + i, 1).font = { bold: true }; });
  finishTable(ws1, 3, 1, 3 + info.length - 1, 2);
  info.forEach((p, i) => { ws1.getCell(3 + i, 1).font = { bold: true }; });
  ws1.getColumn(1).width = 30; ws1.getColumn(2).width = 40; ws1.getColumn(3).width = 12;

  const hr = 3 + info.length + 1;               // group-header row; sub-header = hr + 1
  const GROUPS = [
    { title: 'Total BOM cost', get: sn => sn.bom, key: true },
    { title: 'VENDOR PRICE',   get: sn => sn.vendor, key: true },
  ];
  const nSumCols = 3 + GROUPS.length * 4;
  ['Tab', 'Product', 'CC'].forEach((h, i) => { ws1.getCell(hr, i + 1).value = h; });
  GROUPS.forEach((g, gi) => {
    const c0 = 4 + gi * 4;
    ws1.getCell(hr, c0).value = g.title;
    ['Before', 'After', 'Δ', 'Δ %'].forEach((h, k) => { ws1.getCell(hr + 1, c0 + k).value = h; });
  });
  run.results.forEach((s, i) => {
    const r = hr + 2 + i;
    put(ws1, r, 1, s.sheetName); put(ws1, r, 2, s.productName); put(ws1, r, 3, s.cc);
    GROUPS.forEach((g, gi) => {
      const c0 = 4 + gi * 4;
      const bv = g.get(s.before), av = g.get(s.after);
      ws1.getCell(r, c0).value = bv; ws1.getCell(r, c0 + 1).value = av;
      ws1.getCell(r, c0).numFmt = MONEY; ws1.getCell(r, c0 + 1).numFmt = MONEY;
      putDelta(ws1, r, c0, c0 + 1, c0 + 2, c0 + 3, bv, av);
    });
  });
  const lastSum = hr + 1 + run.results.length;
  finishTable(ws1, hr, 1, lastSum, nSumCols, [1, 3]);
  for (let r = hr + 2; r <= lastSum; r++) [4, 5, 6, 8, 9, 10].forEach(c => { ws1.getCell(r, c).numFmt = DCB_ACCT4; });
  // coloured key groups (Total BOM cost, VENDOR PRICE)
  run.results.forEach((s, i) => {
    GROUPS.forEach((g, gi) => {
      if (!g.key) return;
      for (let k = 0; k < 4; k++) ws1.getCell(hr + 2 + i, 4 + gi * 4 + k).fill = TOTAL_FILL;
    });
  });
  ['Tab', 'Product', 'CC'].forEach((h, i) => { ws1.mergeCells(hr, i + 1, hr + 1, i + 1); });
  GROUPS.forEach((g, gi) => { ws1.mergeCells(hr, 4 + gi * 4, hr, 7 + gi * 4); });
  styleHeader(ws1, hr, 1, nSumCols); styleHeader(ws1, hr + 1, 1, nSumCols);
  for (let c = 4; c <= nSumCols; c++) ws1.getColumn(c).width = ((c - 4) % 4 >= 2) ? 13 : 16;

  /* ---------- Section Costs ---------- */
  const ws2 = wb.addWorksheet('Section Costs');
  ['Tab', 'Line', 'Before (USD)', 'After (USD)', 'Δ (USD)', 'Δ %'].forEach((h, i) => { ws2.getCell(1, i + 1).value = h; });
  let r2 = 2;
  const tabBlocks2 = [];
  for (const s of run.results) {
    const blockStart = r2;
    for (const ln of lineDefs(s)) {
      const b = ln.get(s.before), a = ln.get(s.after);
      ws2.getCell(r2, 2).value = ln.label;
      ws2.getCell(r2, 3).value = b; ws2.getCell(r2, 4).value = a;
      ws2.getCell(r2, 3).numFmt = MONEY; ws2.getCell(r2, 4).numFmt = MONEY;
      putDelta(ws2, r2, 3, 4, 5, 6, b, a);
      ln._row = r2; ln._kind = ln.kind;
      r2++;
    }
    tabBlocks2.push({ s, blockStart, blockEnd: r2 - 1, lines: lineDefs(s).map((ln, i) => ({ kind: ln.kind, row: blockStart + i })) });
  }
  finishTable(ws2, 1, 1, r2 - 1, 6);
  for (let r = 2; r < r2; r++) [3, 4, 5].forEach(c => { ws2.getCell(r, c).numFmt = DCB_ACCT4; });
  for (const tb of tabBlocks2) {
    for (const ln of tb.lines) {
      if (ln.kind !== 'total') continue; // only Total BOM cost & VENDOR PRICE are coloured
      for (let c = 2; c <= 6; c++) { const cell = ws2.getCell(ln.row, c); cell.fill = TOTAL_FILL; cell.font = { bold: true }; }
    }
    put(ws2, tb.blockStart, 1, tb.s.sheetName);
    mergeTabBlock(ws2, tb.blockStart, tb.blockEnd);
  }
  styleHeader(ws2, 1, 1, 6, true);
  [26, 36, 18, 18, 16, 12].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });

  /* ---------- Item Changes ---------- */
  const ws3 = wb.addWorksheet('Item Changes');
  const h3 = ['Tab', 'Section', 'Row', 'Type', 'Designation', 'DSM', 'Item code (last 7)', 'Old price', 'New price', 'Δ price', 'Δ price %',
    'Old supplier', 'New supplier', 'Old usable width', 'New usable width',
    'Row cost before (USD)', 'Row cost after (USD)', 'Cost impact (USD)', 'Status', 'Remark'];
  h3.forEach((h, i) => { ws3.getCell(1, i + 1).value = h; });
  let r3 = 2;
  const tabBlocks3 = [];
  for (const s of run.results) {
    const blockStart = r3;
    for (const it of s.items) {
      const vals = [null, it.sec.label, it.r, it.part, it.des, it.dsm, it.itemCode, it.old.price, it.new.price, null, null,
        it.old.supplier, it.new.supplier, it.old.width, it.new.width,
        it.costBefore, it.costAfter, null, it.status, it.remark || ''];
      vals.forEach((v, j) => { put(ws3, r3, j + 1, v); });
      [8, 9, 10].forEach(n => { ws3.getCell(r3, n).numFmt = DCB_ACCT5; });
      ws3.getCell(r3, 11).numFmt = PCT;
      [16, 17, 18].forEach(n => { ws3.getCell(r3, n).numFmt = DCB_ACCT4; });
      if (typeof it.old.price === 'number' && typeof it.new.price === 'number') {
        ws3.getCell(r3, 10).value = { formula: `I${r3}-H${r3}`, result: it.new.price - it.old.price };
        ws3.getCell(r3, 11).value = { formula: `IF(H${r3}=0,"",J${r3}/H${r3})`, result: it.old.price === 0 ? '' : (it.new.price - it.old.price) / it.old.price };
      }
      if (typeof it.costBefore === 'number' && typeof it.costAfter === 'number') {
        ws3.getCell(r3, 18).value = { formula: `Q${r3}-P${r3}`, result: it.costAfter - it.costBefore };
      }
      r3++;
    }
    tabBlocks3.push({ s, blockStart, blockEnd: r3 - 1 });
  }
  finishTable(ws3, 1, 1, Math.max(r3 - 1, 1), h3.length, [3, 11, 14, 15, 19]);
  for (let r = 2; r < r3; r++) {
    ws3.getCell(r, 6).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    [11, 14, 15].forEach(n => { ws3.getCell(r, n).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }; });
  }
  for (const tb of tabBlocks3) {
    tb.s.items.forEach((it, i) => {
      const fill = dcbStatusFill(it);
      if (fill) ws3.getCell(tb.blockStart + i, 19).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
    });
    if (tb.blockEnd >= tb.blockStart) { put(ws3, tb.blockStart, 1, tb.s.sheetName); mergeTabBlock(ws3, tb.blockStart, tb.blockEnd); }
  }
  styleHeader(ws3, 1, 1, h3.length);
  [24, 18, 7, 22, 52, 13, 14, 11, 11, 11, 10, 20, 20, 16, 16, 16, 16, 16, 24, 60].forEach((w, i) => { ws3.getColumn(i + 1).width = w; });
  ws3.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: h3.length } };

  /* ---------- Needs Attention ---------- */
  const ws4 = wb.addWorksheet('Needs Attention');
  const h4 = ['Tab', 'Section', 'Row', 'Type', 'Designation', 'DSM', 'Item code (last 7)', 'Issue', 'Price kept in sheet', 'Supplier cell now reads'];
  h4.forEach((h, i) => { ws4.getCell(1, i + 1).value = h; });
  let r4 = 2;
  const issueFills = [];
  const blocks4 = [];
  for (const s of run.results) {
    const start4 = r4;
    for (const it of s.items) {
      if (!dcbIsIssue(it)) continue;
      [null, it.sec.label, it.r, it.part, it.des, it.dsm, it.itemCode, it.remark || it.status, it.old.price, it.old.supplier].forEach((v, j) => { put(ws4, r4, j + 1, v); });
      ws4.getCell(r4, 9).numFmt = DCB_ACCT5;
      issueFills.push([r4, DCB_PRICE_COLORS.issue]);
      r4++;
    }
    if (r4 > start4) blocks4.push({ name: s.sheetName, start: start4, end: r4 - 1 });
  }
  if (r4 === 2) { ws4.getCell(2, 1).value = 'Nothing needs attention — every priced row was resolved.'; r4 = 3; }
  finishTable(ws4, 1, 1, r4 - 1, h4.length, [3]);
  issueFills.forEach(([r, argb]) => { ws4.getCell(r, 8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } }; });
  for (const [r] of issueFills) ws4.getCell(r, 6).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
  for (const bl of blocks4) { put(ws4, bl.start, 1, bl.name); mergeTabBlock(ws4, bl.start, bl.end); }
  styleHeader(ws4, 1, 1, h4.length);
  [24, 18, 7, 22, 52, 13, 14, 70, 14, 40].forEach((w, i) => { ws4.getColumn(i + 1).width = w; });
  if (r4 > 3 || ws4.getCell(2, 1).value !== 'Nothing needs attention — every priced row was resolved.') ws4.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: h4.length } };

  return wb;
}

/* ============================================================
   6. UI WIRING
   ============================================================ */
(function wireDecathlonClosedBook() {
  const $ = id => document.getElementById(id);
  document.getElementById('backBtnDecC').addEventListener('click', () => {
    decathlonClosedView.hidden = true;
    decathlonHubView.hidden = false;
  });

  const dz = $('dropzoneDecC'), fileInput = $('fileInputDecC'), filebar = $('filebarDecC');
  const fileNameEl = $('fileNameDecC'), clearBtn = $('clearFileDecC');
  const fabInput = $('fabricPriceFileInputDecC'), fabStatus = $('fabricPriceStatusDecC');
  const accInput = $('priceFileInputDecC'), accStatus = $('priceStatusDecC');
  const goBtn = $('processBtnDecC'), goLabel = $('processLabelDecC'), status = $('statusDecC');
  const results = $('resultsDecC');
  const dlCost = $('downloadCostBtnDecC'), dlReport = $('downloadReportBtnDecC');

  let costFile = null, fabricIdx = null, fabricInfo = '', accIdx = null, accInfo = '';
  let outCost = null, outReport = null;

  const setStatus = (msg, cls) => { status.textContent = msg; status.className = 'status' + (cls ? ' ' + cls : ''); };
  const refresh = () => { goBtn.disabled = !(costFile && (fabricIdx || accIdx)); };

  function setCostFile(file) {
    if (!file || !/\.xlsx$/i.test(file.name)) { setStatus('Please choose an .xlsx cost sheet.', 'err'); return; }
    costFile = file; fileNameEl.textContent = file.name; filebar.classList.add('show');
    results.classList.remove('show'); outCost = outReport = null; setStatus(''); refresh();
  }
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag'));
  dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('drag'); if (e.dataTransfer.files.length) setCostFile(e.dataTransfer.files[0]); });
  fileInput.addEventListener('change', e => { if (e.target.files.length) setCostFile(e.target.files[0]); });
  clearBtn.addEventListener('click', e => {
    e.stopPropagation(); costFile = null; fileInput.value = ''; filebar.classList.remove('show');
    results.classList.remove('show'); outCost = outReport = null; setStatus(''); refresh();
  });

  fabInput.addEventListener('change', async e => {
    const file = e.target.files[0]; if (!file) return;
    fabStatus.textContent = 'Parsing fabric price list…'; fabStatus.className = 'status';
    try {
      const rows = await parseDecathlonFabricPriceFile(file);
      fabricIdx = buildDecathlonFabricPriceIndex(rows);
      fabricInfo = `${file.name} (${rows.length.toLocaleString()} rows, ${rows.filter(r => r.isTrue).length.toLocaleString()} TRUE)`;
      fabStatus.textContent = 'Fabric price list loaded — ' + fabricInfo; fabStatus.className = 'status ok';
    } catch (err) {
      console.error(err); fabricIdx = null; fabricInfo = '';
      fabStatus.textContent = 'Could not read fabric price list: ' + err.message; fabStatus.className = 'status err';
    }
    refresh();
  });
  accInput.addEventListener('change', async e => {
    const file = e.target.files[0]; if (!file) return;
    accStatus.textContent = 'Parsing accessories price list…'; accStatus.className = 'status';
    try {
      const rows = await parseDecathlonPriceFile(file);
      accIdx = buildDecathlonPriceIndex(rows);
      accInfo = `${file.name} (${rows.length.toLocaleString()} rows)`;
      accStatus.textContent = 'Accessories price list loaded — ' + accInfo; accStatus.className = 'status ok';
    } catch (err) {
      console.error(err); accIdx = null; accInfo = '';
      accStatus.textContent = 'Could not read accessories price list: ' + err.message; accStatus.className = 'status err';
    }
    refresh();
  });

  function download(buffer, name) {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = name;
    document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
  }
  const baseName = () => (costFile ? costFile.name.replace(/\.xlsx$/i, '') : 'decathlon_costing');
  dlCost.addEventListener('click', () => { if (outCost) download(outCost, baseName() + '_price_updated.xlsx'); });
  dlReport.addEventListener('click', () => { if (outReport) download(outReport, baseName() + '_price_update_report.xlsx'); });

  goBtn.addEventListener('click', async () => {
    if (!costFile) return;
    if (!fabricIdx || !accIdx) {
      const miss = [];
      if (!fabricIdx) miss.push('Fabric price list not loaded — fabric rows will be left unchanged.');
      if (!accIdx) miss.push('Accessories price list not loaded — accessory rows will be left unchanged (flat-rate items still update).');
      if (!confirm(miss.join('\n\n') + '\n\nContinue anyway?')) return;
    }
    goBtn.disabled = true; goBtn.classList.add('loading'); goLabel.textContent = 'Updating prices…';
    results.classList.remove('show'); setStatus('');
    try {
      const buf = await costFile.arrayBuffer();
      const run = await dcbRunUpdate(buf, {
        fabricIdx, accIdx,
        onProgress: name => setStatus(`Updating tab "${name}"…`),
      });
      setStatus('Building change report…');
      const reportWb = dcbBuildReport(run, {
        generated: new Date().toLocaleString(),
        costSheetName: costFile.name,
        fabricInfo: fabricInfo || '(not loaded)',
        accInfo: accInfo || '(not loaded)',
      });
      outCost = run.buffer;
      outReport = await reportWb.xlsx.writeBuffer();

      results.classList.add('show');
      setStatus(`Done — ${run.results.length} tab(s) updated. Download both files below.`, 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Something went wrong: ' + err.message, 'err');
    } finally {
      goBtn.classList.remove('loading'); goLabel.textContent = 'Update Prices & Generate Report'; refresh();
    }
  });
})();
