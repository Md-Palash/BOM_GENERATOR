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
const DCB = { type: 1, desig: 2, dsm: 3, model: 4, supplier: 6, width: 7, price: 9, usd: 21, localSum: 19, marker: 22, summaryUsd: 12, localTransportVal: 9 };
const DCB_MAXCOL = 27;

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

function dcbMarkerString(ws, r) {
  const cell = ws.findCell(r, DCB.marker);
  if (!cell || cell.value === null || cell.value === undefined) return '';
  if (cell.type === ExcelJS.ValueType.Formula) return '=' + (cell.formula || '');
  return dcbCellText(ws, r, DCB.marker);
}

// "Last 7 digits of the Designation" is the item code; the first token (if
// it is a plain number) is the DSM code.
function dcbParseDesignation(des) {
  const tokens = des.trim().split(/\s+/);
  const dsm = /^\d{6,12}$/.test(tokens[0]) ? tokens[0] : '';
  const m = tokens.length > 1 ? des.trim().match(/(\d{7})$/) : null;
  return { dsm, itemCode: m ? m[1] : '' };
}

function dcbApplyPricing(ws, r, pricing) {
  ws.getCell(r, DCB.price).value = pricing.unitPrice;

  const supCell = ws.getCell(r, DCB.supplier);
  supCell.value = (pricing.supplierText === '' || pricing.supplierText === null || pricing.supplierText === undefined) ? null : pricing.supplierText;
  const argb = supCell.font && supCell.font.color && supCell.font.color.argb;
  if (argb === 'FFCC0000') { // previous run left a red error message here
    supCell.style = Object.assign({}, supCell.style, { font: Object.assign({}, supCell.font, DECATHLON_DEFAULT_FONT) });
  }

  const vCell = ws.getCell(r, DCB.marker);
  if (pricing.needsApiMarker) vCell.value = { formula: `S${r}*0.15` };
  else if (pricing.apiMarkerText) vCell.value = pricing.apiMarkerText;
  else vCell.value = null;

  const dCell = ws.getCell(r, DCB.model);
  if (pricing.highestPriceTaken) dCell.value = 'Highest price taken';
  else if (dcbCellText(ws, r, DCB.model) === 'Highest price taken') dCell.value = null;

  if (pricing.usableWidth !== undefined) ws.getCell(r, DCB.width).value = pricing.usableWidth || null;
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
    it.old = {
      price: p,
      supplier: dcbCellText(ws, it.r, DCB.supplier).trim(),
      width: dcbCellText(ws, it.r, DCB.width).trim(),
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
    const { dsm, itemCode } = dcbParseDesignation(des);
    // Same row shape the Open Book engine prices: Type + DSM + Item code,
    // with the whole Designation kept in `component` so keyword rules
    // (e.g. interlining detection) still see the full text.
    const row = { part, dsm, model: '', component: des, itemCode, gridValue: '', items: '', qty: '', unit: '', comments: '' };
    it.part = part; it.dsm = dsm; it.itemCode = itemCode;

    const fabricStyle = sec.name === 'Fabrics' && !decathlonIsInterlining(row);
    const flat = fabricStyle ? null : decathlonFlatRateFor(row);

    let pricing = null;
    if (!flat && !dsm) {
      it.kind = 'skip'; it.status = 'Skipped – no DSM'; it.remark = 'No DSM code in Designation; row left unchanged.';
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
      dcbApplyPricing(ws, r, pricing);
      it.pricing = pricing;
      it.remark = flat ? 'Flat rate' : (pricing.apiMarkerText ? 'Fabric price, region ' + pricing.apiMarkerText : (pricing.needsApiMarker ? 'Api marker +15% (non-Bangladesh origin)' : ''));
      if (pricing.highestPriceTaken) it.remark += (it.remark ? '; ' : '') + 'Highest price taken (Item code not matched)';
    }
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

  dcbRefreshCachedResults(wb, ws);

  const readCell = (r, c) => dcbCellText(ws, r, c).trim();
  return {
    sheetName: name,
    productName: readCell(1, 2), cc: readCell(2, 2), r3: readCell(4, 2),
    layout: L, items, before, after, counts,
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
  for (const ws of [...wb.worksheets]) {
    const L = dcbAnalyzeSheet(ws);
    if (!L) { skippedTabs.push(ws.name); continue; }
    opts.onProgress && opts.onProgress(ws.name);
    results.push(dcbProcessSheet(wb, ws, L, opts));
  }
  if (!results.length) throw new Error('No Decathlon cost-sheet tabs were recognised in this workbook (looked for FABRICS … Total bom cost).');
  wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
  const buffer = await wb.xlsx.writeBuffer();
  return { buffer, results, skippedTabs };
}

/* ============================================================
   5. CHANGE REPORT (separate workbook)
   ============================================================ */
const DCB_STATUS_FILL = {
  'Price updated': 'FFDCFCE7', 'Price added': 'FFD1FAE5', 'Info updated': 'FFFEF3C7',
  'No change': null, 'Not updated': 'FFFEE2E2',
};
function dcbStatusFill(status) {
  if (status in DCB_STATUS_FILL) return DCB_STATUS_FILL[status];
  return 'FFE5E7EB'; // any "Skipped – …"
}

function dcbBuildReport(run, meta) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'BOM Generator — Decathlon Closed Book';
  const HEAD_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF111827' } };
  const HEAD_FONT = { bold: true, color: { argb: 'FFFFFFFF' } };
  const MONEY = '#,##0.0000';
  const PCT = '0.00%';

  function styleHeader(ws, rowNum, nCols) {
    const row = ws.getRow(rowNum);
    for (let c = 1; c <= nCols; c++) {
      const cell = row.getCell(c);
      cell.fill = HEAD_FILL; cell.font = HEAD_FONT;
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    }
    row.height = 30;
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
    ['All amounts', 'USD price (column U / L of the cost sheet). Δ = After − Before.'],
  ];
  info.forEach((p, i) => {
    ws1.getCell(3 + i, 1).value = p[0]; ws1.getCell(3 + i, 1).font = { bold: true };
    ws1.getCell(3 + i, 2).value = p[1];
  });
  let r0 = 3 + info.length + 1;
  const sumHead = ['Tab', 'Product', 'CC', 'Items scanned', 'Price updated', 'Price added', 'Info updated', 'No change', 'Not updated', 'Skipped',
    'Total BOM cost – before', 'Total BOM cost – after', 'Δ BOM', 'Δ BOM %', 'Vendor price – before', 'Vendor price – after', 'Δ Vendor', 'Δ Vendor %'];
  sumHead.forEach((h, i) => { ws1.getCell(r0, i + 1).value = h; });
  styleHeader(ws1, r0, sumHead.length);
  run.results.forEach((s, i) => {
    const r = r0 + 1 + i;
    const c = s.counts;
    [s.sheetName, s.productName, s.cc, c.total, c.priceUpdated, c.priceAdded, c.infoUpdated, c.noChange, c.notUpdated, c.skipped,
      s.before.bom, s.after.bom].forEach((v, j) => { ws1.getCell(r, j + 1).value = v; });
    putDelta(ws1, r, 11, 12, 13, 14, s.before.bom, s.after.bom);
    ws1.getCell(r, 15).value = s.before.vendor; ws1.getCell(r, 16).value = s.after.vendor;
    putDelta(ws1, r, 15, 16, 17, 18, s.before.vendor, s.after.vendor);
    [11, 12, 15, 16].forEach(n => { ws1.getCell(r, n).numFmt = MONEY; });
    if (c.notUpdated) ws1.getCell(r, 9).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
  });
  const tr = r0 + 1 + run.results.length;
  ws1.getCell(tr, 1).value = 'TOTAL (all tabs)'; ws1.getCell(tr, 1).font = { bold: true };
  for (let n = 4; n <= 10; n++) {
    const col = ws1.getColumn(n).letter;
    ws1.getCell(tr, n).value = { formula: `SUM(${col}${r0 + 1}:${col}${tr - 1})`, result: run.results.reduce((s, x) => s + [x.counts.total, x.counts.priceUpdated, x.counts.priceAdded, x.counts.infoUpdated, x.counts.noChange, x.counts.notUpdated, x.counts.skipped][n - 4], 0) };
    ws1.getCell(tr, n).font = { bold: true };
  }
  if (run.skippedTabs.length) {
    ws1.getCell(tr + 2, 1).value = 'Tabs not treated as cost sheets (left untouched):';
    ws1.getCell(tr + 2, 1).font = { bold: true };
    ws1.getCell(tr + 2, 2).value = run.skippedTabs.join(', ');
  }
  const widths = [26, 28, 12, 12, 12, 12, 12, 12, 12, 12, 18, 18, 14, 12, 18, 18, 14, 12];
  widths.forEach((w, i) => { ws1.getColumn(i + 1).width = w; });
  ws1.getColumn(1).width = 30;
  ws1.views = [{ state: 'frozen', ySplit: r0 }];

  /* ---------- Section Costs ---------- */
  const ws2 = wb.addWorksheet('Section Costs');
  const h2 = ['Tab', 'Line', 'Before (USD)', 'After (USD)', 'Δ (USD)', 'Δ %'];
  h2.forEach((h, i) => { ws2.getCell(1, i + 1).value = h; });
  styleHeader(ws2, 1, h2.length);
  let r2 = 2;
  for (const s of run.results) {
    for (const ln of lineDefs(s)) {
      const b = ln.get(s.before), a = ln.get(s.after);
      ws2.getCell(r2, 1).value = s.sheetName;
      ws2.getCell(r2, 2).value = ln.label;
      ws2.getCell(r2, 3).value = b; ws2.getCell(r2, 4).value = a;
      ws2.getCell(r2, 3).numFmt = MONEY; ws2.getCell(r2, 4).numFmt = MONEY;
      putDelta(ws2, r2, 3, 4, 5, 6, b, a);
      if (ln.kind === 'total') {
        for (let c = 1; c <= 6; c++) { ws2.getCell(r2, c).font = { bold: true }; ws2.getCell(r2, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFF4FF' } }; }
      }
      r2++;
    }
  }
  [26, 36, 18, 18, 16, 12].forEach((w, i) => { ws2.getColumn(i + 1).width = w; });
  ws2.views = [{ state: 'frozen', ySplit: 1 }];
  ws2.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 6 } };

  /* ---------- Item Changes ---------- */
  const ws3 = wb.addWorksheet('Item Changes');
  const h3 = ['Tab', 'Section', 'Row', 'Type', 'Designation', 'DSM', 'Item code (last 7)', 'Old price', 'New price', 'Δ price', 'Δ price %',
    'Old supplier', 'New supplier', 'Old usable width', 'New usable width', 'Old Api marker', 'New Api marker',
    'Row cost before (USD)', 'Row cost after (USD)', 'Cost impact (USD)', 'Status', 'Remark'];
  h3.forEach((h, i) => { ws3.getCell(1, i + 1).value = h; });
  styleHeader(ws3, 1, h3.length);
  let r3 = 2;
  for (const s of run.results) {
    for (const it of s.items) {
      const vals = [s.sheetName, it.sec.label, it.r, it.part, it.des, it.dsm, it.itemCode, it.old.price, it.new.price, null, null,
        it.old.supplier, it.new.supplier, it.old.width, it.new.width, it.old.marker, it.new.marker,
        it.costBefore, it.costAfter, null, it.status, it.remark || ''];
      vals.forEach((v, j) => { if (v !== null && v !== undefined && v !== '') ws3.getCell(r3, j + 1).value = v; });
      ['H', 'I', 'J'].forEach(L => { ws3.getCell(`${L}${r3}`).numFmt = '0.0000'; });
      ws3.getCell(r3, 11).numFmt = PCT;
      [18, 19, 20].forEach(n => { ws3.getCell(r3, n).numFmt = MONEY; });
      if (typeof it.old.price === 'number' && typeof it.new.price === 'number') {
        ws3.getCell(r3, 10).value = { formula: `I${r3}-H${r3}`, result: it.new.price - it.old.price };
        ws3.getCell(r3, 11).value = { formula: `IF(H${r3}=0,"",J${r3}/H${r3})`, result: it.old.price === 0 ? '' : (it.new.price - it.old.price) / it.old.price };
      }
      if (typeof it.costBefore === 'number' && typeof it.costAfter === 'number') {
        ws3.getCell(r3, 20).value = { formula: `S${r3}-R${r3}`, result: it.costAfter - it.costBefore };
      }
      const fill = dcbStatusFill(it.status);
      if (fill) ws3.getCell(r3, 21).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      r3++;
    }
  }
  [24, 18, 7, 22, 52, 13, 14, 11, 11, 11, 10, 20, 20, 16, 16, 16, 16, 16, 16, 16, 24, 60].forEach((w, i) => { ws3.getColumn(i + 1).width = w; });
  ws3.views = [{ state: 'frozen', ySplit: 1, xSplit: 5 }];
  ws3.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: h3.length } };

  /* ---------- Needs Attention ---------- */
  const ws4 = wb.addWorksheet('Needs Attention');
  const h4 = ['Tab', 'Section', 'Row', 'Type', 'Designation', 'DSM', 'Item code (last 7)', 'Issue', 'Price kept in sheet', 'Supplier cell now reads'];
  h4.forEach((h, i) => { ws4.getCell(1, i + 1).value = h; });
  styleHeader(ws4, 1, h4.length);
  let r4 = 2;
  // Packaging sections commonly hold fixed template rows (no DSM) - not an issue.
  const QUIET = new Set(['Sales Packaging', 'Transport Packaging']);
  for (const s of run.results) {
    for (const it of s.items) {
      const isIssue = it.kind === 'fail'
        || (it.kind === 'skip' && !(it.status === 'Skipped – no DSM' && QUIET.has(it.sec.name)));
      if (!isIssue) continue;
      [s.sheetName, it.sec.label, it.r, it.part, it.des, it.dsm, it.itemCode, it.remark || it.status, it.old.price, it.old.supplier].forEach((v, j) => {
        if (v !== null && v !== undefined && v !== '') ws4.getCell(r4, j + 1).value = v;
      });
      ws4.getCell(r4, 9).numFmt = '0.0000';
      ws4.getCell(r4, 8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: it.kind === 'fail' ? 'FFFEE2E2' : 'FFE5E7EB' } };
      r4++;
    }
  }
  if (r4 === 2) { ws4.getCell(2, 1).value = 'Nothing needs attention — every priced row was resolved.'; }
  [24, 18, 7, 22, 52, 13, 14, 70, 14, 40].forEach((w, i) => { ws4.getColumn(i + 1).width = w; });
  ws4.views = [{ state: 'frozen', ySplit: 1 }];
  if (r4 > 2) ws4.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: h4.length } };

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
  const results = $('resultsDecC'), summaryEl = $('summaryDecC');
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

  const money = v => (typeof v === 'number' ? v.toFixed(4) : 'n/a');

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

      const lines = run.results.map(s => {
        const c = s.counts;
        return `<div style="margin-top:8px;"><b>${escapeHtmlDecC(s.sheetName)}</b> — ${c.total} items: ${c.priceUpdated + c.priceAdded} priced, ${c.infoUpdated} info-only, ${c.noChange} unchanged, ${c.notUpdated} not updated, ${c.skipped} skipped. ` +
          `BOM ${money(s.before.bom)} → ${money(s.after.bom)}; Vendor price ${money(s.before.vendor)} → ${money(s.after.vendor)}</div>`;
      });
      summaryEl.innerHTML = lines.join('') + (run.skippedTabs.length ? `<div style="margin-top:8px;color:var(--text-secondary);">Left untouched (not cost sheets): ${escapeHtmlDecC(run.skippedTabs.join(', '))}</div>` : '');
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

function escapeHtmlDecC(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}
