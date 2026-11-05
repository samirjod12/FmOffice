/* FmOffice Sheets — calculation engine (no dependencies). Works in browser and Node. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FmEngine = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ---------- errors & helpers ---------- */
  class FErr { constructor(code) { this.code = code; } toString() { return this.code; } }
  const E = {
    DIV0: new FErr('#DIV/0!'), VALUE: new FErr('#VALUE!'), REF: new FErr('#REF!'),
    NAME: new FErr('#NAME?'), NA: new FErr('#N/A'), NUM: new FErr('#NUM!'), CIRC: new FErr('#CIRC!'), SPILL: new FErr('#SPILL!')
  };
  const ERR_BY_CODE = {}; Object.values(E).forEach(e => ERR_BY_CODE[e.code] = e);

  class Rng {
    constructor(rows, r1, c1, sheet) { this.rows = rows; this.r1 = r1 || 0; this.c1 = c1 || 0; this.sheet = sheet; }
    get h() { return this.rows.length; }
    get w() { return this.rows.length ? this.rows[0].length : 0; }
  }

  const MAXC = 16384, MAXR = 1048576;
  function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; }
  function numToCol(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function addrToStr(r, c) { return numToCol(c) + (r + 1); }
  function parseAddr(a) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(a);
    if (!m) return null;
    return { ac: !!m[1], c: colToNum(m[2]), ar: !!m[3], r: parseInt(m[4], 10) - 1 };
  }

  /* ---------- dates ---------- */
  const DAY = 86400000;
  function serialToDate(s) { return new Date(Math.round((s - 25569) * DAY)); }
  function dateToSerial(d) { return d.getTime() / DAY + 25569; }
  function ymdToSerial(y, m, d) { return Date.UTC(y, m - 1, d) / DAY + 25569; }
  function nowSerial() { const n = new Date(); return (n.getTime() - n.getTimezoneOffset() * 60000) / DAY + 25569; }
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  /* ---------- number formatting ---------- */
  function roundHalfUp(x, d) { const f = Math.pow(10, d); return Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9) / f; }
  function addCommas(s) { return s.replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
  function generalFmt(v) {
    if (v === 0) return '0';
    const a = Math.abs(v);
    if (a >= 1e11 || a < 1e-9) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2');
    let s = String(parseFloat(v.toPrecision(10)));
    if (s.includes('e')) s = String(parseFloat(v.toFixed(9)));
    return s;
  }
  function isDateFmt(f) { return /(^|[^\\"])[ymdhs]/i.test(f.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/E[+-]/i, '')) && !/^0|#/.test(f.replace(/"[^"]*"/g, '')); }
  function formatDate(v, f) {
    const d = serialToDate(v);
    const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), W = d.getUTCDay();
    let H = d.getUTCHours(), Mi = d.getUTCMinutes(), S = d.getUTCSeconds();
    const ampm = /AM\/PM|A\/P/i.test(f);
    const toks = f.match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|\.0+|[^]/gi) || [];
    let out = '';
    for (let i = 0; i < toks.length; i++) {
      const t = toks[i], l = t.toLowerCase();
      if (t[0] === '"') out += t.slice(1, -1);
      else if (t[0] === '\\') out += t[1];
      else if (l === 'yyyy') out += String(Y).padStart(4, '0');
      else if (l === 'yy') out += String(Y % 100).padStart(2, '0');
      else if (l === 'mmmmm') out += MONTHS[M][0];
      else if (l === 'mmmm') out += MONTHS[M];
      else if (l === 'mmm') out += MONTHS[M].slice(0, 3);
      else if (l === 'mm' || l === 'm') {
        let isMin = false;
        for (let j = i - 1; j >= 0; j--) { const p = toks[j].toLowerCase(); if (p === 'h' || p === 'hh') { isMin = true; break; } if (/^(yyyy|yy|d|dd|ddd|dddd|mmm|mmmm)$/.test(p)) break; }
        if (!isMin) for (let j = i + 1; j < toks.length; j++) { const p = toks[j].toLowerCase(); if (p === 's' || p === 'ss') { isMin = true; break; } if (/^(yyyy|yy|d|dd|ddd|dddd|mmm|mmmm)$/.test(p)) break; }
        const val = isMin ? Mi : M + 1;
        out += l === 'mm' ? String(val).padStart(2, '0') : String(val);
      }
      else if (l === 'dddd') out += DAYS[W];
      else if (l === 'ddd') out += DAYS[W].slice(0, 3);
      else if (l === 'dd') out += String(D).padStart(2, '0');
      else if (l === 'd') out += String(D);
      else if (l === 'hh' || l === 'h') { let h = ampm ? (H % 12 || 12) : H; out += l === 'hh' ? String(h).padStart(2, '0') : String(h); }
      else if (l === 'ss') out += String(S).padStart(2, '0');
      else if (l === 's') out += String(S);
      else if (l === 'am/pm') out += H < 12 ? 'AM' : 'PM';
      else if (l === 'a/p') out += H < 12 ? 'A' : 'P';
      else out += t;
    }
    return out;
  }
  const BUILTIN_FMT = {
    'General': null, 'Number': '0.00', 'Currency': '$#,##0.00', 'Accounting': '_($* #,##0.00_)', 'Percent': '0.00%',
    'Short Date': 'm/d/yyyy', 'Long Date': 'dddd, mmmm d, yyyy', 'Time': 'h:mm:ss AM/PM', 'Scientific': '0.00E+00', 'Text': '@'
  };
  function formatNumber(v, f) {
    const secs = f.split(';');
    let sec = secs[0], neg = false;
    if (v < 0 && secs.length > 1) { sec = secs[1]; neg = true; v = -v; }
    else if (v === 0 && secs.length > 2) sec = secs[2];
    const clean = sec.replace(/_./g, '').replace(/\*./g, '').replace(/\[[^\]]*\]/g, '');
    if (/0\.?0*E[+-]0+/i.test(clean)) {
      const m = /0\.(0+)E[+-]0+/i.exec(clean); const dec = m ? m[1].length : 0;
      let s = Math.abs(v).toExponential(dec).toUpperCase().replace(/E([+-])(\d)$/, 'E$10$2');
      return (v < 0 && !neg ? '-' : '') + s;
    }
    let pct = (clean.match(/%/g) || []).length;
    let x = Math.abs(v) * Math.pow(100, pct ? 1 : 0);
    const first = clean.search(/[0#?]/), last = (() => { let k = -1; for (let i = 0; i < clean.length; i++) if (/[0#?]/.test(clean[i])) k = i; return k; })();
    const unq = s => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
    if (first < 0) return (v < 0 && !neg ? '-' : '') + unq(clean);
    const prefix = unq(clean.slice(0, first)), suffix = unq(clean.slice(last + 1));
    const core = clean.slice(first, last + 1);
    const dot = core.indexOf('.');
    const intPat = dot < 0 ? core : core.slice(0, dot), decPat = dot < 0 ? '' : core.slice(dot + 1);
    const dec = (decPat.match(/[0#?]/g) || []).length;
    const useComma = intPat.includes(',');
    let trailingCommas = 0; { const m = /(,+)$/.exec(intPat); if (m && !/[0#?]/.test(intPat.slice(-m[1].length - 1, -m[1].length) || '')) trailingCommas = 0; }
    x = roundHalfUp(x, dec);
    let s = x.toFixed(dec);
    let [ip, dp = ''] = s.split('.');
    const minInt = (intPat.replace(/,/g, '').match(/0/g) || []).length;
    ip = ip.padStart(minInt, '0');
    if (minInt === 0 && ip === '0') ip = '';
    if (useComma) ip = addCommas(ip);
    const optDec = (decPat.match(/#/g) || []).length;
    if (optDec) dp = dp.replace(new RegExp('0{1,' + optDec + '}$'), '');
    let res = ip + (dp ? '.' + dp : '');
    if (res === '') res = '0';
    return (v < 0 && !neg ? '-' : '') + prefix + res + suffix;
  }
  function formatValue(v, nf) {
    if (v === null || v === undefined) return '';
    if (v instanceof FErr) return v.code;
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    if (typeof v !== 'number' || !isFinite(v)) return v instanceof Object ? '' : String(v);
    if (!nf || nf === 'General') return generalFmt(v);
    if (nf === '@') return generalFmt(v);
    if (isDateFmt(nf.split(';')[0])) return formatDate(v, nf.split(';')[0]);
    return formatNumber(v, nf);
  }

  /* ---------- tokenizer ---------- */
  const SHEET = "(?:'(?:[^']|'')+'|[A-Za-z_][\\w.]*)!";
  const RE_RANGE = new RegExp('^(?:' + SHEET + ')?\\$?[A-Za-z]{1,3}\\$?\\d+(?::\\$?[A-Za-z]{1,3}\\$?\\d+)?(?![\\w(])');
  const RE_COLRANGE = new RegExp('^(?:' + SHEET + ')?\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}(?![\\w(])');
  const RE_ROWRANGE = new RegExp('^(?:' + SHEET + ')?\\$?\\d+:\\$?\\d+(?![\\w(])');
  const RE_NUM = /^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/;
  const RE_FN = /^[A-Za-z_][\w.]*(?=\()/;
  const RE_ID = /^[A-Za-z_][\w.]*/;
  const RE_ERR = /^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!|SPILL!)/;

  function tokenize(src) {
    const t = []; let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      const rest = src.slice(i);
      let m;
      if (ch === '"') {
        let j = i + 1, s = '';
        while (j < src.length) { if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; } s += src[j++]; }
        if (j >= src.length) throw E.NAME;
        t.push({ t: 'str', v: s, s: i, e: j + 1 }); i = j + 1; continue;
      }
      if ((m = RE_ERR.exec(rest))) { t.push({ t: 'err', v: m[0], s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_RANGE.exec(rest))) { t.push({ t: 'ref', v: m[0], s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_COLRANGE.exec(rest))) { t.push({ t: 'colr', v: m[0], s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_ROWRANGE.exec(rest))) { t.push({ t: 'rowr', v: m[0], s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_NUM.exec(rest))) { t.push({ t: 'num', v: parseFloat(m[0]), s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_FN.exec(rest))) { t.push({ t: 'fn', v: m[0].toUpperCase(), s: i, e: i + m[0].length }); i += m[0].length; continue; }
      if ((m = RE_ID.exec(rest))) { t.push({ t: 'id', v: m[0], s: i, e: i + m[0].length }); i += m[0].length; continue; }
      const two = src.substr(i, 2);
      if (two === '<=' || two === '>=' || two === '<>') { t.push({ t: 'op', v: two, s: i, e: i + 2 }); i += 2; continue; }
      if ('+-*/^&=<>(),%;:{}'.includes(ch)) { t.push({ t: 'op', v: ch, s: i, e: i + 1 }); i++; continue; }
      throw E.NAME;
    }
    return t;
  }

  /* ---------- parser ---------- */
  function parse(src) {
    const toks = tokenize(src); let p = 0;
    const peek = () => toks[p], next = () => toks[p++];
    const isOp = v => peek() && peek().t === 'op' && peek().v === v;
    function expr() { return cmp(); }
    function cmp() {
      let l = cat();
      while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) { const op = next().v; l = { k: 'bin', op, l, r: cat() }; }
      return l;
    }
    function cat() { let l = add(); while (isOp('&')) { next(); l = { k: 'bin', op: '&', l, r: add() }; } return l; }
    function add() { let l = mul(); while (isOp('+') || isOp('-')) { const op = next().v; l = { k: 'bin', op, l, r: mul() }; } return l; }
    function mul() { let l = pow(); while (isOp('*') || isOp('/')) { const op = next().v; l = { k: 'bin', op, l, r: pow() }; } return l; }
    function pow() { let l = unary(); while (isOp('^')) { next(); l = { k: 'bin', op: '^', l, r: unary() }; } return l; }
    function unary() {
      if (isOp('-')) { next(); return { k: 'un', op: '-', a: unary() }; }
      if (isOp('+')) { next(); return unary(); }
      return postfix();
    }
    function postfix() { let a = primary(); while (isOp('%')) { next(); a = { k: 'pct', a }; } return a; }
    function primary() {
      const t = next();
      if (!t) throw E.NAME;
      if (t.t === 'num') return { k: 'num', v: t.v };
      if (t.t === 'str') return { k: 'str', v: t.v };
      if (t.t === 'err') return { k: 'err', v: ERR_BY_CODE[t.v] || E.NAME };
      if (t.t === 'ref') return { k: 'ref', ...parseRefToken(t.v) };
      if (t.t === 'colr') return { k: 'colr', ...parseRefToken(t.v) };
      if (t.t === 'rowr') return { k: 'rowr', ...parseRefToken(t.v) };
      if (t.t === 'id') {
        const u = t.v.toUpperCase();
        if (u === 'TRUE') return { k: 'bool', v: true };
        if (u === 'FALSE') return { k: 'bool', v: false };
        return { k: 'name', v: t.v };
      }
      if (t.t === 'fn') {
        if (!isOp('(')) throw E.NAME; next();
        const args = [];
        if (isOp(')')) { next(); return { k: 'fn', name: t.v, args }; }
        for (; ;) {
          if (isOp(',') || isOp(';')) { args.push({ k: 'blank' }); }
          else if (isOp(')')) { args.push({ k: 'blank' }); }
          else args.push(expr());
          if (isOp(',') || isOp(';')) { next(); continue; }
          if (isOp(')')) { next(); break; }
          throw E.NAME;
        }
        return { k: 'fn', name: t.v, args };
      }
      if (t.t === 'op' && t.v === '(') { const e = expr(); if (!isOp(')')) throw E.NAME; next(); return { k: 'paren', a: e }; }
      throw E.NAME;
    }
    const ast = expr();
    if (p < toks.length) throw E.NAME;
    return ast;
  }

  function parseRefToken(tok) {
    let sheet = null, body = tok;
    const bang = tok.lastIndexOf('!');
    if (bang >= 0) { sheet = tok.slice(0, bang); body = tok.slice(bang + 1); if (sheet[0] === "'") sheet = sheet.slice(1, -1).replace(/''/g, "'"); }
    const parts = body.split(':');
    const one = s => {
      let m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(s);
      if (m) return { ac: !!m[1], c: colToNum(m[2]), ar: !!m[3], r: parseInt(m[4], 10) - 1 };
      m = /^(\$?)([A-Za-z]{1,3})$/.exec(s);
      if (m) return { ac: !!m[1], c: colToNum(m[2]), col: true };
      m = /^(\$?)(\d+)$/.exec(s);
      if (m) return { ar: !!m[1], r: parseInt(m[2], 10) - 1, row: true };
      throw E.REF;
    };
    const a = one(parts[0]), b = parts[1] ? one(parts[1]) : null;
    return { sheet, a, b };
  }
  function quoteSheet(n) { return /^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; }
  function refPartToStr(p) {
    if (p.col) return (p.ac ? '$' : '') + numToCol(p.c);
    if (p.row) return (p.ar ? '$' : '') + (p.r + 1);
    return (p.ac ? '$' : '') + numToCol(p.c) + (p.ar ? '$' : '') + (p.r + 1);
  }
  function refToStr(ref) {
    return (ref.sheet != null ? quoteSheet(ref.sheet) + '!' : '') + refPartToStr(ref.a) + (ref.b ? ':' + refPartToStr(ref.b) : '');
  }

  /* rewrite every reference token in a formula; fn(refObj, tokenKind) returns new ref obj or null for #REF! */
  function rewriteRefs(text, fn) {
    if (text[0] !== '=') return text;
    let toks;
    try { toks = tokenize(text.slice(1)); } catch (e) { return text; }
    let out = '=', pos = 0; const body = text.slice(1);
    for (const t of toks) {
      if (t.t === 'ref' || t.t === 'colr' || t.t === 'rowr') {
        out += body.slice(pos, t.s);
        let ref; try { ref = parseRefToken(t.v); } catch (e) { out += t.v; pos = t.e; continue; }
        const nr = fn(ref);
        out += nr ? refToStr(nr) : '#REF!';
        pos = t.e;
      }
    }
    return out + body.slice(pos);
  }
  function shiftFormula(text, dr, dc) {
    return rewriteRefs(text, ref => {
      const sh = p => {
        if (!p) return p; const q = { ...p };
        if (q.c !== undefined && !q.row && !q.ac) q.c += dc;
        if (q.r !== undefined && !q.col && !q.ar) q.r += dr;
        if ((q.c !== undefined && (q.c < 0 || q.c >= MAXC)) || (q.r !== undefined && (q.r < 0 || q.r >= MAXR))) return null;
        return q;
      };
      const a = sh(ref.a), b = ref.b ? sh(ref.b) : null;
      if (!a || (ref.b && !b)) return null;
      return { sheet: ref.sheet, a, b };
    });
  }

  /* ---------- coercion ---------- */
  function chk(v) { if (v instanceof FErr) throw v; return v; }
  function scalar(v) { if (v instanceof Rng) return v.rows.length && v.rows[0].length ? v.rows[0][0] : null; return v; }
  function toNum(v) {
    v = scalar(v); chk(v);
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const s = String(v).trim(); const lit = parseLiteral(s);
    if (typeof lit.v === 'number') return lit.v;
    throw E.VALUE;
  }
  function toStr(v) { v = scalar(v); chk(v); if (v === null || v === undefined) return ''; if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'; if (typeof v === 'number') return generalFmt(v); return String(v); }
  function toBool(v) {
    v = scalar(v); chk(v);
    if (v === null || v === undefined || v === '') return false;
    if (typeof v === 'boolean') return v; if (typeof v === 'number') return v !== 0;
    const u = String(v).toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; throw E.VALUE;
  }
  function parseDateText(s) {
    let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
    if (m) return { v: ymdToSerial(+m[1], +m[2], +m[3]), fmt: 'yyyy-mm-dd' };
    m = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(s);
    if (m) { let y = +m[3]; if (y < 100) y += y < 30 ? 2000 : 1900; return { v: ymdToSerial(y, +m[1], +m[2]), fmt: 'm/d/yyyy' }; }
    m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$/i.exec(s);
    if (m) {
      let h = +m[1]; const ap = m[4] && m[4].toUpperCase();
      if (ap === 'PM' && h < 12) h += 12; if (ap === 'AM' && h === 12) h = 0;
      return { v: (h * 3600 + (+m[2]) * 60 + (+(m[3] || 0))) / 86400, fmt: m[3] ? 'h:mm:ss' : 'h:mm' };
    }
    return null;
  }
  /* what a typed (non-formula) string means */
  function parseLiteral(text) {
    if (typeof text !== 'string') return { v: text };
    if (text === '') return { v: null };
    if (text[0] === "'") return { v: text.slice(1) };
    const s = text.trim();
    if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) return { v: parseFloat(s) };
    if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) return { v: parseFloat(s.replace(/,/g, '')), fmt: s.includes('.') ? '#,##0.00' : '#,##0' };
    if (/^[+-]?(\d+\.?\d*|\.\d+)%$/.test(s)) { const d = (s.split('.')[1] || '').replace('%', '').length; return { v: parseFloat(s) / 100, fmt: d ? '0.' + '0'.repeat(d) + '%' : '0%' }; }
    let m = /^([+-]?)\$\s?([\d,]+(?:\.\d+)?)$/.exec(s);
    if (m) return { v: (m[1] === '-' ? -1 : 1) * parseFloat(m[2].replace(/,/g, '')), fmt: '$#,##0.00' };
    const u = s.toUpperCase();
    if (u === 'TRUE') return { v: true }; if (u === 'FALSE') return { v: false };
    if (/^#(DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!)$/.test(u)) return { v: ERR_BY_CODE[u] };
    const d = parseDateText(s); if (d) return d;
    return { v: text };
  }

  function cmpVals(a, b) {
    a = scalar(a); b = scalar(b); chk(a); chk(b);
    if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const rk = v => typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : 3;
    if (rk(a) !== rk(b)) return rk(a) - rk(b);
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  /* wildcard + operator criteria for SUMIF/COUNTIF */
  function makeCriteria(c) {
    c = scalar(c); chk(c);
    let op = '=', val = c;
    if (typeof c === 'string') {
      const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c);
      op = m[1] || '='; val = m[2];
      const lit = parseLiteral(val); if (typeof lit.v === 'number' || typeof lit.v === 'boolean') val = lit.v;
    }
    return v => {
      v = scalar(v);
      if (v instanceof FErr) return false;
      if (typeof val === 'string') {
        if (v === null || v === undefined) v = '';
        if (typeof v !== 'string') { if (op === '=' || op === '<>') { return op === '<>'; } return false; }
        const hasW = /[*?]/.test(val);
        let ok;
        if (op === '=' || op === '<>') {
          if (hasW) { const re = new RegExp('^' + val.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i'); ok = re.test(v); }
          else ok = v.toLowerCase() === val.toLowerCase();
          return op === '=' ? ok : !ok;
        }
        const x = v.toLowerCase(), y = val.toLowerCase();
        return op === '>' ? x > y : op === '<' ? x < y : op === '>=' ? x >= y : x <= y;
      }
      if (typeof v !== typeof val) { if (v === null && typeof val === 'number' && val === 0 && op === '=') return false; return op === '<>'; }
      const d = v < val ? -1 : v > val ? 1 : 0;
      return op === '=' ? d === 0 : op === '<>' ? d !== 0 : op === '>' ? d > 0 : op === '<' ? d < 0 : op === '>=' ? d >= 0 : d <= 0;
    };
  }

  /* ---------- functions ---------- */
  const FN = {}; const LAZY = {};
  const def = (names, f) => names.split(' ').forEach(n => FN[n] = f);
  function nums(args, opts) {
    const out = [];
    for (const a of args) {
      if (a instanceof Rng) { for (const row of a.rows) for (const v of row) { chk(v); if (typeof v === 'number') out.push(v); else if (opts && opts.text0 && typeof v === 'string') out.push(0); } }
      else { chk(a); if (a === null) continue; out.push(toNum(a)); }
    }
    return out;
  }
  function allVals(args) { const o = []; for (const a of args) { if (a instanceof Rng) for (const r of a.rows) for (const v of r) o.push(v); else o.push(a); } return o; }
  const need = (args, n) => { if (args.length < n) throw E.VALUE; };

  def('SUM', a => nums(a).reduce((x, y) => x + y, 0));
  def('PRODUCT', a => nums(a).reduce((x, y) => x * y, 1));
  def('AVERAGE', a => { const n = nums(a); if (!n.length) throw E.DIV0; return n.reduce((x, y) => x + y, 0) / n.length; });
  def('MIN', a => { const n = nums(a); return n.length ? Math.min(...n) : 0; });
  def('MAX', a => { const n = nums(a); return n.length ? Math.max(...n) : 0; });
  def('COUNT', a => { let c = 0; for (const v of allVals(a)) if (typeof v === 'number') c++; return c; });
  def('COUNTA', a => allVals(a).filter(v => v !== null && v !== undefined && v !== '').length);
  def('COUNTBLANK', a => allVals(a).filter(v => v === null || v === undefined || v === '').length);
  def('MEDIAN', a => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw E.NUM; const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; });
  def('MODE MODE.SNGL', a => { const n = nums(a); const c = new Map(); let best = null, bc = 1; for (const x of n) { const k = (c.get(x) || 0) + 1; c.set(x, k); if (k > bc) { bc = k; best = x; } } if (best === null) throw E.NA; return best; });
  const variance = (n, pop) => { const k = n.length - (pop ? 0 : 1); if (k <= 0) throw E.DIV0; const m = n.reduce((x, y) => x + y, 0) / n.length; return n.reduce((s, x) => s + (x - m) * (x - m), 0) / k; };
  def('VAR VAR.S', a => variance(nums(a), false)); def('VARP VAR.P', a => variance(nums(a), true));
  def('STDEV STDEV.S', a => Math.sqrt(variance(nums(a), false))); def('STDEVP STDEV.P', a => Math.sqrt(variance(nums(a), true)));
  def('LARGE', a => { const n = nums([a[0]]).sort((x, y) => y - x), k = toNum(a[1]); if (k < 1 || k > n.length) throw E.NUM; return n[k - 1]; });
  def('SMALL', a => { const n = nums([a[0]]).sort((x, y) => x - y), k = toNum(a[1]); if (k < 1 || k > n.length) throw E.NUM; return n[k - 1]; });
  def('RANK RANK.EQ', a => { const x = toNum(a[0]), n = nums([a[1]]), asc = a[2] !== undefined && toNum(a[2]) !== 0; if (!n.includes(x)) throw E.NA; return 1 + n.filter(v => asc ? v < x : v > x).length; });
  def('SUMPRODUCT', a => { need(a, 1); const rs = a.map(x => x instanceof Rng ? x : new Rng([[x]])); const h = rs[0].h, w = rs[0].w; let s = 0; for (let i = 0; i < h; i++) for (let j = 0; j < w; j++) { let p = 1; for (const r of rs) { if (r.h !== h || r.w !== w) throw E.VALUE; const v = r.rows[i][j]; chk(v); p *= typeof v === 'number' ? v : (typeof v === 'boolean' ? (v ? 1 : 0) : 0); } s += p; } return s; });

  def('ABS', a => Math.abs(toNum(a[0]))); def('SQRT', a => { const x = toNum(a[0]); if (x < 0) throw E.NUM; return Math.sqrt(x); });
  def('POWER', a => { const r = Math.pow(toNum(a[0]), toNum(a[1])); if (!isFinite(r)) throw E.NUM; return r; });
  def('MOD', a => { const x = toNum(a[0]), y = toNum(a[1]); if (y === 0) throw E.DIV0; return x - y * Math.floor(x / y); });
  def('INT', a => Math.floor(toNum(a[0]))); def('SIGN', a => Math.sign(toNum(a[0])));
  def('PI', () => Math.PI); def('EXP', a => Math.exp(toNum(a[0])));
  def('LN', a => { const x = toNum(a[0]); if (x <= 0) throw E.NUM; return Math.log(x); });
  def('LOG', a => { const x = toNum(a[0]), b = a.length > 1 ? toNum(a[1]) : 10; if (x <= 0 || b <= 0 || b === 1) throw E.NUM; return Math.log(x) / Math.log(b); });
  def('LOG10', a => { const x = toNum(a[0]); if (x <= 0) throw E.NUM; return Math.log10(x); });
  def('SIN', a => Math.sin(toNum(a[0]))); def('COS', a => Math.cos(toNum(a[0]))); def('TAN', a => Math.tan(toNum(a[0])));
  def('ASIN', a => Math.asin(toNum(a[0]))); def('ACOS', a => Math.acos(toNum(a[0]))); def('ATAN', a => Math.atan(toNum(a[0]))); def('ATAN2', a => Math.atan2(toNum(a[1]), toNum(a[0])));
  def('DEGREES', a => toNum(a[0]) * 180 / Math.PI); def('RADIANS', a => toNum(a[0]) * Math.PI / 180);
  def('ROUND', a => roundHalfUp(toNum(a[0]), a.length > 1 ? toNum(a[1]) : 0));
  def('ROUNDUP', a => { const d = a.length > 1 ? toNum(a[1]) : 0, f = Math.pow(10, d), x = toNum(a[0]); return Math.sign(x) * Math.ceil(Math.abs(x) * f - 1e-9) / f; });
  def('ROUNDDOWN TRUNC', a => { const d = a.length > 1 ? toNum(a[1]) : 0, f = Math.pow(10, d), x = toNum(a[0]); return Math.sign(x) * Math.floor(Math.abs(x) * f + 1e-9) / f; });
  def('CEILING', a => { const x = toNum(a[0]), s = a.length > 1 ? toNum(a[1]) : 1; return s === 0 ? 0 : Math.ceil(x / s) * s; });
  def('FLOOR', a => { const x = toNum(a[0]), s = a.length > 1 ? toNum(a[1]) : 1; if (s === 0) throw E.DIV0; return Math.floor(x / s) * s; });
  def('RAND', () => Math.random()); def('RANDBETWEEN', a => { const lo = Math.ceil(toNum(a[0])), hi = Math.floor(toNum(a[1])); return lo + Math.floor(Math.random() * (hi - lo + 1)); });
  def('FACT', a => { const n = Math.floor(toNum(a[0])); if (n < 0) throw E.NUM; let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; });
  def('GCD', a => nums(a).reduce((x, y) => { x = Math.abs(x); y = Math.abs(y); while (y) [x, y] = [y, x % y]; return x; }));
  def('LCM', a => nums(a).reduce((x, y) => { const g = ((p, q) => { while (q) [p, q] = [q, p % q]; return p; })(Math.abs(x), Math.abs(y)); return g ? Math.abs(x * y) / g : 0; }));
  def('EVEN', a => { const x = toNum(a[0]); const c = Math.ceil(Math.abs(x)); return Math.sign(x) * (c % 2 ? c + 1 : c); });
  def('ODD', a => { const x = toNum(a[0]); const c = Math.ceil(Math.abs(x)); return Math.sign(x) * (c % 2 ? c : c + 1); });

  LAZY.IF = (n, ev) => { need(n, 2); return toBool(ev(n[0])) ? ev(n[1]) : (n.length > 2 ? ev(n[2]) : false); };
  LAZY.IFERROR = (n, ev) => { try { const v = scalar(ev(n[0])); if (v instanceof FErr) return ev(n[1]); return v; } catch (e) { if (e instanceof FErr) return ev(n[1]); throw e; } };
  LAZY.IFNA = (n, ev) => { try { const v = scalar(ev(n[0])); if (v === E.NA) return ev(n[1]); return v; } catch (e) { if (e === E.NA) return ev(n[1]); throw e; } };
  LAZY.IFS = (n, ev) => { for (let i = 0; i + 1 < n.length; i += 2) if (toBool(ev(n[i]))) return ev(n[i + 1]); throw E.NA; };
  LAZY.SWITCH = (n, ev) => { const x = ev(n[0]); let i = 1; for (; i + 1 < n.length; i += 2) if (cmpVals(x, ev(n[i])) === 0) return ev(n[i + 1]); if (i < n.length) return ev(n[i]); throw E.NA; };
  LAZY.CHOOSE = (n, ev) => { const i = Math.floor(toNum(ev(n[0]))); if (i < 1 || i >= n.length) throw E.VALUE; return ev(n[i]); };
  LAZY.AND = (n, ev) => { let any = false; for (const x of n) { for (const v of allVals([ev(x)])) { if (v === null || typeof v === 'string' && v === '') continue; any = true; if (!toBool(v)) return false; } } if (!any) throw E.VALUE; return true; };
  LAZY.OR = (n, ev) => { let any = false, res = false; for (const x of n) { for (const v of allVals([ev(x)])) { if (v === null || v === '') continue; any = true; if (toBool(v)) res = true; } } if (!any) throw E.VALUE; return res; };
  def('NOT', a => !toBool(a[0])); def('TRUE', () => true); def('FALSE', () => false);
  def('XOR', a => allVals(a).filter(v => v !== null && toBool(v)).length % 2 === 1);

  def('CONCAT CONCATENATE', a => allVals(a).map(toStr).join(''));
  def('TEXTJOIN', a => { const d = toStr(a[0]), skip = toBool(a[1]); let v = allVals(a.slice(2)).map(x => x === null ? '' : toStr(x)); if (skip) v = v.filter(x => x !== ''); return v.join(d); });
  def('LEFT', a => toStr(a[0]).slice(0, a.length > 1 ? Math.max(0, toNum(a[1])) : 1));
  def('RIGHT', a => { const s = toStr(a[0]), n = a.length > 1 ? Math.max(0, toNum(a[1])) : 1; return n === 0 ? '' : s.slice(-n); });
  def('MID', a => { const s = toStr(a[0]), st = toNum(a[1]), n = toNum(a[2]); if (st < 1 || n < 0) throw E.VALUE; return s.substr(st - 1, n); });
  def('LEN', a => toStr(a[0]).length); def('UPPER', a => toStr(a[0]).toUpperCase()); def('LOWER', a => toStr(a[0]).toLowerCase());
  def('PROPER', a => toStr(a[0]).toLowerCase().replace(/(^|[^a-z])([a-z])/g, (m, p, c) => p + c.toUpperCase()));
  def('TRIM', a => toStr(a[0]).trim().replace(/\s+/g, ' '));
  def('CLEAN', a => toStr(a[0]).replace(/[\x00-\x1f]/g, ''));
  def('REPT', a => { const n = toNum(a[1]); if (n < 0) throw E.VALUE; return toStr(a[0]).repeat(n); });
  def('EXACT', a => toStr(a[0]) === toStr(a[1]));
  def('CHAR', a => String.fromCharCode(toNum(a[0]))); def('CODE', a => toStr(a[0]).charCodeAt(0) || 0);
  def('SUBSTITUTE', a => { const s = toStr(a[0]), o = toStr(a[1]), n = toStr(a[2]); if (!o) return s; if (a.length > 3) { let k = toNum(a[3]), i = -1; while (k-- > 0) { i = s.indexOf(o, i + 1); if (i < 0) return s; } return s.slice(0, i) + n + s.slice(i + o.length); } return s.split(o).join(n); });
  def('REPLACE', a => { const s = toStr(a[0]), st = toNum(a[1]) - 1, n = toNum(a[2]); return s.slice(0, st) + toStr(a[3]) + s.slice(st + n); });
  def('FIND', a => { const i = toStr(a[1]).indexOf(toStr(a[0]), a.length > 2 ? toNum(a[2]) - 1 : 0); if (i < 0) throw E.VALUE; return i + 1; });
  def('SEARCH', a => { const re = new RegExp(toStr(a[0]).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.'), 'i'); const s = toStr(a[1]), st = a.length > 2 ? toNum(a[2]) - 1 : 0; const m = re.exec(s.slice(st)); if (!m) throw E.VALUE; return m.index + st + 1; });
  def('VALUE', a => { const v = scalar(a[0]); if (typeof v === 'number') return v; const l = parseLiteral(toStr(v).trim()); if (typeof l.v === 'number') return l.v; throw E.VALUE; });
  def('TEXT', a => formatValue(toNum(a[0]), toStr(a[1])));
  def('N', a => { const v = scalar(a[0]); return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; });
  def('T', a => { const v = scalar(a[0]); return typeof v === 'string' ? v : ''; });
  def('NUMBERVALUE', a => toNum(a[0]));
  def('FIXED', a => { const d = a.length > 1 ? toNum(a[1]) : 2; let s = formatNumber(toNum(a[0]), '#,##0' + (d > 0 ? '.' + '0'.repeat(d) : '')); if (a.length > 2 && toBool(a[2])) s = s.replace(/,/g, ''); return s; });

  def('ISBLANK', a => { const v = scalar(a[0]); return v === null || v === undefined; });
  def('ISNUMBER', a => typeof scalar(a[0]) === 'number'); def('ISTEXT', a => typeof scalar(a[0]) === 'string');
  def('ISLOGICAL', a => typeof scalar(a[0]) === 'boolean'); def('ISERROR', a => scalar(a[0]) instanceof FErr);
  def('ISERR', a => { const v = scalar(a[0]); return v instanceof FErr && v !== E.NA; }); def('ISNA', a => scalar(a[0]) === E.NA);
  def('ISEVEN', a => Math.floor(toNum(a[0])) % 2 === 0); def('ISODD', a => Math.abs(Math.floor(toNum(a[0]))) % 2 === 1);
  def('NA', () => { throw E.NA; });

  def('TODAY', () => Math.floor(nowSerial())); def('NOW', () => nowSerial());
  def('DATE', a => { let y = toNum(a[0]); if (y < 1900) y += 1900; return ymdToSerial(y, toNum(a[1]), toNum(a[2])); });
  def('TIME', a => (toNum(a[0]) * 3600 + toNum(a[1]) * 60 + toNum(a[2])) / 86400);
  def('YEAR', a => serialToDate(toNum(a[0])).getUTCFullYear()); def('MONTH', a => serialToDate(toNum(a[0])).getUTCMonth() + 1);
  def('DAY', a => serialToDate(toNum(a[0])).getUTCDate());
  def('HOUR', a => Math.floor(((toNum(a[0]) % 1) * 86400 + 0.5) / 3600) % 24); def('MINUTE', a => Math.floor(((toNum(a[0]) % 1) * 86400 + 0.5) / 60) % 60); def('SECOND', a => Math.floor((toNum(a[0]) % 1) * 86400 + 0.5) % 60);
  def('WEEKDAY', a => { const w = serialToDate(toNum(a[0])).getUTCDay(), t = a.length > 1 ? toNum(a[1]) : 1; return t === 2 ? (w + 6) % 7 + 1 : t === 3 ? (w + 6) % 7 : w + 1; });
  def('DAYS', a => Math.floor(toNum(a[0])) - Math.floor(toNum(a[1])));
  def('EDATE', a => { const d = serialToDate(toNum(a[0])); return ymdToSerial(d.getUTCFullYear(), d.getUTCMonth() + 1 + toNum(a[1]), d.getUTCDate()); });
  def('EOMONTH', a => { const d = serialToDate(toNum(a[0])); return ymdToSerial(d.getUTCFullYear(), d.getUTCMonth() + 2 + toNum(a[1]), 0); });
  def('DATEDIF', a => { const s = serialToDate(toNum(a[0])), e = serialToDate(toNum(a[1])), u = toStr(a[2]).toUpperCase(); if (e < s) throw E.NUM; const m = (e.getUTCFullYear() - s.getUTCFullYear()) * 12 + e.getUTCMonth() - s.getUTCMonth() - (e.getUTCDate() < s.getUTCDate() ? 1 : 0); if (u === 'D') return Math.floor((e - s) / DAY); if (u === 'M') return m; if (u === 'Y') return Math.floor(m / 12); throw E.NUM; });
  def('DATEVALUE', a => { const d = parseDateText(toStr(a[0]).trim()); if (!d) throw E.VALUE; return Math.floor(d.v); });

  const exactEq = (a, b) => { a = scalar(a); b = scalar(b); if (a instanceof FErr || b instanceof FErr) return false; if (a === null) a = ''; if (b === null) b = ''; if (typeof a !== typeof b) return false; return typeof a === 'string' ? a.toLowerCase() === b.toLowerCase() : a === b; };
  def('VLOOKUP', a => {
    need(a, 3); const t = a[1]; if (!(t instanceof Rng)) throw E.VALUE; const ci = toNum(a[2]) - 1; if (ci < 0 || ci >= t.w) throw E.REF;
    const exact = a.length > 3 && !toBool(a[3]); const key = scalar(a[0]); chk(key);
    if (exact) { for (const row of t.rows) if (exactEq(row[0], key)) return row[ci]; throw E.NA; }
    let best = -1; for (let i = 0; i < t.h; i++) { const v = t.rows[i][0]; if (v === null) continue; if (typeof v !== typeof key) continue; if (cmpVals(v, key) <= 0) best = i; else break; }
    if (best < 0) throw E.NA; return t.rows[best][ci];
  });
  def('HLOOKUP', a => {
    need(a, 3); const t = a[1]; if (!(t instanceof Rng)) throw E.VALUE; const ri = toNum(a[2]) - 1; if (ri < 0 || ri >= t.h) throw E.REF;
    const exact = a.length > 3 && !toBool(a[3]); const key = scalar(a[0]); chk(key);
    if (exact) { for (let j = 0; j < t.w; j++) if (exactEq(t.rows[0][j], key)) return t.rows[ri][j]; throw E.NA; }
    let best = -1; for (let j = 0; j < t.w; j++) { const v = t.rows[0][j]; if (v === null || typeof v !== typeof key) continue; if (cmpVals(v, key) <= 0) best = j; else break; }
    if (best < 0) throw E.NA; return t.rows[ri][best];
  });
  const vec = r => r instanceof Rng ? [].concat(...r.rows) : [r];
  def('XLOOKUP', a => { const ks = vec(a[1]), rs = vec(a[2]); for (let i = 0; i < ks.length; i++) if (exactEq(ks[i], a[0])) return rs[i]; if (a.length > 3) return a[3]; throw E.NA; });
  def('MATCH', a => {
    const ks = vec(a[1]), t = a.length > 2 ? toNum(a[2]) : 1, key = scalar(a[0]);
    if (t === 0) { for (let i = 0; i < ks.length; i++) { if (typeof key === 'string' && /[*?]/.test(key)) { if (makeCriteria(key)(ks[i])) return i + 1; } else if (exactEq(ks[i], key)) return i + 1; } throw E.NA; }
    let best = -1;
    if (t > 0) { for (let i = 0; i < ks.length; i++) { if (ks[i] === null || typeof ks[i] !== typeof key) continue; if (cmpVals(ks[i], key) <= 0) best = i; else break; } }
    else { for (let i = 0; i < ks.length; i++) { if (ks[i] === null || typeof ks[i] !== typeof key) continue; if (cmpVals(ks[i], key) >= 0) best = i; else break; } }
    if (best < 0) throw E.NA; return best + 1;
  });
  def('INDEX', a => {
    const t = a[0] instanceof Rng ? a[0] : new Rng([[a[0]]]); let r = a.length > 1 ? toNum(a[1]) : 0, c = a.length > 2 ? toNum(a[2]) : 0;
    if (t.h === 1 && a.length === 2) { c = r; r = 1; }
    if (r === 0 && c === 0) return t; if (r < 0 || c < 0 || r > t.h || c > t.w) throw E.REF;
    if (r === 0) return new Rng(t.rows.map(x => [x[c - 1]]), t.r1, t.c1 + c - 1, t.sheet);
    if (c === 0) return new Rng([t.rows[r - 1]], t.r1 + r - 1, t.c1, t.sheet);
    return t.rows[r - 1][c - 1];
  });
  def('ROWS', a => a[0] instanceof Rng ? a[0].h : 1); def('COLUMNS', a => a[0] instanceof Rng ? a[0].w : 1);
  def('TRANSPOSE', a => { const t = a[0]; if (!(t instanceof Rng)) return t; const rows = []; for (let j = 0; j < t.w; j++) rows.push(t.rows.map(r => r[j])); return new Rng(rows); });

  const critPairs = (args, from) => { const out = []; for (let i = from; i + 1 < args.length; i += 2) { if (!(args[i] instanceof Rng)) throw E.VALUE; out.push([vec(args[i]), makeCriteria(args[i + 1])]); } return out; };
  const matchAll = (pairs, n) => { const ok = []; for (let i = 0; i < n; i++) ok.push(pairs.every(([v, f]) => f(v[i]))); return ok; };
  def('COUNTIF', a => { need(a, 2); return vec(a[0]).filter(makeCriteria(a[1])).length; });
  def('COUNTIFS', a => { const p = critPairs(a, 0); return matchAll(p, p[0][0].length).filter(Boolean).length; });
  def('SUMIF', a => { need(a, 2); const k = vec(a[0]), f = makeCriteria(a[1]), s = a.length > 2 ? vec(a[2]) : k; let t = 0; k.forEach((v, i) => { if (f(v) && typeof s[i] === 'number') t += s[i]; }); return t; });
  def('SUMIFS', a => { const s = vec(a[0]), p = critPairs(a, 1), ok = matchAll(p, s.length); return s.reduce((t, v, i) => t + (ok[i] && typeof v === 'number' ? v : 0), 0); });
  def('AVERAGEIF', a => { const k = vec(a[0]), f = makeCriteria(a[1]), s = a.length > 2 ? vec(a[2]) : k; let t = 0, n = 0; k.forEach((v, i) => { if (f(v) && typeof s[i] === 'number') { t += s[i]; n++; } }); if (!n) throw E.DIV0; return t / n; });
  def('AVERAGEIFS', a => { const s = vec(a[0]), p = critPairs(a, 1), ok = matchAll(p, s.length); let t = 0, n = 0; s.forEach((v, i) => { if (ok[i] && typeof v === 'number') { t += v; n++; } }); if (!n) throw E.DIV0; return t / n; });
  def('MAXIFS', a => { const s = vec(a[0]), p = critPairs(a, 1), ok = matchAll(p, s.length); const v = s.filter((x, i) => ok[i] && typeof x === 'number'); return v.length ? Math.max(...v) : 0; });
  def('MINIFS', a => { const s = vec(a[0]), p = critPairs(a, 1), ok = matchAll(p, s.length); const v = s.filter((x, i) => ok[i] && typeof x === 'number'); return v.length ? Math.min(...v) : 0; });
  def('UNIQUE', a => { const seen = new Set(), out = []; for (const v of vec(a[0])) { const k = typeof v + ':' + String(v).toLowerCase(); if (!seen.has(k)) { seen.add(k); out.push([v]); } } return new Rng(out); });
  def('PMT', a => { const r = toNum(a[0]), n = toNum(a[1]), pv = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0; if (r === 0) return -(pv + fv) / n; const f = Math.pow(1 + r, n); return -(pv * f + fv) * r / (f - 1); });
  def('FV', a => { const r = toNum(a[0]), n = toNum(a[1]), p = toNum(a[2]), pv = a.length > 3 ? toNum(a[3]) : 0; if (r === 0) return -(pv + p * n); const f = Math.pow(1 + r, n); return -(pv * f + p * (f - 1) / r); });
  def('PV', a => { const r = toNum(a[0]), n = toNum(a[1]), p = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0; if (r === 0) return -(fv + p * n); const f = Math.pow(1 + r, n); return -(fv + p * (f - 1) / r) / f; });
  def('NPV', a => { const r = toNum(a[0]); return nums(a.slice(1)).reduce((s, v, i) => s + v / Math.pow(1 + r, i + 1), 0); });


  /* ---------- dynamic arrays, stats, finance, misc ---------- */
  const R = a => a instanceof Rng ? a : new Rng([[a]]);
  const flat = a => [].concat(...R(a).rows);
  const mk = rows => new Rng(rows);
  def('SEQUENCE', a => { const r = toNum(a[0]), c = a.length > 1 && a[1] !== null ? toNum(a[1]) : 1, s = a.length > 2 && a[2] !== null ? toNum(a[2]) : 1, st = a.length > 3 && a[3] !== null ? toNum(a[3]) : 1; if (r < 1 || c < 1 || r * c > 1e6) throw E.VALUE; const rows = []; let v = s; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) { row.push(v); v += st; } rows.push(row); } return mk(rows); });
  def('RANDARRAY', a => { const r = a.length ? toNum(a[0]) : 1, c = a.length > 1 ? toNum(a[1]) : 1; const rows = []; for (let i = 0; i < r; i++) { const row = []; for (let j = 0; j < c; j++) row.push(Math.random()); rows.push(row); } return mk(rows); });
  def('SORT', a => { const t = R(a[0]), idx = a.length > 1 && a[1] !== null ? toNum(a[1]) - 1 : 0, asc = !(a.length > 2 && a[2] !== null && toNum(a[2]) === -1), byCol = a.length > 3 && toBool(a[3]); if (byCol) { const cols = []; for (let j = 0; j < t.w; j++) cols.push(t.rows.map(r => r[j])); cols.sort((x, y) => cmpVals(x[idx], y[idx]) * (asc ? 1 : -1)); const rows = []; for (let i = 0; i < t.h; i++) rows.push(cols.map(c => c[i])); return mk(rows); } const rows = t.rows.slice().sort((x, y) => cmpVals(x[idx], y[idx]) * (asc ? 1 : -1)); return mk(rows); });
  def('SORTBY', a => { const t = R(a[0]); const keys = []; for (let i = 1; i < a.length; i += 2) keys.push([flat(a[i]), a[i + 1] !== undefined && toNum(a[i + 1]) === -1 ? -1 : 1]); const idx = t.rows.map((_, i) => i); idx.sort((x, y) => { for (const [k, d] of keys) { const c = cmpVals(k[x], k[y]) * d; if (c) return c; } return x - y; }); return mk(idx.map(i => t.rows[i])); });
  def('FILTER', a => { const t = R(a[0]), inc = R(a[1]); const byRow = inc.w === 1 || inc.h === t.h && inc.w === 1; const rows = []; if (inc.h === t.h && inc.w === 1) { t.rows.forEach((r, i) => { if (toBool(inc.rows[i][0])) rows.push(r); }); } else if (inc.w === t.w && inc.h === 1) { const keep = inc.rows[0].map(toBool); t.rows.forEach(r => rows.push(r.filter((_, j) => keep[j]))); if (!rows[0] || !rows[0].length) rows.length = 0; } else throw E.VALUE; if (!rows.length) { if (a.length > 2) return a[2]; throw E.NA; } return mk(rows); });
  def('VSTACK', a => mk([].concat(...a.map(x => R(x).rows.map(r => r.slice())))));
  def('HSTACK', a => { const rs = a.map(R); const h = Math.max(...rs.map(r => r.h)); const rows = []; for (let i = 0; i < h; i++) { let row = []; for (const r of rs) { for (let j = 0; j < r.w; j++) row.push(r.rows[i] ? r.rows[i][j] : E.NA); } rows.push(row); } return mk(rows); });
  def('TAKE', a => { const t = R(a[0]); const n = toNum(a[1]); const c = a.length > 2 && a[2] !== null ? toNum(a[2]) : null; let rows = n >= 0 ? t.rows.slice(0, n) : t.rows.slice(n); if (c !== null) rows = rows.map(r => c >= 0 ? r.slice(0, c) : r.slice(c)); return mk(rows); });
  def('DROP', a => { const t = R(a[0]); const n = toNum(a[1]); const c = a.length > 2 && a[2] !== null ? toNum(a[2]) : 0; let rows = n >= 0 ? t.rows.slice(n) : t.rows.slice(0, n); if (c) rows = rows.map(r => c >= 0 ? r.slice(c) : r.slice(0, c)); return mk(rows); });
  def('TEXTSPLIT', a => { const s = toStr(a[0]), cd = a.length > 1 && a[1] !== null ? toStr(a[1]) : '', rd = a.length > 2 && a[2] !== null ? toStr(a[2]) : ''; const rowsT = rd ? s.split(rd) : [s]; const rows = rowsT.map(x => cd ? x.split(cd) : [x]); const w = Math.max(...rows.map(r => r.length)); return mk(rows.map(r => { while (r.length < w) r.push(E.NA); return r; })); });
  def('TEXTBEFORE', a => { const s = toStr(a[0]), d = toStr(a[1]), n = a.length > 2 ? toNum(a[2]) : 1; let i = -1, k = n; if (n > 0) { while (k-- > 0) { i = s.indexOf(d, i + 1); if (i < 0) throw E.NA; } } else { i = s.length; while (k++ < 0) { i = s.lastIndexOf(d, i - 1); if (i < 0) throw E.NA; } } return s.slice(0, i); });
  def('TEXTAFTER', a => { const s = toStr(a[0]), d = toStr(a[1]), n = a.length > 2 ? toNum(a[2]) : 1; let i = -1, k = n; if (n > 0) { while (k-- > 0) { i = s.indexOf(d, i + 1); if (i < 0) throw E.NA; } } else { i = s.length; while (k++ < 0) { i = s.lastIndexOf(d, i - 1); if (i < 0) throw E.NA; } } return s.slice(i + d.length); });
  LAZY.LET = (n, ev, wb, ctx) => { const sc = {}; wb.scopes.push(sc); try { for (let i = 0; i + 1 < n.length; i += 2) { if (n[i].k !== 'name') throw E.NAME; sc[n[i].v.toUpperCase()] = ev(n[i + 1]); } return ev(n[n.length - 1]); } finally { wb.scopes.pop(); } };
  def('HYPERLINK', a => a.length > 1 && a[1] !== null ? scalar(a[1]) : toStr(a[0]));
  def('PERCENTILE PERCENTILE.INC', a => { const n = nums([a[0]]).sort((x, y) => x - y), k = toNum(a[1]); if (!n.length || k < 0 || k > 1) throw E.NUM; const p = (n.length - 1) * k, lo = Math.floor(p), hi = Math.ceil(p); return n[lo] + (n[hi] - n[lo]) * (p - lo); });
  def('QUARTILE QUARTILE.INC', a => { const q = toNum(a[1]); if (q < 0 || q > 4 || q !== Math.floor(q)) throw E.NUM; return FN.PERCENTILE([a[0], q / 4]); });
  const pairs = (ys, xs) => { const y = flat(ys), x = flat(xs); if (y.length !== x.length) throw E.NA; const o = []; for (let i = 0; i < y.length; i++) if (typeof y[i] === 'number' && typeof x[i] === 'number') o.push([x[i], y[i]]); if (o.length < 2) throw E.DIV0; return o; };
  const lin = p => { const n = p.length, mx = p.reduce((s, v) => s + v[0], 0) / n, my = p.reduce((s, v) => s + v[1], 0) / n; let sxy = 0, sxx = 0, syy = 0; for (const [x, y] of p) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; } if (!sxx) throw E.DIV0; return { m: sxy / sxx, b: my - sxy / sxx * mx, r: syy ? sxy / Math.sqrt(sxx * syy) : 0 }; };
  def('SLOPE', a => lin(pairs(a[0], a[1])).m); def('INTERCEPT', a => lin(pairs(a[0], a[1])).b); def('CORREL', a => lin(pairs(a[0], a[1])).r); def('RSQ', a => lin(pairs(a[0], a[1])).r ** 2);
  def('FORECAST FORECAST.LINEAR', a => { const l = lin(pairs(a[1], a[2])); return l.b + l.m * toNum(a[0]); });
  def('SUMSQ', a => nums(a).reduce((s, v) => s + v * v, 0));
  def('AVERAGEA', a => { const v = allVals(a).filter(x => x !== null && x !== ''); if (!v.length) throw E.DIV0; return v.reduce((s, x) => s + (typeof x === 'number' ? x : typeof x === 'boolean' ? (x ? 1 : 0) : 0), 0) / v.length; });
  def('COMBIN', a => { const n = toNum(a[0]), k = toNum(a[1]); if (k < 0 || k > n) throw E.NUM; let r = 1; for (let i = 1; i <= k; i++) r = r * (n - k + i) / i; return Math.round(r); });
  def('PERMUT', a => { const n = toNum(a[0]), k = toNum(a[1]); if (k < 0 || k > n) throw E.NUM; let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; });
  def('MROUND', a => { const x = toNum(a[0]), m = toNum(a[1]); if (m === 0) return 0; return Math.round(x / m) * m; });
  def('DOLLAR', a => formatNumber(toNum(a[0]), '$#,##0' + (a.length > 1 && toNum(a[1]) > 0 ? '.' + '0'.repeat(toNum(a[1])) : a.length > 1 ? '' : '.00')));
  def('UNICHAR', a => String.fromCodePoint(toNum(a[0])));
  def('NPER', a => { const r = toNum(a[0]), p = toNum(a[1]), pv = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0; if (r === 0) return -(pv + fv) / p; const v = (p - fv * r) / (p + pv * r); if (v <= 0) throw E.NUM; return Math.log(v) / Math.log(1 + r); });
  def('RATE', a => { const n = toNum(a[0]), p = toNum(a[1]), pv = toNum(a[2]), fv = a.length > 3 ? toNum(a[3]) : 0; let r = a.length > 5 ? toNum(a[5]) : 0.1; for (let i = 0; i < 100; i++) { const f = Math.pow(1 + r, n); const y = pv * f + p * (f - 1) / r + fv; const dy = pv * n * Math.pow(1 + r, n - 1) + p * (n * Math.pow(1 + r, n - 1) * r - (f - 1)) / (r * r); const nr = r - y / dy; if (!isFinite(nr)) throw E.NUM; if (Math.abs(nr - r) < 1e-10) return nr; r = nr; } throw E.NUM; });
  def('IRR', a => { const cf = nums([a[0]]); if (!cf.some(x => x > 0) || !cf.some(x => x < 0)) throw E.NUM; const f = r => cf.reduce((s, v, i) => s + v / Math.pow(1 + r, i), 0); let lo = -0.99, hi = 10; if (f(lo) * f(hi) > 0) throw E.NUM; for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; if (f(lo) * f(m) <= 0) hi = m; else lo = m; } return (lo + hi) / 2; });
  def('SLN', a => (toNum(a[0]) - toNum(a[1])) / toNum(a[2]));
  def('IPMT', a => { const r = toNum(a[0]), per = toNum(a[1]), n = toNum(a[2]), pv = toNum(a[3]); const pmt = FN.PMT([r, n, pv]); let bal = pv; for (let i = 1; i < per; i++) bal = bal * (1 + r) + pmt; return -(bal * r); });
  def('PPMT', a => FN.PMT([a[0], a[2], a[3]]) - FN.IPMT(a));
  const isWk = d => { const w = serialToDate(d).getUTCDay(); return w !== 0 && w !== 6; };
  def('NETWORKDAYS', a => { let s = Math.floor(toNum(a[0])), e = Math.floor(toNum(a[1])); const sign = e < s ? -1 : 1; if (e < s) [s, e] = [e, s]; const hol = new Set(a.length > 2 ? nums([a[2]]).map(Math.floor) : []); let n = 0; for (let d = s; d <= e; d++) if (isWk(d) && !hol.has(d)) n++; return n * sign; });
  def('WORKDAY', a => { let d = Math.floor(toNum(a[0])), n = Math.floor(toNum(a[1])); const hol = new Set(a.length > 2 ? nums([a[2]]).map(Math.floor) : []); const st = n < 0 ? -1 : 1; while (n !== 0) { d += st; if (isWk(d) && !hol.has(d)) n -= st; } return d; });
  def('SUBTOTAL', a => { const f = toNum(a[0]) % 100; const rest = a.slice(1); const m = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' }[f]; if (!m) throw E.VALUE; return FN[m](rest); });
  def('LOOKUP', a => { const k = flat(a[1]), r = a.length > 2 ? flat(a[2]) : k; let best = -1; for (let i = 0; i < k.length; i++) { if (k[i] === null || typeof k[i] !== typeof scalar(a[0])) continue; if (cmpVals(k[i], a[0]) <= 0) best = i; else break; } if (best < 0) throw E.NA; return r[best]; });
  def('ADDRESS', a => { const r = toNum(a[0]), c = toNum(a[1]), t = a.length > 2 && a[2] !== null ? toNum(a[2]) : 1; let s = (t === 1 || t === 3 ? '$' : '') + numToCol(c - 1) + (t === 1 || t === 2 ? '$' : '') + r; if (a.length > 4 && a[4] !== null) s = quoteSheet(toStr(a[4])) + '!' + s; return s; });
  def('INDIRECT', (a, ctx) => { const wb = ctx.wb; const p = parseRefToken(toStr(a[0]).trim()); const si = p.sheet ? wb.sheetIndex(p.sheet) : ctx.si; if (si < 0) throw E.REF; if (!p.b) return wb.getValue(si, p.a.r, p.a.c); return wb.getRange(si, p.a.r, p.a.c, p.b.r, p.b.c); });
  def('ISFORMULA', () => false);
  def('TYPE', a => { const v = scalar(a[0]); return typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : v instanceof FErr ? 16 : 1; });
  def('ERROR.TYPE', a => { const v = scalar(a[0]); const m = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7 }; if (v instanceof FErr && m[v.code]) return m[v.code]; throw E.NA; });

  /* ---------- workbook ---------- */
  const CELLKEY = (r, c) => r * 20000 + c;
  class Sheet {
    constructor(name) { this.name = name; this.cells = new Map(); this.colW = {}; this.rowH = {}; this.merges = []; this.freeze = { r: 0, c: 0 }; this.charts = []; this.cf = []; this.hiddenRows = {}; this.hiddenCols = {}; this.filter = null; this.tabColor = null; this.showGrid = true; this.dv = []; this.notes = {}; this.links = {}; }
  }
  class Workbook {
    constructor() { this.sheets = [new Sheet('Sheet1')]; this.cache = new Map(); this.stack = new Set(); this.astCache = new Map(); this.version = 0; this.names = {}; this.spillCells = new Map(); this.spillDirty = true; this.scopes = []; }
    invalidate() { this.cache.clear(); this.version++; this.spillDirty = true; this.spillCells.clear(); }
    sheetIndex(name) { const l = String(name).toLowerCase(); return this.sheets.findIndex(s => s.name.toLowerCase() === l); }
    uniqueName(base) { let n = base, i = 2; while (this.sheetIndex(n) >= 0) n = base.replace(/\d+$/, '') + (i++); return n; }
    addSheet(name, at) { const s = new Sheet(name || this.uniqueName('Sheet' + (this.sheets.length + 1))); if (at === undefined) this.sheets.push(s); else this.sheets.splice(at, 0, s); this.invalidate(); return s; }
    removeSheet(i) { if (this.sheets.length < 2) return false; const nm = this.sheets[i].name; this.sheets.splice(i, 1); this.rewriteAll(ref => ref.sheet && ref.sheet.toLowerCase() === nm.toLowerCase() ? null : ref); this.invalidate(); return true; }
    renameSheet(i, name) {
      const old = this.sheets[i].name; this.sheets[i].name = name;
      this.rewriteAll(ref => (ref.sheet && ref.sheet.toLowerCase() === old.toLowerCase()) ? { ...ref, sheet: name } : ref); this.invalidate();
    }
    duplicateSheet(i) { const src = this.sheets[i], s = new Sheet(this.uniqueName(src.name + ' (2)')); for (const [k, v] of src.cells) s.cells.set(k, { t: v.t, s: v.s ? { ...v.s } : undefined, af: v.af, lit: v.lit }); s.colW = { ...src.colW }; s.rowH = { ...src.rowH }; s.merges = src.merges.map(m => ({ ...m })); s.freeze = { ...src.freeze }; s.tabColor = src.tabColor; s.dv = JSON.parse(JSON.stringify(src.dv)); s.notes = { ...src.notes }; s.links = { ...src.links }; s.cf = JSON.parse(JSON.stringify(src.cf)); this.sheets.splice(i + 1, 0, s); this.invalidate(); return s; }
    rewriteAll(fn) { for (const sh of this.sheets) for (const cell of sh.cells.values()) if (typeof cell.t === 'string' && cell.t[0] === '=') cell.t = rewriteRefs(cell.t, fn); }

    cell(si, r, c) { return this.sheets[si].cells.get(CELLKEY(r, c)); }
    getInput(si, r, c) { const x = this.cell(si, r, c); return x ? x.t : ''; }
    setInput(si, r, c, text) {
      const sh = this.sheets[si], k = CELLKEY(r, c); let cell = sh.cells.get(k);
      if (text === '' || text === null || text === undefined) { if (cell) { if (cell.s && Object.keys(cell.s).length) { cell.t = ''; cell.lit = null; cell.af = null; } else sh.cells.delete(k); } this.invalidate(); return; }
      if (!cell) { cell = {}; sh.cells.set(k, cell); }
      if (typeof text === 'number' || typeof text === 'boolean') { cell.t = String(text).toUpperCase() === 'TRUE' || String(text).toUpperCase() === 'FALSE' ? String(text).toUpperCase() : String(text); cell.lit = { v: text }; cell.af = null; }
      else { cell.t = text; if (text[0] === '=' && text.length > 1) { cell.lit = null; cell.af = null; } else { const l = parseLiteral(text); cell.lit = l; cell.af = l.fmt || null; } }
      this.invalidate();
    }
    getStyle(si, r, c) { const x = this.cell(si, r, c); return x && x.s ? x.s : null; }
    setStyle(si, r, c, patch) {
      const sh = this.sheets[si], k = CELLKEY(r, c); let cell = sh.cells.get(k);
      if (!cell) { cell = { t: '' }; sh.cells.set(k, cell); }
      cell.s = { ...(cell.s || {}) };
      for (const key in patch) { if (patch[key] === null || patch[key] === undefined || patch[key] === false) delete cell.s[key]; else cell.s[key] = patch[key]; }
      if (!Object.keys(cell.s).length) { cell.s = undefined; if (cell.t === '') sh.cells.delete(k); }
    }
    clearCell(si, r, c, keepStyle) { const sh = this.sheets[si], k = CELLKEY(r, c); const cell = sh.cells.get(k); if (!cell) return; if (keepStyle && cell.s) { cell.t = ''; cell.lit = null; cell.af = null; } else sh.cells.delete(k); this.invalidate(); }

    usedRange(si) {
      let r2 = -1, c2 = -1;
      for (const [k, v] of this.sheets[si].cells) { if (v.t === '' && !v.s) continue; if (v.t === '') continue; const r = Math.floor(k / 20000), c = k % 20000; if (r > r2) r2 = r; if (c > c2) c2 = c; }
      return { r2, c2 };
    }
    getValue(si, r, c) {
      if (this.spillDirty) this.computeSpills();
      const key = si + ':' + CELLKEY(r, c);
      if (this.cache.has(key)) return this.cache.get(key);
      const cell = this.cell(si, r, c);
      if (!cell || cell.t === '' || cell.t === undefined) { const sp = this.spillCells.get(key); return sp === undefined ? null : sp; }
      if (typeof cell.t === 'string' && cell.t[0] === '=' && cell.t.length > 1) {
        if (this.stack.has(key)) return E.CIRC;
        this.stack.add(key);
        let val;
        try {
          let ast = this.astCache.get(cell.t);
          if (!ast) { ast = parse(cell.t.slice(1)); this.astCache.set(cell.t, ast); }
          val = this.evalNode(ast, { si, r, c, wb: this });
          if (val instanceof Rng && (val.h > 1 || val.w > 1)) {
            let blocked = false;
            for (let i = 0; i < val.h && !blocked; i++) for (let j = 0; j < val.w; j++) { if (!i && !j) continue; const c2 = this.cell(si, r + i, c + j); if (c2 && c2.t !== '') { blocked = true; break; } }
            if (blocked) val = E.SPILL; else { for (let i = 0; i < val.h; i++) for (let j = 0; j < val.w; j++) if (i || j) { let x = val.rows[i][j]; this.spillCells.set(si + ':' + CELLKEY(r + i, c + j), x === undefined ? null : x); } const sz = { h: val.h, w: val.w }; val = val.rows[0][0]; this.spillSize.set(key, sz); }
          }
          val = scalar(val); if (val === null || val === undefined) val = 0;
        } catch (e) { val = e instanceof FErr ? e : (e instanceof RangeError ? E.NUM : E.NAME); }
        this.stack.delete(key);
        this.cache.set(key, val); return val;
      }
      const l = cell.lit || (cell.lit = parseLiteral(cell.t));
      return l.v;
    }
    computeSpills() {
      this.spillDirty = false; this.spillSize = this.spillSize || new Map(); this.spillSize.clear();
      const re = /^=\s*(FILTER|SORT|SORTBY|UNIQUE|SEQUENCE|TRANSPOSE|TEXTSPLIT|RANDARRAY|VSTACK|HSTACK|TAKE|DROP|LET|INDEX|IF|LOOKUP|XLOOKUP)\s*\(/i;
      for (let si = 0; si < this.sheets.length; si++) for (const [k, cell] of this.sheets[si].cells) if (typeof cell.t === 'string' && cell.t[0] === '=' && re.test(cell.t)) this.getValue(si, Math.floor(k / 20000), k % 20000);
    }
    isSpill(si, r, c) { if (this.spillDirty) this.computeSpills(); return this.spillCells.has(si + ':' + CELLKEY(r, c)); }
    resolveSheet(name, ctx) { if (name == null) return ctx.si; const i = this.sheetIndex(name); if (i < 0) throw E.REF; return i; }
    getRange(si, r1, c1, r2, c2) {
      if (r1 > r2) [r1, r2] = [r2, r1]; if (c1 > c2) [c1, c2] = [c2, c1];
      const used = this.usedRange(si);
      if (r2 >= MAXR - 1 || r2 >= 1048000) r2 = Math.max(used.r2, r1); if (c2 >= MAXC - 1) c2 = Math.max(used.c2, c1);
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 4000000) throw E.NUM;
      const rows = [];
      for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) row.push(this.getValue(si, r, c)); rows.push(row); }
      return new Rng(rows, r1, c1, si);
    }
    evalNode(n, ctx) {
      const ev = x => this.evalNode(x, ctx);
      switch (n.k) {
        case 'num': case 'str': case 'bool': return n.v;
        case 'blank': return null;
        case 'err': throw n.v;
        case 'paren': return ev(n.a);
        case 'ref': {
          const si = this.resolveSheet(n.sheet, ctx);
          if (!n.b) return this.getValue(si, n.a.r, n.a.c);
          return this.getRange(si, n.a.r, n.a.c, n.b.r, n.b.c);
        }
        case 'colr': { const si = this.resolveSheet(n.sheet, ctx); return this.getRange(si, 0, n.a.c, MAXR - 1, (n.b || n.a).c); }
        case 'rowr': { const si = this.resolveSheet(n.sheet, ctx); return this.getRange(si, n.a.r, 0, (n.b || n.a).r, MAXC - 1); }
        case 'name': {
          const up = n.v.toUpperCase();
          for (let i = this.scopes.length - 1; i >= 0; i--) if (up in this.scopes[i]) return this.scopes[i][up];
          const d = this.names[up]; if (d === undefined) throw E.NAME;
          let ast = this.astCache.get('=' + d); if (!ast) { ast = parse(String(d).replace(/^=/, '')); this.astCache.set('=' + d, ast); }
          return this.evalNode(ast, ctx);
        }
        case 'un': { const v = ev(n.a); return arr1(v, x => -toNum(x)); }
        case 'pct': { const v = ev(n.a); return arr1(v, x => toNum(x) / 100); }
        case 'bin': return binop(n.op, ev(n.l), ev(n.r));
        case 'fn': {
          if (LAZY[n.name]) return LAZY[n.name](n.args, ev, this, ctx);
          const f = FN[n.name];
          if (!f) throw E.NAME;
          if (n.name === 'ROW' || n.name === 'COLUMN') return null;
          const args = n.args.map(ev);
          return f(args, ctx);
        }
      }
      throw E.NAME;
    }

    /* structural edits: axis 'r' or 'c' */
    adjust(si, axis, at, n, isInsert) {
      const target = this.sheets[si].name.toLowerCase();
      const own = idx => this.sheets[idx].name.toLowerCase();
      for (let i = 0; i < this.sheets.length; i++) {
        for (const cell of this.sheets[i].cells.values()) {
          if (typeof cell.t !== 'string' || cell.t[0] !== '=') continue;
          cell.t = rewriteRefs(cell.t, ref => {
            const sname = ref.sheet ? ref.sheet.toLowerCase() : own(i);
            if (sname !== target) return ref;
            return adjustRef(ref, axis, at, n, isInsert);
          });
        }
      }
      for (const k of Object.keys(this.names)) { const nr = rewriteRefs('=' + this.names[k], ref => { const sname = ref.sheet ? ref.sheet.toLowerCase() : null; if (sname !== target) return ref; return adjustRef(ref, axis, at, n, isInsert); }); this.names[k] = nr.slice(1); }
      this.astCache.clear();
    }
    insertRows(si, at, n) { this.adjust(si, 'r', at, n, true); this.moveCells(si, 'r', at, n); this.invalidate(); }
    deleteRows(si, at, n) { this.adjust(si, 'r', at, n, false); this.moveCells(si, 'r', at, -n); this.invalidate(); }
    insertCols(si, at, n) { this.adjust(si, 'c', at, n, true); this.moveCells(si, 'c', at, n); this.invalidate(); }
    deleteCols(si, at, n) { this.adjust(si, 'c', at, n, false); this.moveCells(si, 'c', at, -n); this.invalidate(); }
    moveCells(si, axis, at, delta) {
      const sh = this.sheets[si], nm = new Map();
      for (const [k, v] of sh.cells) {
        let r = Math.floor(k / 20000), c = k % 20000, idx = axis === 'r' ? r : c;
        if (delta > 0) { if (idx >= at) idx += delta; }
        else { if (idx >= at && idx < at - delta) continue; if (idx >= at - delta) idx += delta; }
        if (axis === 'r') r = idx; else c = idx;
        nm.set(CELLKEY(r, c), v);
      }
      sh.cells = nm;
      const shiftMap = obj => { const o = {}; for (const key in obj) { let i = +key; if (delta > 0) { if (i >= at) i += delta; } else { if (i >= at && i < at - delta) continue; if (i >= at - delta) i += delta; } o[i] = obj[key]; } return o; };
      if (axis === 'r') { sh.rowH = shiftMap(sh.rowH); sh.hiddenRows = shiftMap(sh.hiddenRows); } else { sh.colW = shiftMap(sh.colW); sh.hiddenCols = shiftMap(sh.hiddenCols); }
      sh.merges = sh.merges.map(m => {
        const lo = axis === 'r' ? m.r1 : m.c1, hi = axis === 'r' ? m.r2 : m.c2;
        let a = lo, b = hi;
        if (delta > 0) { if (a >= at) a += delta; if (b >= at) b += delta; }
        else { const d = -delta; if (a >= at + d) a -= d; else if (a >= at) a = at; if (b >= at + d) b -= d; else if (b >= at) b = at - 1; if (b < a) return null; }
        return axis === 'r' ? { ...m, r1: a, r2: b } : { ...m, c1: a, c2: b };
      }).filter(Boolean).filter(m => m.r1 !== m.r2 || m.c1 !== m.c2);
    }

    display(si, r, c) {
      const cell = this.cell(si, r, c); if (!cell || cell.t === '') { if (this.isSpill(si, r, c)) return formatValue(this.getValue(si, r, c), null); return ''; }
      const v = this.getValue(si, r, c);
      let nf = cell.s && cell.s.nf; if (!nf) nf = cell.af || null;
      if (!nf && typeof v === 'number' && cell.t[0] === '=') { const m = /^=\s*(TODAY|NOW|DATE|EDATE|EOMONTH)\s*\(/i.exec(cell.t); if (m) nf = m[1].toUpperCase() === 'NOW' ? 'm/d/yyyy h:mm' : 'm/d/yyyy'; }
      return formatValue(v, nf);
    }

    toJSON() {
      return {
        v: 1, names: this.names, sheets: this.sheets.map(s => ({
          name: s.name, colW: s.colW, rowH: s.rowH, merges: s.merges, freeze: s.freeze, tab: s.tabColor, grid: s.showGrid, cf: s.cf, charts: s.charts,
          hr: s.hiddenRows, hc: s.hiddenCols, fl: s.filter || null, dv: s.dv, nt: s.notes, lk: s.links,
          cells: [...s.cells].map(([k, v]) => [Math.floor(k / 20000), k % 20000, v.t, v.s || 0]).filter(x => x[2] !== '' || x[3])
        }))
      };
    }
    static fromJSON(j) {
      const wb = new Workbook(); wb.sheets = [];
      for (const s of j.sheets) {
        const sh = new Sheet(s.name); sh.colW = s.colW || {}; sh.rowH = s.rowH || {}; sh.merges = s.merges || []; sh.freeze = s.freeze || { r: 0, c: 0 };
        sh.tabColor = s.tab || null; sh.showGrid = s.grid !== false; sh.cf = s.cf || []; sh.charts = s.charts || []; sh.hiddenRows = s.hr || {}; sh.hiddenCols = s.hc || {}; sh.filter = s.fl || null; sh.dv = s.dv || []; sh.notes = s.nt || {}; sh.links = s.lk || {};
        wb.sheets.push(sh);
        for (const [r, c, t, st] of s.cells) { const cell = { t }; if (st) cell.s = st; if (typeof t === 'string' && t !== '' && t[0] !== '=') { cell.lit = parseLiteral(t); cell.af = cell.lit.fmt || null; } sh.cells.set(CELLKEY(r, c), cell); }
      }
      if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
      wb.names = j.names || {};
      return wb;
    }
  }

  function adjustRef(ref, axis, at, n, isInsert) {
    const prop = axis, flag = axis === 'r' ? 'col' : 'row';
    const part = p => p && p[flag] ? null : p;
    const a = { ...ref.a }, b = ref.b ? { ...ref.b } : null;
    if (a[flag] && (!b || b[flag])) return ref; // whole-column refs on row ops etc.
    const one = (p, isEnd) => {
      if (p[flag]) return p; const i = p[prop];
      if (isInsert) { if (i >= at) p[prop] = i + n; return p; }
      if (i >= at && i < at + n) return null;
      if (i >= at + n) p[prop] = i - n; return p;
    };
    if (!b) { const x = one(a); return x ? { sheet: ref.sheet, a: x, b: null } : null; }
    if (isInsert) { one(a); one(b); return { sheet: ref.sheet, a, b }; }
    let s = a[prop], e = b[prop];
    if (a[flag]) return ref;
    const del0 = at, del1 = at + n - 1;
    let ns = s < del0 ? s : (s <= del1 ? del0 : s - n);
    let ne = e < del0 ? e : (e <= del1 ? del0 - 1 : e - n);
    if (ne < ns) return null;
    a[prop] = ns; b[prop] = ne; return { sheet: ref.sheet, a, b };
  }

  function arr1(v, f) { if (v instanceof Rng) return new Rng(v.rows.map(r => r.map(x => x instanceof FErr ? x : f(x))), v.r1, v.c1, v.sheet); return f(v); }
  function arr2(a, b, f) {
    const A = a instanceof Rng, B = b instanceof Rng;
    if (!A && !B) return f(a, b);
    const h = Math.max(A ? a.h : 1, B ? b.h : 1), w = Math.max(A ? a.w : 1, B ? b.w : 1), rows = [];
    for (let i = 0; i < h; i++) { const row = []; for (let j = 0; j < w; j++) { const x = A ? (a.rows[a.h === 1 ? 0 : i] || [])[a.w === 1 ? 0 : j] : a, y = B ? (b.rows[b.h === 1 ? 0 : i] || [])[b.w === 1 ? 0 : j] : b; try { row.push(f(x === undefined ? E.NA : x, y === undefined ? E.NA : y)); } catch (e) { if (e instanceof FErr) row.push(e); else throw e; } } rows.push(row); }
    return new Rng(rows);
  }
  function binop(op, a, b) {
    return arr2(a, b, (x, y) => {
      switch (op) {
        case '+': return toNum(x) + toNum(y);
        case '-': return toNum(x) - toNum(y);
        case '*': return toNum(x) * toNum(y);
        case '/': { const d = toNum(y); if (d === 0) throw E.DIV0; return toNum(x) / d; }
        case '^': { const r = Math.pow(toNum(x), toNum(y)); if (!isFinite(r) || isNaN(r)) throw E.NUM; return r; }
        case '&': return toStr(x) + toStr(y);
        case '=': return cmpVals(x, y) === 0; case '<>': return cmpVals(x, y) !== 0;
        case '<': return cmpVals(x, y) < 0; case '>': return cmpVals(x, y) > 0;
        case '<=': return cmpVals(x, y) <= 0; case '>=': return cmpVals(x, y) >= 0;
      }
      throw E.NAME;
    });
  }

  /* ROW()/COLUMN() need ctx — patch via FN with ctx param */
  FN.ROW = (a, ctx) => a.length ? (a[0] instanceof Rng ? a[0].r1 + 1 : 1) : ctx.r + 1;
  FN.COLUMN = (a, ctx) => a.length ? (a[0] instanceof Rng ? a[0].c1 + 1 : 1) : ctx.c + 1;
  // evalNode short-circuits ROW/COLUMN above; route properly:
  const _evalNode = Workbook.prototype.evalNode;
  Workbook.prototype.evalNode = function (n, ctx) {
    if (n.k === 'fn' && (n.name === 'ROW' || n.name === 'COLUMN')) {
      const args = n.args.map(x => this.evalNode(x, ctx));
      if (args.length && !(args[0] instanceof Rng)) { const a0 = n.args[0]; if (a0.k === 'ref') return n.name === 'ROW' ? a0.a.r + 1 : a0.a.c + 1; }
      return FN[n.name](args, ctx);
    }
    if (n.k === 'fn' && n.name === 'OFFSET') {
      const a0 = n.args[0]; if (!a0 || (a0.k !== 'ref')) throw E.VALUE; const si = this.resolveSheet(a0.sheet, ctx); const ev = x => this.evalNode(x, ctx);
      const dr = toNum(ev(n.args[1])), dc = toNum(ev(n.args[2])); const b = a0.b || a0.a; const h0 = Math.abs(b.r - a0.a.r) + 1, w0 = Math.abs(b.c - a0.a.c) + 1;
      const h = n.args[3] && n.args[3].k !== 'blank' ? toNum(ev(n.args[3])) : h0, w = n.args[4] && n.args[4].k !== 'blank' ? toNum(ev(n.args[4])) : w0;
      const r1 = Math.min(a0.a.r, b.r) + dr, c1 = Math.min(a0.a.c, b.c) + dc; if (r1 < 0 || c1 < 0 || h < 1 || w < 1) throw E.REF;
      return h === 1 && w === 1 ? this.getValue(si, r1, c1) : this.getRange(si, r1, c1, r1 + h - 1, c1 + w - 1);
    }
    return _evalNode.call(this, n, ctx);
  };

  return {
    Workbook, Sheet, Rng, FErr, E, parse, tokenize, rewriteRefs, shiftFormula, parseRefToken, refToStr, parseLiteral, formatValue, BUILTIN_FMT,
    colToNum, numToCol, addrToStr, parseAddr, adjustRef, toNum, toStr, cmpVals, isDateFmt, serialToDate, dateToSerial, ymdToSerial, generalFmt, CELLKEY, MAXC, MAXR, FN, LAZY
  };
});
