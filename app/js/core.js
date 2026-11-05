/* FmOffice Sheets — core: state, grid rendering, selection, editing, clipboard, undo */
(function () {
  'use strict';
  const E = FmEngine;
  const $ = s => document.querySelector(s);
  const App = window.App = { CMD: {}, hooks: {} };
  const S = App.S = {
    wb: new E.Workbook(), si: 0, sel: { r1: 0, c1: 0, r2: 0, c2: 0 }, act: { r: 0, c: 0 }, edit: null, clip: null,
    zoom: 1, undo: [], redo: [], name: 'Book1', handle: null, dirty: false, showHeadings: true, showFbar: true, chartSel: null, filterOpen: null
  };
  const DW = 84, DH = 24, HW = 50, HH = 26;
  const cur = () => S.wb.sheets[S.si];
  App.cur = cur;
  const cv = $('#cv'), ctx = cv.getContext('2d'), scroller = $('#scroller'), content = $('#content'), editor = $('#editor');
  let dpr = window.devicePixelRatio || 1;

  /* ---------- geometry ---------- */
  const geo = { sheet: null, ver: -1, cx: null, ry: null };
  function buildIdx(map, hidden, def) {
    const keys = [...new Set([...Object.keys(map), ...Object.keys(hidden)].map(Number))].sort((a, b) => a - b);
    const sizes = keys.map(k => hidden[k] ? 0 : (map[k] !== undefined ? map[k] : def)); const cum = [0];
    keys.forEach((k, i) => cum.push(cum[i] + sizes[i] - def));
    const sz = {}; keys.forEach((k, i) => sz[k] = sizes[i]);
    return { keys, cum, sz, def };
  }
  let geomVer = 0; App.geomChanged = () => { geomVer++; };
  function G() {
    const sh = cur();
    if (geo.sheet !== sh || geo.ver !== geomVer) { geo.sheet = sh; geo.ver = geomVer; geo.cx = buildIdx(sh.colW, sh.hiddenCols, DW); geo.ry = buildIdx(sh.rowH, sh.hiddenRows, DH); }
    return geo;
  }
  function pos(ix, i) { let lo = 0, hi = ix.keys.length; while (lo < hi) { const m = (lo + hi) >> 1; if (ix.keys[m] < i) lo = m + 1; else hi = m; } return i * ix.def + ix.cum[lo]; }
  function size(ix, i) { const v = ix.sz[i]; return v === undefined ? ix.def : v; }
  function at(ix, x, max) { let lo = 0, hi = max; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (pos(ix, m) <= x) lo = m; else hi = m - 1; } return lo; }
  const colX = c => pos(G().cx, c), rowY = r => pos(G().ry, r), colW = c => size(G().cx, c), rowH = r => size(G().ry, r);
  const colAt = x => at(G().cx, Math.max(0, x), E.MAXC - 1), rowAt = y => at(G().ry, Math.max(0, y), E.MAXR - 1);
  Object.assign(App, { colX, rowY, colW, rowH, colAt, rowAt, DW, DH, HW, HH });

  /* ---------- theme ---------- */
  const C = {}; function readTheme() { const cs = getComputedStyle(document.documentElement); ['grid', 'cell', 'hdr', 'hdr-text', 'hdr-sel', 'hdr-sel-text', 'sel', 'accent', 'text', 'line', 'muted'].forEach(k => C[k] = cs.getPropertyValue('--' + k).trim()); }
  App.readTheme = () => { readTheme(); draw(); };
  readTheme();
  const REFCOL = ['#2b6fd6', '#d6432b', '#8a3fc0', '#1f8a4c', '#c77d00', '#0a8fa6', '#b3367a', '#5a6b00'];

  /* ---------- selection helpers ---------- */
  const norm = s => ({ r1: Math.min(s.r1, s.r2), c1: Math.min(s.c1, s.c2), r2: Math.max(s.r1, s.r2), c2: Math.max(s.c1, s.c2) });
  App.selNorm = () => norm(S.sel);
  function mergeAt(r, c) { for (const m of cur().merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }
  function expandMerges(s) { let ch = true; s = norm(s); while (ch) { ch = false; for (const m of cur().merges) if (!(m.r1 > s.r2 || m.r2 < s.r1 || m.c1 > s.c2 || m.c2 < s.c1)) { if (m.r1 < s.r1) { s.r1 = m.r1; ch = true; } if (m.r2 > s.r2) { s.r2 = m.r2; ch = true; } if (m.c1 < s.c1) { s.c1 = m.c1; ch = true; } if (m.c2 > s.c2) { s.c2 = m.c2; ch = true; } } } return s; }
  App.mergeAt = mergeAt;
  function setSel(r1, c1, r2, c2, keepAct) {
    const cl = (v, m) => Math.max(0, Math.min(m, v));
    r1 = cl(r1, E.MAXR - 1); r2 = cl(r2, E.MAXR - 1); c1 = cl(c1, E.MAXC - 1); c2 = cl(c2, E.MAXC - 1);
    if (!keepAct) S.act = { r: r1, c: c1 };
    S.ext = { r: r2, c: c2 };
    S.sel = expandMerges({ r1, c1, r2, c2 });
    updateUI();
  }
  App.setSel = setSel;
  function activeRect() { const m = mergeAt(S.act.r, S.act.c); return m || { r1: S.act.r, c1: S.act.c, r2: S.act.r, c2: S.act.c }; }
  function selectCell(r, c, extend) {
    r = Math.max(0, Math.min(E.MAXR - 1, r)); c = Math.max(0, Math.min(E.MAXC - 1, c));
    if (extend) setSel(S.act.r, S.act.c, r, c, true); else { S.act = { r, c }; const m = mergeAt(r, c); if (m) { S.act = { r: m.r1, c: m.c1 }; S.ext = { r: m.r1, c: m.c1 }; S.sel = { ...m }; updateUI(); } else setSel(r, c, r, c); }
    scrollTo(r, c);
  }
  App.selectCell = selectCell;
  function viewport() { return { w: scroller.clientWidth / S.zoom, h: scroller.clientHeight / S.zoom }; }
  function scrollTo(r, c) {
    const sh = cur(), v = viewport(), z = S.zoom, fr = sh.freeze.r, fc = sh.freeze.c;
    const x0 = colX(c), x1 = x0 + colW(c), y0 = rowY(r), y1 = y0 + rowH(r);
    const fw = colX(fc), fh = rowY(fr); let sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z;
    if (c >= fc) { const l = fw + sx, rr = sx + v.w - HW; if (x0 < l) sx = x0 - fw; else if (x1 > rr) sx = x1 - v.w + HW + 16; }
    if (r >= fr) { const t = fh + sy, bb = sy + v.h - HH; if (y0 < t) sy = y0 - fh; else if (y1 > bb) sy = y1 - v.h + HH + 16; }
    scroller.scrollLeft = Math.max(0, sx * z); scroller.scrollTop = Math.max(0, sy * z);
  }
  App.scrollTo = scrollTo;

  /* ---------- style helpers ---------- */
  const px = pt => Math.round(pt * 4 / 3 * 10) / 10;
  function fontOf(s) { return (s && s.i ? 'italic ' : '') + (s && s.b ? '700 ' : '400 ') + px(s && s.fs || 11) + 'px ' + (s && s.ff ? '"' + s.ff + '",' : '') + '"Segoe UI",Calibri,system-ui,sans-serif'; }
  function wrapLines(text, maxW, font) {
    ctx.font = font; const out = [];
    for (const para of String(text).split('\n')) {
      let line = '';
      for (const w of para.split(/(\s+)/)) {
        const t = line + w;
        if (ctx.measureText(t).width > maxW && line) { out.push(line.trimEnd()); line = w.trimStart(); } else line = t;
        while (ctx.measureText(line).width > maxW && line.length > 1) { let k = line.length - 1; while (k > 1 && ctx.measureText(line.slice(0, k)).width > maxW) k--; out.push(line.slice(0, k)); line = line.slice(k); }
      }
      out.push(line);
    }
    return out;
  }
  App.wrapLines = wrapLines;

  /* ---------- conditional formatting ---------- */
  function cfFor(sh, si, r, c, v, disp) {
    let res = null;
    for (const rule of sh.cf) {
      if (r < rule.r1 || r > rule.r2 || c < rule.c1 || c > rule.c2) continue;
      const num = typeof v === 'number';
      let hit = false;
      switch (rule.type) {
        case 'gt': hit = num && v > rule.v; break; case 'lt': hit = num && v < rule.v; break; case 'eq': hit = (num && v === +rule.v) || String(disp).toLowerCase() === String(rule.v).toLowerCase(); break;
        case 'between': hit = num && v >= Math.min(rule.v, rule.v2) && v <= Math.max(rule.v, rule.v2); break;
        case 'text': hit = String(disp).toLowerCase().includes(String(rule.v).toLowerCase()); break;
        case 'dup': { hit = cfDup(sh, si, rule, disp); break; }
        case 'blank': hit = disp === ''; break;
        case 'scale': case 'bar': { if (!num) break; const st = cfStats(si, rule); if (st.max === st.min) break; const t = (v - st.min) / (st.max - st.min); if (rule.type === 'bar') { res = res || {}; res.bar = { t, color: rule.color || '#63be7b' }; } else { res = res || {}; res.bg = lerpColor(rule.c1c || '#f8696b', rule.c2c || '#63be7b', rule.c3c || '#ffeb84', t); } break; }
      }
      if (hit) { res = res || {}; if (rule.bg) res.bg = rule.bg; if (rule.fc) res.fc = rule.fc; if (rule.bold) res.b = 1; }
    }
    return res;
  }
  const cfCache = { ver: -1, m: new Map() };
  function cfStats(si, rule) {
    if (cfCache.ver !== S.wb.version) { cfCache.ver = S.wb.version; cfCache.m.clear(); }
    let st = cfCache.m.get(rule); if (st) return st; let min = Infinity, max = -Infinity;
    for (let r = rule.r1; r <= Math.min(rule.r2, rule.r1 + 5000); r++) for (let c = rule.c1; c <= rule.c2; c++) { const v = S.wb.getValue(si, r, c); if (typeof v === 'number') { if (v < min) min = v; if (v > max) max = v; } }
    st = { min, max, dups: null }; cfCache.m.set(rule, st); return st;
  }
  function cfDup(sh, si, rule, disp) {
    const st = cfStats(si, rule); if (!st.dups) { st.dups = new Map(); for (let r = rule.r1; r <= Math.min(rule.r2, rule.r1 + 5000); r++) for (let c = rule.c1; c <= rule.c2; c++) { const d = S.wb.display(si, r, c).toLowerCase(); if (d !== '') st.dups.set(d, (st.dups.get(d) || 0) + 1); } }
    return (st.dups.get(String(disp).toLowerCase()) || 0) > 1;
  }
  function lerpColor(a, b, m, t) {
    const h = x => [1, 3, 5].map(i => parseInt(x.slice(i, i + 2), 16)); const A = h(a), B = h(b), M = h(m);
    const mix = (p, q, k) => p.map((v, i) => Math.round(v + (q[i] - v) * k));
    const c = t < 0.5 ? mix(A, M, t * 2) : mix(M, B, (t - 0.5) * 2); return 'rgb(' + c.join(',') + ')';
  }

  /* ---------- rendering ---------- */
  let raf = 0;
  function draw() { if (!raf) raf = requestAnimationFrame(() => { raf = 0; render(); }); }
  App.draw = draw;
  function render() {
    const sh = cur(), z = S.zoom, W = scroller.clientWidth, H = scroller.clientHeight; if (!W || !H) return;
    dpr = window.devicePixelRatio || 1;
    if (cv.width !== Math.round(W * dpr) || cv.height !== Math.round(H * dpr)) { cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr); cv.style.width = W + 'px'; cv.style.height = H + 'px'; }
    ctx.setTransform(dpr * z, 0, 0, dpr * z, 0, 0);
    const lw = W / z, lh = H / z, sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z;
    ctx.fillStyle = C.cell; ctx.fillRect(0, 0, lw, lh);
    const hw = S.showHeadings ? HW : 0, hh = S.showHeadings ? HH : 0;
    const fr = sh.freeze.r, fc = sh.freeze.c, fw = colX(fc), fh = rowY(fr);
    const sel = norm(S.sel);
    const quads = [
      { c0: fc, c1: null, r0: fr, r1: null, sx, sy, clip: [hw + fw, hh + fh, lw, lh] },
      { c0: 0, c1: fc - 1, r0: fr, r1: null, sx: 0, sy, clip: [hw, hh + fh, hw + fw, lh] },
      { c0: fc, c1: null, r0: 0, r1: fr - 1, sx, sy: 0, clip: [hw + fw, hh, lw, hh + fh] },
      { c0: 0, c1: fc - 1, r0: 0, r1: fr - 1, sx: 0, sy: 0, clip: [hw, hh, hw + fw, hh + fh] }
    ];
    const editRefs = S.edit && String(S.edit.text).startsWith('=') ? refRects(S.edit.text) : [];
    for (const q of quads) {
      if ((q.c1 !== null && q.c1 < q.c0) || (q.r1 !== null && q.r1 < q.r0)) continue;
      ctx.save(); ctx.beginPath(); ctx.rect(q.clip[0], q.clip[1], q.clip[2] - q.clip[0], q.clip[3] - q.clip[1]); ctx.clip();
      const ox = hw - q.sx, oy = hh - q.sy;
      const c0 = q.c1 === null ? colAt(q.sx + colX(q.c0) * 0 + (q.c0 === 0 ? 0 : colX(q.c0))) : q.c0, r0 = q.r1 === null ? rowAt(q.sy + (q.r0 === 0 ? 0 : rowY(q.r0))) : q.r0;
      const cEnd = q.c1 === null ? colAt(q.sx + lw) + 1 : q.c1, rEnd = q.r1 === null ? rowAt(q.sy + lh) + 1 : q.r1;
      const cs = Math.max(c0, q.c0), rs = Math.max(r0, q.r0);
      drawCells(sh, cs, Math.min(cEnd, E.MAXC - 1), rs, Math.min(rEnd, E.MAXR - 1), ox, oy, q);
      // selection + edit refs
      for (const er of editRefs) { const m = mergeAt(er.r1, er.c1) || er; const x = ox + colX(er.c1), y = oy + rowY(er.r1), w = colX(er.c2 + 1) - colX(er.c1), h = rowY(er.r2 + 1) - rowY(er.r1); ctx.strokeStyle = er.color; ctx.lineWidth = 2; ctx.strokeRect(x + 1, y + 1, w - 2, h - 2); ctx.fillStyle = er.color + '18'; ctx.fillRect(x, y, w, h); }
      const x1 = ox + colX(sel.c1), y1 = oy + rowY(sel.r1), x2 = ox + colX(sel.c2 + 1), y2 = oy + rowY(sel.r2 + 1);
      if (!S.edit) {
        ctx.fillStyle = C.sel; ctx.fillRect(x1, y1, x2 - x1, y2 - y1);
        const ar = activeRect(); const ax = ox + colX(ar.c1), ay = oy + rowY(ar.r1), aw = colX(ar.c2 + 1) - colX(ar.c1), ah = rowY(ar.r2 + 1) - rowY(ar.r1);
        ctx.fillStyle = C.cell; ctx.fillRect(ax, ay, aw, ah); drawCellAt(sh, ar.r1, ar.c1, ax, ay, aw, ah, true);
        ctx.strokeStyle = C.accent; ctx.lineWidth = 2; ctx.strokeRect(x1, y1, x2 - x1, y2 - y1);
        ctx.fillStyle = C.accent; ctx.strokeStyle = C.cell; ctx.lineWidth = 1.5; ctx.fillRect(x2 - 4, y2 - 4, 8, 8); ctx.strokeRect(x2 - 4, y2 - 4, 8, 8);
      }
      if (S.clip && S.clip.si === S.si) { const k = S.clip; const cx1 = ox + colX(k.c1), cy1 = oy + rowY(k.r1), cx2 = ox + colX(k.c1 + k.cols), cy2 = oy + rowY(k.r1 + k.rows); ctx.save(); ctx.setLineDash([5, 4]); ctx.lineDashOffset = -(performance.now() / 60 | 0) % 9; ctx.strokeStyle = C.accent; ctx.lineWidth = 1.6; ctx.strokeRect(cx1, cy1, cx2 - cx1, cy2 - cy1); ctx.restore(); }
      ctx.restore();
    }
    if (S.showHeadings) drawHeaders(sh, lw, lh, sx, sy, fc, fr, fw, fh, sel);
    // freeze lines
    if (fc || fr) { ctx.strokeStyle = C.muted; ctx.lineWidth = 1.2; ctx.beginPath(); if (fc) { ctx.moveTo(hw + fw, hh); ctx.lineTo(hw + fw, lh); } if (fr) { ctx.moveTo(hw, hh + fh); ctx.lineTo(lw, hh + fh); } ctx.stroke(); }
    positionEditor(); positionFilters();
    if (S.clip) { clearTimeout(App._marq); App._marq = setTimeout(draw, 120); }
    const totW = (colX(E.MAXC) + HW) * z, totH = (rowY(Math.min(E.MAXR, 1048576)) + HH) * z;
    content.style.width = totW + 'px'; content.style.height = totH + 'px';
  }
  function drawCells(sh, c0, c1, r0, r1, ox, oy, q) {
    if (c1 < c0 || r1 < r0) return;
    const x0 = ox + colX(c0), y0 = oy + rowY(r0), x1 = ox + colX(c1 + 1), y1 = oy + rowY(r1 + 1);
    if (sh.showGrid) {
      ctx.strokeStyle = C.grid; ctx.lineWidth = 1; ctx.beginPath();
      for (let c = c0; c <= c1 + 1; c++) { const x = Math.round(ox + colX(c)) + .5; ctx.moveTo(x, y0); ctx.lineTo(x, y1); }
      for (let r = r0; r <= r1 + 1; r++) { const y = Math.round(oy + rowY(r)) + .5; ctx.moveTo(x0, y); ctx.lineTo(x1, y); }
      ctx.stroke();
    }
    const drawn = new Set();
    for (let r = r0; r <= r1; r++) {
      const h = rowH(r); if (!h) continue;
      for (let c = c0; c <= c1; c++) {
        const w = colW(c); if (!w) continue;
        let m = null; if (sh.merges.length) { m = mergeAt(r, c); if (m) { const k = m.r1 * 20000 + m.c1; if (drawn.has(k)) continue; drawn.add(k); if (r !== m.r1 || c !== m.c1) { /* origin off-screen: draw from origin position */ } } }
        const rr = m ? m.r1 : r, cc = m ? m.c1 : c;
        const x = ox + colX(cc), y = oy + rowY(rr), ww = m ? colX(m.c2 + 1) - colX(m.c1) : w, hh = m ? rowY(m.r2 + 1) - rowY(m.r1) : h;
        drawCellAt(sh, rr, cc, x, y, ww, hh, false, c0, c1);
      }
    }
  }
  function autoInk(bg) {
    let m = /^#([0-9a-f]{6})$/i.exec(bg); if (!m) { m = /^rgb\((\d+),(\d+),(\d+)\)$/.exec(bg); if (!m) return C.text; const l = (0.299 * m[1] + 0.587 * m[2] + 0.114 * m[3]) / 255; return l > 0.6 ? '#1d2430' : '#ffffff'; }
    const n = parseInt(m[1], 16), l = (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255; return l > 0.6 ? '#1d2430' : '#ffffff';
  }
  function drawCellAt(sh, r, c, x, y, w, h, force, vc0, vc1) {
    const cell = sh.cells.get(r * 20000 + c); const st = cell && cell.s;
    const spill = !(cell && cell.t !== '') && S.wb.isSpill(S.si, r, c);
    const v = (cell && cell.t !== '') || spill ? S.wb.getValue(S.si, r, c) : null;
    const disp = (cell && cell.t !== '') || spill ? S.wb.display(S.si, r, c) : '';
    const key = r * 20000 + c; const isLink = !!sh.links[key] || (cell && typeof cell.t === 'string' && /^=\s*HYPERLINK\(/i.test(cell.t));
    const cf = sh.cf.length ? cfFor(sh, S.si, r, c, v, disp) : null;
    const bg = (cf && cf.bg) || (st && st.bg);
    const isMerge = !!mergeAt(r, c);
    if (bg) { ctx.fillStyle = bg; ctx.fillRect(x, y, w, h); }
    else if (isMerge || force) { ctx.fillStyle = C.cell; ctx.fillRect(x, y, w, h); }
    if (cf && cf.bar) { ctx.fillStyle = cf.bar.color; ctx.fillRect(x + 2, y + 3, Math.max(2, (w - 4) * (0.08 + 0.92 * cf.bar.t)), h - 6); ctx.globalAlpha = 1; }
    if (st && st.bd) {
      const b = st.bd; const line = (s, x1, y1, x2, y2) => { ctx.strokeStyle = s.c || '#000'; ctx.lineWidth = s.w === 2 ? 2 : 1; ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke(); };
      if (b.t) line(b.t, x, y + .5, x + w, y + .5); if (b.b) line(b.b, x, y + h - .5, x + w, y + h - .5); if (b.l) line(b.l, x + .5, y, x + .5, y + h); if (b.r) line(b.r, x + w - .5, y, x + w - .5, y + h);
    }
    if (sh.notes[key]) { ctx.fillStyle = '#e5484d'; ctx.beginPath(); ctx.moveTo(x + w - 7, y); ctx.lineTo(x + w, y); ctx.lineTo(x + w, y + 7); ctx.fill(); }
    if ((!cell && !spill) || disp === '') return;
    const font = fontOf(st && cf && cf.b ? { ...st, b: 1 } : cf && cf.b ? { b: 1 } : st); ctx.font = font;
    let color = (cf && cf.fc) || (st && st.fc) || (bg ? autoInk(bg) : C.text);
    if (isLink && !(st && st.fc)) color = '#2b6fd6';
    if (v instanceof E.FErr) color = '#d6432b';
    const isNum = typeof v === 'number', isStr = typeof v === 'string';
    let ha = st && st.ha; if (!ha || ha === 'general') ha = isStr ? 'left' : 'right'; if (typeof v === 'boolean' || v instanceof E.FErr) ha = (st && st.ha && st.ha !== 'general') ? st.ha : 'center';
    const va = (st && st.va) || 'bottom'; const fs = px(st && st.fs || 11); const pad = 5 + (st && st.ind ? st.ind * 10 : 0);
    ctx.fillStyle = color; ctx.textBaseline = 'alphabetic';
    let lines, lh = fs * 1.25;
    if (st && st.wrap) lines = wrapLines(disp, w - pad * 2, font); else lines = [disp.replace(/\n/g, ' ')];
    let tw = ctx.measureText(lines[0]).width;
    if (isNum && !st?.wrap && tw > w - 6 && !(st && st.nf && /[a-z]/i.test(st.nf) && false)) { lines = ['#'.repeat(Math.max(1, Math.floor((w - 6) / ctx.measureText('#').width)))]; tw = ctx.measureText(lines[0]).width; }
    ctx.save(); let cx1 = x, cx2 = x + w;
    if (!(st && st.wrap) && (tw > w - pad * 2) && !isNum) {
      // overflow into empty neighbours
      if (ha === 'left') { let cc = c + (mergeAt(r, c) ? mergeAt(r, c).c2 - c : 0) + 1, acc = w; while (acc < tw + pad * 2 && cc < c + 60) { const n = sh.cells.get(r * 20000 + cc); if (n && n.t !== '') break; acc += colW(cc); cc++; } cx2 = x + Math.min(acc, tw + pad * 2 + 2); }
      else if (ha === 'right') { let cc = c - 1, acc = w; while (acc < tw + pad * 2 && cc >= 0) { const n = sh.cells.get(r * 20000 + cc); if (n && n.t !== '') break; acc += colW(cc); cc--; } cx1 = x + w - Math.min(acc, tw + pad * 2 + 2); }
    }
    ctx.beginPath(); ctx.rect(cx1, y, cx2 - cx1, h); ctx.clip();
    const total = lines.length * lh; let ty = va === 'top' ? y + 3 + fs : va === 'middle' ? y + (h - total) / 2 + fs : y + h - 4 - (lines.length - 1) * lh;
    if (va === 'bottom' && st && st.wrap) ty = y + h - 4 - (lines.length - 1) * lh;
    for (const ln of lines) {
      const lwid = ctx.measureText(ln).width; let tx = ha === 'left' ? x + pad : ha === 'center' ? x + (w - lwid) / 2 : x + w - pad - lwid;
      ctx.fillText(ln, tx, ty - 1);
      if ((st && (st.u || st.st)) || isLink) { ctx.fillRect(tx, (st && st.st && !st.u && !isLink) ? ty - fs * 0.32 : ty + 1, lwid, 1); }
      ty += lh;
    }
    ctx.restore();
  }
  function drawHeaders(sh, lw, lh, sx, sy, fc, fr, fw, fh, sel) {
    ctx.fillStyle = C.hdr; ctx.fillRect(0, 0, lw, HH); ctx.fillRect(0, 0, HW, lh);
    ctx.font = '600 12px "Segoe UI",system-ui,sans-serif'; ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
    const drawCol = (c, x) => {
      const w = colW(c); if (!w) return; const on = c >= sel.c1 && c <= sel.c2; const full = sel.r1 === 0 && sel.r2 >= 1048575;
      if (on) { ctx.fillStyle = full ? C.accent : C['hdr-sel']; ctx.fillRect(x, 0, w, HH); }
      ctx.fillStyle = on ? (full ? '#fff' : C['hdr-sel-text']) : C['hdr-text']; ctx.fillText(E.numToCol(c), x + w / 2, HH / 2 + 1);
      ctx.fillStyle = C.line; ctx.fillRect(x + w - 1, 0, 1, HH);
    };
    const drawRow = (r, y) => {
      const h = rowH(r); if (!h) return; const on = r >= sel.r1 && r <= sel.r2; const full = sel.c1 === 0 && sel.c2 >= 16383;
      if (on) { ctx.fillStyle = full ? C.accent : C['hdr-sel']; ctx.fillRect(0, y, HW, h); }
      ctx.fillStyle = on ? (full ? '#fff' : C['hdr-sel-text']) : C['hdr-text']; ctx.fillText(String(r + 1), HW / 2, y + h / 2 + 1);
      ctx.fillStyle = C.line; ctx.fillRect(0, y + h - 1, HW, 1);
    };
    ctx.save(); ctx.beginPath(); ctx.rect(HW + fw, 0, lw, HH); ctx.clip();
    for (let c = Math.max(fc, colAt(sx + fw)); c <= colAt(sx + lw) + 1 && c < E.MAXC; c++) drawCol(c, HW + colX(c) - sx); ctx.restore();
    for (let c = 0; c < fc; c++) drawCol(c, HW + colX(c));
    ctx.save(); ctx.beginPath(); ctx.rect(0, HH + fh, HW, lh); ctx.clip();
    for (let r = Math.max(fr, rowAt(sy + fh)); r <= rowAt(sy + lh) + 1 && r < E.MAXR; r++) drawRow(r, HH + rowY(r) - sy); ctx.restore();
    for (let r = 0; r < fr; r++) drawRow(r, HH + rowY(r));
    ctx.fillStyle = C.hdr; ctx.fillRect(0, 0, HW, HH); ctx.fillStyle = C.line; ctx.fillRect(0, HH - 1, lw, 1); ctx.fillRect(HW - 1, 0, 1, lh);
    ctx.fillStyle = C.muted; ctx.beginPath(); ctx.moveTo(HW - 6, HH - 6); ctx.lineTo(HW - 6, HH - 16); ctx.lineTo(HW - 16, HH - 6); ctx.fill();
    ctx.textAlign = 'start'; ctx.textBaseline = 'alphabetic';
  }
  function refRects(text) {
    const out = []; let toks; try { toks = E.tokenize(text.slice(1)); } catch (e) { return out; }
    let k = 0; for (const t of toks) if (t.t === 'ref') {
      try { const p = E.parseRefToken(t.v); if (p.sheet && E.Workbook && S.wb.sheetIndex(p.sheet) !== S.si) { k++; continue; } const a = p.a, b = p.b || p.a; out.push({ r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c), color: REFCOL[k % REFCOL.length] }); } catch (e) { } k++;
    }
    return out;
  }
  function screenRect(r1, c1, r2, c2) {
    const sh = cur(), z = S.zoom, hw = S.showHeadings ? HW : 0, hh = S.showHeadings ? HH : 0;
    const frozenC = c1 < sh.freeze.c, frozenR = r1 < sh.freeze.r; const sx = frozenC ? 0 : scroller.scrollLeft / z, sy = frozenR ? 0 : scroller.scrollTop / z;
    return { x: (hw + colX(c1) - sx), y: (hh + rowY(r1) - sy), w: colX(c2 + 1) - colX(c1), h: rowY(r2 + 1) - rowY(r1) };
  }
  App.screenRect = screenRect;

  /* ---------- editing ---------- */
  const FNAMES = [...new Set([...Object.keys(E.FN), ...Object.keys(E.LAZY)])].sort();
  App.FNAMES = FNAMES;
  function startEdit(r, c, mode, initial) {
    if (S.edit) commitEdit(0, 0);
    const m = mergeAt(r, c); if (m) { r = m.r1; c = m.c1; }
    S.act = { r, c }; const t = initial !== undefined ? initial : S.wb.getInput(S.si, r, c);
    S.edit = { r, c, text: t, mode, orig: S.wb.getInput(S.si, r, c) };
    editor.classList.remove('hidden'); editor.value = t; scrollTo(r, c); positionEditor(); editor.focus();
    if (initial === undefined) editor.setSelectionRange(t.length, t.length);
    $('#fbar').value = t; draw(); acUpdate();
  }
  App.startEdit = startEdit;
  function positionEditor() {
    if (!S.edit) return;
    const { r, c } = S.edit; const m = mergeAt(r, c) || { r1: r, c1: c, r2: r, c2: c };
    const sr = screenRect(m.r1, m.c1, m.r2, m.c2), z = S.zoom; const st = S.wb.getStyle(S.si, r, c);
    const lines = editor.value.split('\n'); let wch = Math.max(...lines.map(l => l.length)) + 2;
    const tmp = ctx; tmp.font = fontOf(st); let tw = 0; for (const l of lines) tw = Math.max(tw, tmp.measureText(l).width); // ctx now has editor font
    const needW = Math.max(sr.w, tw + 24); const needH = Math.max(sr.h, lines.length * px(st && st.fs || 11) * 1.3 + 6);
    editor.style.left = (sr.x * z + scroller.scrollLeft * 0) + 'px';
    const wrapX = scroller.scrollLeft, wrapY = scroller.scrollTop;
    const fx = m.c1 < cur().freeze.c ? 0 : wrapX, fy = m.r1 < cur().freeze.r ? 0 : wrapY;
    editor.style.left = (sr.x * z + fx) + 'px'; editor.style.top = (sr.y * z + fy) + 'px';
    editor.style.width = needW * z + 'px'; editor.style.height = needH * z + 'px';
    editor.style.font = fontOf(st).replace(/(\d+(\.\d+)?)px/, (m0, n) => (parseFloat(n) * z) + 'px'); editor.style.color = (st && st.fc) || C.text; editor.style.background = (st && st.bg) || C.cell;
    editor.style.textAlign = (st && st.ha && st.ha !== 'general') ? st.ha : 'left';
    const ab = $('#acbox'); if (!ab.classList.contains('hidden')) { ab.style.left = parseFloat(editor.style.left) + 'px'; ab.style.top = (parseFloat(editor.style.top) + needH * z + 2) + 'px'; }
  }
  function balance(t) { if (!t.startsWith('=')) return t; let d = 0, q = false; for (const ch of t) { if (ch === '"') q = !q; else if (!q) { if (ch === '(') d++; else if (ch === ')') d--; } } return d > 0 && !q ? t + ')'.repeat(d) : t; }
  function dvAt(sh, r, c) { for (const d of sh.dv) if (r >= d.r1 && r <= d.r2 && c >= d.c1 && c <= d.c2) return d; return null; }
  function dvList(d) {
    const src = String(d.src || ''); if (src[0] === '=') { try { const p = E.parseRefToken(src.slice(1).replace(/^=/, '')); let si = p.sheet ? S.wb.sheetIndex(p.sheet) : S.si; if (p.names) return []; const b = p.b || p.a; if (p.a.row || p.a.col) return []; const out = []; for (let r = Math.min(p.a.r, b.r); r <= Math.max(p.a.r, b.r); r++) for (let c = Math.min(p.a.c, b.c); c <= Math.max(p.a.c, b.c); c++) { const x = S.wb.display(si, r, c); if (x !== '') out.push(x); } return out; } catch (e) { const nm = S.wb.names[src.slice(1).toUpperCase()]; if (nm) { try { const p = E.parseRefToken(nm); const si = p.sheet ? S.wb.sheetIndex(p.sheet) : S.si; const b = p.b || p.a; const out = []; for (let r = p.a.r; r <= b.r; r++) for (let c = p.a.c; c <= b.c; c++) { const x = S.wb.display(si, r, c); if (x !== '') out.push(x); } return out; } catch (e2) { } } return []; } }
    return src.split(',').map(x => x.trim()).filter(x => x !== '');
  }
  function dvCheck(d, text) {
    const bad = { style: d.style || 'stop', title: d.errTitle || 'Invalid entry', msg: d.errMsg || 'This value doesn’t match the data validation rules set for this cell.' };
    if (text === '') return d.allowBlank === false ? bad : null; if (text[0] === '=') return null;
    if (d.type === 'any') return null;
    if (d.type === 'list') return dvList(d).some(x => x.toLowerCase() === text.toLowerCase()) ? null : bad;
    const lit = E.parseLiteral(text).v; let val;
    if (d.type === 'textLength') val = text.length; else { if (typeof lit !== 'number') return bad; val = lit; if (d.type === 'whole' && !Number.isInteger(val)) return bad; }
    const num = x => { const l = E.parseLiteral(String(x)).v; return typeof l === 'number' ? l : parseFloat(x); }; const v1 = num(d.v1), v2 = num(d.v2);
    const ok = { between: val >= Math.min(v1, v2) && val <= Math.max(v1, v2), notBetween: val < Math.min(v1, v2) || val > Math.max(v1, v2), eq: val === v1, ne: val !== v1, gt: val > v1, lt: val < v1, ge: val >= v1, le: val <= v1 }[d.op || 'between'];
    return ok ? null : bad;
  }
  App.dvAt = dvAt; App.dvList = dvList; App.dvCheck = dvCheck;
  function commitEdit(dr, dc, noMove) {
    if (!S.edit) return; const e = S.edit; let text = editor.value; S.edit = null; editor.classList.add('hidden'); acHide();
    const finish = () => { scroller.focus({ preventScroll: true }); if (!noMove) { const m = mergeAt(e.r, e.c); selectCell((dr > 0 && m ? m.r2 : e.r) + dr, (dc > 0 && m ? m.c2 : e.c) + dc); } updateUI(); draw(); };
    const apply = () => { snap(); text = balance(text); S.wb.setInput(S.si, e.r, e.c, text); const st = S.wb.getStyle(S.si, e.r, e.c); if (st && st.wrap) App.CMD.autofitRow(e.r, true); afterChange(); };
    if (text !== e.orig) {
      const dv = dvAt(cur(), e.r, e.c); const bad = dv ? dvCheck(dv, text) : null;
      if (bad) {
        if (bad.style === 'stop') { App.dialog(bad.title, `<div>${bad.msg.replace(/</g, '&lt;')}</div>`, [{ label: 'Cancel' }, { label: 'Retry', primary: 1, fn: () => { setTimeout(() => startEdit(e.r, e.c, 'edit', text), 0); } }]); updateUI(); draw(); return; }
        if (bad.style === 'warning') { App.dialog(bad.title, `<div>${bad.msg.replace(/</g, '&lt;')}</div><div>Continue anyway?</div>`, [{ label: 'No' }, { label: 'Yes', primary: 1, fn: () => { apply(); finish(); } }]); updateUI(); draw(); return; }
        toast(bad.msg);
      }
      apply();
    }
    finish();
  }
  function cancelEdit() { if (!S.edit) return; S.edit = null; editor.classList.add('hidden'); acHide(); updateUI(); draw(); scroller.focus({ preventScroll: true }); }
  App.commitEdit = commitEdit; App.cancelEdit = cancelEdit;
  function canPoint(t, caret) { if (!t.startsWith('=')) return false; const p = t.slice(0, caret).replace(/\s+$/, ''); return p.length > 0 && /[=+\-*\/^&(,<>;:%]$|^=$/.test(p) && !/%$/.test(p) || /^=$/.test(p); }
  function insertRef(txt) {
    const caret = editor.selectionStart; const before = editor.value.slice(0, caret); const after = editor.value.slice(editor.selectionEnd);
    // replace a previously pointed reference
    const pr = S.edit.point; let b2 = before; if (pr && before.endsWith(pr)) b2 = before.slice(0, before.length - pr.length);
    editor.value = b2 + txt + after; S.edit.point = txt; const cp = (b2 + txt).length; editor.setSelectionRange(cp, cp); S.edit.text = editor.value; $('#fbar').value = editor.value; draw();
  }
  editor.addEventListener('input', () => { if (!S.edit) return; S.edit.text = editor.value; S.edit.point = null; $('#fbar').value = editor.value; positionEditor(); acUpdate(); draw(); });
  editor.addEventListener('keydown', e => {
    e.stopPropagation();
    if (!S.edit) return;
    const ab = $('#acbox');
    if (!ab.classList.contains('hidden')) {
      const items = [...ab.querySelectorAll('.mi')]; let i = items.findIndex(x => x.classList.contains('hl'));
      if (e.key === 'ArrowDown') { e.preventDefault(); items[i]?.classList.remove('hl'); items[(i + 1) % items.length].classList.add('hl'); items[(i + 1) % items.length].scrollIntoView({ block: 'nearest' }); return; }
      if (e.key === 'ArrowUp') { e.preventDefault(); items[i]?.classList.remove('hl'); items[(i - 1 + items.length) % items.length].classList.add('hl'); return; }
      if (e.key === 'Tab' || (e.key === 'Enter' && i >= 0 && ab.dataset.explicit)) { e.preventDefault(); acPick(items[Math.max(0, i)].dataset.fn); return; }
      if (e.key === 'Escape') { acHide(); e.preventDefault(); return; }
    }
    if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); commitEdit(e.shiftKey ? -1 : 1, 0); }
    else if (e.key === 'Enter' && e.altKey) { e.preventDefault(); const p = editor.selectionStart; editor.setRangeText('\n', p, editor.selectionEnd, 'end'); editor.dispatchEvent(new Event('input')); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (e.key === 'F2') { S.edit.mode = S.edit.mode === 'enter' ? 'edit' : 'enter'; e.preventDefault(); }
    else if (e.key === 'F4') { e.preventDefault(); cycleAbs(); }
    else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      if (S.edit.mode === 'enter' || (S.edit.point && S.edit.mode !== 'edit')) {
        const caret = editor.selectionStart;
        if (canPoint(editor.value, caret) || S.edit.point) {
          e.preventDefault(); let p = S.edit.pointCell || { r: S.edit.r, c: S.edit.c }; const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
          p = { r: Math.max(0, p.r + d[0]), c: Math.max(0, p.c + d[1]) }; S.edit.pointCell = p; insertRef(E.addrToStr(p.r, p.c)); scrollTo(p.r, p.c); return;
        }
        if (!e.shiftKey && !e.ctrlKey) { e.preventDefault(); const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key]; commitEdit(d[0], d[1]); }
      }
    }
    else if (e.ctrlKey && (e.key === 'b' || e.key === 'i' || e.key === 'u')) { /* allow style shortcuts after edit */ }
  });
  editor.addEventListener('keyup', () => { if (S.edit) { S.edit.pointCell = null; } });
  editor.addEventListener('click', () => { if (S.edit) { S.edit.point = null; S.edit.pointCell = null; } });
  function cycleAbs() {
    const t = editor.value, caret = editor.selectionStart; const re = /\$?[A-Za-z]{1,3}\$?\d+/g; let m;
    while ((m = re.exec(t))) { if (caret >= m.index && caret <= m.index + m[0].length) { const p = E.parseAddr(m[0]); let nx; if (!p.ac && !p.ar) nx = [1, 1]; else if (p.ac && p.ar) nx = [0, 1]; else if (!p.ac && p.ar) nx = [1, 0]; else nx = [0, 0]; const s = (nx[0] ? '$' : '') + E.numToCol(p.c) + (nx[1] ? '$' : '') + (p.r + 1); editor.value = t.slice(0, m.index) + s + t.slice(m.index + m[0].length); editor.setSelectionRange(m.index + s.length, m.index + s.length); editor.dispatchEvent(new Event('input')); return; } }
  }
  /* function autocomplete */
  function acHide() { const ab = $('#acbox'); ab.classList.add('hidden'); ab.dataset.explicit = ''; }
  function acUpdate() {
    const ab = $('#acbox'); if (!S.edit || !editor.value.startsWith('=')) return acHide();
    const caret = editor.selectionStart, before = editor.value.slice(0, caret); const m = /([A-Za-z][A-Za-z0-9_.]*)$/.exec(before);
    if (!m || /\d/.test(m[1][0])) return acHide();
    const pre = m[1].toUpperCase(); const list = FNAMES.filter(f => f.startsWith(pre)).slice(0, 12);
    if (!list.length || (list.length === 1 && list[0] === pre && editor.value[caret] === '(')) return acHide();
    ab.innerHTML = list.map((f, i) => `<button class="mi${i === 0 ? ' hl' : ''}" data-fn="${f}"><b>${f}</b><span class="k">${(App.FDOC[f] || '').split(' — ')[0] || ''}</span></button>`).join(''); ab.dataset.explicit = '1';
    ab.classList.remove('hidden'); ab.querySelectorAll('.mi').forEach(b => b.onmousedown = ev => { ev.preventDefault(); acPick(b.dataset.fn); }); positionEditor();
  }
  function acPick(fn) {
    const caret = editor.selectionStart, before = editor.value.slice(0, caret), m = /([A-Za-z][A-Za-z0-9_.]*)$/.exec(before);
    const nb = before.slice(0, before.length - m[1].length) + fn + '('; const after = editor.value.slice(caret); editor.value = nb + after; editor.setSelectionRange(nb.length, nb.length); acHide(); editor.dispatchEvent(new Event('input'));
  }
  App.FDOC = {};

  /* formula bar & name box */
  const fbar = $('#fbar'), namebox = $('#namebox');
  fbar.addEventListener('focus', () => { if (!S.edit) { startEdit(S.act.r, S.act.c, 'edit'); fbar.focus(); } });
  fbar.addEventListener('input', () => { if (!S.edit) return; editor.value = fbar.value; S.edit.text = fbar.value; positionEditor(); acUpdate(); draw(); });
  fbar.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); commitEdit(1, 0); } else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
    else if (e.key === 'Tab') { e.preventDefault(); commitEdit(0, 1); }
  });
  namebox.addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); gotoRef(namebox.value); scroller.focus(); } else if (e.key === 'Escape') { updateUI(); scroller.focus(); }
  });
  namebox.addEventListener('focus', () => namebox.select());
  function gotoRef(s) {
    s = s.trim(); try {
      const nmv = S.wb.names[s.toUpperCase()]; if (nmv && /^[^!]*!?\$?[A-Z]/i.test(nmv)) s = nmv.replace(/^=/, '');
      let sheetIdx = S.si; if (s.includes('!')) { const i = S.wb.sheetIndex(s.split('!')[0].replace(/^'|'$/g, '')); if (i >= 0) { App.switchSheet(i); } s = s.split('!')[1]; }
      const p = E.parseRefToken(s); const a = p.a, b = p.b || p.a;
      if (a.col) setSel(0, a.c, E.MAXR - 1, b.c); else if (a.row) setSel(a.r, 0, b.r, E.MAXC - 1); else { setSel(a.r, a.c, b.r, b.c); } scrollTo(a.r || 0, a.c || 0);
    } catch (e) { toast('Reference isn’t valid'); updateUI(); }
  }

  /* ---------- undo ---------- */
  function snap() { S.undo.push(JSON.stringify(S.wb.toJSON()) + '|' + S.si); if (S.undo.length > 60) S.undo.shift(); S.redo = []; S.dirty = true; }
  function restore(str) { const i = str.lastIndexOf('|'); const si = +str.slice(i + 1); S.wb = E.Workbook.fromJSON(JSON.parse(str.slice(0, i))); S.si = Math.min(si, S.wb.sheets.length - 1); App.geomChanged(); }
  function undo() { if (S.edit) return cancelEdit(); if (!S.undo.length) return; S.redo.push(JSON.stringify(S.wb.toJSON()) + '|' + S.si); restore(S.undo.pop()); afterChange(true); }
  function redo() { if (!S.redo.length) return; S.undo.push(JSON.stringify(S.wb.toJSON()) + '|' + S.si); restore(S.redo.pop()); afterChange(true); }
  function afterChange(structural) { App.geomChanged(); S.dirty = true; if (structural) { App.renderSheetTabs && App.renderSheetTabs(); } App.renderCharts && App.renderCharts(); updateUI(); draw(); scheduleAutosave(); }
  Object.assign(App, { snap, undo, redo, afterChange });
  App.act = fn => { snap(); fn(); afterChange(); };

  /* ---------- UI sync ---------- */
  function updateUI() {
    const n = norm(S.sel); const one = n.r1 === n.r2 && n.c1 === n.c2 || (mergeAt(n.r1, n.c1) && mergeAt(n.r1, n.c1).r2 === n.r2 && mergeAt(n.r1, n.c1).c2 === n.c2);
    if (document.activeElement !== namebox) namebox.value = one ? E.addrToStr(S.act.r, S.act.c) : (n.r1 === 0 && n.r2 >= 1048575 ? E.numToCol(n.c1) + ':' + E.numToCol(n.c2) : n.c1 === 0 && n.c2 >= 16383 ? (n.r1 + 1) + ':' + (n.r2 + 1) : E.addrToStr(n.r1, n.c1) + ':' + E.addrToStr(n.r2, n.c2));
    if (!S.edit && document.activeElement !== fbar) fbar.value = S.wb.getInput(S.si, S.act.r, S.act.c);
    updateStats(n); App.hooks.sel && App.hooks.sel(); draw();
  }
  App.updateUI = updateUI;
  function updateStats(n) {
    const el = $('#stats'); if (n.r1 === n.r2 && n.c1 === n.c2) { el.innerHTML = ''; return; }
    let cnt = 0, sum = 0, min = Infinity, max = -Infinity, nonEmpty = 0; const rr = Math.min(n.r2, n.r1 + 20000), cc = Math.min(n.c2, n.c1 + 200);
    if ((rr - n.r1 + 1) * (cc - n.c1 + 1) > 300000) { el.innerHTML = ''; return; }
    for (let r = n.r1; r <= rr; r++) for (let c = n.c1; c <= cc; c++) { const v = S.wb.getValue(S.si, r, c); if (v !== null && v !== '') nonEmpty++; if (typeof v === 'number') { cnt++; sum += v; if (v < min) min = v; if (v > max) max = v; } }
    const f = x => E.formatValue(Math.round(x * 1e6) / 1e6);
    el.innerHTML = cnt ? `<span><b>Average</b> ${f(sum / cnt)}</span><span><b>Count</b> ${nonEmpty}</span><span><b>Min</b> ${f(min)}</span><span><b>Max</b> ${f(max)}</span><span><b>Sum</b> ${f(sum)}</span>` : `<span><b>Count</b> ${nonEmpty}</span>`;
  }
  let toastT; function toast(msg) { const t = $('#toast'); t.textContent = msg; t.classList.remove('hidden'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.add('hidden'), 2600); }
  App.toast = toast;

  /* ---------- mouse ---------- */
  function hit(ev) {
    const rc = scroller.getBoundingClientRect(), z = S.zoom, sh = cur(); const px_ = (ev.clientX - rc.left) / z, py_ = (ev.clientY - rc.top) / z;
    const hw = S.showHeadings ? HW : 0, hh = S.showHeadings ? HH : 0; const fw = colX(sh.freeze.c), fh = rowY(sh.freeze.r);
    const sx = scroller.scrollLeft / z, sy = scroller.scrollTop / z; const out = { px: px_, py: py_ };
    const gx = px_ - hw, gy = py_ - hh; const inFx = gx < fw && sh.freeze.c, inFy = gy < fh && sh.freeze.r;
    const lx = inFx ? gx : gx + sx, ly = inFy ? gy : gy + sy;
    out.c = colAt(lx); out.r = rowAt(ly); out.lx = lx; out.ly = ly;
    out.zone = (px_ < hw && py_ < hh) ? 'corner' : py_ < hh ? 'colhdr' : px_ < hw ? 'rowhdr' : 'cell';
    return out;
  }
  function linkAt(r, c) { const sh = cur(); const l = sh.links[r * 20000 + c]; if (l) return l; const cell = sh.cells.get(r * 20000 + c); if (cell && typeof cell.t === 'string') { const m = /^=\s*HYPERLINK\(\s*"((?:[^"]|"")*)"/i.exec(cell.t); if (m) return m[1].replace(/""/g, '"'); } return null; }
  App.linkAt = linkAt;
  App.openUrl = url => { if (!/^(https?:|mailto:)/i.test(url)) { if (/^[\w.-]+\.[a-z]{2,}/i.test(url)) url = 'https://' + url; else { toast('Only web and email links can be opened'); return; } } if (window.__TAURI__ && window.__TAURI__.core) window.__TAURI__.core.invoke('open_url', { url }).catch(() => window.open(url, '_blank')); else window.open(url, '_blank', 'noopener'); };
  let tipEl = null; function noteTip(ev) { const h = hit(ev); let t = null; if (h.zone === 'cell') { const n = cur().notes[h.r * 20000 + h.c]; if (n) t = n; else if (linkAt(h.r, h.c)) t = linkAt(h.r, h.c) + '\nCtrl+click to open'; } if (!t) { if (tipEl) tipEl.style.display = 'none'; return; } if (!tipEl) { tipEl = document.createElement('div'); tipEl.style.cssText = 'position:fixed;z-index:60;max-width:280px;background:#fff8c5;color:#3b3300;border:1px solid #d9c84f;border-radius:8px;padding:7px 10px;box-shadow:0 4px 14px rgba(0,0,0,.2);white-space:pre-wrap;pointer-events:none;font-size:12px'; document.body.appendChild(tipEl); } tipEl.textContent = t; tipEl.style.display = 'block'; tipEl.style.left = (ev.clientX + 14) + 'px'; tipEl.style.top = (ev.clientY + 14) + 'px'; }
  let drag = null;
  scroller.addEventListener('mousedown', ev => {
    if (ev.target !== cv && ev.target !== content && ev.target !== scroller) return;
    if (ev.button === 2) { const h = hit(ev); const n = norm(S.sel); if (h.zone === 'cell' && !(h.r >= n.r1 && h.r <= n.r2 && h.c >= n.c1 && h.c <= n.c2)) selectCell(h.r, h.c); else if (h.zone === 'colhdr' && !(h.c >= n.c1 && h.c <= n.c2 && n.r1 === 0 && n.r2 >= 1048575)) setSel(0, h.c, E.MAXR - 1, h.c); else if (h.zone === 'rowhdr' && !(h.r >= n.r1 && h.r <= n.r2 && n.c1 === 0 && n.c2 >= 16383)) setSel(h.r, 0, h.r, E.MAXC - 1); return; }
    if (ev.button !== 0) return;
    const h = hit(ev); closeMenus(); const sh = cur(); App.chartDeselect && App.chartDeselect();
    scroller.focus({ preventScroll: true });
    if (h.zone === 'corner') { setSel(0, 0, E.MAXR - 1, E.MAXC - 1); S.act = { r: 0, c: 0 }; updateUI(); return; }
    if (S.edit) {
      const canIns = h.zone === 'cell' && S.edit.text.startsWith('=') && canPoint(editor.value, editor.selectionStart);
      if (canIns) { ev.preventDefault(); const a = E.addrToStr(h.r, h.c); insertRef(a); S.edit.pointCell = { r: h.r, c: h.c }; drag = { type: 'point', r0: h.r, c0: h.c }; setTimeout(() => editor.focus(), 0); return; }
      commitEdit(0, 0, true);
    }
    if (h.zone === 'colhdr') {
      const sr = h.lx; const edge = Math.abs(sr - (colX(h.c) + colW(h.c)));
      if (edge < 4 / S.zoom && colW(h.c)) { drag = { type: 'colw', c: h.c, x0: ev.clientX, w0: colW(h.c) }; snap(); return; }
      if (edge < 4 / S.zoom && !colW(h.c)) { }
      if (ev.shiftKey) setSel(0, S.act.c, E.MAXR - 1, h.c, true); else { S.act = { r: 0, c: h.c }; setSel(0, h.c, E.MAXR - 1, h.c); S.act = { r: 0, c: h.c }; }
      drag = { type: 'cols', c0: h.c }; updateUI(); return;
    }
    if (h.zone === 'rowhdr') {
      const edge = Math.abs(h.ly - (rowY(h.r) + rowH(h.r)));
      if (edge < 4 / S.zoom && rowH(h.r)) { drag = { type: 'rowh', r: h.r, y0: ev.clientY, h0: rowH(h.r) }; snap(); return; }
      if (ev.shiftKey) setSel(S.act.r, 0, h.r, E.MAXC - 1, true); else { setSel(h.r, 0, h.r, E.MAXC - 1); S.act = { r: h.r, c: 0 }; }
      drag = { type: 'rows', r0: h.r }; updateUI(); return;
    }
    // fill handle
    const n = norm(S.sel), hr = screenRect(n.r1, n.c1, n.r2, n.c2); const hx = hr.x + hr.w, hy = hr.y + hr.h;
    if (Math.abs(h.px - hx) < 6 && Math.abs(h.py - hy) < 6) { drag = { type: 'fill', src: { ...n } }; return; }
    if (ev.ctrlKey && !ev.shiftKey) { const url = linkAt(h.r, h.c); if (url) { App.openUrl(url); return; } }
    if (ev.shiftKey) { selectCell(h.r, h.c, true); drag = { type: 'cells' }; }
    else { selectCell(h.r, h.c); drag = { type: 'cells' }; }
  });
  window.addEventListener('mousemove', ev => {
    if (!drag) { if (ev.target === cv) { hoverCursor(ev); noteTip(ev); } else if (tipEl) tipEl.style.display = 'none'; return; }
    const rc = scroller.getBoundingClientRect(); const dx = ev.clientX < rc.left + HW ? -30 : ev.clientX > rc.right - 16 ? 30 : 0, dy = ev.clientY < rc.top + HH ? -30 : ev.clientY > rc.bottom - 16 ? 30 : 0;
    if (drag.type === 'colw') { const w = Math.max(6, drag.w0 + (ev.clientX - drag.x0) / S.zoom); cur().colW[drag.c] = Math.round(w); App.geomChanged(); draw(); return; }
    if (drag.type === 'rowh') { const h2 = Math.max(4, drag.h0 + (ev.clientY - drag.y0) / S.zoom); cur().rowH[drag.r] = Math.round(h2); App.geomChanged(); draw(); return; }
    if (dx || dy) { scroller.scrollLeft += dx; scroller.scrollTop += dy; }
    const h = hit({ clientX: Math.min(Math.max(ev.clientX, rc.left + HW + 1), rc.right - 1), clientY: Math.min(Math.max(ev.clientY, rc.top + HH + 1), rc.bottom - 1) });
    if (drag.type === 'cells') setSel(S.act.r, S.act.c, h.r, h.c, true);
    else if (drag.type === 'cols') setSel(0, drag.c0, E.MAXR - 1, h.c, true);
    else if (drag.type === 'rows') setSel(drag.r0, 0, h.r, E.MAXC - 1, true);
    else if (drag.type === 'point') { const a = E.addrToStr(drag.r0, drag.c0), b = E.addrToStr(h.r, h.c); insertRef(a === b ? a : a + ':' + b); }
    else if (drag.type === 'fill') { const n = drag.src; let r2 = n.r2, c2 = n.c2, r1 = n.r1, c1 = n.c1; const dr = h.r > n.r2 ? h.r - n.r2 : h.r < n.r1 ? h.r - n.r1 : 0, dc = h.c > n.c2 ? h.c - n.c2 : h.c < n.c1 ? h.c - n.c1 : 0; if (Math.abs(dr) >= Math.abs(dc)) { if (dr > 0) r2 = h.r; else if (dr < 0) r1 = h.r; } else { if (dc > 0) c2 = h.c; else if (dc < 0) c1 = h.c; } drag.dest = { r1, c1, r2, c2 }; S.sel = { r1, c1, r2, c2 }; draw(); }
  });
  window.addEventListener('mouseup', () => {
    if (!drag) return; const d = drag; drag = null;
    if (d.type === 'fill' && d.dest) { snap(); fillRange(d.src, d.dest); afterChange(); setSel(d.dest.r1, d.dest.c1, d.dest.r2, d.dest.c2); }
    if (d.type === 'colw' || d.type === 'rowh') { afterChange(); }
    if (d.type === 'point') editor.focus();
  });
  function hoverCursor(ev) {
    const h = hit(ev); let cur_ = 'cell';
    if (h.zone === 'colhdr') cur_ = Math.abs(h.lx - (colX(h.c) + colW(h.c))) < 4 / S.zoom ? 'col-resize' : 'default';
    else if (h.zone === 'rowhdr') cur_ = Math.abs(h.ly - (rowY(h.r) + rowH(h.r))) < 4 / S.zoom ? 'row-resize' : 'default';
    else if (h.zone === 'corner') cur_ = 'default';
    else { const n = norm(S.sel), hr = screenRect(n.r1, n.c1, n.r2, n.c2); if (Math.abs(h.px - (hr.x + hr.w)) < 6 && Math.abs(h.py - (hr.y + hr.h)) < 6) cur_ = 'crosshair'; }
    cv.style.cursor = cur_;
  }
  scroller.addEventListener('dblclick', ev => {
    if (ev.target !== cv) return; const h = hit(ev);
    if (h.zone === 'colhdr' && Math.abs(h.lx - (colX(h.c) + colW(h.c))) < 4 / S.zoom) { App.CMD.autofitCol(h.c); return; }
    if (h.zone === 'rowhdr' && Math.abs(h.ly - (rowY(h.r) + rowH(h.r))) < 4 / S.zoom) { App.CMD.autofitRow(h.r); return; }
    if (h.zone === 'cell') startEdit(h.r, h.c, 'edit');
  });
  scroller.addEventListener('scroll', () => { draw(); if (S.edit) positionEditor(); App.closeMenus(); });
  scroller.addEventListener('wheel', ev => { if (ev.ctrlKey) { ev.preventDefault(); App.setZoom(Math.round((S.zoom * 100) + (ev.deltaY < 0 ? 10 : -10))); } }, { passive: false });
  scroller.tabIndex = 0; scroller.style.outline = 'none';
  window.addEventListener('resize', draw);
  new ResizeObserver(draw).observe(scroller);

  /* context menu */
  scroller.addEventListener('contextmenu', ev => {
    ev.preventDefault(); const h = hit(ev); const items = [];
    const mc = (label, cmd, key, dis) => ({ label, cmd, key, dis });
    if (h.zone === 'colhdr') items.push(mc('Cut', 'cut', 'Ctrl+X'), mc('Copy', 'copy', 'Ctrl+C'), mc('Paste', 'paste', 'Ctrl+V'), '-', mc('Insert column', 'insCol'), mc('Delete column', 'delCol'), mc('Clear contents', 'clear'), '-', mc('Column width…', 'colWidth'), mc('Hide', 'hideCols'), mc('Unhide', 'unhideCols'));
    else if (h.zone === 'rowhdr') items.push(mc('Cut', 'cut', 'Ctrl+X'), mc('Copy', 'copy', 'Ctrl+C'), mc('Paste', 'paste', 'Ctrl+V'), '-', mc('Insert row', 'insRow'), mc('Delete row', 'delRow'), mc('Clear contents', 'clear'), '-', mc('Row height…', 'rowHeight'), mc('Hide', 'hideRows'), mc('Unhide', 'unhideRows'));
    else items.push(mc('Cut', 'cut', 'Ctrl+X'), mc('Copy', 'copy', 'Ctrl+C'), mc('Paste', 'paste', 'Ctrl+V'), mc('Paste values only', 'pasteValues'), '-', mc('Insert row above', 'insRow'), mc('Insert column left', 'insCol'), mc('Delete row', 'delRow'), mc('Delete column', 'delCol'), mc('Clear contents', 'clear'), '-', mc('Sort A → Z', 'sortAsc'), mc('Sort Z → A', 'sortDesc'), mc('Filter', 'toggleFilter'), '-', mc('Insert note…', 'note'), mc('Link…', 'link', 'Ctrl+K'), mc('Data validation…', 'dataValidation'), '-', mc('Format cells…', 'formatCells', 'Ctrl+1'));
    App.showMenu(items, ev.clientX, ev.clientY);
  });

  /* ---------- keyboard ---------- */
  function ctrlJump(r, c, dr, dc) {
    const v = (rr, cc) => { const x = S.wb.getValue(S.si, rr, cc); return x !== null && x !== ''; }; const u = S.wb.usedRange(S.si); const maxR = Math.max(u.r2, 0), maxC = Math.max(u.c2, 0);
    let nr = r + dr, nc = c + dc; const inb = (a, b) => a >= 0 && b >= 0 && a < E.MAXR && b < E.MAXC;
    if (!inb(nr, nc)) return { r, c };
    if (v(r, c) && v(nr, nc)) { while (inb(nr + dr, nc + dc) && v(nr + dr, nc + dc) && (dr ? nr < maxR + 1 : nc < maxC + 1)) { nr += dr; nc += dc; } return { r: nr, c: nc }; }
    while (inb(nr, nc) && !v(nr, nc) && (dr > 0 ? nr <= maxR : dr < 0 ? nr > 0 : dc > 0 ? nc <= maxC : nc > 0)) { nr += dr; nc += dc; }
    if (!inb(nr, nc)) return { r, c };
    if (!v(nr, nc)) { if (dr > 0) nr = maxR; else if (dc > 0) nc = maxC; else if (dr < 0) nr = 0; else nc = 0; }
    return { r: Math.max(0, nr), c: Math.max(0, nc) };
  }
  scroller.addEventListener('keydown', ev => {
    if (S.edit) return; const ctrl = ev.ctrlKey || ev.metaKey, k = ev.key; const lk = k.length === 1 ? k.toLowerCase() : k;
    const shortcuts = { c: 'copy', x: 'cut', v: null, z: 'undo', y: 'redo', b: 'bold', i: 'italic', u: 'underline', a: 'selectAll', f: 'find', h: 'replace', s: 'save', o: 'open', n: 'new', p: 'print', d: 'fillDown', r: 'fillRight', k: 'link', '1': 'formatCells', '5': 'strike', '9': 'hideRows', '0': 'hideCols' };
    if (ctrl && !ev.altKey) {
      if (lk === 'v') return; if (lk === 'c' || lk === 'x') return; // handled by clipboard events
      if (lk === 'z' && ev.shiftKey) { ev.preventDefault(); App.CMD.redo(); return; }
      if (lk === 's' && ev.shiftKey) { ev.preventDefault(); App.CMD.saveAs(); return; }
      if (lk === ';') { ev.preventDefault(); const d = new Date(); const s = (d.getMonth() + 1) + '/' + d.getDate() + '/' + d.getFullYear(); App.act(() => S.wb.setInput(S.si, S.act.r, S.act.c, s)); return; }
      if (lk === ':') { ev.preventDefault(); const d = new Date(); App.act(() => S.wb.setInput(S.si, S.act.r, S.act.c, d.getHours() + ':' + String(d.getMinutes()).padStart(2, '0'))); return; }
      if (ev.shiftKey && lk === '!') { ev.preventDefault(); App.CMD.numFmt('#,##0.00'); return; } if (ev.shiftKey && lk === '$') { ev.preventDefault(); App.CMD.numFmt('$#,##0.00'); return; } if (ev.shiftKey && lk === '%') { ev.preventDefault(); App.CMD.numFmt('0%'); return; }
      if (ev.shiftKey && (lk === '+')) { ev.preventDefault(); App.CMD.insCell(); return; } if (lk === '-') { ev.preventDefault(); App.CMD.delCell(); return; }
      if (lk === ' ') { ev.preventDefault(); const n = norm(S.sel); setSel(0, n.c1, E.MAXR - 1, n.c2); return; }
      if (lk === 'Home') { ev.preventDefault(); selectCell(0, 0, ev.shiftKey); return; }
      if (lk === 'End') { ev.preventDefault(); const u = S.wb.usedRange(S.si); selectCell(Math.max(0, u.r2), Math.max(0, u.c2), ev.shiftKey); return; }
      if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(k)) {
        ev.preventDefault(); const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
        const from = ev.shiftKey ? (S.ext || S.act) : S.act; const t = ctrlJump(from.r, from.c, d[0], d[1]); selectCell(t.r, t.c, ev.shiftKey); return;
      }
      const cmd = shortcuts[lk]; if (cmd && App.CMD[cmd]) { ev.preventDefault(); App.CMD[cmd](); return; }
      return;
    }
    if (ev.altKey) { if (k === '=') { ev.preventDefault(); App.CMD.autosum(); } if (k === 'ArrowDown' && cur().filter) { ev.preventDefault(); } return; }
    const n = norm(S.sel);
    const move = (dr, dc) => { ev.preventDefault(); if (ev.shiftKey) { const e = { r: (S.ext || S.act).r + dr, c: (S.ext || S.act).c + dc }; setSel(S.act.r, S.act.c, Math.max(0, Math.min(E.MAXR - 1, e.r)), Math.max(0, Math.min(E.MAXC - 1, e.c)), true); scrollTo(Math.max(0, e.r), Math.max(0, e.c)); } else { let r = S.act.r, c = S.act.c; const m = mergeAt(r, c); if (m) { if (dr > 0) r = m.r2; if (dc > 0) c = m.c2; } selectCell(r + dr, c + dc); } };
    switch (k) {
      case 'ArrowUp': return move(-1, 0); case 'ArrowDown': return move(1, 0); case 'ArrowLeft': return move(0, -1); case 'ArrowRight': return move(0, 1);
      case 'PageDown': ev.preventDefault(); { const rows = Math.max(1, Math.floor(viewport().h / DH) - 2); if (ev.altKey) return; selectCell(S.act.r + rows, S.act.c, ev.shiftKey); } return;
      case 'PageUp': ev.preventDefault(); { const rows = Math.max(1, Math.floor(viewport().h / DH) - 2); selectCell(S.act.r - rows, S.act.c, ev.shiftKey); } return;
      case 'Home': ev.preventDefault(); selectCell(S.act.r, 0, ev.shiftKey); return;
      case 'End': return;
      case 'Tab': { ev.preventDefault(); const nn = norm(S.sel); const multi = nn.r1 !== nn.r2 || nn.c1 !== nn.c2; if (multi) { let r = S.act.r, c = S.act.c + (ev.shiftKey ? -1 : 1); if (c > nn.c2) { c = nn.c1; r++; } if (c < nn.c1) { c = nn.c2; r--; } if (r > nn.r2) r = nn.r1; if (r < nn.r1) r = nn.r2; S.act = { r, c }; updateUI(); scrollTo(r, c); } else selectCell(S.act.r, S.act.c + (ev.shiftKey ? -1 : 1)); return; }
      case 'Enter': { ev.preventDefault(); const nn = norm(S.sel); const multi = nn.r1 !== nn.r2 || nn.c1 !== nn.c2; if (multi) { let r = S.act.r + (ev.shiftKey ? -1 : 1), c = S.act.c; if (r > nn.r2) { r = nn.r1; c++; } if (r < nn.r1) { r = nn.r2; c--; } if (c > nn.c2) c = nn.c1; if (c < nn.c1) c = nn.c2; S.act = { r, c }; updateUI(); scrollTo(r, c); } else selectCell(S.act.r + (ev.shiftKey ? -1 : 1), S.act.c); return; }
      case 'F2': ev.preventDefault(); startEdit(S.act.r, S.act.c, 'edit'); return;
      case 'F4': ev.preventDefault(); App.CMD.repeat && App.CMD.repeat(); return;
      case 'Delete': case 'Backspace': ev.preventDefault(); App.CMD.clearContents(); return;
      case 'Escape': if (S.clip) { S.clip = null; draw(); } return;
      case ' ': if (ev.shiftKey) { ev.preventDefault(); setSel(n.r1, 0, n.r2, E.MAXC - 1); S.act = { r: S.act.r, c: 0 }; return; } break;
    }
    if (k.length === 1 && !ev.ctrlKey && !ev.metaKey) { ev.preventDefault(); startEdit(S.act.r, S.act.c, 'enter', k); editor.value = k; S.edit.text = k; $('#fbar').value = k; acUpdate(); editor.setSelectionRange(k.length, k.length); }
  });
  // IME / composition start in grid
  scroller.addEventListener('compositionstart', () => { if (!S.edit) startEdit(S.act.r, S.act.c, 'enter', ''); });

  /* ---------- clipboard ---------- */
  function rangeData(n, si) {
    const rows = []; const sh = S.wb.sheets[si];
    for (let r = n.r1; r <= n.r2; r++) { const row = []; for (let c = n.c1; c <= n.c2; c++) { const cell = sh.cells.get(r * 20000 + c); row.push(cell ? { t: cell.t, s: cell.s ? JSON.parse(JSON.stringify(cell.s)) : undefined } : null); } rows.push(row); }
    return rows;
  }
  function copyData(cut) {
    const n = expandMerges(norm(S.sel)); const rows = rangeData(n, S.si);
    const tsv = []; for (let r = n.r1; r <= n.r2; r++) { const line = []; for (let c = n.c1; c <= n.c2; c++) { const d = S.wb.display(S.si, r, c); line.push(/[\t\n"]/.test(d) ? '"' + d.replace(/"/g, '""') + '"' : d); } tsv.push(line.join('\t')); }
    S.clip = { rows, r1: n.r1, c1: n.c1, rowsN: n.r2 - n.r1 + 1, rows_: n.r2 - n.r1 + 1, cols: n.c2 - n.c1 + 1, si: S.si, cut: !!cut, text: tsv.join('\n'), merges: cur().merges.filter(m => m.r1 >= n.r1 && m.r2 <= n.r2 && m.c1 >= n.c1 && m.c2 <= n.c2).map(m => ({ ...m })) };
    S.clip.rows_ = S.clip.rows; S.clip.rows = S.clip.rowsN; S.clip.data = rows; draw(); return S.clip;
  }
  document.addEventListener('copy', ev => { if (S.edit || isTyping(ev.target)) return; ev.preventDefault(); const k = copyData(false); ev.clipboardData.setData('text/plain', k.text); });
  document.addEventListener('cut', ev => { if (S.edit || isTyping(ev.target)) return; ev.preventDefault(); const k = copyData(true); ev.clipboardData.setData('text/plain', k.text); });
  document.addEventListener('paste', ev => {
    if (S.edit || isTyping(ev.target)) return; ev.preventDefault(); const text = ev.clipboardData.getData('text/plain');
    if (S.clip && (text === S.clip.text || text === '')) pasteInternal('all'); else pasteText(text);
  });
  function isTyping(t) { return t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.isContentEditable) && t !== editor ? true : false; }
  function pasteInternal(mode) {
    const k = S.clip; if (!k) return; const n = norm(S.sel); let tr = n.r1, tc = n.c1;
    const rep = { r: 1, c: 1 }; const selR = n.r2 - n.r1 + 1, selC = n.c2 - n.c1 + 1; if (selR > k.rowsN && selR % k.rowsN === 0) rep.r = selR / k.rowsN; if (selC > k.cols && selC % k.cols === 0) rep.c = selC / k.cols;
    snap(); const dstSheet = S.si; let data = k.data; let rows = k.rowsN, cols = k.cols;
    if (mode === 'transpose') { const t = []; for (let j = 0; j < cols; j++) { t.push(data.map(row => row[j])); } data = t; [rows, cols] = [cols, rows]; }
    for (let ri = 0; ri < rep.r; ri++) for (let ci = 0; ci < rep.c; ci++) for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) {
      const r = tr + ri * rows + i, c = tc + ci * cols + j; if (r >= E.MAXR || c >= E.MAXC) continue; const src = data[i][j];
      if (mode === 'formats') { if (src && src.s) S.wb.setStyle(dstSheet, r, c, Object.fromEntries(Object.keys(cur().cells.get(r * 20000 + c)?.s || {}).map(x => [x, null]))), S.wb.setStyle(dstSheet, r, c, src.s); continue; }
      let t = src ? src.t : ''; const dr = r - (k.r1 + i), dc = c - (k.c1 + j);
      if (mode === 'transpose') { t = src ? src.t : ''; if (t[0] === '=') { t = E.shiftFormula(t, 0, 0); } }
      else if (typeof t === 'string' && t[0] === '=' && !k.cut) t = E.shiftFormula(t, dr, dc);
      if (mode === 'values') { const v = src ? S.wb.getValue(k.si, k.r1 + (mode === 'transpose' ? j : i), k.c1 + (mode === 'transpose' ? i : j)) : null; t = v === null ? '' : v instanceof E.FErr ? v.code : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'string' ? (E.parseLiteral(v).v !== v ? "'" + v : v) : String(v); S.wb.setInput(dstSheet, r, c, t); continue; }
      if (mode === 'formulas') { S.wb.setInput(dstSheet, r, c, t); continue; }
      S.wb.setInput(dstSheet, r, c, t);
      const cell = S.wb.sheets[dstSheet].cells.get(r * 20000 + c); if (src && src.s) { if (cell) cell.s = JSON.parse(JSON.stringify(src.s)); else { S.wb.setStyle(dstSheet, r, c, src.s); } } else if (cell && cell.s && mode === 'all') cell.s = undefined;
      if (!src && mode === 'all') S.wb.clearCell(dstSheet, r, c);
    }
    if (mode === 'all' && k.merges) for (const m of k.merges) cur().merges.push({ r1: m.r1 - k.r1 + tr, c1: m.c1 - k.c1 + tc, r2: m.r2 - k.r1 + tr, c2: m.c2 - k.c1 + tc });
    if (k.cut && mode === 'all') {
      const src = { r1: k.r1, c1: k.c1, r2: k.r1 + k.rowsN - 1, c2: k.c1 + k.cols - 1 }; const dr = tr - k.r1, dc = tc - k.c1;
      S.wb.rewriteAll(ref => { if (ref.sheet && S.wb.sheetIndex(ref.sheet) !== k.si) return ref; if (!ref.sheet && false) return ref; const inb = p => p && !p.col && !p.row && p.r >= src.r1 && p.r <= src.r2 && p.c >= src.c1 && p.c <= src.c2; if (inb(ref.a) && (!ref.b || inb(ref.b))) { const sh = p => p && ({ ...p, r: p.r + dr, c: p.c + dc }); return { sheet: ref.sheet, a: sh(ref.a), b: sh(ref.b) }; } return ref; });
      for (let r = src.r1; r <= src.r2; r++) for (let c = src.c1; c <= src.c2; c++) { const inDst = r - dr >= 0 && false; const rr = r + dr, cc = c + dc; if (!(rr >= src.r1 && rr <= src.r2 && cc >= src.c1 && cc <= src.c2 && false)) { const insideDst = r >= tr && r < tr + rows && c >= tc && c < tc + cols; if (!insideDst) S.wb.clearCell(k.si, r, c); } }
      cur().merges = cur().merges.filter(m => !(m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2 && !(m.r1 >= tr && m.r2 < tr + rows)));
      S.clip = null;
    }
    S.wb.invalidate(); S.wb.astCache.clear(); afterChange(); setSel(tr, tc, tr + rows * rep.r - 1, tc + cols * rep.c - 1); S.act = { r: tr, c: tc };
  }
  function pasteText(text) {
    if (!text) return; const rows = FmXlsx.parseCsv(text.replace(/\r\n/g, '\n').replace(/\n$/, '').replace(/^/, '')).length ? parseTsv(text) : []; if (!rows.length) return;
    snap(); const n = norm(S.sel); rows.forEach((row, i) => row.forEach((v, j) => S.wb.setInput(S.si, n.r1 + i, n.c1 + j, v))); afterChange(); setSel(n.r1, n.c1, n.r1 + rows.length - 1, n.c1 + Math.max(...rows.map(r => r.length)) - 1); S.act = { r: n.r1, c: n.c1 };
  }
  function parseTsv(t) { t = t.replace(/\r\n?/g, '\n').replace(/\n$/, ''); const rows = []; let row = [], f = '', q = false; for (let i = 0; i < t.length; i++) { const ch = t[i]; if (q) { if (ch === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; } else if (ch === '"' && f === '') q = true; else if (ch === '\t') { row.push(f); f = ''; } else if (ch === '\n') { row.push(f); rows.push(row); row = []; f = ''; } else f += ch; } row.push(f); rows.push(row); return rows; }
  App.pasteInternal = pasteInternal; App.copyData = copyData; App.pasteText = pasteText;
  async function clipPaste(mode) {
    if (mode && mode !== 'all') { if (S.clip) return pasteInternal(mode); }
    try { const t = await navigator.clipboard.readText(); if (S.clip && t === S.clip.text) return pasteInternal(mode || 'all'); if (t) return pasteText(t); } catch (e) { }
    if (S.clip) pasteInternal(mode || 'all'); else toast('Press Ctrl+V to paste');
  }
  App.clipPaste = clipPaste;
  async function clipCopy(cut) { const k = copyData(cut); try { await navigator.clipboard.writeText(k.text); } catch (e) { } }
  App.clipCopy = clipCopy;

  /* ---------- fill ---------- */
  const LISTS = [['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'], ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'], ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']];
  function seriesNext(vals, count, reverse) {
    const out = []; const nums = vals.map(v => v && v.t !== undefined ? v.t : v);
    const allNum = nums.every(t => t !== '' && typeof E.parseLiteral(t).v === 'number' && t[0] !== '=');
    if (allNum && nums.length >= 2) { const a = nums.map(t => E.parseLiteral(t).v); const step = (a[a.length - 1] - a[0]) / (a.length - 1); const fmt = E.parseLiteral(nums[0]).fmt; for (let i = 1; i <= count; i++) out.push(String(Math.round((reverse ? a[0] - step * i : a[a.length - 1] + step * i) * 1e10) / 1e10)); return out; }
    if (nums.length === 1) {
      const t = nums[0]; if (typeof t === 'string' && t[0] !== '=') {
        for (const L of LISTS) { const i = L.findIndex(x => x.toLowerCase() === t.toLowerCase()); if (i >= 0) { for (let k = 1; k <= count; k++) { const v = L[((reverse ? i - k : i + k) % L.length + L.length) % L.length]; out.push(t === t.toUpperCase() ? v.toUpperCase() : v); } return out; } }
        const m = /^(.*?)(\d+)$/.exec(t); if (m && typeof E.parseLiteral(t).v === 'string') { const w = m[2].length; for (let k = 1; k <= count; k++) out.push(m[1] + String(Math.max(0, +m[2] + (reverse ? -k : k))).padStart(w, '0')); return out; }
      }
    }
    return null;
  }
  function fillRange(src, dest) {
    const vert = dest.r1 !== src.r1 || dest.r2 !== src.r2; const sh = S.si;
    if (vert) {
      const rev = dest.r1 < src.r1; const cnt = rev ? src.r1 - dest.r1 : dest.r2 - src.r2; const h = src.r2 - src.r1 + 1;
      for (let c = src.c1; c <= src.c2; c++) {
        const cells = []; for (let r = src.r1; r <= src.r2; r++) cells.push(S.wb.getInput(sh, r, c)); const ser = seriesNext(cells, cnt, rev);
        for (let i = 1; i <= cnt; i++) {
          const tr = rev ? src.r1 - i : src.r2 + i; const sr = rev ? src.r1 + (((h - (i % h)) % h)) : src.r1 + ((i - 1) % h); const srcR = rev ? src.r2 - ((i - 1) % h) : src.r1 + ((i - 1) % h);
          const sc = S.wb.cell(sh, srcR, c); let t = ser ? ser[i - 1] : S.wb.getInput(sh, srcR, c); if (!ser && typeof t === 'string' && t[0] === '=') t = E.shiftFormula(t, tr - srcR, 0);
          S.wb.setInput(sh, tr, c, t); if (sc && sc.s) S.wb.setStyle(sh, tr, c, sc.s);
        }
      }
    } else {
      const rev = dest.c1 < src.c1; const cnt = rev ? src.c1 - dest.c1 : dest.c2 - src.c2; const w = src.c2 - src.c1 + 1;
      for (let r = src.r1; r <= src.r2; r++) {
        const cells = []; for (let c = src.c1; c <= src.c2; c++) cells.push(S.wb.getInput(sh, r, c)); const ser = seriesNext(cells, cnt, rev);
        for (let i = 1; i <= cnt; i++) {
          const tc = rev ? src.c1 - i : src.c2 + i; const srcC = rev ? src.c2 - ((i - 1) % w) : src.c1 + ((i - 1) % w); const sc = S.wb.cell(sh, r, srcC);
          let t = ser ? ser[i - 1] : S.wb.getInput(sh, r, srcC); if (!ser && typeof t === 'string' && t[0] === '=') t = E.shiftFormula(t, 0, tc - srcC);
          S.wb.setInput(sh, r, tc, t); if (sc && sc.s) S.wb.setStyle(sh, r, tc, sc.s);
        }
      }
    }
  }
  App.fillRange = fillRange;

  /* ---------- zoom ---------- */
  App.setZoom = pct => { pct = Math.max(50, Math.min(200, pct)); S.zoom = pct / 100; $('#zoom').value = pct; $('#zoomv').textContent = pct + '%'; draw(); if (S.edit) positionEditor(); App.renderCharts && App.renderCharts(); };
  $('#zoom').oninput = e => App.setZoom(+e.target.value); $('#zoomin').onclick = () => App.setZoom(S.zoom * 100 + 10); $('#zoomout').onclick = () => App.setZoom(S.zoom * 100 - 10);

  /* ---------- filters (auto-filter dropdown buttons) ---------- */
  function positionFilters() {
    const fl = $('#filters'); const sh = cur(); fl.innerHTML = '';
    const dv = !S.edit && dvAt(sh, S.act.r, S.act.c); if (dv && dv.type === 'list' && dv.showDrop !== false) {
      const m = mergeAt(S.act.r, S.act.c) || { r1: S.act.r, c1: S.act.c, r2: S.act.r, c2: S.act.c }; const sr = screenRect(m.r1, m.c1, m.r2, m.c2); const fx = m.c1 < sh.freeze.c ? 0 : scroller.scrollLeft, fy = m.r1 < sh.freeze.r ? 0 : scroller.scrollTop; const b = document.createElement('div'); b.className = 'fbtn'; b.textContent = '▾'; b.style.left = ((sr.x + sr.w) * S.zoom + fx + 2) + 'px'; b.style.top = (sr.y * S.zoom + fy + 2) + 'px'; b.style.width = '20px'; b.style.height = (Math.max(20, sr.h * S.zoom - 4)) + 'px';
      b.onmousedown = ev => { ev.stopPropagation(); ev.preventDefault(); const items = dvList(dv).map(x => ({ label: x, fn: () => App.act(() => S.wb.setInput(S.si, S.act.r, S.act.c, x)) })); if (items.length) App.showMenu(items, ev.clientX, ev.clientY + 12); else toast('This list is empty'); }; fl.appendChild(b);
    }
    if (!sh.filter) return; const f = sh.filter; const z = S.zoom;
    for (let c = f.c1; c <= f.c2; c++) {
      const sr = screenRect(f.r, c, f.r, c); const b = document.createElement('div'); b.className = 'fbtn' + (f.ex && f.ex[c] && f.ex[c].length ? ' on' : ''); b.textContent = '▾';
      const fx = c < sh.freeze.c ? 0 : scroller.scrollLeft, fy = f.r < sh.freeze.r ? 0 : scroller.scrollTop;
      b.style.left = ((sr.x + sr.w) * z + fx - 20 * 1) + 'px'; b.style.top = ((sr.y + sr.h) * z + fy - 20) + 'px'; b.onmousedown = ev => { ev.stopPropagation(); ev.preventDefault(); App.filterMenu(c, ev.clientX, ev.clientY); }; fl.appendChild(b);
    }
  }

  /* ---------- autosave ---------- */
  let asT; function scheduleAutosave() { clearTimeout(asT); asT = setTimeout(() => { try { localStorage.setItem('fmsheets.autosave', JSON.stringify({ name: S.name, wb: S.wb.toJSON(), t: Date.now() })); $('#savestate').textContent = S.handle || !S.dirty ? 'Saved' : 'Auto-saved'; } catch (e) { } }, 800); }
  App.scheduleAutosave = scheduleAutosave;
  App.helpers = { norm, mergeAt, expandMerges, setSel, cur, ctx };
  App.closeMenus = () => { };
  function closeMenus() { App.closeMenus(); } App.hooks.closeMenus = closeMenus;
  App.focusGrid = () => scroller.focus({ preventScroll: true });
  setTimeout(() => { App.updateUI(); draw(); scroller.focus(); }, 0);
})();
