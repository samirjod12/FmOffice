/* FmOffice Sheets — ribbon, commands, dialogs, sheet tabs, file I/O */
(function () {
  'use strict';
  const E = FmEngine, App = window.App, S = App.S, CMD = App.CMD, $ = s => document.querySelector(s);
  const cur = App.cur, H = App.helpers, norm = H.norm;
  const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html !== undefined) e.innerHTML = html; return e; };
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  /* ---------- icons ---------- */
  const P = {
    undo: 'M9 14 4 9l5-5M4 9h10.5a5.5 5.5 0 0 1 0 11H11', redo: 'm15 14 5-5-5-5M20 9H9.5a5.5 5.5 0 0 0 0 11H13',
    save: 'M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2zM17 21v-8H7v8M7 3v5h8',
    paste: 'M8 4h8v3H8zM16 5h2a2 2 0 0 1 2 2v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7a2 2 0 0 1 2-2h2',
    copy: 'M9 9h11v11H9zM5 15H4V4h11v1', cut: 'M20 4 8.1 15.9M14.5 14.5 20 20M8.1 8.1 12 12M6 9a3 3 0 1 0 0-6 3 3 0 0 0 0 6zM6 21a3 3 0 1 0 0-6 3 3 0 0 0 0 6z',
    bold: 'M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z', italic: 'M19 4h-9M14 20H5M15 4 9 20', underline: 'M6 4v6a6 6 0 0 0 12 0V4M4 21h16', strike: 'M16 4H9a3 3 0 0 0-2.8 4M14 12a4 4 0 0 1 0 8H6M4 12h16',
    left: 'M21 6H3M15 12H3M17 18H3', center: 'M21 6H3M17 12H7M19 18H5', right: 'M21 6H3M21 12H9M21 18H7',
    top: 'M4 4h16M12 20V8M8 12l4-4 4 4', mid: 'M4 12h16M12 3v6M9 6l3 3 3-3M12 21v-6M9 18l3-3 3 3', bottom: 'M4 20h16M12 4v12M8 12l4 4 4-4',
    wrap: 'M3 6h18M3 12h15a3 3 0 1 1 0 6h-4M16 16l-2 2 2 2M3 18h7', merge: 'M8 3H5a2 2 0 0 0-2 2v3M21 8V5a2 2 0 0 0-2-2h-3M3 16v3a2 2 0 0 0 2 2h3M16 21h3a2 2 0 0 0 2-2v-3M8 12h8M10 9l-3 3 3 3M14 9l3 3-3 3',
    fill: 'm19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2c.8.8 2 .8 2.8 0zM5 2l5 5M2 13h15', border: 'M3 3h18v18H3zM3 12h18M12 3v18', indent: 'M3 6h18M11 12h10M3 18h18M3 9l3 3-3 3',
    plus: 'M12 5v14M5 12h14', trash: 'M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6', sigma: 'M18 7V4H6l6 8-6 8h12v-3',
    sortasc: 'm3 8 4-4 4 4M7 4v16M13 6h8M13 12h6M13 18h4', sortdesc: 'm3 16 4 4 4-4M7 20V4M13 6h4M13 12h6M13 18h8', filter: 'M22 3H2l8 9.5V19l4 2v-8.5z',
    search: 'M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zM21 21l-4.3-4.3', chart: 'M3 3v18h18M8 17v-5M13 17V8M18 17v-3', line: 'M3 3v18h18M7 15l4-4 3 3 5-6', pie: 'M21 12A9 9 0 1 1 12 3v9zM15 3.5A9 9 0 0 1 20.5 9H15z',
    area: 'M3 3v18h18M7 16l4-5 4 3 4-6v10H7z', scatter: 'M3 3v18h18M8 15h.01M12 9h.01M15 14h.01M18 7h.01M10 18h.01',
    file: 'M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M12 12v6M9 15h6', open: 'M3 7V5a2 2 0 0 1 2-2h4l2 3h8a2 2 0 0 1 2 2v1M3 7h18l-2 12H5z', print: 'M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v7H6z',
    grid: 'M3 3h18v18H3zM3 9h18M3 15h18M9 3v18M15 3v18', freeze: 'M3 3h18v18H3zM3 9h18M9 3v18', moon: 'M21 13A9 9 0 1 1 11 3a7 7 0 0 0 10 10z', sun: 'M12 17a5 5 0 1 0 0-10 5 5 0 0 0 0 10zM12 1v2M12 21v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M1 12h2M21 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4',
    insrow: 'M3 5h18v6H3zM3 15h8M17 14v6M14 17h6', inscol: 'M5 3v18h6V3zM15 3h6M18 3v0M14 13h6M17 10v6', delrow: 'M3 5h18v6H3zM14 17h6', delcol: 'M5 3v18h6V3zM14 12h6',
    clear: 'M20 20H9L3 14l11-11 7 7-8 8M9 20', cf: 'M3 3h18v18H3zM3 9h18M9 3v18M13 13h4v4h-4z', table: 'M3 3h18v18H3zM3 9h18M3 15h18M9 9v12', styles: 'M12 2l3 7h7l-5.5 4.5L18.5 21 12 16.5 5.5 21l2-7.5L2 9h7z',
    fx: 'M4 20c3 0 3-4 4-8s1-8 4-8M5 12h8M14 14l5 6M19 14l-5 6', calc: 'M5 2h14v20H5zM8 6h8M8 11h2M14 11h2M8 15h2M14 15h2M8 19h2M14 19h2', link: 'M10 13a5 5 0 0 0 7 0l3-3a5 5 0 0 0-7-7l-1 1M14 11a5 5 0 0 0-7 0l-3 3a5 5 0 0 0 7 7l1-1',
    check: 'M20 6 9 17l-5-5', zoom: 'M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zM21 21l-4.3-4.3M8 11h6M11 8v6', dec: 'M4 12h6M14 8v8M18 8v8M4 16h.01', inc: 'M4 12h6M14 8v8M18 8v8M7 9v6', sheet: 'M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6M8 13h8M8 17h8',
    car: 'm6 9 6 6 6-6', brush: 'M18 3l3 3-9 9-3-3zM9 12c-3 0-5 2-5 5 0 2-1 3-2 4 4 1 9 0 9-5', outdent: 'M3 6h18M11 12h10M3 18h18M7 9l-3 3 3 3', text: 'M4 7V4h16v3M9 20h6M12 4v16', dup: 'M8 8h12v12H8zM4 16V4h12', hide: 'M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.9 5.1A10 10 0 0 1 12 5c6 0 10 7 10 7a17 17 0 0 1-3 3.8M6.6 6.6A17 17 0 0 0 2 12s4 7 10 7a9.7 9.7 0 0 0 4-.9', col: 'M5 3v18h6V3zM15 8v13h6V8z', image: 'M3 3h18v18H3zM8.5 10a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3zM21 15l-5-5L5 21'
  };
  const ic = (n, cls) => `<svg class="i ${cls || ''}" viewBox="0 0 24 24"><path d="${P[n] || P.sheet}"/></svg>`;
  App.ic = ic;
  $('.qbtn[data-cmd=save]').innerHTML = ic('save'); $('.qbtn[data-cmd=undo]').innerHTML = ic('undo'); $('.qbtn[data-cmd=redo]').innerHTML = ic('redo');
  $('#tellme-ic').innerHTML = ic('search'); $('#addsheet').innerHTML = ic('plus'); $('#fx-btn').textContent = 'fx';
  document.querySelectorAll('.qbtn[data-cmd]').forEach(b => b.onclick = () => CMD[b.dataset.cmd]());

  /* ---------- menus ---------- */
  const menu = $('#menu');
  function closeMenus() { menu.classList.add('hidden'); $('#tellme-list').classList.add('hidden'); }
  App.closeMenus = closeMenus;
  function showMenu(items, x, y, opts) {
    menu.innerHTML = ''; opts = opts || {};
    for (const it of items) {
      if (it === '-') { menu.appendChild(el('div', 'msep')); continue; }
      if (it.header) { menu.appendChild(el('div', 'mh', esc(it.header))); continue; }
      if (it.palette) { const p = el('div', 'pal'); it.palette.forEach(c => { const i = el('i'); i.style.background = c || 'transparent'; i.title = c || 'No color'; if (!c) i.style.backgroundImage = 'linear-gradient(135deg,transparent 45%,#e33 45%,#e33 55%,transparent 55%)'; i.onclick = () => { closeMenus(); it.pick(c); }; p.appendChild(i); }); menu.appendChild(p); continue; }
      const b = el('button', 'mi' + (it.dis ? ' dis' : ''), (it.icon ? ic(it.icon) : '') + `<span>${esc(it.label)}</span>` + (it.key ? `<span class="k">${esc(it.key)}</span>` : '') + (it.check ? '<span class="k">✓</span>' : ''));
      b.onclick = () => { closeMenus(); App.focusGrid(); if (it.fn) it.fn(); else if (it.cmd) CMD[it.cmd](it.arg); }; menu.appendChild(b);
    }
    menu.classList.remove('hidden'); menu.style.left = '0px'; menu.style.top = '0px';
    const r = menu.getBoundingClientRect(); menu.style.left = Math.max(6, Math.min(x, innerWidth - r.width - 8)) + 'px'; menu.style.top = Math.max(6, Math.min(y, innerHeight - r.height - 8)) + 'px';
  }
  App.showMenu = showMenu;
  document.addEventListener('mousedown', e => { if (!e.target.closest('#menu') && !e.target.closest('.dropbtn') && !e.target.closest('.tellme')) closeMenus(); }, true);
  function dropdown(btn, items) { const r = btn.getBoundingClientRect(); showMenu(typeof items === 'function' ? items() : items, r.left, r.bottom + 4); }

  /* ---------- dialogs ---------- */
  const modal = $('#modal');
  function closeModal() { modal.classList.add('hidden'); modal.innerHTML = ''; App.focusGrid(); }
  function dialog(title, body, buttons, opts) {
    modal.innerHTML = ''; const d = el('div', 'dlg'); d.innerHTML = `<h3>${esc(title)}</h3>`; const b = el('div', 'dbody'); if (typeof body === 'string') b.innerHTML = body; else b.appendChild(body); d.appendChild(b);
    const f = el('div', 'dfoot'); (buttons || [{ label: 'OK', primary: 1 }]).forEach(bt => { const x = el('button', bt.primary ? 'pbtn' : 'gbtn', esc(bt.label)); x.onclick = () => { if (!bt.fn || bt.fn(d) !== false) closeModal(); }; f.appendChild(x); }); d.appendChild(f);
    if (opts && opts.width) d.style.width = opts.width + 'px'; modal.appendChild(d); modal.classList.remove('hidden');
    modal.onmousedown = e => { if (e.target === modal) closeModal(); }; d.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON' && !(opts && opts.noEnter)) { const p = f.querySelector('.pbtn'); if (p) p.click(); } });
    const first = d.querySelector('input,select,textarea'); if (first) { first.focus(); if (first.select) first.select(); } return d;
  }
  App.dialog = dialog; App.closeModal = closeModal;
  App.prompt = (title, val, cb, label) => dialog(title, `<label>${esc(label || '')}<input type="text" id="pv" value="${esc(val)}"></label>`, [{ label: 'Cancel' }, { label: 'OK', primary: 1, fn: d => { cb(d.querySelector('#pv').value); } }]);
  App.alert = (title, msg) => dialog(title, `<div>${esc(msg)}</div>`);

  /* ---------- helpers for commands ---------- */
  const sel = () => App.selNorm();
  const MAXR = E.MAXR, MAXC = E.MAXC;
  function eachCell(fn, limit) {
    const n = sel(); let r2 = n.r2, c2 = n.c2; const sh = cur();
    if (r2 >= MAXR - 1 || c2 >= MAXC - 1) { const u = S.wb.usedRange(S.si); const ex = new Set(); for (const k of sh.cells.keys()) { const r = Math.floor(k / 20000), c = k % 20000; if (r >= n.r1 && r <= n.r2 && c >= n.c1 && c <= n.c2) fn(r, c); } return; }
    if ((r2 - n.r1 + 1) * (c2 - n.c1 + 1) > 500000) { App.toast('Selection too large'); return; }
    for (let r = n.r1; r <= r2; r++) for (let c = n.c1; c <= c2; c++) fn(r, c);
  }
  function applyStyle(patch, wholeAxis) {
    const n = sel(); const full = n.r2 >= MAXR - 1 || n.c2 >= MAXC - 1; App.snap();
    if (full) { // style existing cells and set col/row defaults via a bounded block
      const u = S.wb.usedRange(S.si); const r2 = n.r2 >= MAXR - 1 ? Math.max(u.r2, 200) : n.r2, c2 = n.c2 >= MAXC - 1 ? Math.max(u.c2, 25) : n.c2;
      for (let r = n.r1; r <= Math.min(r2, n.r1 + 3000); r++) for (let c = n.c1; c <= Math.min(c2, n.c1 + 200); c++) S.wb.setStyle(S.si, r, c, patch);
    } else eachCell((r, c) => S.wb.setStyle(S.si, r, c, patch));
    App.afterChange(); return;
  }
  const getSt = () => S.wb.getStyle(S.si, S.act.r, S.act.c) || {};
  CMD.style = patch => applyStyle(patch);
  const toggle = key => () => { const on = !getSt()[key]; applyStyle({ [key]: on ? 1 : null }); };
  CMD.bold = toggle('b'); CMD.italic = toggle('i'); CMD.underline = toggle('u'); CMD.strike = toggle('st');
  CMD.align = v => applyStyle({ ha: getSt().ha === v ? null : v }); CMD.valign = v => applyStyle({ va: v });
  CMD.wrap = () => { const on = !getSt().wrap; applyStyle({ wrap: on ? 1 : null }); if (on) { const n = sel(); for (let r = n.r1; r <= Math.min(n.r2, n.r1 + 500); r++) CMD.autofitRow(r, true); App.geomChanged(); App.afterChange(); } };
  CMD.indent = d => applyStyle({ ind: Math.max(0, (getSt().ind || 0) + d) || null });
  CMD.font = v => applyStyle({ ff: v === 'Calibri' ? null : v }); CMD.fsize = v => applyStyle({ fs: +v === 11 ? null : +v });
  CMD.fontColor = c => applyStyle({ fc: c || null }); CMD.fillColor = c => applyStyle({ bg: c || null });
  CMD.numFmt = f => applyStyle({ nf: f === 'General' || !f ? null : f });
  CMD.decimals = d => { const nf = getSt().nf || (typeof S.wb.getValue(S.si, S.act.r, S.act.c) === 'number' ? 'General' : null); if (!nf) return; let cur_ = nf === 'General' ? '0' : nf; if (nf === 'General') { const v = S.wb.getValue(S.si, S.act.r, S.act.c); const m = String(v).split('.')[1]; cur_ = m ? '0.' + '0'.repeat(Math.min(m.length, 8)) : '0'; } const m = /^([^.]*?)(?:\.(0+))?([^0.]*)$/.exec(cur_.split(';')[0]); if (!m) return; const dec = (m[2] || '').length + d; if (dec < 0) return; const nfn = m[1] + (dec ? '.' + '0'.repeat(dec) : '') + m[3]; applyStyle({ nf: nfn }); };
  CMD.currency = () => CMD.numFmt('$#,##0.00'); CMD.percent = () => CMD.numFmt('0%'); CMD.comma = () => CMD.numFmt('#,##0.00');
  CMD.borders = kind => {
    const n = sel(); App.snap(); const line = { c: '#1d2430', w: kind === 'thick' ? 2 : 1 };
    const upd = (r, c, f) => { const cell = S.wb.getStyle(S.si, r, c); const bd = { ...(cell && cell.bd || {}) }; f(bd); S.wb.setStyle(S.si, r, c, { bd: Object.keys(bd).length ? bd : null }); };
    for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) upd(r, c, bd => {
      if (kind === 'none') { delete bd.t; delete bd.b; delete bd.l; delete bd.r; return; }
      if (kind === 'all') { bd.t = bd.b = bd.l = bd.r = line; return; }
      if ((kind === 'outside' || kind === 'thick')) { if (r === n.r1) bd.t = line; if (r === n.r2) bd.b = line; if (c === n.c1) bd.l = line; if (c === n.c2) bd.r = line; return; }
      if (kind === 'top' && r === n.r1) bd.t = line; if (kind === 'bottom' && r === n.r2) bd.b = line; if (kind === 'left' && c === n.c1) bd.l = line; if (kind === 'right' && c === n.c2) bd.r = line;
      if (kind === 'topbottom') { if (r === n.r1) bd.t = line; if (r === n.r2) bd.b = { c: '#1d2430', w: 2 }; }
    });
    App.afterChange();
  };
  CMD.merge = mode => {
    const n = sel(); const sh = cur(); App.snap(); sh.merges = sh.merges.filter(m => m.r1 > n.r2 || m.r2 < n.r1 || m.c1 > n.c2 || m.c2 < n.c1);
    if (mode === 'unmerge') { App.afterChange(); return; }
    if (n.r1 === n.r2 && n.c1 === n.c2) { App.afterChange(); return; }
    if (mode === 'across') { for (let r = n.r1; r <= n.r2; r++) sh.merges.push({ r1: r, c1: n.c1, r2: r, c2: n.c2 }); } else sh.merges.push({ r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 });
    for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) if (!(r === n.r1 && c === n.c1) && !(mode === 'across' && c === n.c1)) S.wb.clearCell(S.si, r, c, true);
    S.wb.setStyle(S.si, n.r1, n.c1, { ha: getSt().ha || 'center' }); App.afterChange();
  };
  CMD.mergeCenter = () => { const n = sel(); const isM = H.mergeAt(n.r1, n.c1); CMD.merge(isM ? 'unmerge' : 'all'); };
  CMD.clearContents = () => { App.snap(); eachCell((r, c) => S.wb.clearCell(S.si, r, c, true)); App.afterChange(); };
  CMD.clear = CMD.clearContents;
  CMD.clearFormats = () => { App.snap(); eachCell((r, c) => { const x = S.wb.cell(S.si, r, c); if (x) { x.s = undefined; if (x.t === '') cur().cells.delete(r * 20000 + c); } }); App.afterChange(); };
  CMD.clearAll = () => { App.snap(); eachCell((r, c) => S.wb.clearCell(S.si, r, c, false)); App.afterChange(); };
  CMD.selectAll = () => { App.setSel(0, 0, MAXR - 1, MAXC - 1); S.act = { r: 0, c: 0 }; };
  CMD.copy = () => App.clipCopy(false); CMD.cut = () => App.clipCopy(true); CMD.paste = () => App.clipPaste('all');
  CMD.pasteValues = () => App.clipPaste('values'); CMD.pasteFormulas = () => App.clipPaste('formulas'); CMD.pasteFormats = () => App.clipPaste('formats'); CMD.pasteTranspose = () => App.clipPaste('transpose');
  CMD.undo = () => { App.undo(); }; CMD.redo = () => { App.redo(); };
  CMD.fillDown = () => { const n = sel(); if (n.r1 === n.r2) return; App.snap(); for (let c = n.c1; c <= n.c2; c++) for (let r = n.r1 + 1; r <= n.r2; r++) { const s = S.wb.cell(S.si, n.r1, c); let t = S.wb.getInput(S.si, n.r1, c); if (t[0] === '=') t = E.shiftFormula(t, r - n.r1, 0); S.wb.setInput(S.si, r, c, t); if (s && s.s) S.wb.setStyle(S.si, r, c, s.s); } App.afterChange(); };
  CMD.fillRight = () => { const n = sel(); if (n.c1 === n.c2) return; App.snap(); for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1 + 1; c <= n.c2; c++) { const s = S.wb.cell(S.si, r, n.c1); let t = S.wb.getInput(S.si, r, n.c1); if (t[0] === '=') t = E.shiftFormula(t, 0, c - n.c1); S.wb.setInput(S.si, r, c, t); if (s && s.s) S.wb.setStyle(S.si, r, c, s.s); } App.afterChange(); };
  CMD.fillSeries = () => { const n = sel(); App.snap(); if (n.r2 > n.r1) App.fillRange({ r1: n.r1, c1: n.c1, r2: n.r1, c2: n.c2 }, n); else if (n.c2 > n.c1) App.fillRange({ r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c1 }, n); App.afterChange(); };

  /* rows/cols */
  const cnt = (a, b) => b - a + 1;
  CMD.insRow = () => { const n = sel(); App.snap(); S.wb.insertRows(S.si, n.r1, cnt(n.r1, n.r2)); App.afterChange(true); };
  CMD.delRow = () => { const n = sel(); App.snap(); S.wb.deleteRows(S.si, n.r1, cnt(n.r1, n.r2)); App.afterChange(true); };
  CMD.insCol = () => { const n = sel(); App.snap(); S.wb.insertCols(S.si, n.c1, cnt(n.c1, n.c2)); App.afterChange(true); };
  CMD.delCol = () => { const n = sel(); App.snap(); S.wb.deleteCols(S.si, n.c1, cnt(n.c1, n.c2)); App.afterChange(true); };
  CMD.insCell = () => { const n = sel(); const full = n.c1 === 0 && n.c2 >= MAXC - 1; if (full) return CMD.insRow(); if (n.r1 === 0 && n.r2 >= MAXR - 1) return CMD.insCol(); CMD.insRow(); };
  CMD.delCell = () => { const n = sel(); if (n.r1 === 0 && n.r2 >= MAXR - 1) return CMD.delCol(); CMD.delRow(); };
  CMD.hideRows = () => { const n = sel(); App.snap(); for (let r = n.r1; r <= Math.min(n.r2, n.r1 + 10000); r++) cur().hiddenRows[r] = 1; App.afterChange(); };
  CMD.hideCols = () => { const n = sel(); App.snap(); for (let c = n.c1; c <= Math.min(n.c2, n.c1 + 1000); c++) cur().hiddenCols[c] = 1; App.afterChange(); };
  CMD.unhideRows = () => { const n = sel(); App.snap(); const lo = Math.max(0, n.r1 - 1), hi = n.r2 + 1; for (let r = lo; r <= hi; r++) delete cur().hiddenRows[r]; App.afterChange(); };
  CMD.unhideCols = () => { const n = sel(); App.snap(); const lo = Math.max(0, n.c1 - 1), hi = n.c2 + 1; for (let c = lo; c <= hi; c++) delete cur().hiddenCols[c]; App.afterChange(); };
  CMD.colWidth = () => App.prompt('Column width', String(Math.round(App.colW(sel().c1) / 7 * 10) / 10), v => { const w = parseFloat(v); if (!(w >= 0)) return; const n = sel(); App.snap(); for (let c = n.c1; c <= Math.min(n.c2, n.c1 + 1000); c++) cur().colW[c] = Math.round(w * 7); App.afterChange(); }, 'Width (characters)');
  CMD.rowHeight = () => App.prompt('Row height', String(Math.round(App.rowH(sel().r1) * 0.75 * 10) / 10), v => { const h = parseFloat(v); if (!(h >= 0)) return; const n = sel(); App.snap(); for (let r = n.r1; r <= Math.min(n.r2, n.r1 + 5000); r++) cur().rowH[r] = Math.round(h / 0.75); App.afterChange(); }, 'Height (points)');
  CMD.autofitCol = c0 => {
    const n = sel(); const cols = c0 !== undefined ? [c0] : Array.from({ length: Math.min(cnt(n.c1, n.c2), 200) }, (_, i) => n.c1 + i); App.snap(); const ctx = H.ctx; const sh = cur();
    for (const c of cols) { let w = 40; for (const [k, cell] of sh.cells) { if (k % 20000 !== c || cell.t === '') continue; const r = Math.floor(k / 20000); const d = S.wb.display(S.si, r, c); const st = cell.s; ctx.font = (st && st.b ? 'bold ' : '') + (st && st.i ? 'italic ' : '') + (Math.round((st && st.fs || 11) * 4 / 3)) + 'px "Segoe UI",sans-serif'; const lines = st && st.wrap ? [d] : d.split('\n'); for (const l of lines) w = Math.max(w, ctx.measureText(l).width + 16 + (st && st.ind ? st.ind * 10 : 0)); } sh.colW[c] = Math.ceil(Math.min(w, 600)); }
    App.afterChange();
  };
  CMD.autofitRow = (r0, silent) => {
    const sh = cur(); const rows = r0 !== undefined ? [r0] : (() => { const n = sel(); return Array.from({ length: Math.min(cnt(n.r1, n.r2), 500) }, (_, i) => n.r1 + i); })(); if (!silent) App.snap();
    for (const r of rows) { let h = 24; for (const [k, cell] of sh.cells) { if (Math.floor(k / 20000) !== r || cell.t === '') continue; const c = k % 20000; const st = cell.s; const fs = Math.round((st && st.fs || 11) * 4 / 3); if (st && st.wrap) { const font = (st.b ? 'bold ' : '') + fs + 'px "Segoe UI",sans-serif'; const lines = App.wrapLines(S.wb.display(S.si, r, c), App.colW(c) - 10, font); h = Math.max(h, lines.length * fs * 1.25 + 8); } else h = Math.max(h, fs * 1.25 + 8); } if (h > 24 || sh.rowH[r]) sh.rowH[r] = Math.ceil(h); if (h <= 24) delete sh.rowH[r]; }
    if (!silent) App.afterChange();
  };
  CMD.freeze = mode => { const sh = cur(); const n = sel(); App.snap(); if (mode === 'none') sh.freeze = { r: 0, c: 0 }; else if (mode === 'row') sh.freeze = { r: 1, c: sh.freeze.c }; else if (mode === 'col') sh.freeze = { r: sh.freeze.r, c: 1 }; else sh.freeze = { r: S.act.r, c: S.act.c }; App.geomChanged(); App.afterChange(); };
  CMD.gridlines = () => { cur().showGrid = !cur().showGrid; App.draw(); refreshRibbonState(); };
  CMD.headings = () => { S.showHeadings = !S.showHeadings; App.draw(); refreshRibbonState(); };
  CMD.fbarToggle = () => { S.showFbar = !S.showFbar; $('#formulabar').style.display = S.showFbar ? '' : 'none'; setTimeout(App.draw, 0); refreshRibbonState(); };

  /* autosum */
  CMD.autosum = fn => {
    fn = typeof fn === 'string' ? fn : 'SUM'; const n = sel(); App.snap(); const has = (r, c) => typeof S.wb.getValue(S.si, r, c) === 'number';
    if (n.r1 === n.r2 && n.c1 === n.c2) {
      let r = n.r1 - 1, c = n.c1; let top = r; while (top >= 0 && has(top, c)) top--; top++;
      if (top <= r) { S.wb.setInput(S.si, n.r1, n.c1, `=${fn}(${E.addrToStr(top, c)}:${E.addrToStr(r, c)})`); }
      else { let cc = n.c1 - 1, l = cc; while (l >= 0 && has(n.r1, l)) l--; l++; if (l <= cc) S.wb.setInput(S.si, n.r1, n.c1, `=${fn}(${E.addrToStr(n.r1, l)}:${E.addrToStr(n.r1, cc)})`); else { App.afterChange(); App.startEdit(n.r1, n.c1, 'enter', `=${fn}()`); const ed = $('#editor'); ed.setSelectionRange(fn.length + 2, fn.length + 2); return; } }
    } else if (n.r2 > n.r1) { for (let c = n.c1; c <= n.c2; c++) S.wb.setInput(S.si, n.r2 + 1, c, `=${fn}(${E.addrToStr(n.r1, c)}:${E.addrToStr(n.r2, c)})`); }
    else { for (let r = n.r1; r <= n.r2; r++) S.wb.setInput(S.si, r, n.c2 + 1, `=${fn}(${E.addrToStr(r, n.c1)}:${E.addrToStr(r, n.c2)})`); }
    App.afterChange();
  };
  CMD.insertFn = name => { App.startEdit(S.act.r, S.act.c, 'edit', '=' + name + '('); const ed = $('#editor'); ed.value = '=' + name + '('; S.edit.text = ed.value; $('#fbar').value = ed.value; ed.setSelectionRange(ed.value.length, ed.value.length); };

  /* sort */
  function currentRegion() {
    let n = sel(); if (n.r1 !== n.r2 || n.c1 !== n.c2) return n; const has = (r, c) => { const v = S.wb.getValue(S.si, r, c); return v !== null && v !== ''; };
    let { r1, c1, r2, c2 } = n; if (!has(r1, c1)) return n;
    const rowHas = r => { for (let c = c1; c <= c2; c++) if (has(r, c)) return true; return false; }, colHas = c => { for (let r = r1; r <= r2; r++) if (has(r, c)) return true; return false; };
    let ch = true; while (ch) { ch = false; if (r1 > 0 && rowHas(r1 - 1)) { r1--; ch = true; } if (r2 < 100000 && rowHas(r2 + 1)) { r2++; ch = true; } if (c1 > 0 && colHas(c1 - 1)) { c1--; ch = true; } if (c2 < 1000 && colHas(c2 + 1)) { c2++; ch = true; } }
    return { r1, c1, r2, c2 };
  }
  function sortRange(n, keyCol, asc, header, extraKeys) {
    const r0 = n.r1 + (header ? 1 : 0); if (r0 >= n.r2) return; const rows = []; const sh = cur();
    for (let r = r0; r <= n.r2; r++) { const cells = []; for (let c = n.c1; c <= n.c2; c++) { const x = sh.cells.get(r * 20000 + c); cells.push(x ? { t: x.t, s: x.s } : null); } rows.push({ r, cells, keys: [keyCol, ...(extraKeys || [])].map(k => S.wb.getValue(S.si, r, k)) }); }
    const dirs = [asc, ...(extraKeys || []).map(() => asc)];
    rows.sort((a, b) => { for (let i = 0; i < a.keys.length; i++) { const va = a.keys[i], vb = b.keys[i]; const ea = va === null || va === '', eb = vb === null || vb === ''; if (ea !== eb) return ea ? 1 : -1; if (ea && eb) continue; const d = E.cmpVals(va, vb) * (dirs[i] ? 1 : -1); if (d) return d; } return a.r - b.r; });
    App.snap(); rows.forEach((row, i) => { const r = r0 + i; row.cells.forEach((cell, j) => { const c = n.c1 + j; if (!cell) S.wb.clearCell(S.si, r, c); else { let t = cell.t; if (t[0] === '=') t = E.shiftFormula(t, r - row.r, 0); S.wb.setInput(S.si, r, c, t); const nc = sh.cells.get(r * 20000 + c); if (nc) nc.s = cell.s ? { ...cell.s } : undefined; } }); }); App.afterChange();
  }
  function hasHeader(n) { if (n.r2 <= n.r1) return false; for (let c = n.c1; c <= n.c2; c++) { const a = S.wb.getValue(S.si, n.r1, c), b = S.wb.getValue(S.si, n.r1 + 1, c); if (typeof a === 'string' && typeof b === 'number') return true; } return false; }
  App.currentRegion = currentRegion;
  CMD.sortAsc = () => { const n = currentRegion(); sortRange(n, Math.min(Math.max(S.act.c, n.c1), n.c2), true, hasHeader(n)); };
  CMD.sortDesc = () => { const n = currentRegion(); sortRange(n, Math.min(Math.max(S.act.c, n.c1), n.c2), false, hasHeader(n)); };
  CMD.sortDialog = () => {
    const n = currentRegion(); const opts = []; for (let c = n.c1; c <= n.c2; c++) { const h = S.wb.display(S.si, n.r1, c); opts.push(`<option value="${c}">Column ${E.numToCol(c)}${h ? ' — ' + esc(h) : ''}</option>`); }
    const d = dialog('Sort', `<div class="prev">Range ${E.addrToStr(n.r1, n.c1)}:${E.addrToStr(n.r2, n.c2)}</div><label class="ck"><input type="checkbox" id="sh" ${hasHeader(n) ? 'checked' : ''}> My data has headers</label><label>Sort by<select id="sc">${opts.join('')}</select></label><label>Then by<select id="sc2"><option value="">(none)</option>${opts.join('')}</select></label><label>Order<select id="so"><option value="1">A to Z / smallest to largest</option><option value="0">Z to A / largest to smallest</option></select></label>`, [{ label: 'Cancel' }, { label: 'Sort', primary: 1, fn: d => { const k2 = d.querySelector('#sc2').value; sortRange(n, +d.querySelector('#sc').value, d.querySelector('#so').value === '1', d.querySelector('#sh').checked, k2 === '' ? [] : [+k2]); } }]);
    d.querySelector('#sc').value = Math.min(Math.max(S.act.c, n.c1), n.c2);
  };
  CMD.removeDuplicates = () => {
    const n = currentRegion(); const seen = new Set(); const header = hasHeader(n); const keep = []; const sh = cur(); App.snap();
    for (let r = n.r1; r <= n.r2; r++) { const key = Array.from({ length: n.c2 - n.c1 + 1 }, (_, i) => S.wb.display(S.si, r, n.c1 + i).toLowerCase()).join('\u0001'); if (header && r === n.r1) { keep.push(r); continue; } if (!seen.has(key)) { seen.add(key); keep.push(r); } }
    const removed = n.r2 - n.r1 + 1 - keep.length; const data = keep.map(r => Array.from({ length: n.c2 - n.c1 + 1 }, (_, i) => { const x = sh.cells.get(r * 20000 + n.c1 + i); return x ? { t: x.t, s: x.s } : null; }));
    for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) S.wb.clearCell(S.si, r, c);
    data.forEach((row, i) => row.forEach((cell, j) => { if (cell) { S.wb.setInput(S.si, n.r1 + i, n.c1 + j, cell.t); const nc = sh.cells.get((n.r1 + i) * 20000 + n.c1 + j); if (nc) nc.s = cell.s; } })); App.afterChange(); App.toast(removed + ' duplicate row' + (removed === 1 ? '' : 's') + ' removed');
  };

  /* filter */
  CMD.toggleFilter = () => {
    const sh = cur(); App.snap(); if (sh.filter) { for (let r = sh.filter.r + 1; r <= sh.filter.r + 100000; r++) { if (!sh.hiddenRows[r] && r > sh.filter.r + 2000 && !S.wb.cell(S.si, r, sh.filter.c1)) break; delete sh.hiddenRows[r]; } sh.filter = null; }
    else { const n = currentRegion(); sh.filter = { r: n.r1, c1: n.c1, c2: n.c2, ex: {} }; }
    App.geomChanged(); App.afterChange(); refreshRibbonState();
  };
  function applyFilter() {
    const sh = cur(), f = sh.filter; if (!f) return; const u = S.wb.usedRange(S.si); let last = f.r; for (let r = f.r + 1; r <= Math.max(u.r2, f.r); r++) { let any = false; for (let c = f.c1; c <= f.c2; c++) { const v = S.wb.getValue(S.si, r, c); if (v !== null && v !== '') { any = true; break; } } if (any) last = r; }
    for (let r = f.r + 1; r <= last; r++) { let hide = false; for (const c in f.ex) { if ((f.ex[c] || []).includes(S.wb.display(S.si, r, +c))) { hide = true; break; } } if (hide) sh.hiddenRows[r] = 1; else delete sh.hiddenRows[r]; }
    f.last = last;
  }
  App.filterMenu = (c, x, y) => {
    const sh = cur(), f = sh.filter; const vals = new Map(); const u = S.wb.usedRange(S.si);
    for (let r = f.r + 1; r <= Math.max(u.r2, f.r + 1); r++) { const d = S.wb.display(S.si, r, c); vals.set(d, (vals.get(d) || 0) + 1); } const ex = new Set(f.ex[c] || []);
    const keys = [...vals.keys()].sort((a, b) => E.cmpVals(E.parseLiteral(a).v, E.parseLiteral(b).v)).slice(0, 400);
    const d = dialog(`Filter column ${E.numToCol(c)}`, `<div style="display:flex;gap:8px"><button class="gbtn" id="fa">Sort A→Z</button><button class="gbtn" id="fz">Sort Z→A</button><button class="gbtn" id="fs">Select all</button><button class="gbtn" id="fn">Clear</button></div><div class="flist" style="padding:6px">${keys.map(k => `<label class="ck" style="padding:3px 4px"><input type="checkbox" data-v="${esc(k)}" ${ex.has(k) ? '' : 'checked'}> ${k === '' ? '(Blanks)' : esc(k)} <span style="color:var(--muted)">(${vals.get(k)})</span></label>`).join('')}</div>`, [{ label: 'Cancel' }, { label: 'Apply', primary: 1, fn: d => { App.snap(); const off = [...d.querySelectorAll('input[data-v]')].filter(i => !i.checked).map(i => i.dataset.v); if (off.length) f.ex[c] = off; else delete f.ex[c]; applyFilter(); App.geomChanged(); App.afterChange(); } }], { width: 420, noEnter: 1 });
    d.querySelector('#fa').onclick = () => { closeModal(); sortRange({ r1: f.r, c1: f.c1, r2: f.last || u.r2, c2: f.c2 }, c, true, true); };
    d.querySelector('#fz').onclick = () => { closeModal(); sortRange({ r1: f.r, c1: f.c1, r2: f.last || u.r2, c2: f.c2 }, c, false, true); };
    d.querySelector('#fs').onclick = () => d.querySelectorAll('input[data-v]').forEach(i => i.checked = true); d.querySelector('#fn').onclick = () => d.querySelectorAll('input[data-v]').forEach(i => i.checked = false);
  };

  /* find & replace */
  CMD.find = () => findDialog(false); CMD.replace = () => findDialog(true);
  let lastFind = { q: '', pos: null };
  function findDialog(rep) {
    const d = dialog(rep ? 'Find & Replace' : 'Find', `<label>Find what<input type="text" id="fq" value="${esc(lastFind.q)}"></label>${rep ? '<label>Replace with<input type="text" id="rq"></label>' : ''}<div class="row2"><label class="ck"><input type="checkbox" id="mc"> Match case</label><label class="ck"><input type="checkbox" id="wc"> Match entire cell</label></div><label class="ck"><input type="checkbox" id="fo"> Search formulas instead of values</label><div id="fres" class="prev" style="min-height:24px;font-size:13px"></div>`, [{ label: 'Close' }, { label: 'Find all', fn: d => { const r = findAll(d); d.querySelector('#fres').textContent = r.length ? r.length + ' match' + (r.length > 1 ? 'es' : '') + ': ' + r.slice(0, 30).map(x => E.addrToStr(x.r, x.c)).join(', ') : 'No matches'; return false; } }, ...(rep ? [{ label: 'Replace all', fn: d => { doReplace(d, true); return false; } }, { label: 'Replace', fn: d => { doReplace(d, false); return false; } }] : []), { label: 'Find next', primary: 1, fn: d => { findNext(d); return false; } }], { width: 460 });
  }
  function matcher(d) { const q = d.querySelector('#fq').value, mc = d.querySelector('#mc').checked, wc = d.querySelector('#wc').checked, fo = d.querySelector('#fo').checked; lastFind.q = q; const Q = mc ? q : q.toLowerCase(); return { q, test: (r, c) => { let t = fo ? S.wb.getInput(S.si, r, c) : S.wb.display(S.si, r, c); if (t === '') return false; if (!mc) t = t.toLowerCase(); return wc ? t === Q : t.includes(Q); } }; }
  function findAll(d) { const m = matcher(d); if (!m.q) return []; const out = []; for (const k of cur().cells.keys()) { const r = Math.floor(k / 20000), c = k % 20000; if (m.test(r, c)) out.push({ r, c }); } out.sort((a, b) => a.r - b.r || a.c - b.c); return out; }
  function findNext(d) { const all = findAll(d); if (!all.length) { d.querySelector('#fres').textContent = 'No matches'; return; } const a = S.act; let nx = all.find(x => x.r > a.r || (x.r === a.r && x.c > a.c)) || all[0]; App.selectCell(nx.r, nx.c); d.querySelector('#fres').textContent = 'Found at ' + E.addrToStr(nx.r, nx.c) + ' (' + all.length + ' total)'; }
  function doReplace(d, all) {
    const m = matcher(d); const rq = d.querySelector('#rq').value; const mc = d.querySelector('#mc').checked; if (!m.q) return;
    const esc_ = m.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const re = new RegExp(d.querySelector('#wc').checked ? '^' + esc_ + '$' : esc_, mc ? 'g' : 'gi'); let n = 0; App.snap();
    const targets = all ? findAll(d) : (m.test(S.act.r, S.act.c) ? [{ r: S.act.r, c: S.act.c }] : []);
    for (const { r, c } of targets) { const t = S.wb.getInput(S.si, r, c); const nt = t.replace(re, () => rq); if (nt !== t) { S.wb.setInput(S.si, r, c, nt); n++; } }
    App.afterChange(); d.querySelector('#fres').textContent = n + ' replaced'; if (!all) findNext(d);
  }

  /* format cells (number) */
  CMD.formatCells = () => {
    const cur_ = getSt().nf || 'General'; const cats = { General: ['General'], Number: ['0', '0.00', '#,##0', '#,##0.00', '#,##0.00;[Red]-#,##0.00'], Currency: ['$#,##0.00', '$#,##0', '€#,##0.00', '£#,##0.00', '¥#,##0', '৳#,##0.00', '₹#,##0.00'], Percentage: ['0%', '0.0%', '0.00%'], Date: ['m/d/yyyy', 'd/m/yyyy', 'yyyy-mm-dd', 'mmm d, yyyy', 'mmmm d, yyyy', 'dddd, mmmm d, yyyy', 'd-mmm-yy'], Time: ['h:mm AM/PM', 'h:mm:ss AM/PM', 'hh:mm', 'hh:mm:ss', 'm/d/yyyy h:mm'], Scientific: ['0.00E+00', '0.0E+00'], Text: ['@'], Custom: [] };
    const found = Object.keys(cats).find(k => cats[k].includes(cur_)) || 'Custom';
    const d = dialog('Format Cells — Number', `<label>Category<select id="nc">${Object.keys(cats).map(k => `<option ${k === found ? 'selected' : ''}>${k}</option>`).join('')}</select></label><label>Format<select id="nl"></select></label><label>Format code (custom)<input type="text" id="nf" value="${esc(cur_)}"></label><div class="prev" id="np"></div>`, [{ label: 'Cancel' }, { label: 'OK', primary: 1, fn: d => { CMD.numFmt(d.querySelector('#nf').value); } }], { width: 460 });
    const sample = (() => { const v = S.wb.getValue(S.si, S.act.r, S.act.c); return typeof v === 'number' ? v : 1234.5678; })();
    const prev = () => { try { d.querySelector('#np').textContent = E.formatValue(sample, d.querySelector('#nf').value || 'General'); } catch (e) { d.querySelector('#np').textContent = '—'; } };
    const fill = () => { const c = d.querySelector('#nc').value; d.querySelector('#nl').innerHTML = cats[c].map(f => `<option ${f === d.querySelector('#nf').value ? 'selected' : ''}>${esc(f)}</option>`).join('') || '<option>(type a code below)</option>'; };
    fill(); prev(); d.querySelector('#nc').onchange = () => { fill(); const f = cats[d.querySelector('#nc').value][0]; if (f) d.querySelector('#nf').value = f; prev(); }; d.querySelector('#nl').onchange = () => { d.querySelector('#nf').value = d.querySelector('#nl').value; prev(); }; d.querySelector('#nf').oninput = prev;
  };

  /* conditional formatting */
  function addCF(rule) { const n = sel(); const sh = cur(); App.snap(); sh.cf.push({ r1: n.r1, c1: n.c1, r2: Math.min(n.r2, 100000), c2: Math.min(n.c2, 2000), ...rule }); App.afterChange(); }
  function cfPrompt(type, title, two) {
    dialog(title, `<label>${two ? 'Between' : 'Value'}<input type="text" id="v1"></label>${two ? '<label>and<input type="text" id="v2"></label>' : ''}<div class="row2"><label>Fill color<input type="color" id="bg" value="#ffc7ce"></label><label>Text color<input type="color" id="fc" value="#9c0006"></label></div>`, [{ label: 'Cancel' }, { label: 'OK', primary: 1, fn: d => { const nv = x => isNaN(parseFloat(x)) ? x : parseFloat(x); addCF({ type, v: nv(d.querySelector('#v1').value), v2: two ? nv(d.querySelector('#v2').value) : undefined, bg: d.querySelector('#bg').value, fc: d.querySelector('#fc').value }); } }]);
  }
  CMD.cf = kind => {
    if (kind === 'gt') cfPrompt('gt', 'Greater than'); else if (kind === 'lt') cfPrompt('lt', 'Less than'); else if (kind === 'eq') cfPrompt('eq', 'Equal to'); else if (kind === 'between') cfPrompt('between', 'Between', true); else if (kind === 'text') cfPrompt('text', 'Text contains');
    else if (kind === 'dup') addCF({ type: 'dup', bg: '#ffc7ce', fc: '#9c0006' }); else if (kind === 'scale') addCF({ type: 'scale' }); else if (kind === 'bar') addCF({ type: 'bar', color: '#63be7b' });
    else if (kind === 'clearSel') { const n = sel(); App.snap(); const sh = cur(); sh.cf = sh.cf.filter(r => r.r1 > n.r2 || r.r2 < n.r1 || r.c1 > n.c2 || r.c2 < n.c1); App.afterChange(); }
    else if (kind === 'clearAll') { App.snap(); cur().cf = []; App.afterChange(); }
  };

  /* format as table / styles */
  const STYLES = { Good: { bg: '#c6efce', fc: '#006100' }, Bad: { bg: '#ffc7ce', fc: '#9c0006' }, Neutral: { bg: '#ffeb9c', fc: '#9c5700' }, 'Heading': { b: 1, fs: 14, fc: '#1f3d2e', bd: { b: { c: '#1f8a4c', w: 2 } } }, Title: { b: 1, fs: 18, fc: '#1f3d2e' }, Total: { b: 1, bd: { t: { c: '#1d2430', w: 1 }, b: { c: '#1d2430', w: 2 } } }, Note: { bg: '#fff2cc', bd: { t: { c: '#b2b2b2' }, b: { c: '#b2b2b2' }, l: { c: '#b2b2b2' }, r: { c: '#b2b2b2' } } }, Input: { bg: '#ffcc99', fc: '#3f3f76' } };
  CMD.cellStyle = name => applyStyle({ ...STYLES[name] });
  CMD.formatTable = scheme => {
    const n = currentRegion(); const colors = { green: ['#1f8a4c', '#e3f4ea'], blue: ['#2b6fd6', '#e6eefb'], orange: ['#d9822b', '#fdf0e1'], gray: ['#4b5563', '#eef0f3'], purple: ['#7c4dbd', '#f0eafa'] }[scheme || 'green']; App.snap();
    for (let r = n.r1; r <= n.r2; r++) for (let c = n.c1; c <= n.c2; c++) { if (r === n.r1) S.wb.setStyle(S.si, r, c, { bg: colors[0], fc: '#ffffff', b: 1 }); else S.wb.setStyle(S.si, r, c, { bg: (r - n.r1) % 2 === 0 ? colors[1] : null, fc: null, b: null }); }
    const sh = cur(); sh.filter = { r: n.r1, c1: n.c1, c2: n.c2, ex: {} }; App.afterChange(); refreshRibbonState();
  };

  /* sheets */
  CMD.addSheet = () => { App.snap(); const s = S.wb.addSheet(); S.si = S.wb.sheets.indexOf(s); App.geomChanged(); resetSel(); App.afterChange(true); };
  function resetSel() { S.sel = { r1: 0, c1: 0, r2: 0, c2: 0 }; S.act = { r: 0, c: 0 }; S.ext = { r: 0, c: 0 }; S.clip = S.clip && S.clip.si === S.si ? S.clip : null; $('#scroller').scrollTop = 0; $('#scroller').scrollLeft = 0; }
  App.switchSheet = i => { if (S.edit) App.commitEdit(0, 0, true); S.si = i; App.geomChanged(); S.sel = { r1: 0, c1: 0, r2: 0, c2: 0 }; S.act = { r: 0, c: 0 }; S.ext = { r: 0, c: 0 }; $('#scroller').scrollTop = 0; $('#scroller').scrollLeft = 0; renderSheetTabs(); App.renderCharts(); App.updateUI(); refreshRibbonState(); };
  function renderSheetTabs() {
    const box = $('#sheettabs'); box.innerHTML = '';
    S.wb.sheets.forEach((s, i) => { const b = el('button', 'stab' + (i === S.si ? ' on' : ''), esc(s.name) + (s.tabColor ? `<span class="tc" style="background:${s.tabColor}"></span>` : '')); b.onclick = () => App.switchSheet(i); b.ondblclick = () => CMD.renameSheet(i); b.oncontextmenu = ev => { ev.preventDefault(); App.switchSheet(i); showMenu([{ label: 'Insert sheet', fn: CMD.addSheet }, { label: 'Rename', fn: () => CMD.renameSheet(i) }, { label: 'Duplicate', fn: () => CMD.dupSheet(i) }, { label: 'Move left', fn: () => moveSheet(i, -1), dis: i === 0 }, { label: 'Move right', fn: () => moveSheet(i, 1), dis: i === S.wb.sheets.length - 1 }, { header: 'Tab color' }, { palette: ['', '#1f8a4c', '#2b6fd6', '#d9822b', '#d6432b', '#7c4dbd', '#0a8fa6', '#4b5563'], pick: c => App.act(() => { s.tabColor = c || null; }) }, '-', { label: 'Delete', fn: () => CMD.delSheet(i), dis: S.wb.sheets.length < 2 }], ev.clientX, ev.clientY - 220); }; box.appendChild(b); });
    const on = box.querySelector('.on'); if (on) on.scrollIntoView({ inline: 'nearest', block: 'nearest' });
  }
  App.renderSheetTabs = renderSheetTabs; $('#addsheet').onclick = () => CMD.addSheet();
  CMD.renameSheet = i => App.prompt('Rename sheet', S.wb.sheets[i].name, v => { v = v.trim().replace(/[\\/?*[\]:]/g, ''); if (!v) return; if (S.wb.sheetIndex(v) >= 0 && S.wb.sheetIndex(v) !== i) { App.toast('A sheet with that name already exists'); return; } App.snap(); S.wb.renameSheet(i, v.slice(0, 31)); App.afterChange(true); });
  CMD.delSheet = i => { if (S.wb.sheets.length < 2) return; dialog('Delete sheet?', `<div>“${esc(S.wb.sheets[i].name)}” will be deleted. You can still undo this.</div>`, [{ label: 'Cancel' }, { label: 'Delete', primary: 1, fn: () => { App.snap(); S.wb.removeSheet(i); S.si = Math.min(S.si, S.wb.sheets.length - 1); App.geomChanged(); resetSel(); App.afterChange(true); } }]); };
  CMD.dupSheet = i => { App.snap(); S.wb.duplicateSheet(i); S.si = i + 1; App.geomChanged(); App.afterChange(true); };
  function moveSheet(i, d) { App.snap(); const a = S.wb.sheets; const [x] = a.splice(i, 1); a.splice(i + d, 0, x); S.si = i + d; S.wb.invalidate(); App.afterChange(true); }

  /* ---------- file ops ---------- */
  const tauri = window.__TAURI__;
  async function saveBlob(name, blob, types) {
    if (tauri && tauri.dialog && tauri.fs) { try { const path = await tauri.dialog.save({ defaultPath: name, filters: types }); if (!path) return null; await tauri.fs.writeFile(path, new Uint8Array(await blob.arrayBuffer())); return { path }; } catch (e) { console.warn(e); } }
    if (window.showSaveFilePicker) { try { const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: types[0].name, accept: { [blob.type || 'application/octet-stream']: types[0].extensions.map(x => '.' + x) } }] }); const w = await h.createWritable(); await w.write(blob); await w.close(); return { handle: h }; } catch (e) { if (e.name === 'AbortError') return null; } }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); return { downloaded: true };
  }
  const markSaved = () => { S.dirty = false; $('#savestate').textContent = 'Saved'; };
  CMD.save = async () => { if (S.edit) App.commitEdit(0, 0, true); const blob = FmXlsx.exportXlsx(S.wb); if (S.handle && S.handle.createWritable) { try { const w = await S.handle.createWritable(); await w.write(blob); await w.close(); markSaved(); App.toast('Saved'); return; } catch (e) { } } return CMD.saveAs(); };
  CMD.saveAs = async () => { if (S.edit) App.commitEdit(0, 0, true); const name = (S.name || 'Book1').replace(/\.(xlsx|csv)$/i, '') + '.xlsx'; const r = await saveBlob(name, FmXlsx.exportXlsx(S.wb), [{ name: 'Excel workbook', extensions: ['xlsx'] }]); if (r) { if (r.handle) { S.handle = r.handle; S.name = r.handle.name.replace(/\.xlsx$/i, ''); $('#docname').value = S.name; } markSaved(); App.toast(r.downloaded ? 'Downloaded ' + name : 'Saved'); } };
  CMD.saveCsv = async () => { const r = await saveBlob((S.name || 'Book1') + '.csv', FmXlsx.exportCsv(S.wb, S.si), [{ name: 'CSV', extensions: ['csv'] }]); if (r) App.toast('CSV saved (current sheet)'); };
  CMD.new = () => { const go = () => { loadWb(new E.Workbook(), 'Book1'); }; if (S.dirty) dialog('Start a new workbook?', '<div>You have unsaved changes. Start over anyway?</div>', [{ label: 'Cancel' }, { label: 'Discard & new', primary: 1, fn: go }]); else go(); };
  CMD.open = async () => {
    if (tauri && tauri.dialog && tauri.fs) { try { const path = await tauri.dialog.open({ multiple: false, filters: [{ name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'csv', 'tsv'] }] }); if (!path) return; const bytes = await tauri.fs.readFile(path); const nm = String(path).split(/[\\/]/).pop(); return openBuffer(bytes.buffer, nm); } catch (e) { console.warn(e); } }
    if (window.showOpenFilePicker) { try { const [h] = await window.showOpenFilePicker({ types: [{ description: 'Spreadsheets', accept: { 'application/octet-stream': ['.xlsx', '.xlsm', '.csv', '.tsv'] } }] }); const f = await h.getFile(); await openFile(f); S.handle = /\.xlsx$/i.test(f.name) ? h : null; return; } catch (e) { if (e.name === 'AbortError') return; } }
    $('#filein').click();
  };
  $('#filein').onchange = e => { const f = e.target.files[0]; if (f) openFile(f); e.target.value = ''; };
  async function openFile(f) { return openBuffer(await f.arrayBuffer(), f.name); }
  async function openBuffer(buf, name) {
    try {
      let wb; if (/\.(csv|tsv|txt)$/i.test(name)) wb = FmXlsx.importCsv(new TextDecoder().decode(buf), name); else wb = await FmXlsx.importXlsx(buf);
      S.handle = null; loadWb(wb, name.replace(/\.[^.]+$/, '')); App.toast('Opened ' + name);
    } catch (e) { console.error(e); App.alert('Couldn’t open that file', String(e.message || e)); }
  }
  function loadWb(wb, name) { S.wb = wb; S.si = 0; S.name = name; S.undo = []; S.redo = []; S.dirty = false; S.clip = null; S.handle = null; $('#docname').value = name; App.geomChanged(); resetSel(); renderSheetTabs(); App.renderCharts(); App.updateUI(); refreshRibbonState(); $('#savestate').textContent = ''; }
  App.loadWb = loadWb;
  window.addEventListener('dragover', e => e.preventDefault()); window.addEventListener('drop', e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) openFile(f); });
  window.addEventListener('beforeunload', e => { if (S.dirty && !(window.__TAURI__)) { e.preventDefault(); e.returnValue = ''; } });
  CMD.print = () => {
    const u = S.wb.usedRange(S.si); let h = '<table style="border-collapse:collapse;font:12px Segoe UI,Arial">'; const sh = cur();
    for (let r = 0; r <= u.r2; r++) { if (sh.hiddenRows[r]) continue; h += '<tr>'; for (let c = 0; c <= u.c2; c++) { if (sh.hiddenCols[c]) continue; const st = S.wb.getStyle(S.si, r, c) || {}; const v = S.wb.getValue(S.si, r, c); h += `<td style="border:${sh.showGrid ? '1px solid #ddd' : '0'};padding:2px 6px;text-align:${st.ha || (typeof v === 'number' ? 'right' : 'left')};${st.b ? 'font-weight:700;' : ''}${st.i ? 'font-style:italic;' : ''}${st.bg ? 'background:' + st.bg + ';' : ''}${st.fc ? 'color:' + st.fc + ';' : ''}min-width:${(sh.colW[c] || 84) * 0.8}px">${esc(S.wb.display(S.si, r, c))}</td>`; } h += '</tr>'; }
    const f = document.createElement('iframe'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0'; document.body.appendChild(f); f.contentDocument.write('<html><head><title>' + esc(S.name) + '</title><style>@page{size:' + S.page.size + ' ' + S.page.orient + ';margin:' + ({ normal: '19mm', narrow: '10mm', wide: '25mm' })[S.page.margins] + '}</style></head><body>' + h + '</table></body></html>'); f.contentDocument.close(); setTimeout(() => { f.contentWindow.print(); setTimeout(() => f.remove(), 1500); }, 200);
  };
  $('#docname').addEventListener('change', e => { S.name = e.target.value.trim() || 'Book1'; });
  $('#docname').addEventListener('keydown', e => { if (e.key === 'Enter') { e.target.blur(); App.focusGrid(); } });

  /* theme */
  function setTheme(t) { document.documentElement.dataset.theme = t; try { localStorage.setItem('fmsheets.theme', t); } catch (e) { } $('#btn-theme').innerHTML = ic(t === 'dark' ? 'sun' : 'moon'); App.readTheme(); }
  CMD.theme = () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'); $('#btn-theme').onclick = CMD.theme;
  { let t = 'light'; try { t = localStorage.getItem('fmsheets.theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); } catch (e) { } setTheme(t); }

  /* ---------- function library ---------- */
  const FDOC = {
    SUM: 'SUM(number1, …) — Adds numbers', AVERAGE: 'AVERAGE(number1, …) — Average of numbers', MIN: 'MIN(number1, …) — Smallest value', MAX: 'MAX(number1, …) — Largest value', COUNT: 'COUNT(value1, …) — Counts numbers', COUNTA: 'COUNTA(value1, …) — Counts non-empty cells', COUNTBLANK: 'COUNTBLANK(range) — Counts empty cells', MEDIAN: 'MEDIAN(number1, …) — Middle value', MODE: 'MODE(number1, …) — Most frequent value', STDEV: 'STDEV(number1, …) — Sample standard deviation', VAR: 'VAR(number1, …) — Sample variance', LARGE: 'LARGE(array, k) — k-th largest', SMALL: 'SMALL(array, k) — k-th smallest', RANK: 'RANK(number, ref, [order]) — Rank in a list', PRODUCT: 'PRODUCT(number1, …) — Multiplies numbers', SUMPRODUCT: 'SUMPRODUCT(array1, …) — Sum of products',
    IF: 'IF(test, then, [else]) — Returns one value if TRUE, another if FALSE', IFS: 'IFS(test1, value1, …) — First TRUE test wins', IFERROR: 'IFERROR(value, value_if_error) — Catches errors', IFNA: 'IFNA(value, value_if_na) — Catches #N/A', AND: 'AND(logical1, …) — TRUE if all TRUE', OR: 'OR(logical1, …) — TRUE if any TRUE', NOT: 'NOT(logical) — Reverses TRUE/FALSE', XOR: 'XOR(logical1, …) — Exclusive OR', SWITCH: 'SWITCH(expr, val1, res1, …, [default]) — Match a value', CHOOSE: 'CHOOSE(index, v1, v2, …) — Pick by position',
    VLOOKUP: 'VLOOKUP(lookup, table, col, [approx]) — Vertical lookup', HLOOKUP: 'HLOOKUP(lookup, table, row, [approx]) — Horizontal lookup', XLOOKUP: 'XLOOKUP(lookup, lookup_array, return_array, [if_not_found]) — Modern lookup', INDEX: 'INDEX(array, row, [col]) — Value at position', MATCH: 'MATCH(lookup, array, [type]) — Position of a value', ROWS: 'ROWS(array) — Number of rows', COLUMNS: 'COLUMNS(array) — Number of columns', ROW: 'ROW([ref]) — Row number', COLUMN: 'COLUMN([ref]) — Column number', TRANSPOSE: 'TRANSPOSE(array) — Swap rows and columns', UNIQUE: 'UNIQUE(array) — Distinct values',
    COUNTIF: 'COUNTIF(range, criteria) — Count matching cells', COUNTIFS: 'COUNTIFS(range1, crit1, …) — Count with several criteria', SUMIF: 'SUMIF(range, criteria, [sum_range]) — Sum matching cells', SUMIFS: 'SUMIFS(sum_range, range1, crit1, …) — Sum with several criteria', AVERAGEIF: 'AVERAGEIF(range, criteria, [avg_range])', AVERAGEIFS: 'AVERAGEIFS(avg_range, range1, crit1, …)', MAXIFS: 'MAXIFS(max_range, range1, crit1, …)', MINIFS: 'MINIFS(min_range, range1, crit1, …)',
    ROUND: 'ROUND(number, digits) — Round to digits', ROUNDUP: 'ROUNDUP(number, digits) — Round away from zero', ROUNDDOWN: 'ROUNDDOWN(number, digits) — Round toward zero', INT: 'INT(number) — Round down to integer', MOD: 'MOD(number, divisor) — Remainder', ABS: 'ABS(number) — Absolute value', SQRT: 'SQRT(number) — Square root', POWER: 'POWER(number, power) — Raise to a power', PI: 'PI() — 3.14159…', RAND: 'RAND() — Random 0–1', RANDBETWEEN: 'RANDBETWEEN(low, high) — Random integer', CEILING: 'CEILING(number, significance)', FLOOR: 'FLOOR(number, significance)', EXP: 'EXP(n)', LN: 'LN(n)', LOG: 'LOG(n, [base])', LOG10: 'LOG10(n)', FACT: 'FACT(n)', GCD: 'GCD(n1, …)', LCM: 'LCM(n1, …)', SIGN: 'SIGN(n)', TRUNC: 'TRUNC(n, [digits])', EVEN: 'EVEN(n)', ODD: 'ODD(n)',
    CONCAT: 'CONCAT(text1, …) — Join text', CONCATENATE: 'CONCATENATE(text1, …) — Join text', TEXTJOIN: 'TEXTJOIN(delim, ignore_empty, text1, …) — Join with delimiter', LEFT: 'LEFT(text, [n]) — First characters', RIGHT: 'RIGHT(text, [n]) — Last characters', MID: 'MID(text, start, n) — Middle characters', LEN: 'LEN(text) — Length', UPPER: 'UPPER(text)', LOWER: 'LOWER(text)', PROPER: 'PROPER(text) — Capitalize Each Word', TRIM: 'TRIM(text) — Remove extra spaces', SUBSTITUTE: 'SUBSTITUTE(text, old, new, [n]) — Replace text', REPLACE: 'REPLACE(text, start, n, new)', FIND: 'FIND(find, within, [start]) — Case-sensitive position', SEARCH: 'SEARCH(find, within, [start]) — Position, wildcards ok', REPT: 'REPT(text, times)', TEXT: 'TEXT(value, format) — Format a number as text', VALUE: 'VALUE(text) — Text to number', EXACT: 'EXACT(a, b)', CHAR: 'CHAR(n)', CODE: 'CODE(text)', CLEAN: 'CLEAN(text)', FIXED: 'FIXED(n, [decimals], [no_commas])',
    TODAY: 'TODAY() — Current date', NOW: 'NOW() — Current date and time', DATE: 'DATE(year, month, day)', TIME: 'TIME(h, m, s)', YEAR: 'YEAR(date)', MONTH: 'MONTH(date)', DAY: 'DAY(date)', HOUR: 'HOUR(time)', MINUTE: 'MINUTE(time)', SECOND: 'SECOND(time)', WEEKDAY: 'WEEKDAY(date, [type])', DAYS: 'DAYS(end, start)', EDATE: 'EDATE(date, months)', EOMONTH: 'EOMONTH(date, months)', DATEDIF: 'DATEDIF(start, end, "Y"|"M"|"D")', DATEVALUE: 'DATEVALUE(text)',
    PMT: 'PMT(rate, nper, pv, [fv]) — Loan payment', FV: 'FV(rate, nper, pmt, [pv]) — Future value', PV: 'PV(rate, nper, pmt, [fv]) — Present value', NPV: 'NPV(rate, v1, …) — Net present value',
    ISBLANK: 'ISBLANK(v)', ISNUMBER: 'ISNUMBER(v)', ISTEXT: 'ISTEXT(v)', ISERROR: 'ISERROR(v)', ISNA: 'ISNA(v)', ISEVEN: 'ISEVEN(n)', ISODD: 'ISODD(n)', NA: 'NA() — #N/A', N: 'N(v)', T: 'T(v)'
  };
  Object.assign(App.FDOC, FDOC);
  const FCATS = { 'Math & Trig': 'SUM PRODUCT SUMPRODUCT ROUND ROUNDUP ROUNDDOWN INT MOD ABS SQRT POWER PI RAND RANDBETWEEN CEILING FLOOR EXP LN LOG LOG10 FACT GCD LCM SIGN TRUNC SUMIF SUMIFS'.split(' '), Statistical: 'AVERAGE MIN MAX COUNT COUNTA COUNTBLANK MEDIAN MODE STDEV VAR LARGE SMALL RANK COUNTIF COUNTIFS AVERAGEIF AVERAGEIFS MAXIFS MINIFS'.split(' '), Logical: 'IF IFS IFERROR IFNA AND OR NOT XOR SWITCH CHOOSE'.split(' '), 'Lookup & Reference': 'VLOOKUP HLOOKUP XLOOKUP INDEX MATCH ROWS COLUMNS ROW COLUMN TRANSPOSE UNIQUE'.split(' '), Text: 'CONCAT TEXTJOIN LEFT RIGHT MID LEN UPPER LOWER PROPER TRIM SUBSTITUTE REPLACE FIND SEARCH REPT TEXT VALUE EXACT'.split(' '), 'Date & Time': 'TODAY NOW DATE TIME YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY DAYS EDATE EOMONTH DATEDIF DATEVALUE'.split(' '), Financial: 'PMT FV PV NPV'.split(' '), Information: 'ISBLANK ISNUMBER ISTEXT ISERROR ISNA ISEVEN ISODD NA'.split(' ') };
  CMD.insertFunctionDialog = () => {
    const d = dialog('Insert Function', `<label>Search<input type="text" id="fsq" placeholder="e.g. lookup, sum, date…"></label><div class="flist" id="fsl"></div>`, [{ label: 'Cancel' }, { label: 'Insert', primary: 1, fn: d => { const hl = d.querySelector('.fi.hl'); if (hl) CMD.insertFn(hl.dataset.n); } }], { width: 520 });
    const draw_ = () => { const q = d.querySelector('#fsq').value.toLowerCase(); const list = App.FNAMES.filter(f => f.toLowerCase().includes(q) || (FDOC[f] || '').toLowerCase().includes(q)); d.querySelector('#fsl').innerHTML = list.map((f, i) => `<div class="fi${i === 0 ? ' hl' : ''}" data-n="${f}"><b>${f}</b><small>${esc(FDOC[f] || '')}</small></div>`).join('') || '<div class="fi">No matching function</div>'; d.querySelectorAll('.fi').forEach(x => { x.onclick = () => { d.querySelectorAll('.fi').forEach(y => y.classList.remove('hl')); x.classList.add('hl'); }; x.ondblclick = () => { closeModal(); CMD.insertFn(x.dataset.n); }; }); };
    d.querySelector('#fsq').oninput = draw_; draw_();
  };
  $('#fx-btn').onclick = () => CMD.insertFunctionDialog();

  S.page = S.page || { margins: 'normal', orient: 'portrait', size: 'A4' };
  CMD.fontStep = d => { const sz = SIZES; const c = getSt().fs || 11; let k = d > 0 ? sz.find(x => x > c) : [...sz].reverse().find(x => x < c); if (k) CMD.fsize(k); };
  CMD.formatPainter = () => { const st = getSt(); const keys = ['b', 'i', 'u', 'st', 'ff', 'fs', 'fc', 'bg', 'ha', 'va', 'wrap', 'ind', 'nf', 'bd']; S.painter = {}; keys.forEach(k => S.painter[k] = st[k] === undefined ? null : st[k]); document.body.classList.add('painting'); App.toast('Format Painter: click or drag over cells to apply'); };
  window.addEventListener('mouseup', ev => { if (!S.painter || !ev.target.closest || !ev.target.closest('#gridwrap')) return; setTimeout(() => { if (!S.painter) return; const p = S.painter; S.painter = null; document.body.classList.remove('painting'); applyStyle(p); }, 0); });
  CMD.deleteNote = () => { const sh = cur(); const k = S.act.r * 20000 + S.act.c; if (sh.notes && sh.notes[k]) { App.snap(); delete sh.notes[k]; App.afterChange(); App.toast('Note deleted'); } else App.toast('No note in this cell'); };
  /* ---------- ribbon definition ---------- */
  const PALETTE = ['', '#000000', '#434343', '#666666', '#999999', '#cccccc', '#efefef', '#ffffff', '#1f8a4c', '#e3f4ea', '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff', '#e6b8af', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#d9d2e9', '#ead1dc', '#dd7e6b', '#f9cb9c'];
  const FONTS = ['Calibri', 'Arial', 'Segoe UI', 'Verdana', 'Tahoma', 'Georgia', 'Times New Roman', 'Courier New', 'Consolas', 'Trebuchet MS'];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72];
  const NUMFMTS = [['General', 'General'], ['Number', '0.00'], ['Currency', '$#,##0.00'], ['Accounting', '_($* #,##0.00_)'], ['Short Date', 'm/d/yyyy'], ['Long Date', 'dddd, mmmm d, yyyy'], ['Time', 'h:mm:ss AM/PM'], ['Percentage', '0.00%'], ['Fraction', '# ?/?'], ['Scientific', '0.00E+00'], ['Text', '@']];
  const reg = [];  // searchable commands
  function btn(o) { // o: {icon,label,cmd,arg,big,tip,toggleKey,menu}
    const b = el('button', 'btn' + (o.big ? ' big' : '') + (o.sm ? ' sm' : '') + (o.menu ? ' dropbtn' : ''), (o.icon ? ic(o.icon, o.big ? 'lg' : '') : '') + (o.label && (o.big || o.sm) ? `<span>${esc(o.label)}</span>` : '') + (o.swatch ? '<span class="sw"></span>' : '') + (o.menu ? `<svg class="car" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><path d="${P.car}"/></svg>` : ''));
    b.title = (o.tip || o.label || '') + (o.key ? ` (${o.key})` : ''); if (o.swatch) b.querySelector('.sw').style.background = o.swatch;
    b.onmousedown = e => e.preventDefault(); // keep focus in grid/editor
    b.onclick = () => { if (S.edit && !o.menu && o.cmd !== 'undo') { } if (o.menu) dropdown(b, o.menu); else { if (o.fn) o.fn(); else CMD[o.cmd](o.arg); App.focusGrid(); } };
    if (o.state) b.dataset.state = o.state; if (o.label && (o.cmd || o.fn)) reg.push({ label: o.label, tip: o.tip, run: () => o.fn ? o.fn() : CMD[o.cmd](o.arg), icon: o.icon });
    return b;
  }
  const colorMenu = (cmd, title) => () => [{ header: title }, { palette: PALETTE, pick: c => CMD[cmd](c) }];
  const group = (label, ...cols) => { const g = el('div', 'group'); const b = el('div', 'body'); cols.forEach(c => b.appendChild(c)); g.appendChild(b); g.appendChild(el('div', 'glabel', label)); return g; };
  const col = (...items) => { const c = el('div', 'col'); items.forEach(i => c.appendChild(i)); return c; };
  const row = (...items) => { const c = el('div', 'row'); items.forEach(i => c.appendChild(i)); return c; };
  const sep = () => el('div', 'sep');
  function select(opts, val, onchange, width, title) { const s = el('select', 'sel'); s.style.width = width + 'px'; s.title = title || ''; s.innerHTML = opts.map(o => Array.isArray(o) ? `<option value="${esc(o[1])}">${esc(o[0])}</option>` : `<option>${esc(o)}</option>`).join(''); s.value = val; s.onchange = () => { onchange(s.value); App.focusGrid(); }; return s; }
  const colorBtn = (icon, cmd, title, def) => { const b = btn({ icon, tip: title, menu: colorMenu(cmd, title), swatch: def }); b.classList.add('colorbtn'); const main = b; return main; };

  const tabs = [
    {
      id: 'home', label: 'Home', build: () => [
        group('Clipboard', btn({ icon: 'paste', label: 'Paste', big: 1, cmd: 'paste', key: 'Ctrl+V', menu: () => [{ label: 'Paste', cmd: 'paste', key: 'Ctrl+V' }, { label: 'Values only', cmd: 'pasteValues' }, { label: 'Formulas', cmd: 'pasteFormulas' }, { label: 'Formatting only', cmd: 'pasteFormats' }, { label: 'Transpose', cmd: 'pasteTranspose' }] }), col(btn({ icon: 'cut', label: 'Cut', sm: 1, cmd: 'cut' }), btn({ icon: 'copy', label: 'Copy', sm: 1, cmd: 'copy' }), btn({ icon: 'brush', label: 'Format Painter', sm: 1, cmd: 'formatPainter' }))),
        group('Font', col(row(select(FONTS, 'Calibri', CMD.font, 100, 'Font'), select(SIZES, 11, CMD.fsize, 50, 'Font size'), btn({ label: 'A\u25B4', tip: 'Increase font size', sm: 1, fn: () => CMD.fontStep(1) }), btn({ label: 'A\u25BE', tip: 'Decrease font size', sm: 1, fn: () => CMD.fontStep(-1) })), row(btn({ icon: 'bold', tip: 'Bold', key: 'Ctrl+B', cmd: 'bold', state: 'b' }), btn({ icon: 'italic', tip: 'Italic', key: 'Ctrl+I', cmd: 'italic', state: 'i' }), btn({ icon: 'underline', tip: 'Underline', key: 'Ctrl+U', cmd: 'underline', state: 'u' }), btn({ icon: 'strike', tip: 'Strikethrough', cmd: 'strike', state: 'st' }), sep(), btn({ icon: 'border', tip: 'Borders', menu: () => [['All borders', 'all'], ['Outside borders', 'outside'], ['Thick outside', 'thick'], ['Top border', 'top'], ['Bottom border', 'bottom'], ['Left border', 'left'], ['Right border', 'right'], ['Top and thick bottom', 'topbottom'], ['No border', 'none']].map(([l, a]) => ({ label: l, cmd: 'borders', arg: a })) }), colorBtn('fill', 'fillColor', 'Fill color', '#ffe066'), colorBtn('text', 'fontColor', 'Font color', '#d6432b')))),
        group('Alignment', col(row(btn({ icon: 'top', tip: 'Top align', fn: () => CMD.valign('top') }), btn({ icon: 'mid', tip: 'Middle align', fn: () => CMD.valign('middle') }), btn({ icon: 'bottom', tip: 'Bottom align', fn: () => CMD.valign('bottom') }), sep(), btn({ icon: 'wrap', label: 'Wrap Text', sm: 1, cmd: 'wrap', tip: 'Wrap text', state: 'wrap' })), row(btn({ icon: 'left', tip: 'Align left', fn: () => CMD.align('left'), state: 'ha:left' }), btn({ icon: 'center', tip: 'Center', fn: () => CMD.align('center'), state: 'ha:center' }), btn({ icon: 'right', tip: 'Align right', fn: () => CMD.align('right'), state: 'ha:right' }), sep(), btn({ icon: 'outdent', tip: 'Decrease indent', fn: () => CMD.indent(-1) }), btn({ icon: 'indent', tip: 'Increase indent', fn: () => CMD.indent(1) }), sep(), btn({ icon: 'merge', label: 'Merge', sm: 1, cmd: 'merge', arg: 'all', menu: () => [{ label: 'Merge & Center', cmd: 'merge', arg: 'all' }, { label: 'Merge Across', cmd: 'merge', arg: 'across' }, { label: 'Unmerge Cells', cmd: 'merge', arg: 'unmerge' }] })))),
        group('Number', col(row(select(NUMFMTS, 'General', v => CMD.numFmt(v), 108, 'Number format')), row(btn({ tip: 'Currency', label: '$', sm: 1, cmd: 'currency' }), btn({ tip: 'Percent', label: '%', sm: 1, cmd: 'percent' }), btn({ tip: 'Comma', label: ',', sm: 1, cmd: 'comma' }), sep(), btn({ label: '.0\u2190', tip: 'Decrease decimals', sm: 1, fn: () => CMD.decimals(-1) }), btn({ label: '.00\u2192', tip: 'Increase decimals', sm: 1, fn: () => CMD.decimals(1) })))),
        group('Styles', col(btn({ icon: 'cf', label: 'Conditional Formatting', sm: 1, menu: () => [{ header: 'Highlight cells' }, { label: 'Greater than…', cmd: 'cf', arg: 'gt' }, { label: 'Less than…', cmd: 'cf', arg: 'lt' }, { label: 'Between…', cmd: 'cf', arg: 'between' }, { label: 'Equal to…', cmd: 'cf', arg: 'eq' }, { label: 'Text contains…', cmd: 'cf', arg: 'text' }, { label: 'Duplicate values', cmd: 'cf', arg: 'dup' }, '-', { label: 'Data bars', cmd: 'cf', arg: 'bar' }, { label: 'Color scale', cmd: 'cf', arg: 'scale' }, '-', { label: 'Clear rules from selection', cmd: 'cf', arg: 'clearSel' }, { label: 'Clear rules from sheet', cmd: 'cf', arg: 'clearAll' }] }), btn({ icon: 'table', label: 'Format as Table', sm: 1, menu: () => ['green', 'blue', 'orange', 'purple', 'gray'].map(s => ({ label: s[0].toUpperCase() + s.slice(1) + ' table', cmd: 'formatTable', arg: s })) }), btn({ icon: 'styles', label: 'Cell Styles', sm: 1, menu: () => Object.keys(STYLES).map(s => ({ label: s, cmd: 'cellStyle', arg: s })) }))),
        group('Cells', col(btn({ icon: 'insrow', label: 'Insert', sm: 1, menu: () => [{ label: 'Insert row above', cmd: 'insRow', icon: 'insrow' }, { label: 'Insert column left', cmd: 'insCol', icon: 'inscol' }, { label: 'Insert sheet', cmd: 'addSheet', icon: 'sheet' }] }), btn({ icon: 'delrow', label: 'Delete', sm: 1, menu: () => [{ label: 'Delete rows', cmd: 'delRow' }, { label: 'Delete columns', cmd: 'delCol' }, { label: 'Delete sheet', fn: () => CMD.delSheet(S.si) }] }), btn({ icon: 'grid', label: 'Format', sm: 1, menu: () => [{ label: 'Row height…', cmd: 'rowHeight' }, { label: 'AutoFit row height', fn: () => CMD.autofitRow() }, { label: 'Column width…', cmd: 'colWidth' }, { label: 'AutoFit column width', fn: () => CMD.autofitCol() }, '-', { label: 'Hide rows', cmd: 'hideRows' }, { label: 'Hide columns', cmd: 'hideCols' }, { label: 'Unhide rows', cmd: 'unhideRows' }, { label: 'Unhide columns', cmd: 'unhideCols' }, '-', { label: 'Rename sheet…', fn: () => CMD.renameSheet(S.si) }, { label: 'Format cells…', cmd: 'formatCells', key: 'Ctrl+1' }] }))),
        group('Editing', col(btn({ icon: 'sigma', label: 'AutoSum', sm: 1, cmd: 'autosum', menu: () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map(f => ({ label: f[0] + f.slice(1).toLowerCase(), cmd: 'autosum', arg: f })) }), btn({ icon: 'fill', label: 'Fill', sm: 1, menu: () => [{ label: 'Down', cmd: 'fillDown', key: 'Ctrl+D' }, { label: 'Right', cmd: 'fillRight', key: 'Ctrl+R' }, { label: 'Series', cmd: 'fillSeries' }] }), btn({ icon: 'clear', label: 'Clear', sm: 1, menu: () => [{ label: 'Clear all', cmd: 'clearAll' }, { label: 'Clear formats', cmd: 'clearFormats' }, { label: 'Clear contents', cmd: 'clearContents', key: 'Del' }] })), btn({ icon: 'sortasc', label: 'Sort & Filter', big: 1, menu: () => [{ label: 'Sort A to Z', cmd: 'sortAsc', icon: 'sortasc' }, { label: 'Sort Z to A', cmd: 'sortDesc', icon: 'sortdesc' }, { label: 'Custom sort…', cmd: 'sortDialog' }, '-', { label: 'Filter', cmd: 'toggleFilter', icon: 'filter', check: !!cur().filter }] }), btn({ icon: 'search', label: 'Find & Select', big: 1, menu: () => [{ label: 'Find…', cmd: 'find', key: 'Ctrl+F' }, { label: 'Replace…', cmd: 'replace', key: 'Ctrl+H' }, { label: 'Go to…', fn: () => { $('#namebox').focus(); } }] }))
      ]
    },
    {
      id: 'insert', label: 'Insert', build: () => [
        group('Tables', btn({ icon: 'table', label: 'PivotTable', big: 1, cmd: 'pivot' }), btn({ icon: 'table', label: 'Table', big: 1, cmd: 'formatTable', arg: 'green' }), btn({ icon: 'filter', label: 'Filter', big: 1, cmd: 'toggleFilter' })),
        group('Cells', btn({ icon: 'insrow', label: 'Row', big: 1, cmd: 'insRow' }), btn({ icon: 'inscol', label: 'Column', big: 1, cmd: 'insCol' }), btn({ icon: 'sheet', label: 'Sheet', big: 1, cmd: 'addSheet' })),
        group('Charts', btn({ icon: 'chart', label: 'Column', big: 1, fn: () => App.insertChart('col') }), btn({ icon: 'chart', label: 'Stacked', big: 1, fn: () => App.insertChart('stack') }), btn({ icon: 'chart', label: 'Bar', big: 1, fn: () => App.insertChart('bar') }), btn({ icon: 'line', label: 'Line', big: 1, fn: () => App.insertChart('line') }), btn({ icon: 'area', label: 'Area', big: 1, fn: () => App.insertChart('area') }), btn({ icon: 'pie', label: 'Pie', big: 1, fn: () => App.insertChart('pie') }), btn({ icon: 'scatter', label: 'Scatter', big: 1, fn: () => App.insertChart('scatter') })),
        group('Links', btn({ icon: 'link', label: 'Link', big: 1, cmd: 'link', key: 'Ctrl+K' }), btn({ icon: 'text', label: 'Note', big: 1, cmd: 'note' })),
        group('Symbols', btn({ icon: 'fx', label: 'Function', big: 1, cmd: 'insertFunctionDialog' }), btn({ icon: 'sigma', label: 'AutoSum', big: 1, cmd: 'autosum' }))
      ]
    },
    {
      id: 'layout', label: 'Page Layout', build: () => [
        group('Page Setup', btn({ icon: 'file', label: 'Margins', big: 1, menu: () => [['Normal', 'normal'], ['Narrow', 'narrow'], ['Wide', 'wide']].map(([l, a]) => ({ label: l, check: S.page.margins === a, fn: () => { S.page.margins = a; App.toast('Margins: ' + l); } })) }), btn({ icon: 'file', label: 'Orientation', big: 1, menu: () => [['Portrait', 'portrait'], ['Landscape', 'landscape']].map(([l, a]) => ({ label: l, check: S.page.orient === a, fn: () => { S.page.orient = a; App.toast('Orientation: ' + l); } })) }), btn({ icon: 'file', label: 'Size', big: 1, menu: () => ['A4', 'Letter', 'Legal', 'A3'].map(l => ({ label: l, check: S.page.size === l, fn: () => { S.page.size = l; App.toast('Paper: ' + l); } })) }), btn({ icon: 'print', label: 'Print', big: 1, cmd: 'print' })),
        group('Sheet Options', col(btn({ label: 'View Gridlines', sm: 1, icon: 'grid', cmd: 'gridlines', state: 'grid' }), btn({ label: 'View Headings', sm: 1, icon: 'col', cmd: 'headings', state: 'head' }))),
        group('Arrange', btn({ icon: 'freeze', label: 'Freeze Panes', big: 1, menu: () => [{ label: 'Freeze panes (at selection)', cmd: 'freeze', arg: 'at' }, { label: 'Freeze top row', cmd: 'freeze', arg: 'row' }, { label: 'Freeze first column', cmd: 'freeze', arg: 'col' }, { label: 'Unfreeze panes', cmd: 'freeze', arg: 'none' }] }))
      ]
    },
    {
      id: 'formulas', label: 'Formulas', build: () => [
        group('Function Library', btn({ icon: 'fx', label: 'Insert Function', big: 1, cmd: 'insertFunctionDialog' }), btn({ icon: 'sigma', label: 'AutoSum', big: 1, cmd: 'autosum', menu: () => ['SUM', 'AVERAGE', 'COUNT', 'MAX', 'MIN'].map(f => ({ label: f[0] + f.slice(1).toLowerCase(), cmd: 'autosum', arg: f })) }),
          ...Object.keys(FCATS).map(k => btn({ icon: ({ 'Math & Trig': 'calc', Statistical: 'chart', Logical: 'check', 'Lookup & Reference': 'search', Text: 'text', 'Date & Time': 'sheet', Financial: 'calc', Information: 'link' })[k], label: k, big: 1, menu: () => FCATS[k].filter(f => App.FNAMES.includes(f)).map(f => ({ label: f, key: (FDOC[f] || '').split(')')[0].split('(')[1] ? '(' + (FDOC[f] || '').split(')')[0].split('(')[1] + ')' : '', fn: () => CMD.insertFn(f) })) }))),
        group('Defined Names', btn({ icon: 'text', label: 'Name Manager', big: 1, cmd: 'nameManager' }), btn({ icon: 'plus', label: 'Define Name', big: 1, cmd: 'defineName' }), btn({ icon: 'fx', label: 'Use in Formula', big: 1, cmd: 'useName' })),
        group('Tips', el('div', null, '<div style="max-width:300px;color:var(--muted);padding:4px 8px;line-height:1.5">Type <b style="color:var(--text)">=</b> in a cell to start a formula. Click cells to add references, press <b style="color:var(--text)">F4</b> to toggle <code>$</code> anchors, and <b style="color:var(--text)">Tab</b> to accept a function suggestion.</div>'))
      ]
    },
    {
      id: 'data', label: 'Data', build: () => [
        group('Sort & Filter', btn({ icon: 'sortasc', label: 'Sort A→Z', big: 1, cmd: 'sortAsc' }), btn({ icon: 'sortdesc', label: 'Sort Z→A', big: 1, cmd: 'sortDesc' }), btn({ icon: 'sortasc', label: 'Sort…', big: 1, cmd: 'sortDialog' }), btn({ icon: 'filter', label: 'Filter', big: 1, cmd: 'toggleFilter', state: 'filter' })),
        group('Data Tools', btn({ icon: 'check', label: 'Data Validation', big: 1, cmd: 'dataValidation' }), btn({ icon: 'table', label: 'PivotTable', big: 1, cmd: 'pivot' }), btn({ icon: 'dup', label: 'Remove Duplicates', big: 1, cmd: 'removeDuplicates' }), btn({ icon: 'fill', label: 'Fill Down', big: 1, cmd: 'fillDown' }), btn({ icon: 'fill', label: 'Fill Series', big: 1, cmd: 'fillSeries' })),
        group('Import / Export', btn({ icon: 'open', label: 'Open file', big: 1, cmd: 'open' }), btn({ icon: 'save', label: 'Export CSV', big: 1, cmd: 'saveCsv' }))
      ]
    },
    {
      id: 'review', label: 'Review', build: () => [
        group('Proofing', btn({ icon: 'check', label: 'Find…', big: 1, cmd: 'find' }), btn({ icon: 'search', label: 'Replace…', big: 1, cmd: 'replace' })),
        group('Comments', btn({ icon: 'text', label: 'New Note', big: 1, cmd: 'note' }), btn({ icon: 'trash', label: 'Delete Note', big: 1, fn: () => CMD.deleteNote() })),
        group('Changes', btn({ icon: 'undo', label: 'Undo', big: 1, cmd: 'undo' }), btn({ icon: 'redo', label: 'Redo', big: 1, cmd: 'redo' }))
      ]
    },
    {
      id: 'view', label: 'View', build: () => [
        group('Show', col(btn({ label: 'Gridlines', sm: 1, icon: 'grid', cmd: 'gridlines', state: 'grid' }), btn({ label: 'Headings', sm: 1, icon: 'col', cmd: 'headings', state: 'head' }), btn({ label: 'Formula bar', sm: 1, icon: 'fx', cmd: 'fbarToggle', state: 'fbar' }))),
        group('Zoom', btn({ icon: 'zoom', label: '100%', big: 1, fn: () => App.setZoom(100) }), btn({ icon: 'zoom', label: 'Zoom in', big: 1, fn: () => App.setZoom(S.zoom * 100 + 10) }), btn({ icon: 'zoom', label: 'Zoom out', big: 1, fn: () => App.setZoom(S.zoom * 100 - 10) })),
        group('Window', btn({ icon: 'freeze', label: 'Freeze Panes', big: 1, menu: () => [{ label: 'Freeze panes (at selection)', cmd: 'freeze', arg: 'at' }, { label: 'Freeze top row', cmd: 'freeze', arg: 'row' }, { label: 'Freeze first column', cmd: 'freeze', arg: 'col' }, { label: 'Unfreeze panes', cmd: 'freeze', arg: 'none' }] })),
        group('Appearance', btn({ icon: 'moon', label: 'Light / Dark', big: 1, cmd: 'theme' }))
      ]
    }
  ];
  const tabsEl = $('#tabs'), ribbon = $('#ribbon'); let activeTab = 'home'; const panels = {};
  const fileBtn = el('button', 'file', 'FILE'); tabsEl.appendChild(fileBtn);
  fileBtn.onclick = () => App.backstage.open('home');
  CMD.about = () => dialog('FmOffice Sheets', `<div style="display:flex;gap:14px;align-items:center"><img src="img/sheets.png" width="64" height="64" alt=""><div><b>FmOffice Sheets</b> <span style="color:var(--muted)">v${window.FM_VERSION || '0.1.0'}</span><div style="color:var(--muted);margin-top:4px">A free, fan-made spreadsheet app inspired by Microsoft Excel. Not affiliated with or endorsed by Microsoft.</div></div></div>`);
  tabs.forEach(t => { const b = el('button', t.id === activeTab ? 'on' : '', t.label); b.dataset.tab = t.id; b.onclick = () => showTab(t.id); tabsEl.appendChild(b); });
  function showTab(id) { activeTab = id; tabsEl.querySelectorAll('button[data-tab]').forEach(b => b.classList.toggle('on', b.dataset.tab === id)); Object.entries(panels).forEach(([k, p]) => p.classList.toggle('on', k === id)); refreshRibbonState(); }
  tabs.forEach(t => { const p = el('div', 'rpanel' + (t.id === activeTab ? ' on' : '')); t.build().forEach(g => p.appendChild(g)); ribbon.appendChild(p); panels[t.id] = p; });
  // command palette registry
  const extra = [['Save', 'save'], ['Save as…', 'saveAs'], ['Open file', 'open'], ['New workbook', 'new'], ['Print', 'print'], ['Undo', 'undo'], ['Redo', 'redo'], ['Dark mode', 'theme'], ['Find', 'find'], ['Replace', 'replace'], ['Freeze top row', 'freeze', 'row'], ['Unfreeze panes', 'freeze', 'none'], ['Sort A to Z', 'sortAsc'], ['Sort Z to A', 'sortDesc'], ['Filter', 'toggleFilter'], ['Remove duplicates', 'removeDuplicates'], ['Insert chart', null], ['Add sheet', 'addSheet'], ['Export CSV', 'saveCsv'], ['Format cells', 'formatCells'], ['Insert function', 'insertFunctionDialog']];
  extra.forEach(([l, c, a]) => { if (c) reg.push({ label: l, run: () => CMD[c](a) }); });
  [['column chart', 'col'], ['bar chart', 'bar'], ['line chart', 'line'], ['pie chart', 'pie'], ['area chart', 'area'], ['scatter chart', 'scatter']].forEach(([l, t]) => reg.push({ label: 'Insert ' + l, run: () => App.insertChart(t) }));
  const tell = $('#tellme'), tlist = $('#tellme-list');
  tell.oninput = () => { const q = tell.value.trim().toLowerCase(); if (!q) { tlist.classList.add('hidden'); return; } const seen = new Set(); const hits = reg.filter(r => { const k = r.label.toLowerCase(); if (seen.has(k)) return false; seen.add(k); return k.includes(q) || (r.tip || '').toLowerCase().includes(q); }).slice(0, 8); tlist.innerHTML = hits.map((h, i) => `<button class="mi${i === 0 ? ' hl' : ''}" data-i="${i}">${h.icon ? ic(h.icon) : ''}<span>${esc(h.label)}</span></button>`).join('') || '<div class="mh">No matching commands</div>'; tlist.classList.remove('hidden'); tlist.querySelectorAll('.mi').forEach((b, i) => b.onclick = () => { hits[i].run(); tell.value = ''; tlist.classList.add('hidden'); App.focusGrid(); }); tlist._hits = hits; };
  tell.onkeydown = e => { if (e.key === 'Enter') { const h = tlist._hits && tlist._hits[0]; if (h) { h.run(); tell.value = ''; tlist.classList.add('hidden'); App.focusGrid(); } } else if (e.key === 'Escape') { tell.value = ''; tlist.classList.add('hidden'); App.focusGrid(); } };
  window.addEventListener('keydown', e => { if (e.altKey && e.key.toLowerCase() === 'q') { e.preventDefault(); tell.focus(); } });

  function refreshRibbonState() {
    const st = getSt(); const act = S.wb.getValue(S.si, S.act.r, S.act.c);
    document.querySelectorAll('.btn[data-state]').forEach(b => { const k = b.dataset.state; let on = false; if (k.startsWith('ha:')) on = st.ha === k.slice(3); else if (k === 'filter') on = !!cur().filter; else if (k === 'grid') on = cur().showGrid; else if (k === 'head') on = S.showHeadings; else if (k === 'fbar') on = S.showFbar; else on = !!st[k]; b.classList.toggle('on', on); });
    const f = panels.home.querySelectorAll('select.sel'); if (f.length >= 3) { f[0].value = st.ff || 'Calibri'; f[1].value = String(st.fs || 11); const nf = st.nf || 'General'; const o = NUMFMTS.find(x => x[1] === nf); f[2].value = o ? o[1] : 'General'; }
    $('.qbtn[data-cmd=undo]').classList.toggle('dis', !S.undo.length); $('.qbtn[data-cmd=redo]').classList.toggle('dis', !S.redo.length);
  }
  App.hooks.sel = refreshRibbonState; App.refreshRibbon = refreshRibbonState;
  window.addEventListener('keydown', e => { if (e.key === 'F11' || (e.ctrlKey && e.key === 'F1')) { } });
  const _after = App.afterChange; App.afterChange = function (s) { _after(s); refreshRibbonState(); const f = cur().filter; if (f && !s) { /* keep filter rows applied when data changes */ } };
  CMD.repeat = () => { };

  /* ---------- sample data & boot ---------- */
  function sample() {
    const wb = S.wb; const rows = [['Month', 'Product A', 'Product B', 'Product C', 'Total'], ['January', 1200, 950, 700], ['February', 1350, 1010, 820], ['March', 1280, 1100, 910], ['April', 1500, 1180, 880], ['May', 1620, 1250, 1020], ['June', 1710, 1340, 1105]];
    rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))); for (let r = 1; r <= 6; r++) wb.setInput(0, r, 4, `=SUM(B${r + 1}:D${r + 1})`);
    wb.setInput(0, 7, 0, 'Total'); for (let c = 1; c <= 4; c++) wb.setInput(0, 7, c, `=SUM(${E.numToCol(c)}2:${E.numToCol(c)}7)`);
    for (let c = 0; c <= 4; c++) { wb.setStyle(0, 0, c, { bg: '#1f8a4c', fc: '#ffffff', b: 1, ha: c ? 'right' : 'left' }); wb.setStyle(0, 7, c, { b: 1, bd: { t: { c: '#1d2430', w: 1 }, b: { c: '#1d2430', w: 2 } } }); }
    for (let r = 1; r <= 7; r++) for (let c = 1; c <= 4; c++) wb.setStyle(0, r, c, { nf: '$#,##0' });
    for (let r = 2; r <= 6; r += 2) for (let c = 0; c <= 4; c++) wb.setStyle(0, r, c, { bg: '#e3f4ea' });
    wb.setInput(0, 9, 0, 'Tip: select A1:D7 and use Insert ▸ Column to chart this data. Press Alt+Q to search commands.'); wb.setStyle(0, 9, 0, { i: 1, fc: '#6a7685' });
    wb.sheets[0].colW[0] = 110; wb.invalidate();
  }
  function boot() {
    let restored = false;
    try { const a = JSON.parse(localStorage.getItem('fmsheets.autosave') || 'null'); if (a && a.wb && Date.now() - a.t < 7 * 864e5 && a.wb.sheets && a.wb.sheets[0].cells.length) { S.wb = E.Workbook.fromJSON(a.wb); S.name = a.name || 'Book1'; $('#docname').value = S.name; restored = true; $('#savestate').textContent = 'Restored'; } } catch (e) { }
    if (!restored) sample();
    renderSheetTabs(); App.renderCharts(); App.updateUI(); refreshRibbonState(); App.draw(); App.focusGrid();
  }
  window.addEventListener('keydown', e => { if (e.key === 'F6') e.preventDefault(); });
  window.FmApp = App; boot();
})();
