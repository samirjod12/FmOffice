/* FmOffice Sheets — extra features: data validation, named ranges, notes, hyperlinks, pivot tables */
(function () {
  'use strict';
  const E = FmEngine, App = window.App, S = App.S, CMD = App.CMD, $ = s => document.querySelector(s);
  const cur = App.cur, esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const sel = () => App.selNorm();
  const absRef = n => { const sh = cur(); const nm = /^[A-Za-z_][\w.]*$/.test(sh.name) ? sh.name : "'" + sh.name.replace(/'/g, "''") + "'"; const a = '$' + E.numToCol(n.c1) + '$' + (n.r1 + 1), b = '$' + E.numToCol(n.c2) + '$' + (n.r2 + 1); return nm + '!' + (a === b ? a : a + ':' + b); };
  const full = n => n.r1 === 0 && n.r2 >= E.MAXR - 1 ? { ...n, r2: Math.min(Math.max(S.wb.usedRange(S.si).r2, 0), 100000) } : n;

  /* ---------- data validation ---------- */
  CMD.dataValidation = () => {
    const n = sel(); const ex = App.dvAt(cur(), n.r1, n.c1) || { type: 'any', op: 'between', allowBlank: true, showDrop: true, style: 'stop' };
    const opt = (arr, v) => arr.map(([k, l]) => `<option value="${k}" ${k === v ? 'selected' : ''}>${l}</option>`).join('');
    const d = App.dialog('Data Validation', `
      <label>Allow<select id="dt">${opt([['any', 'Any value'], ['list', 'List'], ['whole', 'Whole number'], ['decimal', 'Decimal'], ['date', 'Date'], ['textLength', 'Text length']], ex.type)}</select></label>
      <div id="dnum"><label>Data<select id="dop">${opt([['between', 'between'], ['notBetween', 'not between'], ['eq', 'equal to'], ['ne', 'not equal to'], ['gt', 'greater than'], ['lt', 'less than'], ['ge', 'greater than or equal to'], ['le', 'less than or equal to']], ex.op)}</select></label>
        <div class="row2"><label>Minimum / value<input type="text" id="v1" value="${esc(ex.v1 ?? '')}"></label><label id="v2w">Maximum<input type="text" id="v2" value="${esc(ex.v2 ?? '')}"></label></div></div>
      <div id="dlist"><label>Source (comma-separated, or a range like =$A$1:$A$5)<input type="text" id="src" value="${esc(ex.src ?? '')}" placeholder="Yes,No,Maybe"></label><label class="ck"><input type="checkbox" id="drop" ${ex.showDrop !== false ? 'checked' : ''}> In-cell dropdown</label></div>
      <label class="ck"><input type="checkbox" id="blank" ${ex.allowBlank !== false ? 'checked' : ''}> Ignore blank</label>
      <div class="row2"><label>Error alert<select id="est">${opt([['stop', 'Stop'], ['warning', 'Warning'], ['info', 'Information']], ex.style || 'stop')}</select></label><label>Title<input type="text" id="et" value="${esc(ex.errTitle ?? '')}"></label></div>
      <label>Error message<input type="text" id="em" value="${esc(ex.errMsg ?? '')}"></label>`,
      [{ label: 'Cancel' }, { label: 'Clear all', fn: () => { App.snap(); const sh = cur(); sh.dv = sh.dv.filter(x => x.r1 > n.r2 || x.r2 < n.r1 || x.c1 > n.c2 || x.c2 < n.c1); App.afterChange(); } }, {
        label: 'OK', primary: 1, fn: d => {
          const g = id => d.querySelector('#' + id); const type = g('dt').value; App.snap(); const sh = cur(); sh.dv = sh.dv.filter(x => x.r1 > n.r2 || x.r2 < n.r1 || x.c1 > n.c2 || x.c2 < n.c1);
          if (type !== 'any') { const rule = { r1: n.r1, c1: n.c1, r2: Math.min(n.r2, 200000), c2: Math.min(n.c2, 2000), type, op: g('dop').value, v1: g('v1').value, v2: g('v2').value, src: g('src').value, showDrop: g('drop').checked, allowBlank: g('blank').checked, style: g('est').value, errTitle: g('et').value, errMsg: g('em').value }; if (type === 'list' && !rule.src.trim()) { App.toast('Enter the list items'); App.snap(); S.undo.pop(); return false; } sh.dv.push(rule); }
          App.afterChange();
        }
      }], { width: 480 });
    const sync = () => { const t = d.querySelector('#dt').value; d.querySelector('#dnum').style.display = ['whole', 'decimal', 'date', 'textLength'].includes(t) ? '' : 'none'; d.querySelector('#dlist').style.display = t === 'list' ? '' : 'none'; const op = d.querySelector('#dop').value; d.querySelector('#v2w').style.display = op === 'between' || op === 'notBetween' ? '' : 'none'; };
    d.querySelector('#dt').onchange = sync; d.querySelector('#dop').onchange = sync; sync();
  };

  /* ---------- names ---------- */
  const validName = n => /^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) && !/^[rc]$/i.test(n) && !E.FN[n.toUpperCase()] && !E.LAZY[n.toUpperCase()];
  CMD.nameManager = () => {
    const render = d => { const names = S.wb.names; const keys = Object.keys(names).sort(); d.querySelector('#nl').innerHTML = keys.length ? keys.map(k => `<div class="fi" style="display:flex;gap:10px;align-items:center"><b style="min-width:110px">${esc(k)}</b><span class="grow" style="flex:1;color:var(--muted)">${esc(names[k])}</span><button class="gbtn" data-del="${esc(k)}" style="padding:4px 10px">Delete</button></div>`).join('') : '<div class="fi" style="color:var(--muted)">No names defined yet. Select cells, type a name below and press Add.</div>'; d.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { App.snap(); delete S.wb.names[b.dataset.del]; S.wb.invalidate(); S.wb.astCache.clear(); App.afterChange(); render(d); }); };
    const d = App.dialog('Name Manager', `<div class="flist" id="nl"></div><div class="row2"><label>Name<input type="text" id="nn" placeholder="e.g. TaxRate"></label><label>Refers to<input type="text" id="nr" value="${esc(absRef(sel()))}"></label></div>`, [{ label: 'Close' }, { label: 'Add', primary: 1, fn: d => { const n = d.querySelector('#nn').value.trim(), r = d.querySelector('#nr').value.trim().replace(/^=/, ''); if (!validName(n)) { App.toast('Names must start with a letter, can’t look like a cell (A1) or a function name'); return false; } try { E.parse(r); } catch (e) { App.toast('“Refers to” isn’t a valid reference or value'); return false; } App.snap(); S.wb.names[n.toUpperCase() === n ? n : n] = r; S.wb.names[n.toUpperCase()] = r; if (n !== n.toUpperCase()) delete S.wb.names[n]; S.wb.invalidate(); S.wb.astCache.clear(); App.afterChange(); render(d); d.querySelector('#nn').value = ''; return false; } }], { width: 560, noEnter: 0 });
    render(d);
  };
  CMD.defineName = () => App.prompt('Define name for selection', '', v => { v = v.trim(); if (!validName(v)) { App.toast('That isn’t a valid name'); return; } App.snap(); S.wb.names[v.toUpperCase()] = absRef(sel()); S.wb.invalidate(); S.wb.astCache.clear(); App.afterChange(); App.toast(v + ' → ' + absRef(sel())); }, 'Name');
  CMD.useName = () => { const names = Object.keys(S.wb.names).sort(); App.showMenu(names.length ? names.map(n => ({ label: n, fn: () => { if (S.edit) { const ed = $('#editor'); ed.setRangeText(n, ed.selectionStart, ed.selectionEnd, 'end'); ed.dispatchEvent(new Event('input')); } else { App.startEdit(S.act.r, S.act.c, 'edit', '=' + n); } } })) : [{ label: 'No names defined', dis: true }], innerWidth / 2 - 90, 140); };

  /* ---------- notes & links ---------- */
  CMD.note = () => { const a = S.act, key = a.r * 20000 + a.c, sh = cur(); const d = App.dialog('Note — ' + E.addrToStr(a.r, a.c), `<label>Note text (leave empty to delete)<textarea id="nt" rows="5" style="min-width:320px">${esc(sh.notes[key] || '')}</textarea></label>`, [{ label: 'Cancel' }, { label: 'OK', primary: 1, fn: d => { const v = d.querySelector('#nt').value; App.snap(); if (v.trim()) sh.notes[key] = v; else delete sh.notes[key]; App.afterChange(); } }], { noEnter: 1 }); };
  CMD.link = () => {
    const a = S.act, key = a.r * 20000 + a.c, sh = cur(); const old = App.linkAt(a.r, a.c) || ''; const shown = S.wb.display(S.si, a.r, a.c);
    App.dialog('Link', `<label>Text to display<input type="text" id="lt" value="${esc(shown)}"></label><label>Address (web address or email)<input type="text" id="lu" value="${esc(old)}" placeholder="https://…"></label>`, [{ label: 'Cancel' }, { label: 'Remove link', fn: () => { App.snap(); delete sh.links[key]; App.afterChange(); } }, { label: 'OK', primary: 1, fn: d => { let u = d.querySelector('#lu').value.trim(), t = d.querySelector('#lt').value; if (!u) return; if (!/^[a-z]+:/i.test(u)) u = /@/.test(u) && !/\//.test(u) ? 'mailto:' + u : 'https://' + u; App.snap(); sh.links[key] = u; if (S.wb.getInput(S.si, a.r, a.c) === '' || t !== shown) S.wb.setInput(S.si, a.r, a.c, t || u); App.afterChange(); } }]);
  };

  /* ---------- pivot table ---------- */
  CMD.pivot = () => {
    let n = sel(); if (n.r1 === n.r2 && n.c1 === n.c2) { n = App.currentRegion ? App.currentRegion() : n; } n = full(n);
    if (n.r2 - n.r1 < 1) { App.toast('Select your data (with a header row) first'); return; }
    const heads = []; for (let c = n.c1; c <= n.c2; c++) { const h = S.wb.display(S.si, n.r1, c); heads.push(h || 'Column ' + E.numToCol(c)); }
    const opts = (none) => (none ? '<option value="-1">(none)</option>' : '') + heads.map((h, i) => `<option value="${i}">${esc(h)}</option>`).join('');
    const d = App.dialog('PivotTable', `<div class="prev">Source: ${E.addrToStr(n.r1, n.c1)}:${E.addrToStr(n.r2, n.c2)} — summary goes to a new sheet</div><label>Rows<select id="pr">${opts()}</select></label><label>Columns<select id="pc">${opts(1)}</select></label><label>Values<select id="pv">${opts()}</select></label><label>Summarize by<select id="pa"><option value="SUM">Sum</option><option value="COUNT">Count</option><option value="AVERAGE">Average</option><option value="MAX">Max</option><option value="MIN">Min</option></select></label>`, [{ label: 'Cancel' }, { label: 'Create', primary: 1, fn: d => { buildPivot(n, +d.querySelector('#pr').value, +d.querySelector('#pc').value, +d.querySelector('#pv').value, d.querySelector('#pa').value, heads); } }], { width: 440 });
    d.querySelector('#pv').value = String(Math.min(heads.length - 1, 1));
  };
  function buildPivot(n, rf, cf, vf, agg, heads) {
    App.snap(); const wb = S.wb; const rows = new Map(), cols = new Map(); const cells = new Map();
    for (let r = n.r1 + 1; r <= n.r2; r++) {
      const rk = wb.display(S.si, r, n.c1 + rf), ck = cf >= 0 ? wb.display(S.si, r, n.c1 + cf) : 'Total'; const v = wb.getValue(S.si, r, n.c1 + vf);
      rows.set(rk, 1); cols.set(ck, 1); const key = rk + '\u0001' + ck; if (!cells.has(key)) cells.set(key, []); cells.get(key).push(v);
    }
    const cmp = (a, b) => E.cmpVals(E.parseLiteral(a).v, E.parseLiteral(b).v); const rk = [...rows.keys()].sort(cmp), ck = [...cols.keys()].sort(cmp);
    const calc = vals => { const nums = vals.filter(x => typeof x === 'number'); switch (agg) { case 'COUNT': return vals.filter(x => x !== null && x !== '').length; case 'SUM': return nums.reduce((s, x) => s + x, 0); case 'AVERAGE': return nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : null; case 'MAX': return nums.length ? Math.max(...nums) : null; case 'MIN': return nums.length ? Math.min(...nums) : null; } };
    const sh = wb.addSheet(wb.uniqueName('Pivot1')); const si = wb.sheets.indexOf(sh); const put = (r, c, v, st) => { if (v !== null && v !== undefined) wb.setInput(si, r, c, typeof v === 'number' ? String(Math.round(v * 1e9) / 1e9) : "'" + v); if (st) wb.setStyle(si, r, c, st); };
    const H = { bg: '#1f8a4c', fc: '#ffffff', b: 1 }; const T = { b: 1, bg: '#e3f4ea', bd: { t: { c: '#1d2430', w: 1 } } };
    const label = { SUM: 'Sum', COUNT: 'Count', AVERAGE: 'Average', MAX: 'Max', MIN: 'Min' }[agg] + ' of ' + heads[vf];
    put(0, 0, label, H); if (cf >= 0) put(0, 1, heads[cf], H); put(1, 0, heads[rf], H);
    ck.forEach((c, j) => put(1, 1 + j, c, { ...H, ha: 'right' })); if (cf >= 0) put(1, 1 + ck.length, 'Grand Total', { ...H, ha: 'right' });
    const colAll = ck.map(() => []), grand = [];
    rk.forEach((r, i) => { put(2 + i, 0, r); const rowAll = []; ck.forEach((c, j) => { const vals = cells.get(r + '\u0001' + c) || []; rowAll.push(...vals); colAll[j].push(...vals); const v = vals.length ? calc(vals) : null; put(2 + i, 1 + j, v, { nf: '#,##0.##' }); }); grand.push(...rowAll); if (cf >= 0) put(2 + i, 1 + ck.length, calc(rowAll), { b: 1, nf: '#,##0.##' }); });
    const tr = 2 + rk.length; put(tr, 0, 'Grand Total', T); ck.forEach((c, j) => put(tr, 1 + j, calc(colAll[j]), { ...T, nf: '#,##0.##' })); if (cf >= 0) put(tr, 1 + ck.length, calc(grand), { ...T, nf: '#,##0.##' });
    sh.colW[0] = 150; for (let j = 1; j <= ck.length + 1; j++) sh.colW[j] = 100;
    S.si = si; App.geomChanged(); App.switchSheet(si); S.wb.invalidate(); App.afterChange(true); App.toast('Pivot summary created (static snapshot — re-run to refresh)');
  }
  CMD.exportPdf = () => CMD.print();
})();
