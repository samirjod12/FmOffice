/* FmOffice Sheets — Backstage (File ▸ Home / New / Open / Save …) */
(function () {
  'use strict';
  const App = window.App, CMD = App.CMD, S = App.S, E = FmEngine, $ = s => document.querySelector(s);
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const bs = $('#backstage');
  const ic = (n, c) => App.ic(n, c);

  /* ---------- recent list ---------- */
  const KEY = 'fmsheets.recent';
  const recent = () => { try { return JSON.parse(localStorage.getItem(KEY) || '[]'); } catch (e) { return []; } };
  function addRecent(name) { if (!name) return; try { const l = recent().filter(x => x.name !== name); l.unshift({ name, t: Date.now() }); localStorage.setItem(KEY, JSON.stringify(l.slice(0, 12))); } catch (e) { } }
  App.addRecent = addRecent;
  ['save', 'saveAs'].forEach(k => { const f = CMD[k]; CMD[k] = async function () { const r = await f.apply(this, arguments); if (!S.dirty) addRecent(S.name); return r; }; });
  const _toast = App.toast; App.toast = function (m) { if (/^Opened /.test(m)) addRecent(m.slice(7).replace(/\.[^.]+$/, '')); return _toast.apply(this, arguments); };

  /* ---------- templates (original FmOffice content) ---------- */
  const hdr = (wb, r, cols) => cols.forEach((v, c) => { wb.setInput(0, r, c, v); wb.setStyle(0, r, c, { bg: '#1f8a4c', fc: '#ffffff', b: 1 }); });
  const TEMPLATES = [
    { id: 'blank', name: 'Blank workbook', color: '#fff', glyph: '+', make: wb => { } },
    {
      id: 'tour', name: 'Welcome to FmOffice Sheets', color: '#e3f4ea', glyph: '👋', make: wb => {
        const t = [['Welcome to FmOffice Sheets', ''], ['Click a cell and type to enter data.', ''], ['Start a formula with =  e.g.', '=SUM(B5:B8)'], ['', ''], ['Apples', 12], ['Pears', 7], ['Plums', 21], ['Cherries', 9], ['Tip', 'Press Alt+Q to search every command.']];
        t.forEach((r, i) => r.forEach((v, c) => v !== '' && wb.setInput(0, i, c, String(v)))); wb.setStyle(0, 0, 0, { b: 1, fs: 18, fc: '#1f8a4c' }); wb.setInput(0, 2, 1, "=SUM(B5:B8)"); wb.sheets[0].colW[0] = 260; wb.sheets[0].colW[1] = 220;
      }
    },
    {
      id: 'budget', name: 'Monthly budget', color: '#fff2cc', glyph: '$', make: wb => {
        hdr(wb, 0, ['Category', 'Planned', 'Actual', 'Difference']);
        [['Rent', 1200, 1200], ['Groceries', 400, 432], ['Transport', 120, 98], ['Utilities', 180, 205], ['Fun', 150, 170], ['Savings', 300, 300]].forEach((r, i) => { r.forEach((v, c) => wb.setInput(0, i + 1, c, String(v))); wb.setInput(0, i + 1, 3, `=C${i + 2}-B${i + 2}`); });
        wb.setInput(0, 7, 0, 'Total'); ['B', 'C', 'D'].forEach((L, c) => wb.setInput(0, 7, c + 1, `=SUM(${L}2:${L}7)`)); for (let c = 0; c < 4; c++) wb.setStyle(0, 7, c, { b: 1, bd: { t: { c: '#000', w: 1 } } });
        for (let r = 1; r <= 7; r++) for (let c = 1; c <= 3; c++) wb.setStyle(0, r, c, { nf: '$#,##0' }); wb.sheets[0].colW[0] = 140;
      }
    },
    {
      id: 'formulas', name: 'Formula basics', color: '#d9e8f7', glyph: 'fx', make: wb => {
        hdr(wb, 0, ['Try this', 'Formula', 'Result']);
        [['Add numbers', '=SUM(10,20,30)'], ['Average', '=AVERAGE(4,8,15,16,23,42)'], ['If / else', '=IF(5>3,"Yes","No")'], ['Join text', '=CONCAT("Fm","Office")'], ['Look up', '=XLOOKUP(2,{1,2,3},{"a","b","c"})'], ['Today', '=TODAY()'], ['Sequence', '=SEQUENCE(1,5)']].forEach((r, i) => { wb.setInput(0, i + 1, 0, r[0]); wb.setInput(0, i + 1, 1, "'" + r[1]); wb.setInput(0, i + 1, 2, r[1]); });
        wb.sheets[0].colW[0] = 120; wb.sheets[0].colW[1] = 260; wb.sheets[0].colW[2] = 120;
      }
    },
    {
      id: 'chart', name: 'Chart sampler', color: '#fce5cd', glyph: '▮▮', make: wb => {
        hdr(wb, 0, ['Quarter', 'North', 'South', 'West']);
        [['Q1', 120, 90, 70], ['Q2', 135, 101, 82], ['Q3', 128, 110, 91], ['Q4', 150, 118, 88]].forEach((r, i) => r.forEach((v, c) => wb.setInput(0, i + 1, c, String(v))));
        wb.setInput(0, 6, 0, 'Select A1:D5, then Insert ▸ Column to make a chart.'); wb.setStyle(0, 6, 0, { i: 1, fc: '#6a7685' });
      }
    },
    {
      id: 'invoice', name: 'Simple invoice', color: '#ead1dc', glyph: '🧾', make: wb => {
        wb.setInput(0, 0, 0, 'INVOICE'); wb.setStyle(0, 0, 0, { b: 1, fs: 20, fc: '#1f8a4c' });
        hdr(wb, 3, ['Item', 'Qty', 'Price', 'Amount']);
        [['Design work', 10, 45], ['Hosting', 12, 9], ['Support', 4, 60]].forEach((r, i) => { r.forEach((v, c) => wb.setInput(0, i + 4, c, String(v))); wb.setInput(0, i + 4, 3, `=B${i + 5}*C${i + 5}`); });
        wb.setInput(0, 8, 2, 'Total'); wb.setInput(0, 8, 3, '=SUM(D5:D7)'); wb.setStyle(0, 8, 2, { b: 1 }); wb.setStyle(0, 8, 3, { b: 1 });
        for (let r = 4; r <= 8; r++) for (let c = 2; c <= 3; c++) wb.setStyle(0, r, c, { nf: '$#,##0.00' }); wb.sheets[0].colW[0] = 180;
      }
    }
  ];
  function fromTemplate(t) {
    const go = () => { const wb = new E.Workbook(); t.make(wb); wb.invalidate && wb.invalidate(); App.loadWb(wb, t.id === 'blank' ? 'Book1' : t.name); close(); };
    if (S.dirty) App.dialog('Start a new workbook?', '<div>You have unsaved changes. Start over anyway?</div>', [{ label: 'Cancel' }, { label: 'Discard & new', primary: 1, fn: go }]); else go();
  }

  /* ---------- views ---------- */
  const greet = () => { const h = new Date().getHours(); return h < 12 ? 'Good morning' : h < 18 ? 'Good afternoon' : 'Good evening'; };
  const ago = t => { const m = Math.round((Date.now() - t) / 60000); return m < 1 ? 'Just now' : m < 60 ? m + ' min ago' : m < 1440 ? Math.round(m / 60) + ' h ago' : new Date(t).toLocaleDateString(); };
  function recentList(tab) {
    const l = recent(); if (tab === 'pinned') return '<div class="bs-empty">Pin workbooks you want to find easily. Pinning is coming soon.</div>';
    if (!l.length) return '<div class="bs-empty">You haven’t opened any workbooks recently. To browse for a workbook, start by clicking on <a data-go="open">Open</a>.</div>';
    return '<div class="bs-rlist">' + l.map(x => `<button class="bs-r" data-open="1">${ic('sheet')}<span><b>${esc(x.name)}</b><small>${ago(x.t)}</small></span></button>`).join('') + '</div>';
  }
  const VIEWS = {
    home(tab) {
      tab = tab || 'recent';
      return `<div class="bs-greet">${greet()}</div>
      <div class="bs-new"><div class="bs-h">New</div><div class="bs-tiles">${TEMPLATES.map(t => `<button class="bs-tile" data-tpl="${t.id}"><span class="bs-thumb" style="background:${t.color}">${esc(t.glyph)}</span><span class="bs-tn">${esc(t.name)}</span></button>`).join('')}</div></div>
      <div class="bs-tabs"><button class="${tab === 'recent' ? 'on' : ''}" data-rt="recent">Recent</button><button class="${tab === 'pinned' ? 'on' : ''}" data-rt="pinned">Pinned</button><a class="bs-more" data-go="open">Open Other Workbooks →</a></div>
      ${recentList(tab)}`;
    },
    new() { return `<div class="bs-h big">New</div><div class="bs-tiles wrap">${TEMPLATES.map(t => `<button class="bs-tile" data-tpl="${t.id}"><span class="bs-thumb" style="background:${t.color}">${esc(t.glyph)}</span><span class="bs-tn">${esc(t.name)}</span></button>`).join('')}</div>`; },
    open() { return `<div class="bs-h big">Open</div><div class="bs-cards"><button class="bs-card" data-act="open">${ic('open', 'lg')}<span><b>This device</b><small>Browse for an .xlsx, .xlsm, .csv or .tsv file</small></span></button></div><div class="bs-h">Recent</div>${recentList('recent')}`; },
    save() { return `<div class="bs-h big">Save</div><div class="bs-cards"><button class="bs-card" data-act="save">${ic('save', 'lg')}<span><b>Save</b><small>Save changes to “${esc(S.name)}”</small></span></button><button class="bs-card" data-act="saveAs">${ic('save', 'lg')}<span><b>Save As .xlsx…</b><small>Excel workbook</small></span></button><button class="bs-card" data-act="saveCsv">${ic('file', 'lg')}<span><b>Export as CSV</b><small>Current sheet only</small></span></button></div>`; },
    print() { return `<div class="bs-h big">Print</div><div class="bs-cards"><button class="bs-card" data-act="print">${ic('print', 'lg')}<span><b>Print / Save as PDF</b><small>${esc(S.page.size)} · ${esc(S.page.orient)} · ${esc(S.page.margins)} margins — change under Page Layout</small></span></button></div>`; },
    account() { return `<div class="bs-h big">Account</div><div class="bs-about"><img src="img/sheets.png" width="72" height="72" alt=""><div><b>FmOffice Sheets</b> <span class="mut">v${window.FM_VERSION || '0.1.0'}</span><p>A free, fan-made spreadsheet app inspired by Microsoft Excel. Not affiliated with or endorsed by Microsoft.</p><p class="mut">Your files stay on your device. There is no sign-in and no cloud upload.</p></div></div>`; },
    feedback() { return `<div class="bs-h big">Feedback</div><div class="bs-about"><div><p>Found a bug or have an idea? Open an issue on the project’s GitHub page.</p></div></div>`; },
    options() { const dark = document.documentElement.dataset.theme === 'dark'; return `<div class="bs-h big">Options</div><div class="bs-cards"><button class="bs-card" data-act="theme">${ic(dark ? 'sun' : 'moon', 'lg')}<span><b>Theme: ${dark ? 'Dark' : 'Light'}</b><small>Click to switch to ${dark ? 'light' : 'dark'}</small></span></button></div>`; }
  };
  const NAV = [['home', 'Home', 'sheet'], ['new', 'New', 'file'], ['open', 'Open', 'open'], '-', ['save', 'Save', 'save'], ['saveAs', 'Save As', 'save'], ['print', 'Print', 'print'], '-', ['account', 'Account', 'text'], ['feedback', 'Feedback', 'text'], ['options', 'Options', 'grid']];
  let view = 'home', tab = 'recent';
  function render() {
    bs.innerHTML = `<div class="bs-side"><button class="bs-back" title="Back (Esc)">←</button>${NAV.map(n => n === '-' ? '<div class="bs-sep"></div>' : `<button class="bs-nav${n[0] === view ? ' on' : ''}" data-v="${n[0]}">${ic(n[2])}<span>${n[1]}</span></button>`).join('')}<div class="bs-sep"></div><button class="bs-nav" data-v="close">${ic('plus')}<span>Close</span></button></div>
    <div class="bs-main"><div class="bs-top"><span class="signin">Sign in</span></div><div class="bs-body">${(VIEWS[view] || VIEWS.home)(tab)}</div></div>`;
    bs.querySelector('.bs-back').onclick = close;
    bs.querySelectorAll('[data-v]').forEach(b => b.onclick = () => { const v = b.dataset.v; if (v === 'close') { close(); CMD.new(); } else if (v === 'saveAs') { close(); CMD.saveAs(); } else { view = v; render(); } });
    bs.querySelectorAll('[data-tpl]').forEach(b => b.onclick = () => fromTemplate(TEMPLATES.find(t => t.id === b.dataset.tpl)));
    bs.querySelectorAll('[data-rt]').forEach(b => b.onclick = () => { tab = b.dataset.rt; render(); });
    bs.querySelectorAll('[data-go]').forEach(b => b.onclick = () => { view = b.dataset.go; render(); });
    bs.querySelectorAll('[data-open]').forEach(b => b.onclick = () => { close(); CMD.open(); });
    bs.querySelectorAll('[data-act]').forEach(b => b.onclick = () => { const a = b.dataset.act; if (a === 'theme') { CMD.theme(); render(); } else { close(); CMD[a](); } });
  }
  function open(v) { view = v || 'home'; tab = 'recent'; render(); bs.classList.remove('hidden'); bs.focus && bs.setAttribute('tabindex', '-1'); bs.focus(); }
  function close() { bs.classList.add('hidden'); App.focusGrid(); }
  bs.addEventListener('keydown', e => { if (e.key === 'Escape') close(); });
  App.backstage = { open, close };
  CMD.backstage = () => open('home');
  // Ctrl+Alt? keep Alt+F like Excel
  window.addEventListener('keydown', e => { if (e.altKey && e.key.toLowerCase() === 'f' && !e.ctrlKey) { e.preventDefault(); open('home'); } });
  // show the start screen on very first launch (no autosave yet)
  try { if (!localStorage.getItem('fmsheets.autosave') && !localStorage.getItem('fmsheets.started')) { localStorage.setItem('fmsheets.started', '1'); open('home'); } } catch (e) { }
})();
