/* FmOffice shared UI helpers for Docs & Slides (menus, dialogs, toast, file save/open, theme, icons) */
(function () {
  'use strict';
  const $ = s => document.querySelector(s); const tauri = window.__TAURI__;
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const P = {
    undo: 'M9 14 4 9l5-5M4 9h10.5a5.5 5.5 0 0 1 0 11H11', redo: 'm15 14 5-5-5-5M20 9H9.5a5.5 5.5 0 0 0 0 11H13', save: 'M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2zM17 21v-8H7v8M7 3v5h8',
    bold: 'M7 4h6a4 4 0 0 1 0 8H7zM7 12h7a4 4 0 0 1 0 8H7z', italic: 'M19 4h-9M14 20H5M15 4 9 20', underline: 'M6 4v6a6 6 0 0 0 12 0V4M4 21h16', strike: 'M16 4H9a3 3 0 0 0-2.8 4M14 12a4 4 0 0 1 0 8H6M4 12h16',
    left: 'M21 6H3M15 12H3M17 18H3', center: 'M21 6H3M17 12H7M19 18H5', right: 'M21 6H3M21 12H9M21 18H7', justify: 'M21 6H3M21 12H3M21 18H3', ul: 'M9 6h12M9 12h12M9 18h12M4 6h.01M4 12h.01M4 18h.01', ol: 'M10 6h11M10 12h11M10 18h11M4 6h1v4M4 10h2M6 18H4c0-1 2-2 2-3s-1-1.5-2-1',
    indent: 'M3 6h18M11 12h10M3 18h18M3 9l3 3-3 3', outdent: 'M3 6h18M11 12h10M3 18h18M6 9l-3 3 3 3', link: 'M10 13a5 5 0 0 0 7 0l3-3a5 5 0 0 0-7-7l-1 1M14 11a5 5 0 0 0-7 0l-3 3a5 5 0 0 0 7 7l1-1', image: 'M3 3h18v18H3zM8.5 10a1.5 1.5 0 1 0 0-3 1.5 1.5 0 0 0 0 3zM21 15l-5-5L5 21',
    table: 'M3 3h18v18H3zM3 9h18M3 15h18M9 9v12', hr: 'M3 12h18', search: 'M11 19a8 8 0 1 0 0-16 8 8 0 0 0 0 16zM21 21l-4.3-4.3', moon: 'M21 13A9 9 0 1 1 11 3a7 7 0 0 0 10 10z', sun: 'M12 17a5 5 0 1 0 0-10 5 5 0 0 0 0 10zM12 1v2M12 21v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M1 12h2M21 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4',
    print: 'M6 9V3h12v6M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2M6 14h12v7H6z', plus: 'M12 5v14M5 12h14', trash: 'M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6M10 11v6M14 11v6', copy: 'M9 9h11v11H9zM5 15H4V4h11v1',
    fill: 'm19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2c.8.8 2 .8 2.8 0zM5 2l5 5M2 13h15', text: 'M4 7V4h16v3M9 20h6M12 4v16', play: 'M6 4l14 8-14 8z', rect: 'M4 5h16v14H4z', ellipse: 'M12 5c5 0 9 3 9 7s-4 7-9 7-9-3-9-7 4-7 9-7z', tri: 'M12 4l9 16H3z', line: 'M4 20L20 4', up: 'M12 19V5M5 12l7-7 7 7', down: 'M12 5v14M19 12l-7 7-7-7',
    clear: 'M20 20H9L3 14l11-11 7 7-8 8M9 20', open: 'M3 7V5a2 2 0 0 1 2-2h4l2 3h8a2 2 0 0 1 2 2v1M3 7h18l-2 12H5z', file: 'M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9zM14 3v6h6', quote: 'M3 21c3 0 7-1 7-8V5H3v8h4M14 21c3 0 7-1 7-8V5h-7v8h4', pb: 'M3 12h4M10 12h4M17 12h4M6 5v4M6 15v4', note: 'M4 4h16v12H8l-4 4z'
  };
  const ic = (n, c) => `<svg class="i ${c || ''}" viewBox="0 0 24 24"><path d="${P[n] || P.file}"/></svg>`;
  const menu = $('#menu');
  function closeMenus() { if (menu) menu.classList.add('hidden'); }
  function showMenu(items, x, y) {
    menu.innerHTML = '';
    for (const it of items) {
      if (it === '-') { const d = document.createElement('div'); d.className = 'msep'; menu.appendChild(d); continue; }
      if (it.palette) { const p = document.createElement('div'); p.className = 'pal'; it.palette.forEach(c => { const i = document.createElement('i'); i.style.background = c || 'transparent'; if (!c) i.style.backgroundImage = 'linear-gradient(135deg,transparent 45%,#e33 45%,#e33 55%,transparent 55%)'; i.onmousedown = e => e.preventDefault(); i.onclick = () => { closeMenus(); it.pick(c); }; p.appendChild(i); }); menu.appendChild(p); continue; }
      const b = document.createElement('button'); b.className = 'mi' + (it.dis ? ' dis' : ''); b.innerHTML = (it.icon ? ic(it.icon) : '') + `<span>${esc(it.label)}</span>` + (it.key ? `<span class="k">${esc(it.key)}</span>` : ''); b.onmousedown = e => e.preventDefault(); b.onclick = () => { closeMenus(); it.fn && it.fn(); }; menu.appendChild(b);
    }
    menu.classList.remove('hidden'); menu.style.left = '0px'; menu.style.top = '0px'; const r = menu.getBoundingClientRect();
    menu.style.left = Math.max(6, Math.min(x, innerWidth - r.width - 8)) + 'px'; menu.style.top = Math.max(6, Math.min(y, innerHeight - r.height - 8)) + 'px';
  }
  document.addEventListener('mousedown', e => { if (!e.target.closest('#menu') && !e.target.closest('.dropbtn')) closeMenus(); }, true);
  const modal = $('#modal');
  function closeModal() { modal.classList.add('hidden'); modal.innerHTML = ''; }
  function dialog(title, body, buttons, opts) {
    modal.innerHTML = ''; const d = document.createElement('div'); d.className = 'dlg'; d.innerHTML = `<h3>${esc(title)}</h3>`; const b = document.createElement('div'); b.className = 'dbody'; b.innerHTML = body; d.appendChild(b);
    const f = document.createElement('div'); f.className = 'dfoot'; (buttons || [{ label: 'OK', primary: 1 }]).forEach(bt => { const x = document.createElement('button'); x.className = bt.primary ? 'pbtn' : 'gbtn'; x.textContent = bt.label; x.onclick = () => { if (!bt.fn || bt.fn(d) !== false) closeModal(); }; f.appendChild(x); }); d.appendChild(f);
    if (opts && opts.width) d.style.width = opts.width + 'px'; modal.appendChild(d); modal.classList.remove('hidden'); modal.onmousedown = e => { if (e.target === modal) closeModal(); };
    d.addEventListener('keydown', e => { if (e.key === 'Escape') closeModal(); if (e.key === 'Enter' && e.target.tagName !== 'TEXTAREA' && e.target.tagName !== 'BUTTON') { const p = f.querySelector('.pbtn'); if (p) p.click(); } });
    const first = d.querySelector('input,select,textarea'); if (first) { first.focus(); if (first.select) first.select(); } return d;
  }
  const prompt_ = (title, val, cb, label) => dialog(title, `<label>${esc(label || '')}<input type="text" id="pv" value="${esc(val)}"></label>`, [{ label: 'Cancel' }, { label: 'OK', primary: 1, fn: d => { cb(d.querySelector('#pv').value); } }]);
  let tt; function toast(m) { const t = $('#toast'); t.textContent = m; t.classList.remove('hidden'); clearTimeout(tt); tt = setTimeout(() => t.classList.add('hidden'), 2600); }
  async function saveBlob(name, blob, ext, label) {
    if (tauri && tauri.dialog && tauri.fs) { try { const path = await tauri.dialog.save({ defaultPath: name, filters: [{ name: label || ext, extensions: [ext] }] }); if (!path) return null; await tauri.fs.writeFile(path, new Uint8Array(await blob.arrayBuffer())); return { path }; } catch (e) { console.warn(e); } }
    if (window.showSaveFilePicker) { try { const h = await window.showSaveFilePicker({ suggestedName: name, types: [{ description: label || ext, accept: { [blob.type || 'application/octet-stream']: ['.' + ext] } }] }); const w = await h.createWritable(); await w.write(blob); await w.close(); return { handle: h }; } catch (e) { if (e.name === 'AbortError') return null; } }
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500); return { downloaded: true };
  }
  async function pickFile(accept, exts) {
    if (tauri && tauri.dialog && tauri.fs) { try { const path = await tauri.dialog.open({ multiple: false, filters: [{ name: 'Documents', extensions: exts }] }); if (!path) return null; const b = await tauri.fs.readFile(path); return { name: String(path).split(/[\\/]/).pop(), buf: b.buffer }; } catch (e) { console.warn(e); } }
    return new Promise(res => { const i = document.createElement('input'); i.type = 'file'; i.accept = accept; i.onchange = async () => { const f = i.files[0]; res(f ? { name: f.name, buf: await f.arrayBuffer() } : null); }; i.click(); });
  }
  function setTheme(t) { document.documentElement.dataset.theme = t; try { localStorage.setItem('fmoffice.theme', t); } catch (e) { } const b = $('#btn-theme'); if (b) b.innerHTML = ic(t === 'dark' ? 'sun' : 'moon'); }
  function initTheme() { let t = 'light'; try { t = localStorage.getItem('fmoffice.theme') || localStorage.getItem('fmsheets.theme') || (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light'); } catch (e) { } setTheme(t); const b = $('#btn-theme'); if (b) b.onclick = () => setTheme(document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'); }
  const dataUrlToBytes = u => { const m = /^data:([^;,]+)(;base64)?,(.*)$/s.exec(u); if (!m) return null; const bin = m[2] ? atob(m[3]) : decodeURIComponent(m[3]); const b = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) b[i] = bin.charCodeAt(i); return { mime: m[1], bytes: b }; };
  const bytesToDataUrl = (bytes, mime) => { let s = ''; for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192)); return 'data:' + mime + ';base64,' + btoa(s); };
  function openLink(url) { if (!/^(https?:|mailto:)/i.test(url)) return; if (tauri && tauri.core) tauri.core.invoke('open_url', { url }).catch(() => window.open(url, '_blank')); else window.open(url, '_blank', 'noopener'); }
  window.Fm = { $, esc, ic, P, showMenu, closeMenus, dialog, closeModal, prompt: prompt_, toast, saveBlob, pickFile, initTheme, dataUrlToBytes, bytesToDataUrl, openLink, tauri };
})();
