/* FmOffice Sheets — charts (SVG, no dependencies) */
(function () {
  'use strict';
  const E = FmEngine, App = window.App, S = App.S, $ = s => document.querySelector(s);
  const PAL = ['#2b8a57', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16'];
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const fmtN = v => { const a = Math.abs(v); if (a >= 1e9) return (v / 1e9).toFixed(1) + 'B'; if (a >= 1e6) return (v / 1e6).toFixed(1) + 'M'; if (a >= 1e4) return (v / 1e3).toFixed(0) + 'K'; return String(Math.round(v * 100) / 100); };

  function readData(ch) {
    const r = ch.range, si = ch.si !== undefined ? ch.si : S.si; const wb = S.wb; const rows = [];
    for (let i = r.r1; i <= r.r2; i++) { const row = []; for (let j = r.c1; j <= r.c2; j++) row.push({ v: wb.getValue(si, i, j), d: wb.display(si, i, j) }); rows.push(row); }
    if (!rows.length) return { cats: [], series: [] };
    const isNum = x => typeof x.v === 'number';
    const hasHdrRow = rows.length > 1 && rows[0].slice(1).every(x => !isNum(x)) && rows[1].slice(1).some(isNum);
    const hasCatCol = rows[0].length > 1 && rows.slice(hasHdrRow ? 1 : 0).every(rw => !isNum(rw[0]));
    const body = rows.slice(hasHdrRow ? 1 : 0); const c0 = hasCatCol ? 1 : 0;
    const cats = body.map((rw, i) => hasCatCol ? rw[0].d : String(i + 1));
    const series = []; for (let j = c0; j < rows[0].length; j++) series.push({ name: hasHdrRow ? rows[0][j].d : 'Series ' + (j - c0 + 1), vals: body.map(rw => isNum(rw[j]) ? rw[j].v : 0) });
    return { cats, series };
  }
  function niceMax(m) { if (m <= 0) return 1; const p = Math.pow(10, Math.floor(Math.log10(m))); const f = m / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; }
  function render(ch, W, H) {
    const d = readData(ch); const t = ch.type; let svg = `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="${esc(ch.title || 'Chart')}">`;
    svg += `<text x="${W / 2}" y="22" text-anchor="middle" style="font-weight:600;font-size:14px">${esc(ch.title || 'Chart Title')}</text>`;
    const multi = d.series.length > 1 || t === 'pie'; const legendH = multi ? 22 : 0;
    const L = 46, R = 16, T = 36, B = 30 + legendH; const pw = W - L - R, ph = H - T - B;
    if (!d.series.length || !d.cats.length) return svg + '<text x="20" y="60">No data</text></svg>';
    if (t === 'pie') {
      const vals = d.series[0].vals.map(v => Math.max(0, v)); const tot = vals.reduce((a, b) => a + b, 0) || 1; const cx = W * 0.4, cy = T + (H - T - 14) / 2, rad = Math.min(W * 0.32, (H - T - 20) / 2); let a0 = -Math.PI / 2;
      vals.forEach((v, i) => { const a1 = a0 + v / tot * Math.PI * 2; const large = a1 - a0 > Math.PI ? 1 : 0; const p = v / tot >= 0.9999 ? `<circle cx="${cx}" cy="${cy}" r="${rad}" fill="${PAL[i % 8]}"/>` : `<path d="M${cx},${cy} L${cx + rad * Math.cos(a0)},${cy + rad * Math.sin(a0)} A${rad},${rad} 0 ${large} 1 ${cx + rad * Math.cos(a1)},${cy + rad * Math.sin(a1)} Z" fill="${PAL[i % 8]}" stroke="var(--panel)" stroke-width="2"/>`; svg += p; if (v / tot > 0.05) { const am = (a0 + a1) / 2; svg += `<text x="${cx + rad * 0.65 * Math.cos(am)}" y="${cy + rad * 0.65 * Math.sin(am) + 4}" text-anchor="middle" style="fill:#fff;font-weight:600">${Math.round(v / tot * 100)}%</text>`; } a0 = a1; });
      d.cats.slice(0, 12).forEach((c, i) => { svg += `<rect x="${W * 0.75}" y="${T + 6 + i * 20}" width="11" height="11" rx="3" fill="${PAL[i % 8]}"/><text x="${W * 0.75 + 17}" y="${T + 16 + i * 20}">${esc(c.slice(0, 14))}</text>`; });
      return svg + '</svg>';
    }
    if (t === 'scatter') {
      const xs = d.cats.map(Number).map((x, i) => isNaN(x) ? i + 1 : x); const all = d.series.flatMap(s => s.vals); const xmin = Math.min(...xs), xmax = Math.max(...xs) || 1, ymax = niceMax(Math.max(...all, 0)), ymin = Math.min(0, ...all);
      for (let i = 0; i <= 4; i++) { const y = T + ph - ph * i / 4; svg += `<line x1="${L}" x2="${L + pw}" y1="${y}" y2="${y}" stroke="var(--line)"/><text x="${L - 6}" y="${y + 4}" text-anchor="end">${fmtN(ymin + (ymax - ymin) * i / 4)}</text>`; }
      d.series.forEach((s, k) => s.vals.forEach((v, i) => { const x = L + (xmax === xmin ? pw / 2 : (xs[i] - xmin) / (xmax - xmin) * pw), y = T + ph - (v - ymin) / ((ymax - ymin) || 1) * ph; svg += `<circle cx="${x}" cy="${y}" r="4.5" fill="${PAL[k % 8]}" opacity=".85"/>`; }));
      svg += `<text x="${L}" y="${T + ph + 18}">${fmtN(xmin)}</text><text x="${L + pw}" y="${T + ph + 18}" text-anchor="end">${fmtN(xmax)}</text>`; return svg + '</svg>';
    }
    const all = d.series.flatMap(s => s.vals); const stacked = t === 'stack'; let ymax = niceMax(stacked ? Math.max(...d.cats.map((_, i) => d.series.reduce((a, s) => a + Math.max(0, s.vals[i]), 0))) : Math.max(...all, 0)); const ymin = Math.min(0, ...all) < 0 ? -niceMax(-Math.min(...all)) : 0;
    const horiz = t === 'bar';
    const sy = v => T + ph - (v - ymin) / (ymax - ymin) * ph; const n = d.cats.length;
    if (!horiz) {
      for (let i = 0; i <= 4; i++) { const v = ymin + (ymax - ymin) * i / 4, y = sy(v); svg += `<line x1="${L}" x2="${L + pw}" y1="${y}" y2="${y}" stroke="var(--line)"/><text x="${L - 6}" y="${y + 4}" text-anchor="end">${fmtN(v)}</text>`; }
      const band = pw / n;
      d.cats.forEach((c, i) => { if (n <= 14 || i % Math.ceil(n / 14) === 0) svg += `<text x="${L + band * (i + .5)}" y="${T + ph + 16}" text-anchor="middle">${esc(c.slice(0, 10))}</text>`; });
      if (t === 'line' || t === 'area') {
        d.series.forEach((s, k) => { const pts = s.vals.map((v, i) => [L + band * (i + .5), sy(v)]); const path = pts.map((p, i) => (i ? 'L' : 'M') + p[0] + ',' + p[1]).join(' '); if (t === 'area') svg += `<path d="${path} L${pts[pts.length - 1][0]},${sy(0)} L${pts[0][0]},${sy(0)} Z" fill="${PAL[k % 8]}" opacity=".25"/>`; svg += `<path d="${path}" fill="none" stroke="${PAL[k % 8]}" stroke-width="2.5" stroke-linejoin="round"/>`; pts.forEach(p => svg += `<circle cx="${p[0]}" cy="${p[1]}" r="3.5" fill="${PAL[k % 8]}"/>`); });
      } else {
        const ns = stacked ? 1 : d.series.length, bw = Math.min(46, band * 0.72 / ns);
        d.cats.forEach((c, i) => { let acc = 0; d.series.forEach((s, k) => { const v = s.vals[i]; const x = L + band * (i + .5) - bw * ns / 2 + (stacked ? 0 : k * bw); const y0 = stacked ? sy(acc) : sy(0), y1 = stacked ? sy(acc + Math.max(0, v)) : sy(v); if (stacked) acc += Math.max(0, v); svg += `<rect x="${x}" y="${Math.min(y0, y1)}" width="${bw - 1}" height="${Math.abs(y0 - y1)}" rx="3" fill="${PAL[k % 8]}"/>`; }); });
      }
    } else {
      const band = ph / n; const sx = v => L + (v - ymin) / (ymax - ymin) * pw;
      for (let i = 0; i <= 4; i++) { const v = ymin + (ymax - ymin) * i / 4, x = sx(v); svg += `<line y1="${T}" y2="${T + ph}" x1="${x}" x2="${x}" stroke="var(--line)"/><text x="${x}" y="${T + ph + 16}" text-anchor="middle">${fmtN(v)}</text>`; }
      const bw = Math.min(30, band * 0.72 / d.series.length);
      d.cats.forEach((c, i) => { svg += `<text x="${L - 6}" y="${T + band * (i + .5) + 4}" text-anchor="end">${esc(c.slice(0, 8))}</text>`; d.series.forEach((s, k) => { const y = T + band * (i + .5) - bw * d.series.length / 2 + k * bw; svg += `<rect x="${sx(0)}" y="${y}" width="${Math.max(0, sx(s.vals[i]) - sx(0))}" height="${bw - 1}" rx="3" fill="${PAL[k % 8]}"/>`; }); });
    }
    svg += `<line x1="${L}" x2="${L + pw}" y1="${T + ph}" y2="${T + ph}" stroke="var(--muted)"/>`;
    if (multi) d.series.forEach((s, k) => { const x = L + k * 110; svg += `<rect x="${x}" y="${H - 20}" width="11" height="11" rx="3" fill="${PAL[k % 8]}"/><text x="${x + 16}" y="${H - 10}">${esc(s.name.slice(0, 12))}</text>`; });
    return svg + '</svg>';
  }

  let drag = null;
  App.renderCharts = function () {
    const host = $('#charts'); const sh = App.cur(); const z = S.zoom; host.innerHTML = '';
    host.style.left = (App.HW * z) + 'px'; host.style.top = (App.HH * z) + 'px';
    for (const ch of sh.charts) {
      const el = document.createElement('div'); el.className = 'chart' + (S.chartSel === ch.id ? ' sel' : ''); el.dataset.id = ch.id;
      el.style.cssText += `left:${ch.x * z}px;top:${ch.y * z}px;width:${ch.w * z}px;height:${ch.h * z}px`;
      el.innerHTML = render(ch, ch.w, ch.h) + '<div class="rz"></div>';
      el.addEventListener('mousedown', ev => {
        ev.stopPropagation(); S.chartSel = ch.id; document.querySelectorAll('.chart').forEach(x => x.classList.toggle('sel', x === el)); App.focusGrid();
        if (ev.button === 2) return; drag = { ch, el, resize: ev.target.classList.contains('rz'), x0: ev.clientX, y0: ev.clientY, ox: ch.x, oy: ch.y, ow: ch.w, oh: ch.h, snapped: false };
      });
      el.addEventListener('contextmenu', ev => { ev.preventDefault(); ev.stopPropagation(); S.chartSel = ch.id; App.showMenu([{ label: 'Edit title…', fn: () => editTitle(ch) }, { label: 'Column chart', fn: () => setType(ch, 'col') }, { label: 'Stacked column', fn: () => setType(ch, 'stack') }, { label: 'Bar chart', fn: () => setType(ch, 'bar') }, { label: 'Line chart', fn: () => setType(ch, 'line') }, { label: 'Area chart', fn: () => setType(ch, 'area') }, { label: 'Pie chart', fn: () => setType(ch, 'pie') }, { label: 'Scatter chart', fn: () => setType(ch, 'scatter') }, '-', { label: 'Delete chart', fn: () => deleteChart(ch.id) }], ev.clientX, ev.clientY); });
      el.addEventListener('dblclick', () => editTitle(ch));
      host.appendChild(el);
    }
  };
  const setType = (ch, t) => App.act(() => { ch.type = t; });
  function editTitle(ch) { App.prompt('Chart title', ch.title || '', v => App.act(() => { ch.title = v; })); }
  function deleteChart(id) { const sh = App.cur(); App.act(() => { sh.charts = sh.charts.filter(c => c.id !== id); }); S.chartSel = null; }
  App.chartDeselect = () => { if (S.chartSel) { S.chartSel = null; document.querySelectorAll('.chart.sel').forEach(x => x.classList.remove('sel')); } };
  window.addEventListener('mousemove', ev => {
    if (!drag) return; const z = S.zoom; const dx = (ev.clientX - drag.x0) / z, dy = (ev.clientY - drag.y0) / z; if (!drag.snapped && Math.abs(dx) + Math.abs(dy) > 2) { App.snap(); drag.snapped = true; }
    if (!drag.snapped) return;
    if (drag.resize) { drag.ch.w = Math.max(180, drag.ow + dx); drag.ch.h = Math.max(120, drag.oh + dy); } else { drag.ch.x = Math.max(0, drag.ox + dx); drag.ch.y = Math.max(0, drag.oy + dy); }
    drag.el.style.left = drag.ch.x * z + 'px'; drag.el.style.top = drag.ch.y * z + 'px'; drag.el.style.width = drag.ch.w * z + 'px'; drag.el.style.height = drag.ch.h * z + 'px';
    if (drag.resize) drag.el.querySelector('svg').outerHTML = render(drag.ch, drag.ch.w, drag.ch.h);
  });
  window.addEventListener('mouseup', () => { if (drag && drag.snapped) { App.scheduleAutosave(); } drag = null; });
  window.addEventListener('keydown', ev => { if ((ev.key === 'Delete' || ev.key === 'Backspace') && S.chartSel && !S.edit && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) { ev.preventDefault(); deleteChart(S.chartSel); } }, true);

  App.insertChart = function (type) {
    let n = App.selNorm(); const wb = S.wb, si = S.si;
    if (n.r1 === n.r2 && n.c1 === n.c2) { // current region
      let r1 = n.r1, r2 = n.r1, c1 = n.c1, c2 = n.c1; const has = (r, c) => { const v = wb.getValue(si, r, c); return v !== null && v !== ''; };
      let g = true; while (g) { g = false; for (let c = c1 - 1; c <= c2 + 1; c++) { if (c >= 0) { if (r1 > 0 && has(r1 - 1, c) && !(r1 - 1 < r1 && false)) { } } } break; }
      const rowHas = (r) => { for (let c = c1; c <= c2; c++) if (has(r, c)) return true; return false; }, colHas = (c) => { for (let r = r1; r <= r2; r++) if (has(r, c)) return true; return false; };
      let ch = true; while (ch) { ch = false; if (r1 > 0 && rowHas(r1 - 1)) { r1--; ch = true; } if (rowHas(r2 + 1) && r2 < 100000) { r2++; ch = true; } if (c1 > 0 && colHas(c1 - 1)) { c1--; ch = true; } if (colHas(c2 + 1) && c2 < 500) { c2++; ch = true; } }
      n = { r1, c1, r2, c2 };
    }
    if (n.r1 === n.r2 && n.c1 === n.c2 && !wb.getValue(si, n.r1, n.c1)) { App.toast('Select some data first, then insert a chart'); return; }
    const sh = App.cur(); const id = 'ch' + Date.now(); const x = App.colX(n.c2 + 1) + 20, y = App.rowY(n.r1);
    App.act(() => { sh.charts.push({ id, type, range: { ...n }, title: '', x, y, w: 460, h: 290 }); });
    S.chartSel = id; App.renderCharts();
  };
})();
