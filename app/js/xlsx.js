/* FmOffice Sheets — .xlsx / .csv import & export. No dependencies (uses CompressionStream when available). */
(function (root) {
  'use strict';
  const Eng = root.FmEngine;
  const enc = new TextEncoder(), dec = new TextDecoder();

  /* ---------- CRC32 + ZIP ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(b) { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zipStore(files, mime) { // files: [{name, data:Uint8Array}]
    const parts = [], central = []; let off = 0;
    const u16 = n => [n & 255, (n >> 8) & 255], u32 = n => [n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255];
    for (const f of files) {
      const name = enc.encode(f.name), crc = crc32(f.data), sz = f.data.length;
      const head = new Uint8Array([0x50, 0x4b, 3, 4, 20, 0, 0, 8, 0, 0, 0, 0, 0x21, 0, ...u32(crc), ...u32(sz), ...u32(sz), ...u16(name.length), 0, 0]);
      parts.push(head, name, f.data);
      central.push(new Uint8Array([0x50, 0x4b, 1, 2, 20, 0, 20, 0, 0, 8, 0, 0, 0, 0, 0x21, 0, ...u32(crc), ...u32(sz), ...u32(sz), ...u16(name.length), 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, ...u32(off)]), name);
      off += head.length + name.length + sz;
    }
    let csize = 0; central.forEach(p => csize += p.length);
    const end = new Uint8Array([0x50, 0x4b, 5, 6, 0, 0, 0, 0, ...u16(files.length), ...u16(files.length), ...u32(csize), ...u32(off), 0, 0]);
    return new Blob([...parts, ...central, end], { type: mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }
  async function inflateRaw(data) {
    const ds = new DecompressionStream('deflate-raw');
    const w = ds.writable.getWriter(); w.write(data); w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
  }
  async function unzip(buf) {
    const b = new Uint8Array(buf), dv = new DataView(buf);
    let e = b.length - 22; while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('Not a valid .xlsx file');
    const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true); const out = {};
    for (let i = 0; i < n; i++) {
      const method = dv.getUint16(p + 10, true), csz = dv.getUint32(p + 20, true), nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), lo = dv.getUint32(p + 42, true);
      const name = dec.decode(b.subarray(p + 46, p + 46 + nl)); p += 46 + nl + el + cl;
      const lnl = dv.getUint16(lo + 26, true), lel = dv.getUint16(lo + 28, true), ds = lo + 30 + lnl + lel;
      const raw = b.subarray(ds, ds + csz);
      out[name] = async () => method === 0 ? raw : inflateRaw(raw);
    }
    return out;
  }

  /* ---------- xml helpers ---------- */
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const parseXml = s => new DOMParser().parseFromString(s, 'application/xml');
  const kids = (el, tag) => el ? [...el.getElementsByTagName(tag)] : [];
  const argb = c => c ? 'FF' + c.replace('#', '').toUpperCase() : null;
  const hexOf = a => a ? '#' + a.slice(-6).toLowerCase() : null;
  const INDEXED = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF'];
  const BUILTIN_NF = { 0: null, 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)', 44: '_($* #,##0.00_)', 49: '@' };

  function chartXml(wb, si, sh, ch) {
    const r = ch.range; const q = n => /^[A-Za-z_][\w.]*$/.test(n) ? n : "'" + n.replace(/'/g, "''") + "'"; const sn = q(sh.name); const A = (rr, cc) => '$' + Eng.numToCol(cc) + '$' + (rr + 1);
    const isNum = (rr, cc) => typeof wb.getValue(si, rr, cc) === 'number';
    const hasHdr = r.r2 > r.r1 && (() => { for (let c = r.c1 + 1; c <= r.c2; c++) if (isNum(r.r1, c)) return false; return true; })() && (() => { for (let c = r.c1 + 1; c <= r.c2; c++) if (isNum(r.r1 + 1, c)) return true; return false; })();
    const hasCat = r.c2 > r.c1 && (() => { for (let rr = r.r1 + (hasHdr ? 1 : 0); rr <= r.r2; rr++) if (isNum(rr, r.c1)) return false; return true; })();
    const r0 = r.r1 + (hasHdr ? 1 : 0), c0 = r.c1 + (hasCat ? 1 : 0); const t = ch.type;
    const catF = hasCat ? `${sn}!${A(r0, r.c1)}:${A(r.r2, r.c1)}` : null;
    const ser = []; for (let c = c0; c <= r.c2; c++) {
      const tx = hasHdr ? `<c:tx><c:strRef><c:f>${sn}!${A(r.r1, c)}</c:f></c:strRef></c:tx>` : `<c:tx><c:v>Series ${c - c0 + 1}</c:v></c:tx>`;
      const val = `${sn}!${A(r0, c)}:${A(r.r2, c)}`;
      if (t === 'scatter') ser.push(`<c:ser><c:idx val="${c - c0}"/><c:order val="${c - c0}"/>${tx}${catF ? `<c:xVal><c:numRef><c:f>${catF}</c:f></c:numRef></c:xVal>` : ''}<c:yVal><c:numRef><c:f>${val}</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>`);
      else ser.push(`<c:ser><c:idx val="${c - c0}"/><c:order val="${c - c0}"/>${tx}${catF ? `<c:cat><c:strRef><c:f>${catF}</c:f></c:strRef></c:cat>` : ''}<c:val><c:numRef><c:f>${val}</c:f></c:numRef></c:val></c:ser>`);
    }
    const sx = ser.join(''); const ax = '<c:axId val="111"/><c:axId val="222"/>';
    let plot; if (t === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${ser.slice(0, 1).join('')}<c:firstSliceAng val="0"/></c:pieChart>`;
    else if (t === 'line') plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sx}<c:marker val="1"/>${ax}</c:lineChart>`;
    else if (t === 'area') plot = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${sx}${ax}</c:areaChart>`;
    else if (t === 'scatter') plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${sx}${ax}</c:scatterChart>`;
    else plot = `<c:barChart><c:barDir val="${t === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${t === 'stack' ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${sx}<c:gapWidth val="80"/>${t === 'stack' ? '<c:overlap val="100"/>' : ''}${ax}</c:barChart>`;
    const axes = t === 'pie' ? '' : (t === 'scatter' ? '<c:valAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="222"/></c:valAx>' : `<c:catAx><c:axId val="111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'l' : 'b'}"/><c:crossAx val="222"/></c:catAx>`) + `<c:valAx><c:axId val="222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'b' : 'l'}"/><c:majorGridlines/><c:crossAx val="111"/></c:valAx>`;
    const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart>${title}<c:plotArea><c:layout/>${plot}${axes}</c:plotArea><c:legend><c:legendPos val="b"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
  }

  /* ---------- EXPORT ---------- */
  function exportXlsx(wb) {
    const fonts = [], fills = [], borders = [], nfs = [], xfs = []; const idx = (arr, key) => { let i = arr.indexOf(key); if (i < 0) { i = arr.length; arr.push(key); } return i; };
    fonts.push('<font><sz val="11"/><name val="Calibri"/></font>'); fills.push('<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'); borders.push('<border><left/><right/><top/><bottom/><diagonal/></border>');
    xfs.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    const styleId = s => {
      if (!s) return 0;
      const f = '<font>' + (s.b ? '<b/>' : '') + (s.i ? '<i/>' : '') + (s.st ? '<strike/>' : '') + (s.u ? '<u/>' : '') + `<sz val="${s.fs || 11}"/>` + (s.fc ? `<color rgb="${argb(s.fc)}"/>` : '') + `<name val="${esc(s.ff || 'Calibri')}"/></font>`;
      const fi = idx(fonts, f);
      const fl = s.bg ? idx(fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(s.bg)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
      let bi = 0;
      if (s.bd) { const side = (n, v) => v ? `<${n} style="${v.w === 2 ? 'medium' : 'thin'}"><color rgb="${argb(v.c || '#000000')}"/></${n}>` : `<${n}/>`; bi = idx(borders, '<border>' + side('left', s.bd.l) + side('right', s.bd.r) + side('top', s.bd.t) + side('bottom', s.bd.b) + '<diagonal/></border>'); }
      let nid = 0; if (s.nf) { const rev = Object.entries(BUILTIN_NF).find(([k, v]) => v === s.nf); if (rev) nid = +rev[0]; else nid = 164 + idx(nfs, s.nf); }
      const al = (s.ha || s.va || s.wrap || s.ind) ? `<alignment${s.ha ? ` horizontal="${s.ha}"` : ''}${s.va ? ` vertical="${s.va === 'middle' ? 'center' : s.va}"` : ''}${s.wrap ? ' wrapText="1"' : ''}${s.ind ? ` indent="${s.ind}"` : ''}/>` : '';
      return idx(xfs, `<xf numFmtId="${nid}" fontId="${fi}" fillId="${fl}" borderId="${bi}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"${al ? ' applyAlignment="1">' + al + '</xf>' : '/>'}`);
    };
    const sst = [], sstMap = new Map(); const sidx = s => { let i = sstMap.get(s); if (i === undefined) { i = sst.length; sst.push(s); sstMap.set(s, i); } return i; };
    const files = []; const add = (name, text) => files.push({ name, data: enc.encode(text) }); const dxfs = []; const ctOver = []; let chartN = 0, drawN = 0, cmtN = 0;
    const dxfId = (bg, fc, b) => { const x = '<dxf>' + (fc || b ? '<font>' + (b ? '<b/>' : '') + (fc ? `<color rgb="${argb(fc)}"/>` : '') + '</font>' : '') + (bg ? `<fill><patternFill><bgColor rgb="${argb(bg)}"/></patternFill></fill>` : '') + '</dxf>'; let i = dxfs.indexOf(x); if (i < 0) { i = dxfs.length; dxfs.push(x); } return i; };
    const colPx = (sh, c) => sh.hiddenCols[c] ? 0 : (sh.colW[c] !== undefined ? sh.colW[c] : 84), rowPx = (sh, r) => sh.hiddenRows[r] ? 0 : (sh.rowH[r] !== undefined ? sh.rowH[r] : 24);
    const cellAtPx = (sh, px, isCol) => { let acc = 0, i = 0; for (; i < 20000; i++) { const w = isCol ? colPx(sh, i) : rowPx(sh, i); if (acc + w > px) return { i, off: px - acc }; acc += w; } return { i, off: 0 }; };
    const sheetXml = wb.sheets.map((sh, si) => {
      const srels = []; let rid = 0; const addRel = (type, target, ext) => { const id = 'rId' + (++rid); srels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`); return id; };
      const rows = new Map();
      for (const [k, cell] of sh.cells) { const r = Math.floor(k / 20000), c = k % 20000; if (!rows.has(r)) rows.set(r, []); rows.get(r).push([c, cell]); }
      let xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
      if (sh.tabColor) xml += `<sheetPr><tabColor rgb="${argb(sh.tabColor)}"/></sheetPr>`;
      xml += `<sheetViews><sheetView workbookViewId="0"${sh.showGrid === false ? ' showGridLines="0"' : ''}${si === 0 ? ' tabSelected="1"' : ''}>`;
      if (sh.freeze.r || sh.freeze.c) xml += `<pane${sh.freeze.c ? ` xSplit="${sh.freeze.c}"` : ''}${sh.freeze.r ? ` ySplit="${sh.freeze.r}"` : ''} topLeftCell="${Eng.addrToStr(sh.freeze.r, sh.freeze.c)}" activePane="${sh.freeze.r && sh.freeze.c ? 'bottomRight' : sh.freeze.r ? 'bottomLeft' : 'topRight'}" state="frozen"/>`;
      xml += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>';
      const cw = Object.keys(sh.colW); if (cw.length) xml += '<cols>' + cw.sort((a, b) => a - b).map(c => `<col min="${+c + 1}" max="${+c + 1}" width="${(sh.colW[c] / 7).toFixed(2)}" customWidth="1"${sh.hiddenCols[c] ? ' hidden="1"' : ''}/>`).join('') + '</cols>';
      xml += '<sheetData>';
      const rkeys = [...new Set([...rows.keys(), ...Object.keys(sh.rowH).map(Number)])].sort((a, b) => a - b);
      for (const r of rkeys) {
        const cs = (rows.get(r) || []).sort((a, b) => a[0] - b[0]);
        xml += `<row r="${r + 1}"${sh.rowH[r] ? ` ht="${(sh.rowH[r] * 0.75).toFixed(1)}" customHeight="1"` : ''}${sh.hiddenRows[r] ? ' hidden="1"' : ''}>`;
        for (const [c, cell] of cs) {
          const ref = Eng.addrToStr(r, c), sid = styleId(cell.s), sa = sid ? ` s="${sid}"` : '';
          if (cell.t === '' || cell.t === undefined) { xml += `<c r="${ref}"${sa}/>`; continue; }
          if (cell.t[0] === '=' && cell.t.length > 1) {
            const v = wb.getValue(si, r, c); const f = esc(cell.t.slice(1));
            if (typeof v === 'number') xml += `<c r="${ref}"${sa}><f>${f}</f><v>${v}</v></c>`;
            else if (typeof v === 'boolean') xml += `<c r="${ref}"${sa} t="b"><f>${f}</f><v>${v ? 1 : 0}</v></c>`;
            else if (v instanceof Eng.FErr) xml += `<c r="${ref}"${sa} t="e"><f>${f}</f><v>${esc(v.code)}</v></c>`;
            else xml += `<c r="${ref}"${sa} t="str"><f>${f}</f><v>${esc(v == null ? '' : v)}</v></c>`;
            continue;
          }
          const lit = cell.lit || Eng.parseLiteral(cell.t); const v = lit.v;
          if (typeof v === 'number') { let sid2 = sid; if (!cell.s && lit.fmt) sid2 = styleId({ nf: lit.fmt }); xml += `<c r="${ref}"${sid2 ? ` s="${sid2}"` : ''}><v>${v}</v></c>`; }
          else if (typeof v === 'boolean') xml += `<c r="${ref}"${sa} t="b"><v>${v ? 1 : 0}</v></c>`;
          else if (v instanceof Eng.FErr) xml += `<c r="${ref}"${sa} t="e"><v>${esc(v.code)}</v></c>`;
          else if (v === null) xml += `<c r="${ref}"${sa}/>`;
          else xml += `<c r="${ref}"${sa} t="s"><v>${sidx(String(v))}</v></c>`;
        }
        xml += '</row>';
      }
      xml += '</sheetData>';
      if (sh.filter) { const f = sh.filter; const last = Math.max(f.last || 0, f.r + 1, wb.usedRange(si).r2); xml += `<autoFilter ref="${Eng.addrToStr(f.r, f.c1)}:${Eng.addrToStr(last, f.c2)}"/>`; }
      if (sh.merges.length) xml += `<mergeCells count="${sh.merges.length}">` + sh.merges.map(m => `<mergeCell ref="${Eng.addrToStr(m.r1, m.c1)}:${Eng.addrToStr(m.r2, m.c2)}"/>`).join('') + '</mergeCells>';
      // conditional formatting
      let prio = 1;
      for (const r of sh.cf) {
        const ref = `${Eng.addrToStr(r.r1, r.c1)}:${Eng.addrToStr(Math.min(r.r2, 1048575), Math.min(r.c2, 16383))}`, tl = Eng.addrToStr(r.r1, r.c1); let rule = '';
        const d = () => ` dxfId="${dxfId(r.bg, r.fc, r.bold)}"`;
        const q = v => typeof v === 'number' ? v : '"' + String(v).replace(/"/g, '""') + '"';
        if (r.type === 'gt') rule = `<cfRule type="cellIs"${d()} priority="${prio++}" operator="greaterThan"><formula>${esc(q(r.v))}</formula></cfRule>`;
        else if (r.type === 'lt') rule = `<cfRule type="cellIs"${d()} priority="${prio++}" operator="lessThan"><formula>${esc(q(r.v))}</formula></cfRule>`;
        else if (r.type === 'eq') rule = `<cfRule type="cellIs"${d()} priority="${prio++}" operator="equal"><formula>${esc(q(r.v))}</formula></cfRule>`;
        else if (r.type === 'between') rule = `<cfRule type="cellIs"${d()} priority="${prio++}" operator="between"><formula>${esc(q(r.v))}</formula><formula>${esc(q(r.v2))}</formula></cfRule>`;
        else if (r.type === 'text') rule = `<cfRule type="containsText"${d()} priority="${prio++}" operator="containsText" text="${esc(r.v)}"><formula>NOT(ISERROR(SEARCH("${esc(String(r.v).replace(/"/g, '""'))}",${tl})))</formula></cfRule>`;
        else if (r.type === 'dup') rule = `<cfRule type="duplicateValues"${d()} priority="${prio++}"/>`;
        else if (r.type === 'blank') rule = `<cfRule type="containsBlanks"${d()} priority="${prio++}"><formula>LEN(TRIM(${tl}))=0</formula></cfRule>`;
        else if (r.type === 'scale') rule = `<cfRule type="colorScale" priority="${prio++}"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/><color rgb="${argb(r.c1c || '#f8696b')}"/><color rgb="${argb(r.c3c || '#ffeb84')}"/><color rgb="${argb(r.c2c || '#63be7b')}"/></colorScale></cfRule>`;
        else if (r.type === 'bar') rule = `<cfRule type="dataBar" priority="${prio++}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(r.color || '#63be7b')}"/></dataBar></cfRule>`;
        if (rule) xml += `<conditionalFormatting sqref="${ref}">${rule}</conditionalFormatting>`;
      }
      if (sh.dv.length) xml += `<dataValidations count="${sh.dv.length}">` + sh.dv.map(v => { const rng = `${Eng.addrToStr(v.r1, v.c1)}:${Eng.addrToStr(Math.min(v.r2, 1048575), Math.min(v.c2, 16383))}`; const f1 = v.type === 'list' ? (String(v.src)[0] === '=' ? esc(String(v.src).slice(1)) : '&quot;' + esc(v.src) + '&quot;') : esc(v.v1); return `<dataValidation type="${v.type}" errorStyle="${v.style === 'warning' ? 'warning' : v.style === 'info' ? 'information' : 'stop'}"${v.type !== 'list' && v.op && v.op !== 'between' ? ` operator="${{ notBetween: 'notBetween', eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual' }[v.op]}"` : ''} allowBlank="${v.allowBlank === false ? 0 : 1}" showDropDown="${v.showDrop === false ? 1 : 0}" showErrorMessage="1" errorTitle="${esc(v.errTitle || '')}" error="${esc(v.errMsg || '')}" sqref="${rng}"><formula1>${f1}</formula1>${['between', 'notBetween'].includes(v.op || 'between') && v.type !== 'list' ? `<formula2>${esc(v.v2)}</formula2>` : ''}</dataValidation>`; }).join('') + '</dataValidations>';
      const lk = Object.entries(sh.links); if (lk.length) xml += '<hyperlinks>' + lk.map(([k, u]) => `<hyperlink ref="${Eng.addrToStr(Math.floor(k / 20000), k % 20000)}" r:id="${addRel('hyperlink', u, true)}"/>`).join('') + '</hyperlinks>';
      xml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>';
      // charts -> drawing
      if (sh.charts.length) {
        drawN++; let dx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'; const drels = [];
        sh.charts.forEach((ch, k) => {
          chartN++; const a = cellAtPx(sh, ch.x, true), b = cellAtPx(sh, ch.y, false), a2 = cellAtPx(sh, ch.x + ch.w, true), b2 = cellAtPx(sh, ch.y + ch.h, false);
          dx += `<xdr:twoCellAnchor><xdr:from><xdr:col>${a.i}</xdr:col><xdr:colOff>${Math.round(a.off * 9525)}</xdr:colOff><xdr:row>${b.i}</xdr:row><xdr:rowOff>${Math.round(b.off * 9525)}</xdr:rowOff></xdr:from><xdr:to><xdr:col>${a2.i}</xdr:col><xdr:colOff>${Math.round(a2.off * 9525)}</xdr:colOff><xdr:row>${b2.i}</xdr:row><xdr:rowOff>${Math.round(b2.off * 9525)}</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${k + 2}" name="Chart ${k + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId${k + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
          drels.push(`<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${chartN}.xml"/>`);
          add(`xl/charts/chart${chartN}.xml`, chartXml(wb, si, sh, ch)); ctOver.push([`/xl/charts/chart${chartN}.xml`, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml']);
        });
        dx += '</xdr:wsDr>'; add(`xl/drawings/drawing${drawN}.xml`, dx); add(`xl/drawings/_rels/drawing${drawN}.xml.rels`, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + drels.join('') + '</Relationships>');
        ctOver.push([`/xl/drawings/drawing${drawN}.xml`, 'application/vnd.openxmlformats-officedocument.drawing+xml']); xml += `<drawing r:id="${addRel('drawing', `../drawings/drawing${drawN}.xml`)}"/>`;
      }
      // notes -> comments
      const nt = Object.entries(sh.notes); if (nt.length) {
        cmtN++; add(`xl/comments${cmtN}.xml`, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>FmOffice</author></authors><commentList>' + nt.map(([k, t]) => `<comment ref="${Eng.addrToStr(Math.floor(k / 20000), k % 20000)}" authorId="0"><text><r><t xml:space="preserve">${esc(t)}</t></r></text></comment>`).join('') + '</commentList></comments>');
        add(`xl/drawings/vmlDrawing${cmtN}.vml`, '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + cmtN + '"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' + nt.map(([k, t], i) => { const r = Math.floor(k / 20000), c = k % 20000; return `<v:shape id="_x0000_s${cmtN * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:5pt;width:140pt;height:60pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 15, ${r + 3}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`; }).join('') + '</xml>');
        ctOver.push([`/xl/comments${cmtN}.xml`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml']);
        addRel('comments', `../comments${cmtN}.xml`); xml += `<legacyDrawing r:id="${addRel('vmlDrawing', `../drawings/vmlDrawing${cmtN}.vml`)}"/>`;
      }
      xml += '</worksheet>';
      if (srels.length) add(`xl/worksheets/_rels/sheet${si + 1}.xml.rels`, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + srels.join('') + '</Relationships>');
      return xml;
    });
    const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + wb.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' + ctOver.map(([p, t]) => `<Override PartName="${p}" ContentType="${t}"/>`).join('') + '</Types>';
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const wbx = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets>' + (Object.keys(wb.names).length ? '<definedNames>' + Object.entries(wb.names).map(([k, v]) => `<definedName name="${esc(k)}">${esc(v)}</definedName>`).join('') + '</definedNames>' : '') + '</workbook>';
    const wbr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wb.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') + `<Relationship Id="rId${wb.sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId${wb.sheets.length + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
    const styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' + (nfs.length ? `<numFmts count="${nfs.length}">` + nfs.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${esc(f)}"/>`).join('') + '</numFmts>' : '') + `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="${borders.length}">${borders.join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>${dxfs.length ? `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>` : ''}</styleSheet>`;
    const sstx = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sst.length}" uniqueCount="${sst.length}">` + sst.map(s => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('') + '</sst>';
    add('[Content_Types].xml', ct); add('_rels/.rels', rels); add('xl/workbook.xml', wbx); add('xl/_rels/workbook.xml.rels', wbr);
    sheetXml.forEach((x, i) => add(`xl/worksheets/sheet${i + 1}.xml`, x)); add('xl/styles.xml', styles); add('xl/sharedStrings.xml', sstx);
    return zipStore(files);
  }

  /* ---------- IMPORT ---------- */
  async function importDrawing(path, sh, readRels, get) {
    const x = await get(path); if (!x) return; const rels = await readRels(path); const doc = parseXml(x);
    const px = (idx, off, isCol) => { let acc = 0; for (let i = 0; i < idx; i++) acc += isCol ? (sh.hiddenCols[i] ? 0 : (sh.colW[i] !== undefined ? sh.colW[i] : 84)) : (sh.hiddenRows[i] ? 0 : (sh.rowH[i] !== undefined ? sh.rowH[i] : 24)); return acc + off / 9525; };
    for (const an of [...doc.getElementsByTagName('*')].filter(e => e.localName === 'twoCellAnchor')) {
      const ce = an.getElementsByTagName('c:chart')[0] || [...an.getElementsByTagName('*')].find(e => e.localName === 'chart'); if (!ce) continue; const rid = ce.getAttribute('r:id') || ce.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'); const rel = rels[rid]; if (!rel) continue;
      const cx = await get(rel.target); if (!cx) continue; const cd = parseXml(cx); const loc = (n) => [...cd.getElementsByTagName('*')].filter(e => e.localName === n);
      let type = 'col'; if (loc('pieChart').length || loc('doughnutChart').length) type = 'pie'; else if (loc('lineChart').length) type = 'line'; else if (loc('areaChart').length) type = 'area'; else if (loc('scatterChart').length) type = 'scatter'; else if (loc('barChart').length) { const bd = loc('barDir')[0]; const gp = loc('grouping')[0]; type = bd && bd.getAttribute('val') === 'bar' ? 'bar' : (gp && /stacked|percentStacked/i.test(gp.getAttribute('val')) ? 'stack' : 'col'); }
      const refs = e => { const f = [...e.getElementsByTagName('*')].find(n => n.localName === 'f'); if (!f) return null; const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f.textContent.trim()); if (!m) return null; return { c1: Eng.colToNum(m[3]), r1: +m[4] - 1, c2: Eng.colToNum(m[5] || m[3]), r2: +(m[6] || m[4]) - 1 }; };
      const sers = loc('ser'); if (!sers.length) continue; let r1 = 1e9, r2 = -1, c1 = 1e9, c2 = -1; let hasTx = false;
      for (const se of sers) { const v = [...se.children].find(n => n.localName === (type === 'scatter' ? 'yVal' : 'val')), c = [...se.children].find(n => n.localName === (type === 'scatter' ? 'xVal' : 'cat')), t = [...se.children].find(n => n.localName === 'tx'); const vr = v && refs(v); if (!vr) continue; r1 = Math.min(r1, vr.r1); r2 = Math.max(r2, vr.r2); c1 = Math.min(c1, vr.c1); c2 = Math.max(c2, vr.c2); const cr = c && refs(c); if (cr) c1 = Math.min(c1, cr.c1); const tr = t && refs(t); if (tr) { hasTx = true; r1 = Math.min(r1, tr.r1); } }
      if (r2 < 0) continue; const fr = an.getElementsByTagName('xdr:from')[0] || [...an.children].find(n => n.localName === 'from'), to = [...an.children].find(n => n.localName === 'to'); const g = (p, n) => +([...p.children].find(e => e.localName === n).textContent);
      const x1 = px(g(fr, 'col'), g(fr, 'colOff'), true), y1 = px(g(fr, 'row'), g(fr, 'rowOff'), false), x2 = px(g(to, 'col'), g(to, 'colOff'), true), y2 = px(g(to, 'row'), g(to, 'rowOff'), false);
      const tt = loc('title')[0]; const title = tt ? [...tt.getElementsByTagName('*')].filter(e => e.localName === 't').map(e => e.textContent).join('') : '';
      sh.charts.push({ id: 'ch' + Date.now() + sh.charts.length, type, range: { r1, c1, r2, c2 }, title, x: x1, y: y1, w: Math.max(180, x2 - x1), h: Math.max(120, y2 - y1) });
    }
  }
  async function importXlsx(buf) {
    const z = await unzip(buf); const get = async n => z[n] ? dec.decode(await z[n]()) : null;
    const wbDoc = parseXml(await get('xl/workbook.xml')), relDoc = parseXml(await get('xl/_rels/workbook.xml.rels'));
    const rels = {}; kids(relDoc, 'Relationship').forEach(r => rels[r.getAttribute('Id')] = r.getAttribute('Target'));
    const ss = []; const sx = await get('xl/sharedStrings.xml');
    if (sx) kids(parseXml(sx), 'si').forEach(si => ss.push(kids(si, 't').filter(t => !t.parentNode || t.parentNode.nodeName !== 'rPh').map(t => t.textContent).join('')));
    // styles
    const xfList = []; const dxfList = []; const stx = await get('xl/styles.xml');
    if (stx) {
      const sd = parseXml(stx); const nmap = {}; kids(sd, 'numFmt').forEach(n => nmap[n.getAttribute('numFmtId')] = n.getAttribute('formatCode'));
      const fontsEl = kids(sd.getElementsByTagName('fonts')[0], 'font'), fillsEl = kids(sd.getElementsByTagName('fills')[0], 'fill'), bordEl = sd.getElementsByTagName('borders')[0] ? kids(sd.getElementsByTagName('borders')[0], 'border') : [];
      const colorOf = el => { if (!el) return null; const rgb = el.getAttribute('rgb'); if (rgb) return hexOf(rgb); const ix = el.getAttribute('indexed'); if (ix && INDEXED[ix]) return '#' + INDEXED[ix].toLowerCase(); return null; };
      const dxfsEl = sd.getElementsByTagName('dxfs')[0]; if (dxfsEl) for (const dx of [...dxfsEl.children]) { const o = {}; const f = dx.getElementsByTagName('font')[0]; if (f) { if (f.getElementsByTagName('b')[0]) o.bold = 1; const c = colorOf(f.getElementsByTagName('color')[0]); if (c) o.fc = c; } const pf = dx.getElementsByTagName('fill')[0]; if (pf) { const c = colorOf(pf.getElementsByTagName('bgColor')[0]) || colorOf(pf.getElementsByTagName('fgColor')[0]); if (c) o.bg = c; } dxfList.push(o); }
      const cxfs = sd.getElementsByTagName('cellXfs')[0];
      if (cxfs) for (const xf of [...cxfs.children]) {
        const s = {}; const nid = xf.getAttribute('numFmtId'); const nf = nmap[nid] || BUILTIN_NF[nid]; if (nf) s.nf = nf;
        const f = fontsEl[+xf.getAttribute('fontId') || 0];
        if (f) {
          if (f.getElementsByTagName('b')[0]) s.b = 1; if (f.getElementsByTagName('i')[0]) s.i = 1; if (f.getElementsByTagName('u')[0]) s.u = 1; if (f.getElementsByTagName('strike')[0]) s.st = 1;
          const sz = f.getElementsByTagName('sz')[0]; if (sz && +sz.getAttribute('val') !== 11) s.fs = +sz.getAttribute('val');
          const nm = f.getElementsByTagName('name')[0]; if (nm && nm.getAttribute('val') !== 'Calibri') s.ff = nm.getAttribute('val');
          const col = colorOf(f.getElementsByTagName('color')[0]); if (col && col !== '#000000') s.fc = col;
        }
        const fl = fillsEl[+xf.getAttribute('fillId') || 0];
        if (fl) { const pf = fl.getElementsByTagName('patternFill')[0]; if (pf && pf.getAttribute('patternType') === 'solid') { const c = colorOf(pf.getElementsByTagName('fgColor')[0]); if (c) s.bg = c; } }
        const bd = bordEl[+xf.getAttribute('borderId') || 0];
        if (bd) { const bo = {}; ['left:l', 'right:r', 'top:t', 'bottom:b'].forEach(p => { const [tag, k] = p.split(':'); const e = bd.getElementsByTagName(tag)[0]; if (e && e.getAttribute('style')) bo[k] = { c: colorOf(e.getElementsByTagName('color')[0]) || '#000000', w: /medium|thick/.test(e.getAttribute('style')) ? 2 : 1 }; }); if (Object.keys(bo).length) s.bd = bo; }
        const al = xf.getElementsByTagName('alignment')[0];
        if (al) { if (al.getAttribute('horizontal') && al.getAttribute('horizontal') !== 'general') s.ha = al.getAttribute('horizontal'); const v = al.getAttribute('vertical'); if (v) s.va = v === 'center' ? 'middle' : v; if (al.getAttribute('wrapText') === '1') s.wrap = 1; if (al.getAttribute('indent')) s.ind = +al.getAttribute('indent'); }
        xfList.push(Object.keys(s).length ? s : null);
      }
    }
    const resolve = (from, target) => { if (target[0] === '/') return target.slice(1); const parts = from.split('/'); parts.pop(); for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); } return parts.join('/'); };
    const readRels = async path => { const rp = path.replace(/([^/]+)$/, '_rels/$1.rels'); const x = await get(rp); const m = {}; if (x) kids(parseXml(x), 'Relationship').forEach(r => m[r.getAttribute('Id')] = { type: r.getAttribute('Type').split('/').pop(), target: r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target') : resolve(path, r.getAttribute('Target')), ext: r.getAttribute('TargetMode') === 'External' }); return m; };
    const wb = new Eng.Workbook(); wb.sheets = [];
    for (const dn of kids(wbDoc, 'definedName')) { const nm = dn.getAttribute('name'); if (!nm || nm.startsWith('_xlnm') || dn.getAttribute('localSheetId') !== null || dn.getAttribute('hidden') === '1') continue; wb.names[nm.toUpperCase()] = dn.textContent.replace(/^=/, ''); }
    for (const sEl of kids(wbDoc, 'sheet')) {
      const name = sEl.getAttribute('name'); const rid = sEl.getAttribute('r:id') || sEl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
      let tgt = rels[rid]; if (!tgt) continue; tgt = tgt.replace(/^\//, ''); const path = tgt.startsWith('xl/') ? tgt : 'xl/' + tgt;
      const sh = new Eng.Sheet(name); wb.sheets.push(sh);
      const xml = await get(path); if (!xml) continue; const d = parseXml(xml);
      const tc = d.getElementsByTagName('tabColor')[0]; if (tc && tc.getAttribute('rgb')) sh.tabColor = hexOf(tc.getAttribute('rgb'));
      const sv = d.getElementsByTagName('sheetView')[0]; if (sv && sv.getAttribute('showGridLines') === '0') sh.showGrid = false;
      const pane = d.getElementsByTagName('pane')[0]; if (pane && pane.getAttribute('state') === 'frozen') sh.freeze = { r: +pane.getAttribute('ySplit') || 0, c: +pane.getAttribute('xSplit') || 0 };
      kids(d, 'col').forEach(c => { const w = Math.round(parseFloat(c.getAttribute('width')) * 7); for (let i = +c.getAttribute('min'); i <= Math.min(+c.getAttribute('max'), 200); i++) { sh.colW[i - 1] = w; if (c.getAttribute('hidden') === '1') sh.hiddenCols[i - 1] = 1; } });
      kids(d, 'mergeCell').forEach(m => { const [a, b] = m.getAttribute('ref').split(':'); const pa = Eng.parseAddr(a), pb = Eng.parseAddr(b); if (pa && pb) sh.merges.push({ r1: pa.r, c1: pa.c, r2: pb.r, c2: pb.c }); });
      const srels = await readRels(path); const sIdx = wb.sheets.length - 1;
      const qv = t => { t = String(t); return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : t.replace(/^"|"$/g, '').replace(/""/g, '"'); };
      const af = d.getElementsByTagName('autoFilter')[0]; if (af && af.getAttribute('ref')) { const [a, b] = af.getAttribute('ref').split(':'); const pa = Eng.parseAddr(a), pb = Eng.parseAddr(b || a); if (pa && pb) sh.filter = { r: pa.r, c1: pa.c, c2: pb.c, ex: {} }; }
      for (const cfEl of kids(d, 'conditionalFormatting')) { const rg = (cfEl.getAttribute('sqref') || '').split(' ')[0].split(':'); const pa = Eng.parseAddr(rg[0]), pb = Eng.parseAddr(rg[1] || rg[0]); if (!pa || !pb) continue; const base = { r1: pa.r, c1: pa.c, r2: pb.r, c2: pb.c };
        for (const rule of kids(cfEl, 'cfRule')) { const t = rule.getAttribute('type'); const dx = dxfList[+rule.getAttribute('dxfId')] || {}; const fm = [...rule.getElementsByTagName('formula')].map(x => x.textContent); const st = { bg: dx.bg, fc: dx.fc, bold: dx.bold };
          if (t === 'cellIs') { const op = rule.getAttribute('operator'); const map = { greaterThan: 'gt', lessThan: 'lt', equal: 'eq', between: 'between', greaterThanOrEqual: 'gt', lessThanOrEqual: 'lt' }; if (map[op]) sh.cf.push({ ...base, type: map[op], v: qv(fm[0]), v2: fm[1] !== undefined ? qv(fm[1]) : undefined, ...st }); }
          else if (t === 'containsText') sh.cf.push({ ...base, type: 'text', v: rule.getAttribute('text') || '', ...st }); else if (t === 'duplicateValues') sh.cf.push({ ...base, type: 'dup', ...st }); else if (t === 'containsBlanks') sh.cf.push({ ...base, type: 'blank', ...st });
          else if (t === 'colorScale') { const cs = [...rule.getElementsByTagName('color')].map(x => hexOf(x.getAttribute('rgb'))); if (cs.length >= 2) sh.cf.push({ ...base, type: 'scale', c1c: cs[0], c2c: cs[cs.length - 1], c3c: cs.length > 2 ? cs[1] : '#ffeb84' }); }
          else if (t === 'dataBar') { const c = rule.getElementsByTagName('color')[0]; sh.cf.push({ ...base, type: 'bar', color: c && c.getAttribute('rgb') ? hexOf(c.getAttribute('rgb')) : '#63be7b' }); } } }
      for (const v of kids(d, 'dataValidation')) { const rg = (v.getAttribute('sqref') || '').split(' ')[0].split(':'); const pa = Eng.parseAddr(rg[0]), pb = Eng.parseAddr(rg[1] || rg[0]); const type = v.getAttribute('type'); if (!pa || !pb || !['list', 'whole', 'decimal', 'date', 'textLength'].includes(type)) continue; const f1 = v.getElementsByTagName('formula1')[0], f2 = v.getElementsByTagName('formula2')[0]; const opm = { notBetween: 'notBetween', equal: 'eq', notEqual: 'ne', greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le' };
        const rule = { r1: pa.r, c1: pa.c, r2: pb.r, c2: pb.c, type, op: opm[v.getAttribute('operator')] || 'between', allowBlank: v.getAttribute('allowBlank') !== '0', showDrop: v.getAttribute('showDropDown') !== '1', style: { warning: 'warning', information: 'info' }[v.getAttribute('errorStyle')] || 'stop', errTitle: v.getAttribute('errorTitle') || '', errMsg: v.getAttribute('error') || '' };
        if (type === 'list') { const t = f1 ? f1.textContent : ''; rule.src = /^".*"$/.test(t) ? t.slice(1, -1) : '=' + t; } else { rule.v1 = f1 ? f1.textContent : ''; rule.v2 = f2 ? f2.textContent : ''; } sh.dv.push(rule); }
      for (const h of kids(d, 'hyperlink')) { const a = Eng.parseAddr((h.getAttribute('ref') || '').split(':')[0]); if (!a) continue; const rid = h.getAttribute('r:id') || h.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'); const loc = h.getAttribute('location'); if (rid && srels[rid] && srels[rid].ext) sh.links[a.r * 20000 + a.c] = srels[rid].target; else if (loc) { /* internal link: ignored */ } }
      for (const rid in srels) {
        if (srels[rid].type === 'comments') { const cx = await get(srels[rid].target); if (cx) for (const cm of kids(parseXml(cx), 'comment')) { const a = Eng.parseAddr(cm.getAttribute('ref')); if (a) sh.notes[a.r * 20000 + a.c] = kids(cm, 't').map(t => t.textContent).join(''); } }
        if (srels[rid].type === 'drawing') { try { await importDrawing(srels[rid].target, sh, readRels, get); } catch (e) { console.warn('chart import', e); } }
      }
      for (const row of kids(d, 'row')) {
        const rn = +row.getAttribute('r') - 1; if (row.getAttribute('ht') && row.getAttribute('customHeight') === '1') sh.rowH[rn] = Math.round(parseFloat(row.getAttribute('ht')) / 0.75); if (row.getAttribute('hidden') === '1') sh.hiddenRows[rn] = 1;
        for (const c of [...row.children].filter(x => x.nodeName === 'c')) {
          const a = Eng.parseAddr(c.getAttribute('r')); if (!a) continue; const t = c.getAttribute('t'), sid = c.getAttribute('s');
          const fEl = c.getElementsByTagName('f')[0], vEl = c.getElementsByTagName('v')[0]; let text = '';
          const style = sid !== null && xfList[+sid] ? { ...xfList[+sid] } : null;
          if (fEl && fEl.textContent) text = '=' + fEl.textContent;
          else if (t === 's' && vEl) text = ss[+vEl.textContent] || '';
          else if (t === 'inlineStr') text = kids(c, 't').map(x => x.textContent).join('');
          else if (t === 'b' && vEl) text = vEl.textContent === '1' ? 'TRUE' : 'FALSE';
          else if (t === 'e' && vEl) text = vEl.textContent;
          else if (t === 'str' && vEl) text = vEl.textContent;
          else if (vEl) text = vEl.textContent;
          if (text !== '' && !/^[=]/.test(text) && (t === 's' || t === 'inlineStr' || t === 'str') && Eng.parseLiteral(text).v !== text) text = "'" + text;
          if (text === '' && !style) continue;
          wb.setInput(wb.sheets.length - 1, a.r, a.c, text);
          if (style) { wb.setStyle(wb.sheets.length - 1, a.r, a.c, style); }
        }
      }
    }
    if (!wb.sheets.length) wb.sheets.push(new Eng.Sheet('Sheet1'));
    wb.invalidate(); return wb;
  }

  /* ---------- CSV ---------- */
  function parseCsv(text) {
    const rows = []; let row = [], f = '', q = false; const delim = (text.split('\n')[0].match(/;/g) || []).length > (text.split('\n')[0].match(/,/g) || []).length ? ';' : (text.split('\n')[0].includes('\t') ? '\t' : ',');
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true; else if (ch === delim) { row.push(f); f = ''; } else if (ch === '\n') { row.push(f.replace(/\r$/, '')); rows.push(row); row = []; f = ''; } else f += ch;
    }
    if (f !== '' || row.length) { row.push(f.replace(/\r$/, '')); rows.push(row); } return rows;
  }
  function importCsv(text, name) {
    const wb = new Eng.Workbook(); wb.sheets[0].name = (name || 'Sheet1').replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1';
    parseCsv(text.replace(/^﻿/, '')).forEach((r, i) => r.forEach((v, j) => { if (v !== '') wb.setInput(0, i, j, v); })); return wb;
  }
  function exportCsv(wb, si) {
    const u = wb.usedRange(si), out = [];
    for (let r = 0; r <= u.r2; r++) { const row = []; for (let c = 0; c <= u.c2; c++) { const s = wb.display(si, r, c); row.push(/[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s); } out.push(row.join(',')); }
    return new Blob(['﻿' + out.join('\r\n')], { type: 'text/csv' });
  }

  root.FmXlsx = { exportXlsx, importXlsx, importCsv, exportCsv, parseCsv, zipStore, unzip, crc32 };
})(typeof self !== 'undefined' ? self : this);
