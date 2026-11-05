/* FmOffice Docs — .docx export/import (HTML <-> WordprocessingML), no dependencies */
(function () {
  'use strict';
  const enc = new TextEncoder(), dec = new TextDecoder();
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const hex = c => { if (!c) return null; let m = /^#([0-9a-f]{6})$/i.exec(c); if (m) return m[1].toUpperCase(); m = /^#([0-9a-f]{3})$/i.exec(c); if (m) return m[1].split('').map(x => x + x).join('').toUpperCase(); m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c); if (m) return [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); return null; };
  const pxToHalfPt = px => Math.round(parseFloat(px) * 0.75 * 2);

  function exportDocx(root, opts) {
    opts = opts || {}; const rels = []; const media = []; let rid = 100, imgN = 0, hasNum = false;
    const addRel = (type, target, ext) => { const id = 'rId' + (++rid); rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${esc(target)}"${ext ? ' TargetMode="External"' : ''}/>`); return id; };
    const run = (text, st) => { const pr = (st.rf ? `<w:rFonts w:ascii="${esc(st.rf)}" w:hAnsi="${esc(st.rf)}" w:cs="${esc(st.rf)}"/>` : '') + (st.b ? '<w:b/>' : '') + (st.i ? '<w:i/>' : '') + (st.s ? '<w:strike/>' : '') + (st.color ? `<w:color w:val="${st.color}"/>` : '') + (st.sz ? `<w:sz w:val="${st.sz}"/><w:szCs w:val="${st.sz}"/>` : '') + (st.hl ? `<w:shd w:val="clear" w:color="auto" w:fill="${st.hl}"/>` : '') + (st.u ? '<w:u w:val="single"/>' : '') + (st.link ? '<w:rStyle w:val="Hyperlink"/>' : ''); return `<w:r>${pr ? `<w:rPr>${pr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`; };
    const imgXml = el => {
      const src = el.getAttribute('src') || ''; const d = Fm.dataUrlToBytes(src); if (!d) return ''; const ext = /png/.test(d.mime) ? 'png' : /gif/.test(d.mime) ? 'gif' : 'jpeg'; imgN++; const name = `image${imgN}.${ext}`; media.push({ name: 'word/media/' + name, data: d.bytes }); const id = addRel('image', 'media/' + name);
      const w = Math.round((el.width || el.naturalWidth || 300) * 9525), h = Math.round((el.height || el.naturalHeight || 200) * 9525);
      return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${w}" cy="${h}"/><wp:docPr id="${imgN}" name="Picture ${imgN}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${imgN}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    };
    const inline = (node, st) => {
      let out = '';
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) { const t = ch.nodeValue.replace(/[ ]/g, ' '); if (t) out += run(t, st); continue; }
        if (ch.nodeType !== 1) continue; const tag = ch.tagName.toLowerCase(); const ns = { ...st }; const cs = ch.style || {};
        if (tag === 'br') { out += '<w:r><w:br/></w:r>'; continue; }
        if (tag === 'img') { out += imgXml(ch); continue; }
        if (tag === 'b' || tag === 'strong') ns.b = 1; if (tag === 'i' || tag === 'em') ns.i = 1; if (tag === 'u') ns.u = 1; if (tag === 's' || tag === 'strike' || tag === 'del') ns.s = 1;
        if (cs.fontWeight === 'bold' || +cs.fontWeight >= 600) ns.b = 1; if (cs.fontStyle === 'italic') ns.i = 1; if ((cs.textDecoration || '').includes('underline') || (cs.textDecorationLine || '').includes('underline')) ns.u = 1; if ((cs.textDecoration || '').includes('line-through')) ns.s = 1;
        if (cs.color) ns.color = hex(cs.color); if (tag === 'font' && ch.getAttribute('color')) ns.color = hex(ch.getAttribute('color')); if (cs.backgroundColor && hex(cs.backgroundColor)) ns.hl = hex(cs.backgroundColor); if (cs.fontSize && /px$/.test(cs.fontSize)) ns.sz = pxToHalfPt(cs.fontSize); if (cs.fontFamily) ns.rf = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
        if (tag === 'a' && /^(https?:|mailto:)/i.test(ch.getAttribute('href') || '')) { const id = addRel('hyperlink', ch.getAttribute('href'), true); out += `<w:hyperlink r:id="${id}">${inline(ch, { ...ns, link: 1, u: 1, color: '2B6FD6' })}</w:hyperlink>`; continue; }
        out += inline(ch, ns);
      }
      return out;
    };
    const jc = el => { const a = (el.style && el.style.textAlign) || el.getAttribute('align') || ''; return { center: 'center', right: 'right', justify: 'both', left: 'left' }[a] ? `<w:jc w:val="${{ center: 'center', right: 'right', justify: 'both', left: 'left' }[a]}"/>` : ''; };
    const para = (el, extra, st0) => {
      const tag = el.tagName.toLowerCase(); let style = ''; const h = /^h([1-6])$/.exec(tag); if (h) style = `<w:pStyle w:val="Heading${h[1]}"/>`; else if (tag === 'blockquote') style = '<w:pStyle w:val="Quote"/>';
      const lh = el.style && parseFloat(el.style.lineHeight); const sp = lh ? `<w:spacing w:line="${Math.round(lh * 240)}" w:lineRule="auto"/>` : ''; const ml = el.style && parseFloat(el.style.marginLeft); const ind = ml ? `<w:ind w:left="${Math.round(ml * 15)}"/>` : '';
      return `<w:p><w:pPr>${style}${extra || ''}${sp}${ind}${jc(el)}</w:pPr>${inline(el, st0 || {})}</w:p>`;
    };
    const list = (el, lvl, body) => {
      const ordered = el.tagName.toLowerCase() === 'ol'; hasNum = true;
      for (const li of el.children) { if (li.tagName.toLowerCase() !== 'li') continue; const clone = li.cloneNode(true); const subs = [...clone.children].filter(c => /^(ul|ol)$/i.test(c.tagName)); subs.forEach(s => s.remove()); body.push(`<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="${ordered ? 2 : 1}"/></w:numPr>${jc(li)}</w:pPr>${inline(clone, {})}</w:p>`); subs.forEach(s => list(s, lvl + 1, body)); }
    };
    const table = el => {
      let x = '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b => `<w:${b} w:val="single" w:sz="4" w:space="0" w:color="999999"/>`).join('') + '</w:tblBorders></w:tblPr>';
      for (const tr of el.querySelectorAll('tr')) { x += '<w:tr>'; for (const td of tr.children) { const hdr = td.tagName.toLowerCase() === 'th'; const bg = td.style && hex(td.style.backgroundColor); x += `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>${bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ''}</w:tcPr><w:p><w:pPr>${jc(td)}</w:pPr>${inline(td, hdr ? { b: 1 } : {})}</w:p></w:tc>`; } x += '</w:tr>'; }
      return x + '</w:tbl><w:p/>';
    };
    const body = [];
    const walk = parent => {
      for (const el of parent.childNodes) {
        if (el.nodeType === 3) { if (el.nodeValue.trim()) body.push(`<w:p>${run(el.nodeValue, {})}</w:p>`); continue; } if (el.nodeType !== 1) continue; const tag = el.tagName.toLowerCase();
        if (/^(p|h[1-6]|blockquote|pre)$/.test(tag)) body.push(para(el, '', tag === 'pre' ? { rf: 'Consolas' } : {}));
        else if (tag === 'div') { if (el.classList.contains('pb')) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>'); else if ([...el.children].some(c => /^(p|div|h[1-6]|ul|ol|table|blockquote)$/i.test(c.tagName))) walk(el); else body.push(para(el)); }
        else if (tag === 'ul' || tag === 'ol') list(el, 0, body); else if (tag === 'table') body.push(table(el));
        else if (tag === 'hr') body.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="999999"/></w:pBdr></w:pPr></w:p>');
        else if (tag === 'img') body.push(`<w:p>${imgXml(el)}</w:p>`); else if (tag === 'br') body.push('<w:p/>');
        else body.push(`<w:p>${inline({ childNodes: [el] }, {})}</w:p>`);
      }
    };
    walk(root);
    const pg = opts.page === 'letter' ? { w: 12240, h: 15840 } : { w: 11906, h: 16838 }; const mg = Math.round((opts.margin || 1) * 1440);
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="${pg.w}" w:h="${pg.h}"/><w:pgMar w:top="${mg}" w:right="${mg}" w:bottom="${mg}" w:left="${mg}" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const hs = (n, sz, col) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="${n === 1 ? 360 : 240}" w:after="80"/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:color w:val="${col}"/><w:sz w:val="${sz}"/></w:rPr></w:style>`;
    const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${hs(1, 40, '1F8A4C')}${hs(2, 32, '1F8A4C')}${hs(3, 26, '2F5D43')}${hs(4, 24, '2F5D43')}${hs(5, 22, '2F5D43')}${hs(6, 22, '2F5D43')}<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:ind w:left="720"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="2B6FD6"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`;
    const lvls = (fmt, txt) => Array.from({ length: 9 }, (_, i) => `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt === 'bullet' ? 'bullet' : (['decimal', 'lowerLetter', 'lowerRoman'][i % 3])}"/><w:lvlText w:val="${fmt === 'bullet' ? (['•', '◦', '▪'][i % 3]) : '%' + (i + 1) + '.'}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + i * 360}" w:hanging="360"/></w:pPr>${fmt === 'bullet' ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' : ''}</w:lvl>`).join('');
    const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvls('bullet')}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvls('decimal')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`;
    const files = [
      { name: '[Content_Types].xml', data: enc.encode(ct) },
      { name: '_rels/.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>') },
      { name: 'word/document.xml', data: enc.encode(doc) }, { name: 'word/styles.xml', data: enc.encode(styles) }, { name: 'word/numbering.xml', data: enc.encode(numbering) },
      { name: 'word/_rels/document.xml.rels', data: enc.encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' + rels.join('') + '</Relationships>') },
      ...media.map(m => ({ name: m.name, data: m.data }))
    ];
    return FmXlsx.zipStore(files, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  }

  /* ---------- import ---------- */
  async function importDocx(buf) {
    const z = await FmXlsx.unzip(buf); const get = async n => z[n] ? dec.decode(await z[n]()) : null; const xml = await get('word/document.xml'); if (!xml) throw new Error('Not a valid .docx file');
    const P_ = s => new DOMParser().parseFromString(s, 'application/xml'); const doc = P_(xml); const W = el => el ? el.localName : '';
    const rels = {}; const rx = await get('word/_rels/document.xml.rels'); if (rx) [...P_(rx).getElementsByTagName('Relationship')].forEach(r => rels[r.getAttribute('Id')] = { type: r.getAttribute('Type').split('/').pop(), target: r.getAttribute('Target'), ext: r.getAttribute('TargetMode') === 'External' });
    const numFmt = {}; const nx = await get('word/numbering.xml'); if (nx) { const nd = P_(nx); const abs = {}; [...nd.getElementsByTagName('*')].filter(e => W(e) === 'abstractNum').forEach(a => { abs[a.getAttribute('w:abstractNumId')] = {}; [...a.children].filter(l => W(l) === 'lvl').forEach(l => { const f = [...l.children].find(c => W(c) === 'numFmt'); abs[a.getAttribute('w:abstractNumId')][l.getAttribute('w:ilvl')] = f ? f.getAttribute('w:val') : 'decimal'; }); }); [...nd.getElementsByTagName('*')].filter(e => W(e) === 'num').forEach(n => { const a = [...n.children].find(c => W(c) === 'abstractNumId'); if (a) numFmt[n.getAttribute('w:numId')] = abs[a.getAttribute('w:val')] || {}; }); }
    const q = (el, name) => [...el.children].find(c => W(c) === name); const val = (el, name) => { const c = q(el, name); return c ? c.getAttribute('w:val') : null; };
    const imgs = {}; for (const id in rels) if (rels[id].type === 'image') { const p = 'word/' + rels[id].target.replace(/^\//, '').replace(/^word\//, ''); if (z[p]) { const b = await z[p](); const ext = p.split('.').pop().toLowerCase(); imgs[id] = Fm.bytesToDataUrl(b, ext === 'png' ? 'image/png' : ext === 'gif' ? 'image/gif' : ext === 'svg' ? 'image/svg+xml' : 'image/jpeg'); } }
    const runHtml = r => {
      const pr = q(r, 'rPr'); let css = ''; let open = '', close = '';
      if (pr) { if (q(pr, 'b') && val(pr, 'b') !== '0') { open += '<b>'; close = '</b>' + close; } if (q(pr, 'i') && val(pr, 'i') !== '0') { open += '<i>'; close = '</i>' + close; } if (q(pr, 'u') && val(pr, 'u') !== 'none') { open += '<u>'; close = '</u>' + close; } if (q(pr, 'strike')) { open += '<s>'; close = '</s>' + close; } const c = val(pr, 'color'); if (c && c !== 'auto') css += `color:#${c};`; const sz = val(pr, 'sz'); if (sz) css += `font-size:${(+sz / 2 / 0.75).toFixed(1)}px;`; const f = q(pr, 'rFonts'); if (f && f.getAttribute('w:ascii')) css += `font-family:'${f.getAttribute('w:ascii')}';`; const hl = val(pr, 'highlight'); if (hl && hl !== 'none') css += `background:${{ yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', lightGray: '#d3d3d3' }[hl] || '#ffff00'};`; const sh = q(pr, 'shd'); if (sh && sh.getAttribute('w:fill') && sh.getAttribute('w:fill') !== 'auto') css += `background:#${sh.getAttribute('w:fill')};`; }
      let inner = ''; for (const c of r.children) { const n = W(c); if (n === 't') inner += esc(c.textContent); else if (n === 'tab') inner += '&emsp;'; else if (n === 'br') inner += c.getAttribute('w:type') === 'page' ? '</span><div class="pb"></div><span>' : '<br>'; else if (n === 'drawing' || n === 'pict') { const b = [...c.getElementsByTagName('*')].find(e => W(e) === 'blip'); const ext = [...c.getElementsByTagName('*')].find(e => W(e) === 'extent'); if (b) { const id = b.getAttribute('r:embed'); if (imgs[id]) inner += `<img src="${imgs[id]}"${ext ? ` style="width:${Math.round(ext.getAttribute('cx') / 9525)}px;max-width:100%"` : ''}>`; } } }
      if (!inner) return ''; return (css ? `<span style="${css}">` : '') + open + inner + close + (css ? '</span>' : '');
    };
    const inlineHtml = p => { let h = ''; for (const c of p.children) { const n = W(c); if (n === 'r') h += runHtml(c); else if (n === 'hyperlink') { const rid = c.getAttribute('r:id'); const t = rid && rels[rid] ? rels[rid].target : null; const inner = [...c.children].filter(x => W(x) === 'r').map(runHtml).join(''); h += t ? `<a href="${esc(t)}">${inner}</a>` : inner; } else if (n === 'sdt' || n === 'smartTag' || n === 'ins') h += inlineHtml(c); } return h; };
    const pHtml = p => {
      const ppr = q(p, 'pPr'); let tag = 'p', style = ''; let list = null;
      if (ppr) { const ps = val(ppr, 'pStyle') || ''; const m = /^Heading(\d)$/i.exec(ps) || /^heading (\d)$/i.exec(ps); if (m) tag = 'h' + Math.min(6, +m[1]); else if (/^Title$/i.test(ps)) { tag = 'h1'; style += 'text-align:center;'; } else if (/^Quote/i.test(ps)) tag = 'blockquote'; const jc = val(ppr, 'jc'); if (jc) style += `text-align:${{ both: 'justify', center: 'center', right: 'right', left: 'left', end: 'right' }[jc] || 'left'};`; const sp = q(ppr, 'spacing'); if (sp && sp.getAttribute('w:line') && sp.getAttribute('w:lineRule') !== 'exact') style += `line-height:${(sp.getAttribute('w:line') / 240).toFixed(2)};`; const np = q(ppr, 'numPr'); if (np) { const id = val(np, 'numId'), lv = val(np, 'ilvl') || '0'; if (id && id !== '0') list = { id, lv: +lv, fmt: (numFmt[id] || {})[lv] || 'bullet' }; } const ind = q(ppr, 'ind'); if (ind && ind.getAttribute('w:left') && !list) { const l = +ind.getAttribute('w:left') / 15; if (l > 0) style += `margin-left:${Math.round(l)}px;`; } }
      let inner = inlineHtml(p); const hasPb = inner.includes('class="pb"'); if (!inner.replace(/<[^>]*>/g, '').trim() && !/<img/.test(inner) && !hasPb) inner = '<br>';
      return { tag, style, inner, list };
    };
    const tblHtml = t => { let h = '<table>'; for (const tr of [...t.children].filter(c => W(c) === 'tr')) { h += '<tr>'; for (const tc of [...tr.children].filter(c => W(c) === 'tc')) { const pr = q(tc, 'tcPr'); const sh = pr && q(pr, 'shd'); const bg = sh && sh.getAttribute('w:fill') && sh.getAttribute('w:fill') !== 'auto' ? ` style="background:#${sh.getAttribute('w:fill')}"` : ''; const cells = [...tc.children].filter(c => W(c) === 'p').map(p => { const o = pHtml(p); return o.inner; }).join('<br>'); h += `<td${bg}>${cells}</td>`; } h += '</tr>'; } return h + '</table><p><br></p>'; };
    const body = [...doc.getElementsByTagName('*')].find(e => W(e) === 'body'); let html = ''; let curList = null;
    const closeList = () => { if (curList) { html += `</${curList.t}>`.repeat(1); curList = null; } };
    for (const el of body.children) {
      const n = W(el);
      if (n === 'p') { const o = pHtml(el); if (o.list) { const t = /bullet/.test(o.list.fmt) ? 'ul' : 'ol'; if (!curList || curList.t !== t || curList.id !== o.list.id) { closeList(); curList = { t, id: o.list.id }; html += `<${t}>`; } html += `<li style="${o.list.lv ? 'margin-left:' + o.list.lv * 24 + 'px;' : ''}${o.style}">${o.inner}</li>`; } else { closeList(); html += `<${o.tag}${o.style ? ` style="${o.style}"` : ''}>${o.inner}</${o.tag}>`; } }
      else if (n === 'tbl') { closeList(); html += tblHtml(el); }
    }
    closeList(); return html || '<p><br></p>';
  }
  window.FmDocx = { exportDocx, importDocx };
})();
