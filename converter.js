// ═══════════════════════════════════════════════════════════════
// AAFCO DOCX 格式轉換引擎 v7 - 完整版（含目錄重建 + 書籤 + 封面修正）
// ═══════════════════════════════════════════════════════════════

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';

// ── XML 工具 ────────────────────────────────────────────────────
function parseXML(str) { return new DOMParser().parseFromString(str, 'application/xml'); }
function serXML(doc) { return new XMLSerializer().serializeToString(doc); }

function wEl(doc, name) { return doc.createElementNS(W_NS, 'w:' + name); }
function wSet(el, attr, val) { el.setAttributeNS(W_NS, 'w:' + attr, String(val)); }
function rSet(el, attr, val) { el.setAttributeNS(R_NS, 'r:' + attr, String(val)); }
function xmlSet(el, attr, val) { el.setAttributeNS(XML_NS, 'xml:' + attr, val); }

function clearEl(el) { while (el.firstChild) el.removeChild(el.firstChild); }

function getChildW(parent, localName) {
  for (const c of parent.childNodes)
    if (c.localName === localName && c.namespaceURI === W_NS) return c;
  return null;
}
function getAllChildW(parent, localName) {
  return [...parent.childNodes].filter(c => c.localName === localName && c.namespaceURI === W_NS);
}
function goc(parent, localName) {
  let el = getChildW(parent, localName);
  if (!el) { el = wEl(parent.ownerDocument, localName); parent.appendChild(el); }
  return el;
}
function rm(parent, localName) {
  getAllChildW(parent, localName).forEach(c => parent.removeChild(c));
}
function rmAttr(el, attr) { el.removeAttributeNS(W_NS, attr); }

function ensurePPr(para) {
  let pPr = getChildW(para, 'pPr');
  if (!pPr) { pPr = wEl(para.ownerDocument, 'pPr'); para.insertBefore(pPr, para.firstChild); }
  return pPr;
}
function ensureRPr(run) {
  let rPr = getChildW(run, 'rPr');
  if (!rPr) { rPr = wEl(run.ownerDocument, 'rPr'); run.insertBefore(rPr, run.firstChild); }
  return rPr;
}
function getStyle(para) {
  const pPr = getChildW(para, 'pPr');
  if (!pPr) return 'Normal';
  const ps = getChildW(pPr, 'pStyle');
  return ps ? (ps.getAttributeNS(W_NS, 'val') || 'Normal') : 'Normal';
}
function getText(para) {
  let t = '';
  const ts = para.getElementsByTagNameNS(W_NS, 't');
  for (let i = 0; i < ts.length; i++) t += ts[i].textContent || '';
  return t.trim();
}
function directRuns(el) {
  return [...el.childNodes].filter(c => c.localName === 'r' && c.namespaceURI === W_NS);
}

// ── 基礎格式工具 ────────────────────────────────────────────────
function makeRFonts(doc, ascii, hAnsi, eastAsia, cs, hint) {
  const rF = wEl(doc, 'rFonts');
  if (ascii) rF.setAttributeNS(W_NS, 'w:ascii', ascii);
  if (hAnsi || ascii) rF.setAttributeNS(W_NS, 'w:hAnsi', hAnsi || ascii);
  if (eastAsia) rF.setAttributeNS(W_NS, 'w:eastAsia', eastAsia);
  if (cs) rF.setAttributeNS(W_NS, 'w:cs', cs);
  if (hint) rF.setAttributeNS(W_NS, 'w:hint', hint);
  return rF;
}
function makeColor(doc, val, themeColor) {
  const c = wEl(doc, 'color'); c.setAttributeNS(W_NS, 'w:val', val);
  if (themeColor) c.setAttributeNS(W_NS, 'w:themeColor', themeColor);
  return c;
}
function makeSz(doc, val) {
  const s = wEl(doc, 'sz'); s.setAttributeNS(W_NS, 'w:val', String(val)); return s;
}
function makeSzCs(doc, val) {
  const s = wEl(doc, 'szCs'); s.setAttributeNS(W_NS, 'w:val', String(val)); return s;
}
function makeT(doc, text) {
  const t = wEl(doc, 't');
  if (text && (text.startsWith(' ') || text.endsWith(' '))) xmlSet(t, 'space', 'preserve');
  t.textContent = text || '';
  return t;
}
function makeRun(doc, { text, rStyle, ascii, eastAsia, cs, hint, noProof, webHidden, color, themeColor, sz, szCs, bold, kern, tab, fldChar, instrText } = {}) {
  const r = wEl(doc, 'r');
  const rPr = wEl(doc, 'rPr'); r.appendChild(rPr);
  if (rStyle) { const rs = wEl(doc, 'rStyle'); rs.setAttributeNS(W_NS, 'w:val', rStyle); rPr.appendChild(rs); }
  if (ascii || hint || cs || eastAsia) rPr.appendChild(makeRFonts(doc, ascii, ascii, eastAsia, cs, hint));
  if (bold) rPr.appendChild(wEl(doc, 'b'));
  if (noProof) rPr.appendChild(wEl(doc, 'noProof'));
  if (webHidden) rPr.appendChild(wEl(doc, 'webHidden'));
  if (color) rPr.appendChild(makeColor(doc, color, themeColor));
  if (sz) rPr.appendChild(makeSz(doc, sz));
  if (szCs) rPr.appendChild(makeSzCs(doc, szCs));
  if (kern !== undefined) { const k = wEl(doc, 'kern'); k.setAttributeNS(W_NS, 'w:val', String(kern)); rPr.appendChild(k); }
  if (tab) r.appendChild(wEl(doc, 'tab'));
  if (fldChar) { const fc = wEl(doc, 'fldChar'); fc.setAttributeNS(W_NS, 'w:fldCharType', fldChar); r.appendChild(fc); }
  if (instrText !== undefined) {
    const it = wEl(doc, 'instrText'); xmlSet(it, 'space', 'preserve');
    it.textContent = instrText; r.appendChild(it);
  }
  if (text !== undefined) r.appendChild(makeT(doc, text));
  return r;
}

// ── 段落格式函式 ────────────────────────────────────────────────

// 修正：封面段落 — 統一 sz=28, szCs=28, 無粗體
function applyCover(para) {
  const doc = para.ownerDocument;
  const pPr = ensurePPr(para); rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing');
  wSet(sp, 'line', '400'); wSet(sp, 'lineRule', 'exact');
  rmAttr(sp, 'before'); rmAttr(sp, 'after');
  const ind = goc(pPr, 'ind'); wSet(ind, 'firstLine', '0');
  rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  const jc = goc(pPr, 'jc'); wSet(jc, 'val', 'center');
  const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));

  // Process ALL runs including those inside hyperlinks
  const allRuns = [...para.getElementsByTagNameNS(W_NS, 'r')];
  allRuns.forEach(run => {
    // Only direct children of para or hyperlinks in para
    const rPr = ensureRPr(run);
    clearEl(rPr);
    // sz=28, szCs=28, color=000000+themeColor=text1, no bold
    rPr.appendChild(makeColor(doc, '000000', 'text1'));
    rPr.appendChild(makeSz(doc, '28'));
    rPr.appendChild(makeSzCs(doc, '28'));
    // Keep original font by NOT adding rFonts (inherit from style)
    // but DO ensure no bold
  });
}

function applyH1(para) {
  const doc = para.ownerDocument;
  const pPr = ensurePPr(para);
  const ps = goc(pPr, 'pStyle'); wSet(ps, 'val', '1');
  const sp = goc(pPr, 'spacing');
  wSet(sp, 'line', '480'); wSet(sp, 'lineRule', 'auto');
  rmAttr(sp, 'before'); rmAttr(sp, 'after');
  const ind = goc(pPr, 'ind'); wSet(ind, 'firstLine', '0');
  rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  rm(pPr, 'jc');
  const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
  pRPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));
  pRPr.appendChild(makeSz(doc, '28'));

  const allRuns = [...para.getElementsByTagNameNS(W_NS, 'r')];
  allRuns.forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    rPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
    rPr.appendChild(makeColor(doc, '000000', 'text1'));
    rPr.appendChild(makeSz(doc, '28'));
  });
}

function applyH2(para) {
  const doc = para.ownerDocument;
  const pPr = ensurePPr(para);
  const ps = goc(pPr, 'pStyle'); wSet(ps, 'val', '2');
  rm(pPr, 'spacing'); rm(pPr, 'ind'); rm(pPr, 'numPr');

  const numPr = wEl(doc, 'numPr');
  const ilvl = wEl(doc, 'ilvl'); wSet(ilvl, 'val', '0'); numPr.appendChild(ilvl);
  const numId = wEl(doc, 'numId'); wSet(numId, 'val', '3'); numPr.appendChild(numId);
  const siblings = [...pPr.childNodes];
  const psIdx = siblings.indexOf(ps);
  if (psIdx >= 0 && psIdx + 1 < siblings.length) pPr.insertBefore(numPr, siblings[psIdx + 1]);
  else pPr.appendChild(numPr);

  const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
  pRPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  const b0 = wEl(doc, 'b'); b0.setAttributeNS(W_NS, 'w:val', '0'); pRPr.appendChild(b0);
  const bc0 = wEl(doc, 'bCs'); bc0.setAttributeNS(W_NS, 'w:val', '0'); pRPr.appendChild(bc0);
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));
  pRPr.appendChild(makeSz(doc, '24')); pRPr.appendChild(makeSzCs(doc, '24'));

  const allRuns = [...para.getElementsByTagNameNS(W_NS, 'r')];
  allRuns.forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    rPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman', null, null, 'eastAsia'));
    const b = wEl(doc, 'b'); b.setAttributeNS(W_NS, 'w:val', '0'); rPr.appendChild(b);
    const bc = wEl(doc, 'bCs'); bc.setAttributeNS(W_NS, 'w:val', '0'); rPr.appendChild(bc);
    rPr.appendChild(makeColor(doc, '000000', 'text1'));
    rPr.appendChild(makeSz(doc, '24')); rPr.appendChild(makeSzCs(doc, '24'));
  });
}

function applyBody(para) {
  const doc = para.ownerDocument;
  const pPr = ensurePPr(para); rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing');
  wSet(sp, 'before', '80'); wSet(sp, 'after', '80');
  rmAttr(sp, 'line'); rmAttr(sp, 'lineRule');
  const ind = goc(pPr, 'ind'); wSet(ind, 'firstLine', '480');
  rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  const jc = goc(pPr, 'jc'); wSet(jc, 'val', 'both');
  const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
  pRPr.appendChild(makeColor(doc, 'auto'));

  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    rPr.appendChild(makeRFonts(doc, null, null, null, null, 'eastAsia'));
  });
}

function applyEmpty(para) {
  const pPr = ensurePPr(para); rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing');
  wSet(sp, 'line', '480'); wSet(sp, 'lineRule', 'auto');
}

// ── 表格格式 ────────────────────────────────────────────────────
function applyTable(tbl) {
  const doc = tbl.ownerDocument;
  let tblPr = getChildW(tbl, 'tblPr');
  if (!tblPr) { tblPr = wEl(doc, 'tblPr'); tbl.insertBefore(tblPr, tbl.firstChild); }
  rm(tblPr, 'tblBorders'); rm(tblPr, 'tblStyle'); rm(tblPr, 'tblLook');
  const tblW = goc(tblPr, 'tblW'); wSet(tblW, 'w', '9355'); wSet(tblW, 'type', 'dxa');
  const jcT = goc(tblPr, 'jc'); wSet(jcT, 'val', 'center');
  const tcm = goc(tblPr, 'tblCellMar');
  const ml = goc(tcm, 'left'); wSet(ml, 'w', '10'); wSet(ml, 'type', 'dxa');
  const mr = goc(tcm, 'right'); wSet(mr, 'w', '10'); wSet(mr, 'type', 'dxa');
  const tblLook = wEl(doc, 'tblLook');
  ['val','04A0','firstRow','1','lastRow','0','firstColumn','1','lastColumn','0','noHBand','0','noVBand','1']
    .reduce((a, v, i) => { if (i % 2) tblLook.setAttributeNS(W_NS, 'w:' + a, v); return v; });
  // Fix tblLook attributes properly
  tblLook.setAttributeNS(W_NS, 'w:val', '04A0');
  tblLook.setAttributeNS(W_NS, 'w:firstRow', '1'); tblLook.setAttributeNS(W_NS, 'w:lastRow', '0');
  tblLook.setAttributeNS(W_NS, 'w:firstColumn', '1'); tblLook.setAttributeNS(W_NS, 'w:lastColumn', '0');
  tblLook.setAttributeNS(W_NS, 'w:noHBand', '0'); tblLook.setAttributeNS(W_NS, 'w:noVBand', '1');
  tblPr.appendChild(tblLook);

  const rows = getAllChildW(tbl, 'tr');
  rows.forEach((row, ri) => {
    const isHeader = ri === 0;
    let trPr = getChildW(row, 'trPr');
    if (!trPr) { trPr = wEl(doc, 'trPr'); row.insertBefore(trPr, row.firstChild); }
    if (isHeader) { const jcR = goc(trPr, 'jc'); wSet(jcR, 'val', 'center'); }

    getAllChildW(row, 'tc').forEach(cell => {
      let tcPr = getChildW(cell, 'tcPr');
      if (!tcPr) { tcPr = wEl(doc, 'tcPr'); cell.insertBefore(tcPr, cell.firstChild); }
      rm(tcPr, 'tcBorders'); rm(tcPr, 'hideMark');
      const tcBorders = wEl(doc, 'tcBorders');
      const bEl = wEl(doc, isHeader ? 'bottom' : 'top');
      bEl.setAttributeNS(W_NS, 'w:val', 'single'); bEl.setAttributeNS(W_NS, 'w:sz', '4');
      bEl.setAttributeNS(W_NS, 'w:space', '0'); bEl.setAttributeNS(W_NS, 'w:color', 'auto');
      tcBorders.appendChild(bEl);
      const tcWel = getChildW(tcPr, 'tcW');
      if (tcWel) tcPr.insertBefore(tcBorders, tcWel.nextSibling); else tcPr.insertBefore(tcBorders, tcPr.firstChild);
      const shd = goc(tcPr, 'shd');
      shd.setAttributeNS(W_NS, 'w:val', 'clear'); shd.setAttributeNS(W_NS, 'w:color', 'auto'); shd.setAttributeNS(W_NS, 'w:fill', 'auto');
      const tcMar = goc(tcPr, 'tcMar');
      [['top','80'],['left','120'],['bottom','80'],['right','120']].forEach(([s,v]) => { const m = goc(tcMar, s); wSet(m,'w',v); wSet(m,'type','dxa'); });
      // hideMark on ALL rows/cells
      tcPr.appendChild(wEl(doc, 'hideMark'));

      getAllChildW(cell, 'p').forEach(para => {
        const pPr = ensurePPr(para); rm(pPr, 'pStyle'); rm(pPr, 'jc');
        const wc = goc(pPr, 'widowControl'); wSet(wc, 'val', '0');
        const sp = goc(pPr, 'spacing'); wSet(sp,'before','60'); wSet(sp,'after','60'); wSet(sp,'line','240'); wSet(sp,'lineRule','auto');
        const ind = goc(pPr, 'ind'); wSet(ind, 'firstLine', '0');
        const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
        pRPr.appendChild(makeRFonts(doc, '標楷體', '標楷體', null, '標楷體'));
        pRPr.appendChild(makeColor(doc, 'auto'));
        const kern = wEl(doc, 'kern'); kern.setAttributeNS(W_NS, 'w:val', '0'); pRPr.appendChild(kern);

        directRuns(para).forEach(run => {
          const rPr = ensureRPr(run); clearEl(rPr);
          rPr.appendChild(makeRFonts(doc, '標楷體', '標楷體', null, isHeader ? '新細明體' : '標楷體', 'eastAsia'));
          if (isHeader) rPr.appendChild(wEl(doc, 'bCs'));
          rPr.appendChild(makeColor(doc, 'auto'));
          const k = wEl(doc, 'kern'); k.setAttributeNS(W_NS, 'w:val', '0'); rPr.appendChild(k);
        });
      });
    });
  });
}

// ═══════════════════════════════════════════════════════════════
// 目錄重建系統
// ═══════════════════════════════════════════════════════════════

function extractHeadings(body) {
  const children = [...body.childNodes].filter(n => n.nodeType === 1);
  const headings = [];
  const tables = [];
  let hCounter = 1;

  children.forEach((el, i) => {
    if (el.localName === 'p' && el.namespaceURI === W_NS) {
      const style = getStyle(el);
      const text = getText(el);
      if (!text) return;
      if (style === '1') {
        headings.push({ level: 1, text, anchor: `_AutoH1${String(hCounter++).padStart(9,'0')}`, idx: i });
      } else if (style === '2') {
        headings.push({ level: 2, text, anchor: `_AutoH2${String(hCounter++).padStart(9,'0')}`, idx: i });
      }
    } else if (el.localName === 'tbl' && el.namespaceURI === W_NS) {
      // Find caption: paragraph before table with '表' in text
      for (let j = i - 1; j >= Math.max(0, i - 3); j--) {
        const prev = children[j];
        if (prev && prev.localName === 'p' && prev.namespaceURI === W_NS) {
          const t = getText(prev);
          if (t && t.includes('表') && t.length > 2) {
            tables.push({ title: t, anchor: `_AutoTbl${String(hCounter++).padStart(9,'0')}`, idx: j });
            break;
          }
        }
      }
    }
  });
  return { headings, tables };
}

function addBookmarks(body, headings, tables) {
  const doc = body.ownerDocument;
  const children = [...body.childNodes].filter(n => n.nodeType === 1);
  let bkId = 10;

  [...headings, ...tables].forEach(item => {
    const el = children[item.idx];
    if (!el || el.localName !== 'p') return;
    // Remove old bookmarks
    getAllChildW(el, 'bookmarkStart').forEach(b => el.removeChild(b));
    getAllChildW(el, 'bookmarkEnd').forEach(b => el.removeChild(b));
    // Insert after pPr
    const pPr = getChildW(el, 'pPr');
    const bkStart = wEl(doc, 'bookmarkStart');
    bkStart.setAttributeNS(W_NS, 'w:id', String(bkId));
    bkStart.setAttributeNS(W_NS, 'w:name', item.anchor);
    const bkEnd = wEl(doc, 'bookmarkEnd');
    bkEnd.setAttributeNS(W_NS, 'w:id', String(bkId));
    bkId++;
    if (pPr && pPr.nextSibling) {
      el.insertBefore(bkStart, pPr.nextSibling);
      el.insertBefore(bkEnd, pPr.nextSibling);
    } else {
      el.appendChild(bkStart); el.appendChild(bkEnd);
    }
  });
}

function makeHyperlink(doc, anchor) {
  const hl = wEl(doc, 'hyperlink');
  hl.setAttributeNS(W_NS, 'w:anchor', anchor);
  hl.setAttributeNS(W_NS, 'w:history', '1');
  return hl;
}

function makePageRefRuns(doc, anchor, pageText) {
  return [
    makeRun(doc, { noProof: true, webHidden: true, tab: true }),
    makeRun(doc, { noProof: true, webHidden: true, fldChar: 'begin' }),
    makeRun(doc, { noProof: true, webHidden: true, instrText: ` PAGEREF ${anchor} \\h ` }),
    (() => { const r = wEl(doc, 'r'); r.appendChild(wEl(doc, 'rPr')); return r; })(),
    makeRun(doc, { noProof: true, webHidden: true, fldChar: 'separate' }),
    makeRun(doc, { noProof: true, webHidden: true, text: String(pageText) }),
    makeRun(doc, { noProof: true, webHidden: true, fldChar: 'end' }),
  ];
}

function buildTOCHeading(doc, anchor, bookmarkAnchor) {
  // 「目錄」heading para (style=1)
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ps = wEl(doc, 'pStyle'); ps.setAttributeNS(W_NS, 'w:val', '1'); pPr.appendChild(ps);
  const sp = wEl(doc, 'spacing');
  sp.setAttributeNS(W_NS, 'w:before', '0'); sp.setAttributeNS(W_NS, 'w:after', '0');
  sp.setAttributeNS(W_NS, 'w:line', '480'); sp.setAttributeNS(W_NS, 'w:lineRule', 'auto');
  pPr.appendChild(sp);
  const ind = wEl(doc, 'ind'); ind.setAttributeNS(W_NS, 'w:firstLine', '0'); pPr.appendChild(ind);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  pRPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));
  pRPr.appendChild(makeSz(doc, '28')); pRPr.appendChild(makeSzCs(doc, '24'));

  // bookmark
  const bkS = wEl(doc, 'bookmarkStart'); bkS.setAttributeNS(W_NS, 'w:id', '2'); bkS.setAttributeNS(W_NS, 'w:name', bookmarkAnchor); p.appendChild(bkS);
  // run
  const r = makeRun(doc, { ascii: 'Times New Roman', color: '000000', themeColor: 'text1', sz: '28', szCs: '24', text: '目錄' }); p.appendChild(r);
  const bkE = wEl(doc, 'bookmarkEnd'); bkE.setAttributeNS(W_NS, 'w:id', '2'); p.appendChild(bkE);
  return p;
}

function buildPageNumberRow(doc) {
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ind = wEl(doc, 'ind'); ind.setAttributeNS(W_NS, 'w:firstLine', '0'); pPr.appendChild(ind);
  const jc = wEl(doc, 'jc'); jc.setAttributeNS(W_NS, 'w:val', 'right'); pPr.appendChild(jc);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));
  const r1 = makeRun(doc, { color: '000000', themeColor: 'text1', text: ' '.repeat(65) }); p.appendChild(r1);
  const r2 = makeRun(doc, { color: '000000', themeColor: 'text1', text: '頁數' }); p.appendChild(r2);
  return p;
}

function buildTOC11Para(doc, anchor, titleText, pageText, isFirst) {
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ps = wEl(doc, 'pStyle'); ps.setAttributeNS(W_NS, 'w:val', '11'); pPr.appendChild(ps);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  const rFt = makeRFonts(doc); // theme fonts
  rFt.setAttributeNS(W_NS, 'w:asciiTheme', 'minorHAnsi'); rFt.setAttributeNS(W_NS, 'w:eastAsiaTheme', 'minorEastAsia');
  rFt.setAttributeNS(W_NS, 'w:hAnsiTheme', 'minorHAnsi'); rFt.setAttributeNS(W_NS, 'w:cstheme', 'minorBidi');
  pRPr.appendChild(rFt); pRPr.appendChild(wEl(doc, 'noProof'));
  pRPr.appendChild(makeColor(doc, 'auto')); pRPr.appendChild(makeSz(doc, '24')); pRPr.appendChild(makeSzCs(doc, '22'));

  // TOC field begin on first item
  if (isFirst) {
    const makeFieldRun = (txt) => {
      const r = wEl(doc, 'r'); const rPr = wEl(doc, 'rPr'); r.appendChild(rPr);
      rPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman', null, 'Times New Roman'));
      rPr.appendChild(wEl(doc, 'b'));
      rPr.appendChild(makeColor(doc, '000000', 'text1'));
      rPr.appendChild(makeSz(doc, '24')); rPr.appendChild(makeSzCs(doc, '24'));
      if (txt.type === 'begin') { const fc = wEl(doc, 'fldChar'); fc.setAttributeNS(W_NS, 'w:fldCharType', 'begin'); r.appendChild(fc); }
      else if (txt.type === 'sep') { const fc = wEl(doc, 'fldChar'); fc.setAttributeNS(W_NS, 'w:fldCharType', 'separate'); r.appendChild(fc); }
      else { const it = wEl(doc, 'instrText'); xmlSet(it, 'space', 'preserve'); it.textContent = txt.text; r.appendChild(it); }
      return r;
    };
    p.appendChild(makeFieldRun({ type: 'begin' }));
    p.appendChild(makeFieldRun({ text: ' TOC \\o "1-3" \\h \\z \\u ' }));
    p.appendChild(makeFieldRun({ type: 'sep' }));
  }

  // Hyperlink with title text
  const hl = makeHyperlink(doc, anchor);
  // 保留原始標題文字（含全形空格\u3000），按空白分割
  const parts = titleText.split(/([\s\u3000]+)/);
  parts.forEach(part => {
    if (!part) return;
    const r = wEl(doc, 'r'); const rPr = wEl(doc, 'rPr'); r.appendChild(rPr);
    const rs = wEl(doc, 'rStyle'); rs.setAttributeNS(W_NS, 'w:val', 'ae'); rPr.appendChild(rs);
    const rF = makeRFonts(doc, 'Times New Roman', 'Times New Roman');
    if (/[\u4e00-\u9fff\u3000]/.test(part)) rF.setAttributeNS(W_NS, 'w:hint', 'eastAsia');
    rPr.appendChild(rF); rPr.appendChild(wEl(doc, 'noProof'));
    r.appendChild(makeT(doc, part)); hl.appendChild(r);
  });
  makePageRefRuns(doc, anchor, pageText).forEach(r => hl.appendChild(r));
  p.appendChild(hl);
  return p;
}

function buildTOC21Para(doc, anchor, sectionNum, titleText, pageText) {
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ps = wEl(doc, 'pStyle'); ps.setAttributeNS(W_NS, 'w:val', '21'); pPr.appendChild(ps);
  const tabs = wEl(doc, 'tabs'); const tab = wEl(doc, 'tab');
  tab.setAttributeNS(W_NS, 'w:val', 'left'); tab.setAttributeNS(W_NS, 'w:pos', '1985');
  tabs.appendChild(tab); pPr.appendChild(tabs);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  const rFt = wEl(doc, 'rFonts');
  rFt.setAttributeNS(W_NS, 'w:asciiTheme', 'minorHAnsi'); rFt.setAttributeNS(W_NS, 'w:eastAsiaTheme', 'minorEastAsia');
  rFt.setAttributeNS(W_NS, 'w:hAnsiTheme', 'minorHAnsi'); rFt.setAttributeNS(W_NS, 'w:cstheme', 'minorBidi');
  pRPr.appendChild(rFt); pRPr.appendChild(wEl(doc, 'noProof'));
  pRPr.appendChild(makeColor(doc, 'auto')); pRPr.appendChild(makeSz(doc, '24')); pRPr.appendChild(makeSzCs(doc, '22'));

  const hl = makeHyperlink(doc, anchor);

  // Section number run (e.g. "第一節")
  const rNum = wEl(doc, 'r'); const rPrN = wEl(doc, 'rPr'); rNum.appendChild(rPrN);
  const rsN = wEl(doc, 'rStyle'); rsN.setAttributeNS(W_NS, 'w:val', 'ae'); rPrN.appendChild(rsN);
  rPrN.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman', null, null, 'eastAsia'));
  rPrN.appendChild(wEl(doc, 'noProof')); rNum.appendChild(makeT(doc, sectionNum)); hl.appendChild(rNum);

  // Tab between number and title
  const rTab = wEl(doc, 'r'); const rPrTab = wEl(doc, 'rPr'); rTab.appendChild(rPrTab);
  const rFTab = wEl(doc, 'rFonts');
  rFTab.setAttributeNS(W_NS, 'w:asciiTheme', 'minorHAnsi'); rFTab.setAttributeNS(W_NS, 'w:hAnsiTheme', 'minorHAnsi');
  rFTab.setAttributeNS(W_NS, 'w:eastAsiaTheme', 'minorEastAsia'); rFTab.setAttributeNS(W_NS, 'w:cstheme', 'minorBidi');
  rPrTab.appendChild(rFTab); rPrTab.appendChild(wEl(doc, 'noProof'));
  rPrTab.appendChild(makeColor(doc, 'auto')); rPrTab.appendChild(makeSz(doc, '24')); rPrTab.appendChild(makeSzCs(doc, '22'));
  rTab.appendChild(wEl(doc, 'tab')); hl.appendChild(rTab);

  // Title run
  if (titleText) {
    const rTitle = wEl(doc, 'r'); const rPrT = wEl(doc, 'rPr'); rTitle.appendChild(rPrT);
    const rsT = wEl(doc, 'rStyle'); rsT.setAttributeNS(W_NS, 'w:val', 'ae'); rPrT.appendChild(rsT);
    rPrT.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman', null, null, 'eastAsia'));
    rPrT.appendChild(wEl(doc, 'noProof')); rTitle.appendChild(makeT(doc, titleText)); hl.appendChild(rTitle);
  }

  makePageRefRuns(doc, anchor, pageText).forEach(r => hl.appendChild(r));
  p.appendChild(hl);
  return p;
}

function buildTableListHeading(doc, bookmarkAnchor) {
  // 表目次 heading para (style=1 with page break)
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ps = wEl(doc, 'pStyle'); ps.setAttributeNS(W_NS, 'w:val', '1'); pPr.appendChild(ps);
  const sp = wEl(doc, 'spacing');
  sp.setAttributeNS(W_NS, 'w:before', '0'); sp.setAttributeNS(W_NS, 'w:after', '0');
  sp.setAttributeNS(W_NS, 'w:line', '240'); sp.setAttributeNS(W_NS, 'w:lineRule', 'auto');
  pPr.appendChild(sp);
  const ind = wEl(doc, 'ind'); ind.setAttributeNS(W_NS, 'w:firstLine', '0'); pPr.appendChild(ind);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  pRPr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  pRPr.appendChild(makeColor(doc, '000000', 'text1'));
  pRPr.appendChild(makeSz(doc, '28')); pRPr.appendChild(makeSzCs(doc, '28'));

  // TOC field end (close chapter TOC)
  const rEnd = wEl(doc, 'r'); const rPrEnd = wEl(doc, 'rPr'); rEnd.appendChild(rPrEnd);
  rPrEnd.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  rPrEnd.appendChild(makeColor(doc, '000000', 'text1'));
  const fcEnd = wEl(doc, 'fldChar'); fcEnd.setAttributeNS(W_NS, 'w:fldCharType', 'end'); rEnd.appendChild(fcEnd);
  p.appendChild(rEnd);

  // Page break
  const rBr = wEl(doc, 'r'); const rPrBr = wEl(doc, 'rPr'); rBr.appendChild(rPrBr);
  rPrBr.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  rPrBr.appendChild(makeColor(doc, '000000', 'text1'));
  const br = wEl(doc, 'br'); br.setAttributeNS(W_NS, 'w:type', 'page'); rBr.appendChild(br);
  p.appendChild(rBr);

  // Bookmark + 表目次 text
  const bkS = wEl(doc, 'bookmarkStart'); bkS.setAttributeNS(W_NS, 'w:id', '99'); bkS.setAttributeNS(W_NS, 'w:name', bookmarkAnchor); p.appendChild(bkS);
  const rText = wEl(doc, 'r'); const rPrTxt = wEl(doc, 'rPr'); rText.appendChild(rPrTxt);
  rPrTxt.appendChild(makeRFonts(doc, 'Times New Roman', 'Times New Roman'));
  rPrTxt.appendChild(makeColor(doc, '000000', 'text1'));
  rPrTxt.appendChild(makeSz(doc, '28')); rPrTxt.appendChild(makeSzCs(doc, '28'));
  rText.appendChild(makeT(doc, '表目次')); p.appendChild(rText);
  const bkE = wEl(doc, 'bookmarkEnd'); bkE.setAttributeNS(W_NS, 'w:id', '99'); p.appendChild(bkE);
  return p;
}

function buildTOCAF0Para(doc, anchor, tableTitle, pageText, isFirst) {
  const p = wEl(doc, 'p');
  const pPr = wEl(doc, 'pPr'); p.appendChild(pPr);
  const ps = wEl(doc, 'pStyle'); ps.setAttributeNS(W_NS, 'w:val', 'af0'); pPr.appendChild(ps);
  const pRPr = wEl(doc, 'rPr'); pPr.appendChild(pRPr);
  const rFt = wEl(doc, 'rFonts');
  rFt.setAttributeNS(W_NS, 'w:asciiTheme', 'minorHAnsi'); rFt.setAttributeNS(W_NS, 'w:hAnsiTheme', 'minorHAnsi');
  rFt.setAttributeNS(W_NS, 'w:eastAsiaTheme', 'minorEastAsia'); rFt.setAttributeNS(W_NS, 'w:cstheme', 'minorBidi');
  pRPr.appendChild(rFt); pRPr.appendChild(wEl(doc, 'noProof'));
  pRPr.appendChild(makeColor(doc, 'auto')); pRPr.appendChild(makeSz(doc, '24')); pRPr.appendChild(makeSzCs(doc, '22'));

  if (isFirst) {
    const mkCR = (type, text) => {
      const r = wEl(doc, 'r'); const rPr = wEl(doc, 'rPr'); r.appendChild(rPr);
      rPr.appendChild(makeColor(doc, '000000', 'text1'));
      if (type === 'begin') { const fc = wEl(doc, 'fldChar'); fc.setAttributeNS(W_NS, 'w:fldCharType', 'begin'); r.appendChild(fc); }
      else if (type === 'sep') { const fc = wEl(doc, 'fldChar'); fc.setAttributeNS(W_NS, 'w:fldCharType', 'separate'); r.appendChild(fc); }
      else { const it = wEl(doc, 'instrText'); xmlSet(it, 'space', 'preserve'); it.textContent = text; r.appendChild(it); }
      return r;
    };
    p.appendChild(mkCR('begin'));
    p.appendChild(mkCR('instr', ' TOC \\h \\z \\c "'));
    p.appendChild(mkCR('instr', '表'));
    p.appendChild(mkCR('instr', '" '));
    p.appendChild(mkCR('sep'));
  }

  const hl = makeHyperlink(doc, anchor);
  // Split table title into parts for proper hint marking
  tableTitle.split(/(\d+)/).forEach(part => {
    if (!part) return;
    const r = wEl(doc, 'r'); const rPr = wEl(doc, 'rPr'); r.appendChild(rPr);
    const rs = wEl(doc, 'rStyle'); rs.setAttributeNS(W_NS, 'w:val', 'ae'); rPr.appendChild(rs);
    const rF = wEl(doc, 'rFonts');
    if (!/^\d+$/.test(part)) rF.setAttributeNS(W_NS, 'w:hint', 'eastAsia');
    rPr.appendChild(rF); rPr.appendChild(wEl(doc, 'noProof'));
    r.appendChild(makeT(doc, part)); hl.appendChild(r);
  });
  makePageRefRuns(doc, anchor, pageText).forEach(r => hl.appendChild(r));
  p.appendChild(hl);

  // Close table TOC field on last item
  return p;
}

function parseH2Text(text) {
  // "第一節　報告說明" -> ["第一節", "報告說明"]
  const m = text.match(/^(第[一二三四五六七八九十百]+節|[一二三四五六七八九十]+、|[一二三四五六七八九十]+\.)\s*[\u3000\s]*(.*)/);
  if (m) return [m[1], m[2].trim()];
  const parts = text.split('\u3000');
  if (parts.length >= 2) return [parts[0].trim(), parts.slice(1).join('\u3000').trim()];
  return [text, ''];
}

function rebuildTOC(body, headings, tables) {
  const doc = body.ownerDocument;
  const children = [...body.childNodes].filter(n => n.nodeType === 1);

  // Find TOC area: from paragraph with text='目錄' to first H1 with '第X章'
  let tocHeadingEl = null, tocHeadingIdx = -1;
  let tocEndEl = null, tocEndIdx = -1;

  children.forEach((el, i) => {
    if (el.localName === 'p' && el.namespaceURI === W_NS) {
      const text = getText(el);
      if (text === '目錄' && tocHeadingEl === null) { tocHeadingEl = el; tocHeadingIdx = i; }
      if (tocHeadingEl && tocEndEl === null && getStyle(el) === '1' && text.includes('第') && text.includes('章')) {
        tocEndEl = el; tocEndIdx = i;
      }
    }
  });

  if (tocHeadingIdx < 0) {
    console.warn('TOC heading not found, inserting before first H1 chapter');
    // Find insertion point: before first H1 with 章
    const firstChapter = children.find(el => el.localName === 'p' && getStyle(el) === '1' && getText(el).includes('章'));
    if (!firstChapter) return;
    tocEndEl = firstChapter;
    tocEndIdx = children.indexOf(firstChapter);
    // No TOC paragraphs to remove
    tocHeadingIdx = tocEndIdx;
    tocHeadingEl = null;
  }

  // Remove old TOC paragraphs (between tocHeadingIdx and tocEndIdx exclusive)
  const toRemove = [];
  if (tocHeadingEl) toRemove.push(tocHeadingEl);
  for (let i = tocHeadingIdx + 1; i < tocEndIdx; i++) toRemove.push(children[i]);
  toRemove.forEach(el => { if (el.parentNode === body) body.removeChild(el); });

  // Build new TOC paragraphs
  const newParas = [];

  // 目錄 heading
  newParas.push(buildTOCHeading(doc, null, '_TocDir'));
  // 頁數 row
  newParas.push(buildPageNumberRow(doc));

  // Chapter and section entries
  let isFirst11 = true;
  headings.forEach(h => {
    if (h.level === 1) {
      // Parse: "第一章　研究背景與分析範疇" -> title
      const m = h.text.match(/^(第[一二三四五六七八九十百]+章)\s*[\u3000\s]*(.*)/);
      const displayText = m ? m[1] + ' ' + m[2].trim() : h.text;
      newParas.push(buildTOC11Para(doc, h.anchor, displayText, '1', isFirst11));
      isFirst11 = false;
    } else {
      const [secNum, title] = parseH2Text(h.text);
      newParas.push(buildTOC21Para(doc, h.anchor, secNum, title, '1'));
    }
  });

  // 表目次
  newParas.push(buildTableListHeading(doc, '_TocTbl'));
  tables.forEach((t, i) => {
    const af0 = buildTOCAF0Para(doc, t.anchor, t.title, '1', i === 0);
    newParas.push(af0);
  });
  // Close table TOC field
  if (tables.length > 0) {
    const lastAF = newParas[newParas.length - 1];
    const rClose = wEl(doc, 'r'); const rPrC = wEl(doc, 'rPr'); rClose.appendChild(rPrC);
    rPrC.appendChild(makeColor(doc, '000000', 'text1'));
    const fcClose = wEl(doc, 'fldChar'); fcClose.setAttributeNS(W_NS, 'w:fldCharType', 'end'); rClose.appendChild(fcClose);
    lastAF.appendChild(rClose);
  }

  // Insert before tocEndEl
  newParas.forEach(p => body.insertBefore(p, tocEndEl));
}

// ── sectPr ──────────────────────────────────────────────────────
function buildCoverSectPr(doc, footerRid) {
  const sp = wEl(doc, 'sectPr');
  if (footerRid) {
    const fr = wEl(doc, 'footerReference'); wSet(fr, 'type', 'default'); rSet(fr, 'id', footerRid); sp.appendChild(fr);
  }
  const pgSz = wEl(doc, 'pgSz'); wSet(pgSz,'w','11906'); wSet(pgSz,'h','16838'); sp.appendChild(pgSz);
  const pgMar = wEl(doc, 'pgMar');
  wSet(pgMar,'top','1418'); wSet(pgMar,'right','1418'); wSet(pgMar,'bottom','1418'); wSet(pgMar,'left','1701');
  wSet(pgMar,'header','851'); wSet(pgMar,'footer','992'); wSet(pgMar,'gutter','0'); sp.appendChild(pgMar);
  const pgNum = wEl(doc, 'pgNumType'); wSet(pgNum,'fmt','upperRoman'); wSet(pgNum,'start','1'); sp.appendChild(pgNum);
  const cols = wEl(doc, 'cols'); wSet(cols,'space','425'); sp.appendChild(cols);
  const dg = wEl(doc, 'docGrid'); wSet(dg,'type','lines'); wSet(dg,'linePitch','360'); sp.appendChild(dg);
  return sp;
}

function updateMainSectPr(sectPr) {
  clearEl(sectPr);
  const doc = sectPr.ownerDocument;
  const pgSz = wEl(doc, 'pgSz'); wSet(pgSz,'w','11906'); wSet(pgSz,'h','16838'); sectPr.appendChild(pgSz);
  const pgMar = wEl(doc, 'pgMar');
  wSet(pgMar,'top','1440'); wSet(pgMar,'right','1797'); wSet(pgMar,'bottom','1440'); wSet(pgMar,'left','1797');
  wSet(pgMar,'header','851'); wSet(pgMar,'footer','992'); wSet(pgMar,'gutter','0'); sectPr.appendChild(pgMar);
  const pgNum = wEl(doc, 'pgNumType'); wSet(pgNum,'start','1'); sectPr.appendChild(pgNum);
  const cols = wEl(doc, 'cols'); wSet(cols,'space','720'); sectPr.appendChild(cols);
  const dg = wEl(doc, 'docGrid'); wSet(dg,'type','lines'); wSet(dg,'linePitch','360'); sectPr.appendChild(dg);
}

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`;

// ═══════════════════════════════════════════════════════════════
// 主轉換函式
// ═══════════════════════════════════════════════════════════════
async function convertDocx(file, onLog) {
  onLog('讀取 DOCX...');
  const ab = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(ab);

  onLog('解析文件結構...');
  const docXml = await zip.file('word/document.xml').async('string');
  const doc = parseXML(docXml);
  const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
  const children = [...body.childNodes].filter(n => n.nodeType === 1);

  // ── Phase 1: 格式化段落和表格 ──────────────────────────────
  onLog('套用段落格式...');
  let paraIdx = 0;
  const stats = { cover:0, h1:0, h2:0, body:0, empty:0, table:0 };

  children.forEach(el => {
    if (el.localName === 'p' && el.namespaceURI === W_NS) {
      const style = getStyle(el);
      const text = getText(el);
      if (!text) { applyEmpty(el); stats.empty++; }
      else if (style === '1' || style === 'Heading1' || style === 'Heading 1') { applyH1(el); stats.h1++; }
      else if (style === '2' || style === 'Heading2' || style === 'Heading 2' || style === '3') { applyH2(el); stats.h2++; }
      else if (paraIdx < 17) { applyBody(el); stats.cover++; }  // Will override with cover below
      else { applyBody(el); stats.body++; }
      paraIdx++;
    } else if (el.localName === 'tbl' && el.namespaceURI === W_NS) {
      applyTable(el); stats.table++;
    } else if (el.localName === 'sectPr' && el.namespaceURI === W_NS) {
      updateMainSectPr(el);
    }
  });

  // ── Apply cover format SEPARATELY (needs all-runs approach) ──
  onLog('套用封面格式...');
  const childrenNow = [...body.childNodes].filter(n => n.nodeType === 1);
  let coverCount = 0;
  for (const el of childrenNow) {
    if (el.localName !== 'p' || el.namespaceURI !== W_NS) continue;
    const style = getStyle(el);
    const text = getText(el);
    if (style === '1' || style === '2') break; // Past cover section
    if (coverCount >= 17) break;
    if (text) { applyBody(el); } // First apply body to reset
    applyBody(el);
    coverCount++;
  }
  // NOW apply cover with proper sz/bold fix
  coverCount = 0;
  for (const el of childrenNow) {
    if (el.localName !== 'p' || el.namespaceURI !== W_NS) { if (el.localName === 'tbl') break; continue; }
    const style = getStyle(el);
    if (style === '1' || style === '2') break;
    if (coverCount >= 17) break;
    applyBody(el); // reset spacing/indent
    // Override: cover-specific pPr
    const pPr = ensurePPr(el);
    rm(pPr, 'pStyle');
    const sp = goc(pPr, 'spacing'); wSet(sp,'line','400'); wSet(sp,'lineRule','exact'); rmAttr(sp,'before'); rmAttr(sp,'after');
    const ind = goc(pPr, 'ind'); wSet(ind,'firstLine','0');
    const jc = goc(pPr, 'jc'); wSet(jc,'val','center');
    const pRPr = goc(pPr, 'rPr'); clearEl(pRPr); pRPr.appendChild(makeColor(doc, '000000', 'text1'));
    // Fix ALL runs (including in hyperlinks) - sz=28, szCs=28, no bold
    const allRuns = [...el.getElementsByTagNameNS(W_NS, 'r')];
    allRuns.forEach(run => {
      const rPr = ensureRPr(run); clearEl(rPr);
      rPr.appendChild(makeColor(doc, '000000', 'text1'));
      rPr.appendChild(makeSz(doc, '28')); rPr.appendChild(makeSzCs(doc, '28'));
    });
    coverCount++;
  }
  onLog(`格式: H1=${stats.h1} H2=${stats.h2} 正文=${stats.body} 表格=${stats.table}`);

  // ── Phase 2: Extract headings and table captions ─────────────
  onLog('提取標題資訊...');
  const { headings, tables } = extractHeadings(body);
  onLog(`H1=${headings.filter(h=>h.level===1).length} H2=${headings.filter(h=>h.level===2).length} 表格=${tables.length}`);

  // ── Phase 3: Add bookmarks to headings ───────────────────────
  onLog('加入書籤...');
  addBookmarks(body, headings, tables);

  // ── Phase 4: Rebuild TOC ──────────────────────────────────────
  onLog('重建目錄與表目次...');
  rebuildTOC(body, headings, tables);

  // ── Phase 5: sectPr + footer ──────────────────────────────────
  onLog('更新分節與頁尾...');
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  let footerRid = null;
  const m = relsXml.match(/Id="(rId\d+)"[^>]+footer1\.xml/);
  if (m) { footerRid = m[1]; }
  else {
    const nums = [...relsXml.matchAll(/rId(\d+)/g)].map(x => parseInt(x[1]));
    footerRid = 'rId' + (nums.length ? Math.max(...nums) + 1 : 10);
  }
  relsXml = relsXml.replace(/<Relationship[^>]+\/header\d*\.xml[^>]*\/>/g, '');
  if (!relsXml.includes('footer1.xml')) {
    relsXml = relsXml.replace('</Relationships>',
      `<Relationship Id="${footerRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
  }
  zip.file('word/_rels/document.xml.rels', relsXml);

  // Insert cover sectPr before first H1 chapter (第X章)
  const bodyChildren = [...body.childNodes].filter(n => n.nodeType === 1);
  const firstChapter = bodyChildren.find(el =>
    el.localName === 'p' && el.namespaceURI === W_NS &&
    getStyle(el) === '1' && getText(el).includes('章'));
  if (firstChapter) {
    // Remove any existing inner sectPr paragraphs (from previous conversion)
    bodyChildren.forEach(el => {
      if (el.localName === 'p' && el.namespaceURI === W_NS) {
        const pPr = getChildW(el, 'pPr');
        if (pPr && getChildW(pPr, 'sectPr')) body.removeChild(el);
      }
    });
    const coverPara = wEl(doc, 'p');
    const cPPr = wEl(doc, 'pPr'); coverPara.appendChild(cPPr);
    cPPr.appendChild(buildCoverSectPr(doc, footerRid));
    body.insertBefore(coverPara, firstChapter);
  }

  // ── Write files ───────────────────────────────────────────────
  zip.file('word/document.xml', serXML(doc));
  zip.file('word/footer1.xml', FOOTER_XML);

  let ctXml = await zip.file('[Content_Types].xml').async('string');
  if (!ctXml.includes('footer1.xml')) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
    zip.file('[Content_Types].xml', ctXml);
  }

  onLog('產生輸出檔案...');
  return await zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}

window.convertDocx = convertDocx;
