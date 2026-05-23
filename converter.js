/**
 * AAFCO 精算報告格式轉換工具 — converter.js v19
 * JavaScript port of format_converter_v19.py
 * 依賴：JSZip (loaded via CDN in index.html)
 */

const W  = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml';
const XML_SPACE = 'http://www.w3.org/XML/1998/namespace';

// ── XML helpers ──────────────────────────────────────────────

function wt(n) { return `{${W}}${n}`; }

function parseXML(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}

function serializeXML(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function wElem(doc, tag) {
  return doc.createElementNS(W, `w:${tag}`);
}

function wAttr(el, attr, val) {
  el.setAttributeNS(W, `w:${attr}`, String(val));
}

function wGet(el, attr) {
  return el.getAttributeNS(W, attr);
}

function findChild(el, tag) {
  for (const c of el.childNodes) {
    if (c.localName === tag && c.namespaceURI === W) return c;
  }
  return null;
}

function findAll(el, tag) {
  return Array.from(el.childNodes).filter(c => c.localName === tag && c.namespaceURI === W);
}

function goc(doc, parent, tag) {
  let el = findChild(parent, tag);
  if (!el) { el = wElem(doc, tag); parent.appendChild(el); }
  return el;
}

function rm(parent, tag) {
  for (const c of findAll(parent, tag)) parent.removeChild(c);
}

function rmAttr(el, attr) {
  if (el.hasAttributeNS(W, attr)) el.removeAttributeNS(W, attr);
}

function getText(para) {
  return Array.from(para.getElementsByTagNameNS(W, 't'))
    .map(t => t.textContent || '').join('').trim();
}

function getStyle(para) {
  const pPr = findChild(para, 'pPr');
  if (!pPr) return 'Normal';
  const ps = findChild(pPr, 'pStyle');
  return ps ? (wGet(ps, 'val') || 'Normal') : 'Normal';
}

function directRuns(para) {
  return findAll(para, 'r');
}

function ensurePPr(doc, para) {
  let pPr = findChild(para, 'pPr');
  if (!pPr) { pPr = wElem(doc, 'pPr'); para.insertBefore(pPr, para.firstChild); }
  return pPr;
}

function ensureRPr(doc, run) {
  let rPr = findChild(run, 'rPr');
  if (!rPr) { rPr = wElem(doc, 'rPr'); run.insertBefore(rPr, run.firstChild); }
  return rPr;
}

function clearChildren(el) {
  while (el.firstChild) el.removeChild(el.firstChild);
}

function insertAfter(parent, newNode, refNode) {
  const next = refNode.nextSibling;
  if (next) parent.insertBefore(newNode, next);
  else parent.appendChild(newNode);
}

function insertBefore(parent, newNode, refNode) {
  parent.insertBefore(newNode, refNode);
}

function bodyChildren(body) {
  return Array.from(body.childNodes).filter(n => n.nodeType === 1);
}

// ── Paragraph formatters ──────────────────────────────────────

function applyH1(doc, para) {
  const pPr = ensurePPr(doc, para);
  wAttr(goc(doc, pPr, 'pStyle'), 'val', '1');
  const sp = goc(doc, pPr, 'spacing');
  wAttr(sp, 'line', '480'); wAttr(sp, 'lineRule', 'auto');
  rmAttr(sp, 'before'); rmAttr(sp, 'after');
  const ind = goc(doc, pPr, 'ind');
  wAttr(ind, 'firstLine', '0');
  rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  rm(pPr, 'jc');
  const pPrRpr = goc(doc, pPr, 'rPr');
  clearChildren(pPrRpr);
  const rF0 = wElem(doc, 'rFonts');
  rF0.setAttributeNS(W, 'w:ascii', 'Times New Roman');
  rF0.setAttributeNS(W, 'w:hAnsi', 'Times New Roman');
  pPrRpr.appendChild(rF0);
  const c0 = wElem(doc, 'color');
  wAttr(c0, 'val', '000000'); wAttr(c0, 'themeColor', 'text1');
  pPrRpr.appendChild(c0);
  const sz = wElem(doc, 'sz'); wAttr(sz, 'val', '28'); pPrRpr.appendChild(sz);
  const szCs = wElem(doc, 'szCs'); wAttr(szCs, 'val', '28'); pPrRpr.appendChild(szCs);

  for (const run of directRuns(para)) {
    const rPr = ensureRPr(doc, run);
    clearChildren(rPr);
    const rF2 = wElem(doc, 'rFonts');
    rF2.setAttributeNS(W, 'w:ascii', 'Times New Roman');
    rF2.setAttributeNS(W, 'w:hAnsi', 'Times New Roman');
    rPr.appendChild(rF2);
    const c2 = wElem(doc, 'color');
    wAttr(c2, 'val', '000000'); wAttr(c2, 'themeColor', 'text1');
    rPr.appendChild(c2);
    const sz2 = wElem(doc, 'sz'); wAttr(sz2, 'val', '28'); rPr.appendChild(sz2);
    const szCs2 = wElem(doc, 'szCs'); wAttr(szCs2, 'val', '28'); rPr.appendChild(szCs2);
  }
}

function applyH2(doc, para, numId) {
  const text = getText(para);
  const hasJiePrefix = /^第[一二三四五六七八九十百]+節[\s\u3000]/.test(text);

  const pPr = ensurePPr(doc, para);
  wAttr(goc(doc, pPr, 'pStyle'), 'val', '2');
  rm(pPr, 'spacing'); rm(pPr, 'ind'); rm(pPr, 'numPr');

  if (hasJiePrefix && numId != null) {
    for (const run of directRuns(para)) {
      const t = findChild(run, 't');
      if (t && t.textContent) {
        t.textContent = t.textContent.replace(/^第[一二三四五六七八九十百]+節[\s\u3000]+/, '');
      }
    }
    const numPr = wElem(doc, 'numPr');
    const ilvl = wElem(doc, 'ilvl'); wAttr(ilvl, 'val', '0'); numPr.appendChild(ilvl);
    const numIdEl = wElem(doc, 'numId'); wAttr(numIdEl, 'val', String(numId)); numPr.appendChild(numIdEl);
    const psEl = findChild(pPr, 'pStyle');
    insertAfter(pPr, numPr, psEl);
  }

  const pPrRpr = goc(doc, pPr, 'rPr');
  clearChildren(pPrRpr);
  const rF = wElem(doc, 'rFonts');
  rF.setAttributeNS(W, 'w:ascii', 'Times New Roman');
  rF.setAttributeNS(W, 'w:hAnsi', 'Times New Roman');
  pPrRpr.appendChild(rF);
  const bv = wElem(doc, 'b'); wAttr(bv, 'val', '0'); pPrRpr.appendChild(bv);
  const bcsv = wElem(doc, 'bCs'); wAttr(bcsv, 'val', '0'); pPrRpr.appendChild(bcsv);
  const cv = wElem(doc, 'color'); wAttr(cv, 'val', '000000'); wAttr(cv, 'themeColor', 'text1'); pPrRpr.appendChild(cv);
  const szv = wElem(doc, 'sz'); wAttr(szv, 'val', '24'); pPrRpr.appendChild(szv);
  const szcsv = wElem(doc, 'szCs'); wAttr(szcsv, 'val', '24'); pPrRpr.appendChild(szcsv);

  for (const run of directRuns(para)) {
    const rPr = ensureRPr(doc, run);
    clearChildren(rPr);
    const rF2 = wElem(doc, 'rFonts');
    rF2.setAttributeNS(W, 'w:ascii', 'Times New Roman');
    rF2.setAttributeNS(W, 'w:hAnsi', 'Times New Roman');
    rF2.setAttributeNS(W, 'w:hint', 'eastAsia');
    rPr.appendChild(rF2);
    const b2 = wElem(doc, 'b'); wAttr(b2, 'val', '0'); rPr.appendChild(b2);
    const bcs2 = wElem(doc, 'bCs'); wAttr(bcs2, 'val', '0'); rPr.appendChild(bcs2);
    const c2 = wElem(doc, 'color'); wAttr(c2, 'val', '000000'); wAttr(c2, 'themeColor', 'text1'); rPr.appendChild(c2);
    const sz2 = wElem(doc, 'sz'); wAttr(sz2, 'val', '24'); rPr.appendChild(sz2);
    const szCs2 = wElem(doc, 'szCs'); wAttr(szCs2, 'val', '24'); rPr.appendChild(szCs2);
  }
}

function applyBody(doc, para) {
  const pPr = ensurePPr(doc, para);
  rm(pPr, 'pStyle');
  const sp = goc(doc, pPr, 'spacing');
  wAttr(sp, 'before', '80'); wAttr(sp, 'after', '80');
  rmAttr(sp, 'line'); rmAttr(sp, 'lineRule');
  const ind = goc(doc, pPr, 'ind');
  rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  const textForJc = getText(para);
  const isCaption = /^表\s*\d/.test(textForJc);
  if (isCaption) {
    wAttr(goc(doc, pPr, 'jc'), 'val', 'left');
    wAttr(ind, 'firstLine', '0');
  } else {
    wAttr(goc(doc, pPr, 'jc'), 'val', 'both');
    wAttr(ind, 'firstLine', '480');
  }
  const pPrRpr = goc(doc, pPr, 'rPr');
  clearChildren(pPrRpr);
  wAttr(wElem(doc, 'color'), 'val', 'auto'); // placeholder — appendChild below
  const cEl = wElem(doc, 'color'); wAttr(cEl, 'val', 'auto'); pPrRpr.appendChild(cEl);

  for (const run of directRuns(para)) {
    const rPr = findChild(run, 'rPr');
    if (rPr) clearChildren(rPr);
  }
}

function makeEmptyCoverPara(doc) {
  const p = wElem(doc, 'p');
  const pPr = wElem(doc, 'pPr'); p.appendChild(pPr);
  const sp = wElem(doc, 'spacing'); wAttr(sp, 'line', '400'); wAttr(sp, 'lineRule', 'exact'); pPr.appendChild(sp);
  const ind = wElem(doc, 'ind'); wAttr(ind, 'firstLine', '0'); pPr.appendChild(ind);
  const jc = wElem(doc, 'jc'); wAttr(jc, 'val', 'center'); pPr.appendChild(jc);
  const rPr = wElem(doc, 'rPr'); pPr.appendChild(rPr);
  const c = wElem(doc, 'color'); wAttr(c, 'val', '000000'); wAttr(c, 'themeColor', 'text1'); rPr.appendChild(c);
  const sz = wElem(doc, 'sz'); wAttr(sz, 'val', '28'); rPr.appendChild(sz);
  const szCs = wElem(doc, 'szCs'); wAttr(szCs, 'val', '28'); rPr.appendChild(szCs);
  return p;
}

function applyCover(doc, para) {
  const paraText = getText(para);

  // Remove leading spaces and 第二部分 from cover runs
  for (const run of directRuns(para)) {
    const t = findChild(run, 't');
    if (t && t.textContent) {
      t.textContent = t.textContent.replace('（第二部分）', '').replace('(第二部分)', '');
      if (['專案編號', '研發配方', '目標對象'].some(kw => t.textContent.includes(kw))) {
        t.textContent = t.textContent.replace(/^\s+/, '');
      }
    }
  }

  const LEFT_KW = ['專案編號', '研發配方', '目標對象'];
  const isLeft = LEFT_KW.some(kw => paraText.includes(kw));

  const pPr = ensurePPr(doc, para);
  rm(pPr, 'pStyle'); rm(pPr, 'pBdr');
  const sp = goc(doc, pPr, 'spacing'); wAttr(sp, 'line', '400'); wAttr(sp, 'lineRule', 'exact');
  rmAttr(sp, 'before'); rmAttr(sp, 'after');
  const ind = goc(doc, pPr, 'ind'); wAttr(ind, 'firstLine', '0');
  rmAttr(ind, 'firstLineChars'); rmAttr(ind, 'hanging'); rmAttr(ind, 'left');
  wAttr(goc(doc, pPr, 'jc'), 'val', isLeft ? 'left' : 'center');
  const pPrRpr = goc(doc, pPr, 'rPr');
  clearChildren(pPrRpr);
  const c0 = wElem(doc, 'color'); wAttr(c0, 'val', '000000'); wAttr(c0, 'themeColor', 'text1'); pPrRpr.appendChild(c0);
  const sz = wElem(doc, 'sz'); wAttr(sz, 'val', '28'); pPrRpr.appendChild(sz);
  const szCs = wElem(doc, 'szCs'); wAttr(szCs, 'val', '28'); pPrRpr.appendChild(szCs);

  for (const run of directRuns(para)) {
    const rPr = ensureRPr(doc, run);
    clearChildren(rPr);
    const c = wElem(doc, 'color'); wAttr(c, 'val', '000000'); wAttr(c, 'themeColor', 'text1'); rPr.appendChild(c);
    const sz2 = wElem(doc, 'sz'); wAttr(sz2, 'val', '28'); rPr.appendChild(sz2);
    const szCs2 = wElem(doc, 'szCs'); wAttr(szCs2, 'val', '28'); rPr.appendChild(szCs2);
  }
}

function applyEmpty(doc, para) {
  const pPr = ensurePPr(doc, para);
  rm(pPr, 'pStyle');
  const sp = goc(doc, pPr, 'spacing'); wAttr(sp, 'line', '480'); wAttr(sp, 'lineRule', 'auto');
}

function applyTocLine(doc, para) {
  const pPr = ensurePPr(doc, para);
  rm(pPr, 'pStyle');
  const sp = goc(doc, pPr, 'spacing'); wAttr(sp, 'line', '480'); wAttr(sp, 'lineRule', 'auto');
}

// ── Table formatter ───────────────────────────────────────────

function applyTable(doc, tbl, ti) {
  let tblPr = findChild(tbl, 'tblPr');
  if (!tblPr) { tblPr = wElem(doc, 'tblPr'); tbl.insertBefore(tblPr, tbl.firstChild); }
  rm(tblPr, 'tblBorders'); rm(tblPr, 'tblStyle'); rm(tblPr, 'tblLook');
  const tblW = goc(doc, tblPr, 'tblW'); wAttr(tblW, 'w', '9746'); wAttr(tblW, 'type', 'dxa');
  wAttr(goc(doc, tblPr, 'jc'), 'val', 'center');
  const tblCellMar = goc(doc, tblPr, 'tblCellMar');
  const left = goc(doc, tblCellMar, 'left'); wAttr(left, 'w', '10'); wAttr(left, 'type', 'dxa');
  const right = goc(doc, tblCellMar, 'right'); wAttr(right, 'w', '10'); wAttr(right, 'type', 'dxa');
  const tblLook = wElem(doc, 'tblLook');
  wAttr(tblLook, 'val', '04A0'); wAttr(tblLook, 'firstRow', '1'); wAttr(tblLook, 'lastRow', '0');
  wAttr(tblLook, 'firstColumn', '1'); wAttr(tblLook, 'lastColumn', '0');
  wAttr(tblLook, 'noHBand', '0'); wAttr(tblLook, 'noVBand', '1');
  tblPr.appendChild(tblLook);

  const TABLE3_SUBHEADERS = ['蛋白質與胺基酸', '脂肪與脂肪酸', '維生素', '礦物質'];
  const TABLE3_NO_BORDER_ROWS = [
    '維生素A Vitamin A', '維生素D Vitamin D3', '維生素E Vitamin E',
    '維生素B1 Thiamine', '維生素B2 Riboflavin', '菸鹼酸 Niacin (B3)', '維生素B6 Pyridoxine'
  ];

  const rows = findAll(tbl, 'tr');
  const totalRows = rows.length;

  rows.forEach((row, ri) => {
    const isHeader = ri === 0;
    const isLast = ri === totalRows - 1;
    let trPr = findChild(row, 'trPr');
    if (!trPr) { trPr = wElem(doc, 'trPr'); row.insertBefore(trPr, row.firstChild); }
    if (isHeader) wAttr(goc(doc, trPr, 'jc'), 'val', 'center');

    const cells = findAll(row, 'tc');
    cells.forEach(cell => {
      let tcPr = findChild(cell, 'tcPr');
      if (!tcPr) { tcPr = wElem(doc, 'tcPr'); cell.insertBefore(tcPr, cell.firstChild); }
      rm(tcPr, 'tcBorders');
      const tcBorders = wElem(doc, 'tcBorders');
      const tcW = findChild(tcPr, 'tcW');
      if (tcW) insertAfter(tcPr, tcBorders, tcW);
      else tcPr.insertBefore(tcBorders, tcPr.firstChild);

      const rowText = Array.from(cell.getElementsByTagNameNS(W, 't')).map(t => t.textContent).join('');
      const isNoBorder = ti === 2 && !isHeader && TABLE3_NO_BORDER_ROWS.some(s => rowText.includes(s));
      const isSubhdr = ti === 2 && !isHeader && !isNoBorder && TABLE3_SUBHEADERS.some(s => rowText.includes(s));

      function addBorder(tag, sz) {
        const b = wElem(doc, tag);
        wAttr(b, 'val', 'single'); wAttr(b, 'sz', sz); wAttr(b, 'space', '0'); wAttr(b, 'color', 'auto');
        tcBorders.appendChild(b);
      }

      if (isHeader) { addBorder('top', '12'); addBorder('bottom', '4'); }
      else if (isLast) { addBorder('bottom', '12'); }
      else if (isSubhdr) { addBorder('top', '4'); addBorder('bottom', '4'); }

      const shd = goc(doc, tcPr, 'shd');
      wAttr(shd, 'val', 'clear'); wAttr(shd, 'color', 'auto'); wAttr(shd, 'fill', 'auto');
      const tcMar = goc(doc, tcPr, 'tcMar');
      for (const [side, val] of [['top','80'],['left','120'],['bottom','80'],['right','120']]) {
        const m = goc(doc, tcMar, side); wAttr(m, 'w', val); wAttr(m, 'type', 'dxa');
      }
      rm(tcPr, 'hideMark');
      tcPr.appendChild(wElem(doc, 'hideMark'));

      for (const para of findAll(cell, 'p')) {
        const pPr = ensurePPr(doc, para);
        rm(pPr, 'pStyle'); rm(pPr, 'jc');
        wAttr(goc(doc, pPr, 'widowControl'), 'val', '0');
        const sp = goc(doc, pPr, 'spacing');
        wAttr(sp, 'before', '60'); wAttr(sp, 'after', '60'); wAttr(sp, 'line', '240'); wAttr(sp, 'lineRule', 'auto');
        wAttr(goc(doc, pPr, 'ind'), 'firstLine', '0');
        const pPrRpr = goc(doc, pPr, 'rPr');
        clearChildren(pPrRpr);
        const cEl = wElem(doc, 'color'); wAttr(cEl, 'val', 'auto'); pPrRpr.appendChild(cEl);
        const kernEl = wElem(doc, 'kern'); wAttr(kernEl, 'val', '0'); pPrRpr.appendChild(kernEl);

        for (const run of findAll(para, 'r')) {
          const origRpr = findChild(run, 'rPr');
          let origColor = null;
          if (origRpr) {
            const origC = findChild(origRpr, 'color');
            if (origC) {
              const v = wGet(origC, 'val') || '';
              if (v && !['AUTO', '000000', ''].includes(v.toUpperCase())) origColor = v;
            }
          }
          const rPr = ensureRPr(doc, run);
          clearChildren(rPr);
          if (isHeader) {
            rPr.appendChild(wElem(doc, 'bCs'));
            const c = wElem(doc, 'color'); wAttr(c, 'val', 'auto'); rPr.appendChild(c);
          } else if (origColor) {
            const c = wElem(doc, 'color'); wAttr(c, 'val', origColor); rPr.appendChild(c);
          } else {
            const cellText = getText(para);
            let judgeColor = null;
            if (cellText.includes('符合') || cellText.includes('PASS')) judgeColor = '006400';
            else if (cellText.includes('不足') || cellText.includes('DEF')) judgeColor = 'CC0000';
            else if (cellText.includes('邊緣') || cellText.includes('MAR')) judgeColor = '8B6914';
            else if (cellText.includes('整體合規判定') || cellText.includes('DOES NOT MEET AAFCO')) judgeColor = 'FF0000';
            if (judgeColor) { const c = wElem(doc, 'color'); wAttr(c, 'val', judgeColor); rPr.appendChild(c); }
          }
          const kEl = wElem(doc, 'kern'); wAttr(kEl, 'val', '0'); rPr.appendChild(kEl);
        }
      }
    });
  });
}

// ── Cover blank line insertion ────────────────────────────────

function insertCoverBlanks(doc, body) {
  const children = bodyChildren(body);
  const coverParas = [];
  for (const el of children) {
    if (el.localName === 'p' && el.namespaceURI === W) {
      const s = getStyle(el);
      if (s === '1' || s === '2') break;
      coverParas.push(el);
    }
  }

  const textParas = coverParas
    .map(el => ({ el, text: getText(el) }))
    .filter(({ text }) => text.length > 0);

  const RULES = [
    [t => t.startsWith('寵物餐食'), 1],
    [t => t.includes('精算報告') || t.includes('Part'), 2],
    [t => t.includes('Development Project'), 2],
    [t => t.includes('已絕育成年犬消化之原型食材模組研發'), 1],
    [t => t.includes('Effects of Whole') || t.includes('Neutered'), 3],
    [t => t.startsWith('目標對象'), 3],
    [t => t.includes('Ph.D.') && t.includes('（'), 3],
    [t => (t.includes('年') && t.includes('月') && t.includes('民國')) || (t.includes('年') && t.includes('月') && t.length < 15), 1],
  ];

  const insertions = [];
  for (const [fn, count] of RULES) {
    for (const { el, text } of textParas) {
      if (fn(text)) { insertions.push({ el, count }); break; }
    }
  }

  // De-dup and process from bottom to top
  const seen = new Set();
  const unique = [];
  for (const ins of [...insertions].sort((a, b) => {
    const ai = Array.from(body.childNodes).indexOf(a.el);
    const bi = Array.from(body.childNodes).indexOf(b.el);
    return bi - ai;
  })) {
    if (!seen.has(ins.el)) { unique.push(ins); seen.add(ins.el); }
  }

  for (const { el, count } of unique) {
    for (let i = 0; i < count; i++) {
      insertAfter(body, makeEmptyCoverPara(doc), el);
    }
  }
}

// ── Table 2 column width fix ──────────────────────────────────

function fixTable2Widths(doc, body) {
  const tables = findAll(body, 'tbl');
  if (tables.length < 2) return;
  const tbl2 = tables[1];
  const tblGrid = findChild(tbl2, 'tblGrid');
  if (tblGrid) {
    const cols = findAll(tblGrid, 'gridCol');
    if (cols.length === 3) { wAttr(cols[1], 'w', '900'); wAttr(cols[2], 'w', '1500'); }
  }
  for (const row of findAll(tbl2, 'tr')) {
    const cells = findAll(row, 'tc');
    if (cells.length === 3) {
      for (const [idx, newW] of [[1, '900'], [2, '1500']]) {
        const tcPr = findChild(cells[idx], 'tcPr');
        if (tcPr) { const tcW = findChild(tcPr, 'tcW'); if (tcW) wAttr(tcW, 'w', newW); }
      }
    }
  }
}

// ── Table title normalization ─────────────────────────────────

const TABLE_TITLE_MAP = {
  '表 1、FMLA-N-D10-08 單餐食材組成與巨量營養素': '表1、單餐食材組成與巨量營養素',
  '表 2、10 kg 已絕育成年犬能量需求計算': '表2、10 kg 已絕育成年犬能量需求計算',
  '表 3、FMLA-N-D10-08 AAFCO 成年犬維持期營養素全項對照表（DM 基準，AAFCO最低值已套用PF9校正係數 ×1.032201）': '表3、成年犬維持期營養素全項對照表',
  '表 4、FMLA-N-D10-08 AAFCO 合規統計摘要': '表4、合規統計摘要',
  '表 5、FMLA-N-D10-08 補充劑處方——每餐及每日精確劑量': '表5、補充劑處方——每餐及每日精確劑量',
};

function normalizeTableTitles(doc, body) {
  for (const el of findAll(body, 'p')) {
    const paraText = getText(el);
    for (const [oldT, newT] of Object.entries(TABLE_TITLE_MAP)) {
      if (paraText.includes(oldT)) {
        for (const run of directRuns(el)) {
          const t = findChild(run, 't');
          if (t && t.textContent && t.textContent.includes(oldT)) {
            t.textContent = t.textContent.replace(oldT, newT);
            t.removeAttributeNS(XML_SPACE, 'space');
          }
        }
        break;
      }
    }
  }
}

// ── Disclaimer / copyright special formatting ─────────────────

function applySpecialBodyFormatting(doc, body) {
  let disclaimerBodyIdx = -1;
  const children = bodyChildren(body);

  children.forEach((el, i) => {
    if (el.localName !== 'p') return;
    const text = getText(el);
    if (text.includes('聲明事項') && text.includes('Disclaimer')) {
      const pPr = ensurePPr(doc, el);
      wAttr(goc(doc, pPr, 'jc'), 'val', 'left');
      wAttr(goc(doc, pPr, 'ind'), 'firstLine', '0');
      for (const run of directRuns(el)) {
        const rPr = ensureRPr(doc, run);
        if (!findChild(rPr, 'b')) rPr.appendChild(wElem(doc, 'b'));
      }
    } else if (text.includes('本報告採用食材平均組成值')) {
      for (const run of directRuns(el)) {
        const rPr = ensureRPr(doc, run);
        if (!findChild(rPr, 'i')) rPr.appendChild(wElem(doc, 'i'));
      }
      disclaimerBodyIdx = i;
    }
  });

  // Insert copyright paragraph after disclaimer body
  if (disclaimerBodyIdx >= 0) {
    const refEl = children[disclaimerBodyIdx];
    const cp = wElem(doc, 'p');
    const cpPr = wElem(doc, 'pPr'); cp.appendChild(cpPr);
    const sp = wElem(doc, 'spacing'); wAttr(sp, 'before', '80'); wAttr(sp, 'after', '80'); cpPr.appendChild(sp);
    const jc = wElem(doc, 'jc'); wAttr(jc, 'val', 'both'); cpPr.appendChild(jc);
    const ind = wElem(doc, 'ind'); wAttr(ind, 'firstLine', '480'); cpPr.appendChild(ind);
    const pPrRpr = wElem(doc, 'rPr'); const cA = wElem(doc, 'color'); wAttr(cA, 'val', 'auto'); pPrRpr.appendChild(cA); cpPr.appendChild(pPrRpr);
    const cpR = wElem(doc, 'r'); cp.appendChild(cpR);
    const cpRpr = wElem(doc, 'rPr'); cpR.appendChild(cpRpr);
    cpRpr.appendChild(wElem(doc, 'b'));
    const cRed = wElem(doc, 'color'); wAttr(cRed, 'val', 'FF0000'); cpRpr.appendChild(cRed);
    const cpT = wElem(doc, 't');
    cpT.setAttributeNS(XML_SPACE, 'xml:space', 'preserve');
    cpT.textContent = '© 2026 Yi-Cheng Hou, Ph.D. All Rights Reserved. 未經書面授權，不得重製、轉載或轉讓。';
    cpR.appendChild(cpT);
    insertAfter(body, cp, refEl);
  }
}

// ── Simplified TOC rebuild (update existing entries) ──────────

function rebuildTOC(doc, body) {
  // Instead of full TOC rebuild, update the table-of-figures entries to use short titles
  // Full heading TOC is preserved from source (bookmarks and PAGEREF fields are kept)
  const allText = Array.from(body.getElementsByTagNameNS(W, 't'));
  const tocTitleMap = {
    'FMLA-N-D10-08 單餐食材組成與巨量營養素': '單餐食材組成與巨量營養素',
    'FMLA-N-D10-08 AAFCO 成年犬維持期營養素全項對照表（DM 基準，AAFCO最低值已套用PF9校正係數 ×1.032201）': '成年犬維持期營養素全項對照表',
    'FMLA-N-D10-08 AAFCO 合規統計摘要': '合規統計摘要',
    'FMLA-N-D10-08 補充劑處方——每餐及每日精確劑量': '補充劑處方——每餐及每日精確劑量',
    // Handle split runs — partial matches
    '、FMLA-N-D': '、',
    ' AAFCO 成年犬維持期營養素全項對照表（DM 基準，AAFCO最低值已套用PF': ' 成年犬維持期營養素全項對照表',
    ' AAFCO 合規統計摘要': ' 合規統計摘要',
    ' 補充劑處方——每餐及每日精確劑量': ' 補充劑處方——每餐及每日精確劑量',
    ' 單餐食材組成與巨量營養素': ' 單餐食材組成與巨量營養素',
  };

  // For runs that contain numbers like "10", "-", "08" after FMLA-N-D, we need to
  // track and eliminate them. Use a state machine approach on consecutive runs in hyperlinks.
  const hyperlinks = Array.from(body.getElementsByTagNameNS(W, 'hyperlink'));
  for (const hlink of hyperlinks) {
    const runs = findAll(hlink, 'r');
    // Detect if this hyperlink is a table TOC entry (contains 表 and digits)
    const hlinkText = getText(hlink);
    if (!hlinkText.match(/^表\s*\d/) && !hlinkText.includes('表1') && !hlinkText.includes('表2')
        && !hlinkText.includes('表3') && !hlinkText.includes('表4') && !hlinkText.includes('表5')) continue;

    // Rebuild the title part: find all text runs before the tab/PAGEREF
    const titleRuns = [];
    let foundPageRef = false;
    for (const run of runs) {
      if (findChild(run, 'tab') || findChild(run, 'fldChar') || findChild(run, 'instrText')) { foundPageRef = true; break; }
      titleRuns.push(run);
    }
    if (titleRuns.length === 0 || foundPageRef === false) continue;

    // Reconstruct short title
    const fullText = titleRuns.map(r => { const t = findChild(r, 't'); return t ? t.textContent : ''; }).join('');
    let shortText = fullText;
    // Map long → short
    shortText = shortText.replace(/表\s*(\d)、FMLA-N-D10-08\s*(AAFCO\s*)?成年犬維持期營養素全項對照表.*/, '表$1、成年犬維持期營養素全項對照表');
    shortText = shortText.replace(/表\s*(\d)、FMLA-N-D10-08\s*(AAFCO\s*)?合規統計摘要/, '表$1、合規統計摘要');
    shortText = shortText.replace(/表\s*(\d)、FMLA-N-D10-08\s*補充劑處方/, '表$1、補充劑處方');
    shortText = shortText.replace(/表\s*(\d)、FMLA-N-D10-08\s*單餐食材組成/, '表$1、單餐食材組成');
    shortText = shortText.replace(/表\s*(\d)、(\d+\s*kg)/, '表$1、$2');
    shortText = shortText.replace(/表\s+(\d)、/, '表$1、');

    if (shortText === fullText) continue;

    // Replace all title runs with a single run containing the short text
    const firstRun = titleRuns[0];
    const rPr = findChild(firstRun, 'rPr');
    const newT = findChild(firstRun, 't') || wElem(doc, 't');
    newT.textContent = shortText;
    newT.removeAttributeNS(XML_SPACE, 'space');
    if (!findChild(firstRun, 't')) firstRun.appendChild(newT);
    // Remove the remaining title runs
    for (let i = 1; i < titleRuns.length; i++) hlink.removeChild(titleRuns[i]);
  }
}

// ── sectPr update ─────────────────────────────────────────────

function updateMainSectPr(doc, sectPr, tplSectPr) {
  if (!tplSectPr) return;
  // Copy page size and margins from template
  for (const tag of ['pgSz', 'pgMar', 'pgNumType', 'cols', 'docGrid']) {
    rm(sectPr, tag);
    const tplEl = findChild(tplSectPr, tag);
    if (tplEl) sectPr.appendChild(tplEl.cloneNode(true));
  }
}

// ── Main convert function ─────────────────────────────────────

async function convert(srcBytes, tplBytes, onProgress) {
  onProgress('解壓縮檔案…', 10);

  const srcZip = await JSZip.loadAsync(srcBytes);
  const tplZip = await JSZip.loadAsync(tplBytes);

  // Load all files from source
  const out = {};
  for (const [name, file] of Object.entries(srcZip.files)) {
    if (!file.dir) out[name] = await file.async('uint8array');
  }

  // Copy template styles, numbering, settings, fonts, footer
  const TPL_COPY = ['word/styles.xml', 'word/numbering.xml', 'word/settings.xml', 'word/fontTable.xml'];
  for (const k of TPL_COPY) {
    const f = tplZip.file(k);
    if (f) out[k] = await f.async('uint8array');
  }
  const tplFooter = tplZip.file('word/footer1.xml');
  if (tplFooter) out['word/footer1.xml'] = await tplFooter.async('uint8array');

  // Fix relationships - add footer if missing, remove headers
  let relsXml = new TextDecoder().decode(out['word/_rels/document.xml.rels']);
  let footerRid = (relsXml.match(/Id="(rId\d+)"[^>]+footer1\.xml/) || [])[1];
  if (!footerRid) {
    const nums = [...relsXml.matchAll(/rId(\d+)/g)].map(m => parseInt(m[1]));
    footerRid = `rId${Math.max(...nums, 9) + 1}`;
    relsXml = relsXml.replace('</Relationships>',
      `<Relationship Id="${footerRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
  }
  relsXml = relsXml.replace(/<Relationship[^>]+\/header[^>]+\/>/g, '');
  out['word/_rels/document.xml.rels'] = new TextEncoder().encode(relsXml);

  // Remove header files
  for (const k of Object.keys(out)) {
    if (/\/header\d*\.xml/.test(k)) delete out[k];
  }

  // Fix Content_Types
  let ct = new TextDecoder().decode(out['[Content_Types].xml']);
  ct = ct.replace(/<Override[^>]+header\d*\.xml[^>]*\/>/g, '');
  if (!ct.includes('footer1.xml')) {
    ct = ct.replace('</Types>', '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
  }
  out['[Content_Types].xml'] = new TextEncoder().encode(ct);

  onProgress('解析文件 XML…', 20);

  const srcDocXml = new TextDecoder('utf-8').decode(out['word/document.xml']);
  const doc = parseXML(srcDocXml);
  const body = doc.getElementsByTagNameNS(W, 'body')[0];

  // Get template sectPr for page layout
  const tplDocBytes = tplZip.file('word/document.xml') ? await tplZip.file('word/document.xml').async('uint8array') : null;
  let tplSectPr = null;
  if (tplDocBytes) {
    const tplDoc = parseXML(new TextDecoder().decode(tplDocBytes));
    tplSectPr = tplDoc.getElementsByTagNameNS(W, 'sectPr')[0] || null;
  }

  onProgress('套用段落格式…', 35);

  // Phase 1: Format all paragraphs and tables
  let paraIdx = 0;
  let currentChapterIdx = 0;
  let chapterH1Count = 0;

  for (const el of bodyChildren(body)) {
    if (el.localName === 'p') {
      const style = getStyle(el);
      const text = getText(el);

      if (!text) {
        applyEmpty(doc, el);
      } else if (['1', 'Heading1', 'Heading 1'].includes(style)) {
        applyH1(doc, el);
        if (text.includes('第') && text.includes('章')) {
          chapterH1Count++;
          currentChapterIdx = chapterH1Count;
        }
      } else if (['2', 'Heading2', 'Heading 2', '3', 'Heading3'].includes(style)) {
        const numId = { 1: 3, 2: 4 }[currentChapterIdx] || null;
        applyH2(doc, el, numId);
      } else if (paraIdx < 17) {
        const ct = text.trim();
        if (['©', '版權所有', '本報告為委託', '僅授權委託方'].some(k => ct.includes(k))) {
          body.removeChild(el);
          continue;
        }
        applyCover(doc, el);
      } else if (text.includes('……')) {
        applyTocLine(doc, el);
      } else {
        if (['©', '版權所有', '本報告為委託', '僅授權委託方'].some(k => text.includes(k))) {
          body.removeChild(el);
          continue;
        }
        applyBody(doc, el);
        if (text.includes('聲明事項') && text.includes('Disclaimer')) {
          const pPr = ensurePPr(doc, el);
          wAttr(goc(doc, pPr, 'jc'), 'val', 'left');
          wAttr(goc(doc, pPr, 'ind'), 'firstLine', '0');
          for (const run of directRuns(el)) {
            const rPr = ensureRPr(doc, run);
            if (!findChild(rPr, 'b')) rPr.appendChild(wElem(doc, 'b'));
          }
        } else if (text.includes('本報告採用食材平均組成值')) {
          for (const run of directRuns(el)) {
            const rPr = ensureRPr(doc, run);
            if (!findChild(rPr, 'i')) rPr.appendChild(wElem(doc, 'i'));
          }
        }
      }
      paraIdx++;
    } else if (el.localName === 'tbl') {
      const tblIndex = findAll(body, 'tbl').indexOf(el);
      applyTable(doc, el, tblIndex);
    } else if (el.localName === 'sectPr') {
      if (tplSectPr) updateMainSectPr(doc, el, tplSectPr);
    }
  }

  // Insert copyright paragraph
  applySpecialBodyFormatting(doc, body);

  onProgress('修正封面空行…', 50);
  insertCoverBlanks(doc, body);

  onProgress('修正表格欄位…', 60);
  fixTable2Widths(doc, body);
  normalizeTableTitles(doc, body);

  onProgress('更新表目次…', 70);
  rebuildTOC(doc, body);

  onProgress('序列化輸出…', 85);

  const serialized = serializeXML(doc);
  // Ensure XML declaration
  const finalXml = serialized.startsWith('<?xml')
    ? serialized
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${serialized}`;

  out['word/document.xml'] = new TextEncoder().encode(finalXml);

  onProgress('壓縮輸出檔案…', 92);

  const outZip = new JSZip();
  for (const [name, data] of Object.entries(out)) {
    outZip.file(name, data);
  }
  const blob = await outZip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 6 } });

  onProgress('完成！', 100);
  return blob;
}

// Export
window.AAFCOConverter = { convert };
