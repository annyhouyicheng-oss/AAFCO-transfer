// AAFCO DOCX 格式轉換引擎 v3
// 完全對應「排版用這份檔案格式.docx」模板

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function wt(n) { return `{${W_NS}}${n}`; }
function wa(n) { return `w:${n}`; }

function parseXML(str) {
  return new DOMParser().parseFromString(str, 'application/xml');
}
function serXML(doc) {
  return new XMLSerializer().serializeToString(doc);
}

function goc(parent, localName) {
  let el = null;
  const children = parent.childNodes;
  for (let i = 0; i < children.length; i++) {
    if (children[i].localName === localName && children[i].namespaceURI === W_NS) {
      el = children[i]; break;
    }
  }
  if (!el) {
    el = parent.ownerDocument.createElementNS(W_NS, 'w:' + localName);
    parent.appendChild(el);
  }
  return el;
}

function rm(parent, localName) {
  const children = [...parent.childNodes];
  for (const c of children) {
    if (c.localName === localName && c.namespaceURI === W_NS) parent.removeChild(c);
  }
}

function sw(el, attr, val) { el.setAttributeNS(W_NS, 'w:' + attr, String(val)); }

function ensurePPr(para) {
  let pPr = null;
  for (const c of para.childNodes) {
    if (c.localName === 'pPr' && c.namespaceURI === W_NS) { pPr = c; break; }
  }
  if (!pPr) {
    pPr = para.ownerDocument.createElementNS(W_NS, 'w:pPr');
    para.insertBefore(pPr, para.firstChild);
  }
  return pPr;
}

function ensureRPr(run) {
  let rPr = null;
  for (const c of run.childNodes) {
    if (c.localName === 'rPr' && c.namespaceURI === W_NS) { rPr = c; break; }
  }
  if (!rPr) {
    rPr = run.ownerDocument.createElementNS(W_NS, 'w:rPr');
    run.insertBefore(rPr, run.firstChild);
  }
  return rPr;
}

function setFont(rPr, ascii, eastAsia) {
  const rFonts = goc(rPr, 'rFonts');
  rFonts.setAttributeNS(W_NS, 'w:ascii', ascii);
  rFonts.setAttributeNS(W_NS, 'w:hAnsi', ascii);
  rFonts.setAttributeNS(W_NS, 'w:eastAsia', eastAsia || ascii);
  rFonts.setAttributeNS(W_NS, 'w:cs', ascii);
}

function setColorAuto(rPr) {
  const c = goc(rPr, 'color');
  c.setAttributeNS(W_NS, 'w:val', 'auto');
  // remove theme attrs
  [...c.attributes].forEach(a => { if (a.name.includes('theme') || a.name.includes('Theme')) c.removeAttribute(a.name); });
}

function setSz(rPr, sz) {
  goc(rPr, 'sz').setAttributeNS(W_NS, 'w:val', String(sz));
  goc(rPr, 'szCs').setAttributeNS(W_NS, 'w:val', String(sz));
}

function getDirectRuns(para) {
  return [...para.childNodes].filter(c => c.localName === 'r' && c.namespaceURI === W_NS);
}

function getStyle(para) {
  for (const c of para.childNodes) {
    if (c.localName === 'pPr' && c.namespaceURI === W_NS) {
      for (const cc of c.childNodes) {
        if (cc.localName === 'pStyle' && cc.namespaceURI === W_NS)
          return cc.getAttributeNS(W_NS, 'val') || 'Normal';
      }
    }
  }
  return 'Normal';
}

function getText(para) {
  let t = '';
  const ts = para.getElementsByTagNameNS(W_NS, 't');
  for (let i = 0; i < ts.length; i++) t += ts[i].textContent || '';
  return t.trim();
}

function stripRunColors(para) {
  const runs = para.getElementsByTagNameNS(W_NS, 'r');
  for (let i = 0; i < runs.length; i++) {
    const rPr = runs[i].getElementsByTagNameNS(W_NS, 'rPr')[0];
    if (!rPr) continue;
    const color = rPr.getElementsByTagNameNS(W_NS, 'color')[0];
    if (color) {
      color.setAttributeNS(W_NS, 'w:val', 'auto');
      [...color.attributes].forEach(a => { if (a.name.includes('theme')) color.removeAttribute(a.name); });
    }
    rm(rPr, 'u'); rm(rPr, 'highlight'); rm(rPr, 'shd'); rm(rPr, 'rStyle');
  }
}

function applyRunBase(rPr, ascii, eastAsia, sz) {
  setFont(rPr, ascii, eastAsia);
  setColorAuto(rPr);
  if (sz) setSz(rPr, sz);
  rm(rPr, 'u'); rm(rPr, 'highlight'); rm(rPr, 'shd'); rm(rPr, 'rStyle');
}

function applyH1(para) {
  const pPr = ensurePPr(para);
  const ps = goc(pPr, 'pStyle'); sw(ps, 'val', '1');
  const sp = goc(pPr, 'spacing'); sw(sp, 'line', '480'); sw(sp, 'lineRule', 'auto'); sw(sp, 'before', '0'); sw(sp, 'after', '0');
  const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '0');
  rm(pPr, 'jc');
  const pRPr = goc(pPr, 'rPr'); setFont(pRPr, 'Times New Roman', '標楷體'); setColorAuto(pRPr); setSz(pRPr, 28);
  getDirectRuns(para).forEach(r => { const rPr = ensureRPr(r); applyRunBase(rPr, 'Times New Roman', '標楷體', 28); rm(rPr, 'b'); rm(rPr, 'bCs'); });
}

function applyH2(para) {
  const pPr = ensurePPr(para);
  const ps = goc(pPr, 'pStyle'); sw(ps, 'val', '2');
  const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '0');
  const pRPr = goc(pPr, 'rPr'); setFont(pRPr, 'Times New Roman', '標楷體'); setColorAuto(pRPr); setSz(pRPr, 24);
  getDirectRuns(para).forEach(r => { const rPr = ensureRPr(r); applyRunBase(rPr, 'Times New Roman', '標楷體', 24); rm(rPr, 'b'); rm(rPr, 'bCs'); });
}

function applyBody(para) {
  const pPr = ensurePPr(para);
  rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing'); sw(sp, 'line', '480'); sw(sp, 'lineRule', 'auto'); sw(sp, 'before', '80'); sw(sp, 'after', '80');
  const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '480'); ind.removeAttributeNS(W_NS, 'hanging');
  const jc = goc(pPr, 'jc'); sw(jc, 'val', 'both');
  const pRPr = goc(pPr, 'rPr'); setFont(pRPr, 'Times New Roman', '標楷體'); setColorAuto(pRPr); setSz(pRPr, 24);
  getDirectRuns(para).forEach(r => { const rPr = ensureRPr(r); applyRunBase(rPr, 'Times New Roman', '標楷體', 24); rm(rPr, 'b'); rm(rPr, 'bCs'); });
}

function applyCover(para) {
  const pPr = ensurePPr(para);
  rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing'); sw(sp, 'line', '400'); sw(sp, 'lineRule', 'exact');
  const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '0');
  const jc = goc(pPr, 'jc'); sw(jc, 'val', 'center');
  stripRunColors(para);
}

function applyTOC(para) {
  const pPr = ensurePPr(para);
  rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing'); sw(sp, 'line', '480'); sw(sp, 'lineRule', 'auto'); sw(sp, 'before', '40'); sw(sp, 'after', '40');
  const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '0');
  getDirectRuns(para).forEach(r => { const rPr = ensureRPr(r); applyRunBase(rPr, 'Times New Roman', '標楷體', 24); rm(rPr, 'b'); rm(rPr, 'bCs'); });
}

function applyEmpty(para) {
  const pPr = ensurePPr(para);
  rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing'); sw(sp, 'line', '480'); sw(sp, 'lineRule', 'auto');
}

function applyTable(tbl) {
  // tblPr
  let tblPr = null;
  for (const c of tbl.childNodes) { if (c.localName === 'tblPr' && c.namespaceURI === W_NS) { tblPr = c; break; } }
  if (!tblPr) { tblPr = tbl.ownerDocument.createElementNS(W_NS, 'w:tblPr'); tbl.insertBefore(tblPr, tbl.firstChild); }
  rm(tblPr, 'tblBorders'); rm(tblPr, 'tblStyle');
  const tblW = goc(tblPr, 'tblW'); sw(tblW, 'w', '9355'); sw(tblW, 'type', 'dxa');
  const jcT = goc(tblPr, 'jc'); sw(jcT, 'val', 'center');
  const tblCellMar = goc(tblPr, 'tblCellMar');
  const ml = goc(tblCellMar, 'left'); sw(ml, 'w', '10'); sw(ml, 'type', 'dxa');
  const mr = goc(tblCellMar, 'right'); sw(mr, 'w', '10'); sw(mr, 'type', 'dxa');

  const rows = [...tbl.childNodes].filter(c => c.localName === 'tr' && c.namespaceURI === W_NS);
  rows.forEach((row, ri) => {
    const isHeader = ri === 0;
    // trPr
    let trPr = null;
    for (const c of row.childNodes) { if (c.localName === 'trPr' && c.namespaceURI === W_NS) { trPr = c; break; } }
    if (!trPr) { trPr = row.ownerDocument.createElementNS(W_NS, 'w:trPr'); row.insertBefore(trPr, row.firstChild); }
    if (isHeader) { const jcR = goc(trPr, 'jc'); sw(jcR, 'val', 'center'); }

    const cells = [...row.childNodes].filter(c => c.localName === 'tc' && c.namespaceURI === W_NS);
    cells.forEach(cell => {
      let tcPr = null;
      for (const c of cell.childNodes) { if (c.localName === 'tcPr' && c.namespaceURI === W_NS) { tcPr = c; break; } }
      if (!tcPr) { tcPr = cell.ownerDocument.createElementNS(W_NS, 'w:tcPr'); cell.insertBefore(tcPr, cell.firstChild); }

      // tcBorders: header=bottom only, data=top only
      rm(tcPr, 'tcBorders');
      const tcBorders = tcPr.ownerDocument.createElementNS(W_NS, 'w:tcBorders');
      tcPr.appendChild(tcBorders);
      const bSide = isHeader ? 'bottom' : 'top';
      const bEl = tcBorders.ownerDocument.createElementNS(W_NS, 'w:' + bSide);
      bEl.setAttributeNS(W_NS, 'w:val', 'single'); bEl.setAttributeNS(W_NS, 'w:sz', '4');
      bEl.setAttributeNS(W_NS, 'w:space', '0'); bEl.setAttributeNS(W_NS, 'w:color', 'auto');
      tcBorders.appendChild(bEl);

      // shd: clear, fill=auto (remove all color)
      const shd = goc(tcPr, 'shd');
      shd.setAttributeNS(W_NS, 'w:val', 'clear');
      shd.setAttributeNS(W_NS, 'w:color', 'auto');
      shd.setAttributeNS(W_NS, 'w:fill', 'auto');

      // tcMar
      const tcMar = goc(tcPr, 'tcMar');
      [['top','80'],['left','120'],['bottom','80'],['right','120']].forEach(([side, val]) => {
        const m = goc(tcMar, side); sw(m, 'w', val); sw(m, 'type', 'dxa');
      });

      // Paragraphs
      const paras = [...cell.childNodes].filter(c => c.localName === 'p' && c.namespaceURI === W_NS);
      paras.forEach(para => {
        const pPr = ensurePPr(para);
        rm(pPr, 'pStyle'); rm(pPr, 'jc');
        const wc = goc(pPr, 'widowControl'); sw(wc, 'val', '0');
        const sp = goc(pPr, 'spacing'); sw(sp, 'before', '60'); sw(sp, 'after', '60'); sw(sp, 'line', '240'); sw(sp, 'lineRule', 'auto');
        const ind = goc(pPr, 'ind'); sw(ind, 'firstLine', '0');
        const pRPr = goc(pPr, 'rPr');
        setFont(pRPr, '標楷體', '標楷體'); setColorAuto(pRPr);
        goc(pRPr, 'kern').setAttributeNS(W_NS, 'w:val', '0');

        getDirectRuns(para).forEach(run => {
          const rPr = ensureRPr(run);
          setFont(rPr, '標楷體', '標楷體'); setColorAuto(rPr);
          goc(rPr, 'kern').setAttributeNS(W_NS, 'w:val', '0');
          rm(rPr, 'b'); rm(rPr, 'bCs'); rm(rPr, 'u'); rm(rPr, 'highlight'); rm(rPr, 'shd'); rm(rPr, 'rStyle');
          rm(rPr, 'sz'); rm(rPr, 'szCs');
          const c = rPr.getElementsByTagNameNS(W_NS, 'color')[0];
          if (c) c.setAttributeNS(W_NS, 'w:val', 'auto');
        });
      });
    });
  });
}

function updateSectPr(sectPr, cfg, footerRid) {
  const doc = sectPr.ownerDocument;
  const pgSz = goc(sectPr, 'pgSz');
  sw(pgSz, 'w', '11906'); sw(pgSz, 'h', '16838');
  pgSz.removeAttributeNS(W_NS, 'orient');
  const pgMar = goc(sectPr, 'pgMar');
  Object.entries(cfg).forEach(([k, v]) => sw(pgMar, k, String(v)));
  rm(sectPr, 'footerReference');
  const fr = doc.createElementNS(W_NS, 'w:footerReference');
  fr.setAttributeNS(W_NS, 'w:type', 'default');
  fr.setAttributeNS(R_NS, 'r:id', footerRid);
  sectPr.insertBefore(fr, sectPr.firstChild);
  rm(sectPr, 'headerReference');
}

const COVER_CFG = {top:'1418',right:'1418',bottom:'1418',left:'1701',header:'851',footer:'992',gutter:'0'};
const MAIN_CFG  = {top:'1440',right:'1797',bottom:'1440',left:'1797',header:'851',footer:'992',gutter:'0'};

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>`;

const STYLES_OVERRIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/>
</w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="425"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
  <w:name w:val="Normal"/>
  <w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="425"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
  <w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="1">
  <w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
  <w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/>
  <w:color w:val="000000"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
</w:style>
<w:style w:type="paragraph" w:styleId="2">
  <w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>
  <w:pPr><w:spacing w:line="480" w:lineRule="auto"/><w:ind w:firstLine="0"/><w:outlineLvl w:val="1"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="標楷體" w:hAnsi="Times New Roman"/>
  <w:color w:val="000000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
</w:style>
</w:styles>`;

async function convertDocx(file, onLog) {
  onLog('讀取 DOCX...');
  const ab = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(ab);

  onLog('解析文件結構...');
  const docXml = await zip.file('word/document.xml').async('string');
  const doc = parseXML(docXml);
  const body = doc.getElementsByTagNameNS(W_NS, 'body')[0];
  const children = [...body.childNodes].filter(n => n.nodeType === 1);

  onLog('套用段落格式...');
  const stats = {cover:0, h1:0, h2:0, body:0, toc:0, empty:0, table:0};
  let paraIdx = 0;

  for (const el of children) {
    if (el.localName === 'p' && el.namespaceURI === W_NS) {
      const style = getStyle(el);
      const text = getText(el);
      if (!text) { applyEmpty(el); stats.empty++; }
      else if (style === 'Heading1' || style === 'Heading 1') { applyH1(el); stats.h1++; }
      else if (style === 'Heading2' || style === 'Heading 2' || style === 'Heading3') { applyH2(el); stats.h2++; }
      else if (paraIdx < 17) { applyCover(el); stats.cover++; }
      else if (text.includes('……')) { applyTOC(el); stats.toc++; }
      else { applyBody(el); stats.body++; }
      paraIdx++;
    } else if (el.localName === 'tbl' && el.namespaceURI === W_NS) {
      applyTable(el); stats.table++;
    }
  }
  onLog(`段落: H1=${stats.h1} H2=${stats.h2} 正文=${stats.body} 表格=${stats.table}`);

  // sectPr
  onLog('更新頁面設定...');
  let footerRid = 'rId_f1';
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
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

  const allSect = [...doc.getElementsByTagNameNS(W_NS, 'sectPr')];
  allSect.forEach((sp, i) => {
    const cfg = (allSect.length > 1 && i < allSect.length - 1) ? COVER_CFG : MAIN_CFG;
    updateSectPr(sp, cfg, footerRid);
  });

  onLog('寫入樣式與頁尾...');
  zip.file('word/document.xml', serXML(doc));
  zip.file('word/styles.xml', STYLES_OVERRIDE);
  zip.file('word/footer1.xml', FOOTER_XML);

  let ctXml = await zip.file('[Content_Types].xml').async('string');
  if (!ctXml.includes('footer1.xml')) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
    zip.file('[Content_Types].xml', ctXml);
  }

  onLog('產生輸出檔案...');
  const blob = await zip.generateAsync({type:'blob', compression:'DEFLATE'});
  return blob;
}

window.convertDocx = convertDocx;
