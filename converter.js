// AAFCO DOCX 格式轉換引擎 v4 - 完全精確修正版
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function parseXML(str) { return new DOMParser().parseFromString(str, 'application/xml'); }
function serXML(doc) { return new XMLSerializer().serializeToString(doc); }

function goc(parent, localName) {
  for (const c of parent.childNodes)
    if (c.localName === localName && c.namespaceURI === W_NS) return c;
  const el = parent.ownerDocument.createElementNS(W_NS, 'w:' + localName);
  parent.appendChild(el);
  return el;
}
function rm(parent, localName) {
  [...parent.childNodes].forEach(c => { if (c.localName === localName && c.namespaceURI === W_NS) parent.removeChild(c); });
}
function sw(el, attr, val) { el.setAttributeNS(W_NS, 'w:' + attr, String(val)); }
function rmAttr(el, attr) { el.removeAttributeNS(W_NS, attr); }

function ensurePPr(para) {
  for (const c of para.childNodes) if (c.localName === 'pPr' && c.namespaceURI === W_NS) return c;
  const pPr = para.ownerDocument.createElementNS(W_NS, 'w:pPr');
  para.insertBefore(pPr, para.firstChild);
  return pPr;
}
function ensureRPr(run) {
  for (const c of run.childNodes) if (c.localName === 'rPr' && c.namespaceURI === W_NS) return c;
  const rPr = run.ownerDocument.createElementNS(W_NS, 'w:rPr');
  run.insertBefore(rPr, run.firstChild);
  return rPr;
}
function getStyle(para) {
  for (const c of para.childNodes)
    if (c.localName === 'pPr' && c.namespaceURI === W_NS)
      for (const cc of c.childNodes)
        if (cc.localName === 'pStyle' && cc.namespaceURI === W_NS)
          return cc.getAttributeNS(W_NS, 'val') || 'Normal';
  return 'Normal';
}
function getText(para) {
  let t = '';
  const ts = para.getElementsByTagNameNS(W_NS, 't');
  for (let i=0;i<ts.length;i++) t += ts[i].textContent||'';
  return t.trim();
}
function directRuns(para) { return [...para.childNodes].filter(c=>c.localName==='r'&&c.namespaceURI===W_NS); }
function clearEl(el) { while(el.firstChild) el.removeChild(el.firstChild); }
function makeEl(doc, localName) { return doc.createElementNS(W_NS, 'w:'+localName); }

// ── 修正1：正文（hint=eastAsia 最簡化，移除 i/iCs/b/bCs/color/sz）
function applyBody(para) {
  const pPr = ensurePPr(para);
  rm(pPr, 'pStyle');
  const sp = goc(pPr, 'spacing'); sw(sp,'before','80'); sw(sp,'after','80');
  rmAttr(sp,'line'); rmAttr(sp,'lineRule');
  const ind = goc(pPr, 'ind'); sw(ind,'firstLine','480'); rmAttr(ind,'hanging'); rmAttr(ind,'left');
  const jc = goc(pPr, 'jc'); sw(jc,'val','both');
  const pRPr = goc(pPr, 'rPr'); clearEl(pRPr);
  goc(pRPr,'color').setAttributeNS(W_NS,'w:val','auto');
  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    const rF = makeEl(run.ownerDocument,'rFonts');
    rF.setAttributeNS(W_NS,'w:hint','eastAsia');
    rPr.appendChild(rF);
  });
}

// ── 修正2：H1（顏色 000000+themeColor=text1，無 i/iCs）
function applyH1(para) {
  const pPr = ensurePPr(para);
  const ps = goc(pPr,'pStyle'); sw(ps,'val','1');
  // 修正1: spacing 只保留 line/lineRule，無 before/after
  const sp = goc(pPr,'spacing'); sw(sp,'line','480'); sw(sp,'lineRule','auto');
  rmAttr(sp,'before'); rmAttr(sp,'after');
  const ind = goc(pPr,'ind'); sw(ind,'firstLine','0');
  rm(pPr,'jc');
  const pRPr = goc(pPr,'rPr'); clearEl(pRPr);
  const rF0 = makeEl(para.ownerDocument,'rFonts');
  rF0.setAttributeNS(W_NS,'w:ascii','Times New Roman'); rF0.setAttributeNS(W_NS,'w:hAnsi','Times New Roman');
  pRPr.appendChild(rF0);
  const c0 = makeEl(para.ownerDocument,'color');
  c0.setAttributeNS(W_NS,'w:val','000000'); c0.setAttributeNS(W_NS,'w:themeColor','text1');
  pRPr.appendChild(c0);
  // 修正2: sz 只加 sz，不加 szCs（模板無 szCs）
  const s0 = makeEl(para.ownerDocument,'sz'); s0.setAttributeNS(W_NS,'w:val','28'); pRPr.appendChild(s0);
  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    const rF = makeEl(run.ownerDocument,'rFonts');
    rF.setAttributeNS(W_NS,'w:ascii','Times New Roman'); rF.setAttributeNS(W_NS,'w:hAnsi','Times New Roman');
    rPr.appendChild(rF);
    const c = makeEl(run.ownerDocument,'color');
    c.setAttributeNS(W_NS,'w:val','000000'); c.setAttributeNS(W_NS,'w:themeColor','text1');
    rPr.appendChild(c);
    // 修正2: 不加 szCs
    const s = makeEl(run.ownerDocument,'sz'); s.setAttributeNS(W_NS,'w:val','28'); rPr.appendChild(s);
  });
}

// ── 修正3：H2（numPr numId=3，移除 spacing，bold=0，顏色 000000+themeColor）
function applyH2(para) {
  const pPr = ensurePPr(para);
  const ps = goc(pPr,'pStyle'); sw(ps,'val','2');
  rm(pPr,'spacing'); rm(pPr,'numPr'); rm(pPr,'ind');
  // insert numPr after pStyle
  const numPr = makeEl(para.ownerDocument,'numPr');
  const ilvl = makeEl(para.ownerDocument,'ilvl'); ilvl.setAttributeNS(W_NS,'w:val','0'); numPr.appendChild(ilvl);
  const numId = makeEl(para.ownerDocument,'numId'); numId.setAttributeNS(W_NS,'w:val','3'); numPr.appendChild(numId);
  let psIdx = 0;
  [...pPr.childNodes].forEach((c,i) => { if(c.localName==='pStyle') psIdx=i+1; });
  const siblings = [...pPr.childNodes];
  if (psIdx < siblings.length) pPr.insertBefore(numPr, siblings[psIdx]);
  else pPr.appendChild(numPr);

  const pRPr = goc(pPr,'rPr'); clearEl(pRPr);
  const rF0 = makeEl(para.ownerDocument,'rFonts');
  rF0.setAttributeNS(W_NS,'w:ascii','Times New Roman'); rF0.setAttributeNS(W_NS,'w:hAnsi','Times New Roman');
  pRPr.appendChild(rF0);
  const b0 = makeEl(para.ownerDocument,'b'); b0.setAttributeNS(W_NS,'w:val','0'); pRPr.appendChild(b0);
  const bc0 = makeEl(para.ownerDocument,'bCs'); bc0.setAttributeNS(W_NS,'w:val','0'); pRPr.appendChild(bc0);
  const c0 = makeEl(para.ownerDocument,'color'); c0.setAttributeNS(W_NS,'w:val','000000'); c0.setAttributeNS(W_NS,'w:themeColor','text1'); pRPr.appendChild(c0);
  const s0 = makeEl(para.ownerDocument,'sz'); s0.setAttributeNS(W_NS,'w:val','24'); pRPr.appendChild(s0);
  const sc0 = makeEl(para.ownerDocument,'szCs'); sc0.setAttributeNS(W_NS,'w:val','24'); pRPr.appendChild(sc0);

  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    const rF = makeEl(run.ownerDocument,'rFonts');
    rF.setAttributeNS(W_NS,'w:ascii','Times New Roman'); rF.setAttributeNS(W_NS,'w:hAnsi','Times New Roman'); rF.setAttributeNS(W_NS,'w:hint','eastAsia');
    rPr.appendChild(rF);
    const b = makeEl(run.ownerDocument,'b'); b.setAttributeNS(W_NS,'w:val','0'); rPr.appendChild(b);
    const bc = makeEl(run.ownerDocument,'bCs'); bc.setAttributeNS(W_NS,'w:val','0'); rPr.appendChild(bc);
    const c = makeEl(run.ownerDocument,'color'); c.setAttributeNS(W_NS,'w:val','000000'); c.setAttributeNS(W_NS,'w:themeColor','text1'); rPr.appendChild(c);
    const s = makeEl(run.ownerDocument,'sz'); s.setAttributeNS(W_NS,'w:val','24'); rPr.appendChild(s);
    const sc = makeEl(run.ownerDocument,'szCs'); sc.setAttributeNS(W_NS,'w:val','24'); rPr.appendChild(sc);
  });
}

function applyCover(para) {
  const pPr = ensurePPr(para); rm(pPr,'pStyle');
  const sp = goc(pPr,'spacing'); sw(sp,'line','400'); sw(sp,'lineRule','exact');
  const ind = goc(pPr,'ind'); sw(ind,'firstLine','0');
  const jc = goc(pPr,'jc'); sw(jc,'val','center');
  const pRPr = goc(pPr,'rPr'); clearEl(pRPr);
  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run);
    ['i','iCs','u','highlight','shd','rStyle'].forEach(t => rm(rPr,t));
    const color = [...rPr.childNodes].find(c=>c.localName==='color');
    if (color) { color.setAttributeNS(W_NS,'w:val','000000'); color.setAttributeNS(W_NS,'w:themeColor','text1'); }
  });
}

function applyTOC(para) {
  const pPr = ensurePPr(para); rm(pPr,'pStyle');
  const sp = goc(pPr,'spacing'); sw(sp,'line','480'); sw(sp,'lineRule','auto'); sw(sp,'before','40'); sw(sp,'after','40');
  const ind = goc(pPr,'ind'); sw(ind,'firstLine','0');
  directRuns(para).forEach(run => {
    const rPr = ensureRPr(run); clearEl(rPr);
    const rF = makeEl(run.ownerDocument,'rFonts');
    rF.setAttributeNS(W_NS,'w:ascii','Times New Roman'); rF.setAttributeNS(W_NS,'w:hAnsi','Times New Roman'); rF.setAttributeNS(W_NS,'w:hint','eastAsia');
    rPr.appendChild(rF);
  });
}

function applyEmpty(para) {
  const pPr = ensurePPr(para); rm(pPr,'pStyle');
  const sp = goc(pPr,'spacing'); sw(sp,'line','480'); sw(sp,'lineRule','auto');
}

// ── 修正4：表格（tblLook, hideMark, cs=新細明體, bCs）
function applyTable(tbl) {
  let tblPr = [...tbl.childNodes].find(c=>c.localName==='tblPr'&&c.namespaceURI===W_NS);
  if (!tblPr) { tblPr = tbl.ownerDocument.createElementNS(W_NS,'w:tblPr'); tbl.insertBefore(tblPr,tbl.firstChild); }
  rm(tblPr,'tblBorders'); rm(tblPr,'tblStyle'); rm(tblPr,'tblLook');
  const tblW = goc(tblPr,'tblW'); sw(tblW,'w','9355'); sw(tblW,'type','dxa');
  const jcT = goc(tblPr,'jc'); sw(jcT,'val','center');
  const tcm = goc(tblPr,'tblCellMar');
  const ml = goc(tcm,'left'); sw(ml,'w','10'); sw(ml,'type','dxa');
  const mr = goc(tcm,'right'); sw(mr,'w','10'); sw(mr,'type','dxa');
  const tblLook = tbl.ownerDocument.createElementNS(W_NS,'w:tblLook');
  tblLook.setAttributeNS(W_NS,'w:val','04A0'); tblLook.setAttributeNS(W_NS,'w:firstRow','1');
  tblLook.setAttributeNS(W_NS,'w:lastRow','0'); tblLook.setAttributeNS(W_NS,'w:firstColumn','1');
  tblLook.setAttributeNS(W_NS,'w:lastColumn','0'); tblLook.setAttributeNS(W_NS,'w:noHBand','0');
  tblLook.setAttributeNS(W_NS,'w:noVBand','1'); tblPr.appendChild(tblLook);

  const rows = [...tbl.childNodes].filter(c=>c.localName==='tr'&&c.namespaceURI===W_NS);
  rows.forEach((row,ri) => {
    const isHeader = ri===0;
    let trPr = [...row.childNodes].find(c=>c.localName==='trPr'&&c.namespaceURI===W_NS);
    if (!trPr) { trPr = row.ownerDocument.createElementNS(W_NS,'w:trPr'); row.insertBefore(trPr,row.firstChild); }
    if (isHeader) { const jcR = goc(trPr,'jc'); sw(jcR,'val','center'); }

    const cells = [...row.childNodes].filter(c=>c.localName==='tc'&&c.namespaceURI===W_NS);
    cells.forEach(cell => {
      let tcPr = [...cell.childNodes].find(c=>c.localName==='tcPr'&&c.namespaceURI===W_NS);
      if (!tcPr) { tcPr = cell.ownerDocument.createElementNS(W_NS,'w:tcPr'); cell.insertBefore(tcPr,cell.firstChild); }
      rm(tcPr,'tcBorders'); rm(tcPr,'hideMark');
      const tcBorders = cell.ownerDocument.createElementNS(W_NS,'w:tcBorders');
      const bSide = isHeader?'bottom':'top';
      const bEl = cell.ownerDocument.createElementNS(W_NS,'w:'+bSide);
      bEl.setAttributeNS(W_NS,'w:val','single'); bEl.setAttributeNS(W_NS,'w:sz','4');
      bEl.setAttributeNS(W_NS,'w:space','0'); bEl.setAttributeNS(W_NS,'w:color','auto');
      tcBorders.appendChild(bEl);
      const tcW = [...tcPr.childNodes].find(c=>c.localName==='tcW');
      if (tcW) tcPr.insertBefore(tcBorders, tcW.nextSibling);
      else tcPr.insertBefore(tcBorders, tcPr.firstChild);
      const shd = goc(tcPr,'shd');
      shd.setAttributeNS(W_NS,'w:val','clear'); shd.setAttributeNS(W_NS,'w:color','auto'); shd.setAttributeNS(W_NS,'w:fill','auto');
      const tcMar = goc(tcPr,'tcMar');
      [['top','80'],['left','120'],['bottom','80'],['right','120']].forEach(([s,v])=>{ const m=goc(tcMar,s);sw(m,'w',v);sw(m,'type','dxa'); });
      // 修正3: 所有列所有儲存格都加 hideMark（模板行為）
      rm(tcPr,'hideMark');
      tcPr.appendChild(cell.ownerDocument.createElementNS(W_NS,'w:hideMark'));

      const paras = [...cell.childNodes].filter(c=>c.localName==='p'&&c.namespaceURI===W_NS);
      paras.forEach(para => {
        const pPr = ensurePPr(para); rm(pPr,'pStyle'); rm(pPr,'jc');
        const wc = goc(pPr,'widowControl'); sw(wc,'val','0');
        const sp = goc(pPr,'spacing'); sw(sp,'before','60'); sw(sp,'after','60'); sw(sp,'line','240'); sw(sp,'lineRule','auto');
        const ind = goc(pPr,'ind'); sw(ind,'firstLine','0');
        const pRPr = goc(pPr,'rPr'); clearEl(pRPr);
        const rFp = makeEl(para.ownerDocument,'rFonts');
        rFp.setAttributeNS(W_NS,'w:ascii','標楷體'); rFp.setAttributeNS(W_NS,'w:hAnsi','標楷體'); rFp.setAttributeNS(W_NS,'w:cs','標楷體');
        pRPr.appendChild(rFp);
        goc(pRPr,'color').setAttributeNS(W_NS,'w:val','auto');
        goc(pRPr,'kern').setAttributeNS(W_NS,'w:val','0');

        directRuns(para).forEach(run => {
          const rPr = ensureRPr(run); clearEl(rPr);
          const rF = makeEl(run.ownerDocument,'rFonts');
          rF.setAttributeNS(W_NS,'w:ascii','標楷體'); rF.setAttributeNS(W_NS,'w:hAnsi','標楷體');
          rF.setAttributeNS(W_NS,'w:cs', isHeader?'新細明體':'標楷體');
          rF.setAttributeNS(W_NS,'w:hint','eastAsia');
          rPr.appendChild(rF);
          if (isHeader) { const bCs=makeEl(run.ownerDocument,'bCs'); rPr.appendChild(bCs); }
          goc(rPr,'color').setAttributeNS(W_NS,'w:val','auto');
          goc(rPr,'kern').setAttributeNS(W_NS,'w:val','0');
        });
      });
    });
  });
}

// ── 修正5：sectPr（2個，封面區 upperRoman + top=1418，主文 start=1 + top=1440）
function buildCoverSectPr(doc, footerRid) {
  const sp = doc.createElementNS(W_NS,'w:sectPr');
  if (footerRid) {
    const fr = doc.createElementNS(W_NS,'w:footerReference');
    fr.setAttributeNS(W_NS,'w:type','default'); fr.setAttributeNS(R_NS,'r:id',footerRid); sp.appendChild(fr);
  }
  const pgSz = doc.createElementNS(W_NS,'w:pgSz');
  pgSz.setAttributeNS(W_NS,'w:w','11906'); pgSz.setAttributeNS(W_NS,'w:h','16838'); sp.appendChild(pgSz);
  const pgMar = doc.createElementNS(W_NS,'w:pgMar');
  pgMar.setAttributeNS(W_NS,'w:top','1418'); pgMar.setAttributeNS(W_NS,'w:right','1418');
  pgMar.setAttributeNS(W_NS,'w:bottom','1418'); pgMar.setAttributeNS(W_NS,'w:left','1701');
  pgMar.setAttributeNS(W_NS,'w:header','851'); pgMar.setAttributeNS(W_NS,'w:footer','992');
  pgMar.setAttributeNS(W_NS,'w:gutter','0'); sp.appendChild(pgMar);
  const pgNum = doc.createElementNS(W_NS,'w:pgNumType');
  pgNum.setAttributeNS(W_NS,'w:fmt','upperRoman'); pgNum.setAttributeNS(W_NS,'w:start','1'); sp.appendChild(pgNum);
  const cols = doc.createElementNS(W_NS,'w:cols'); cols.setAttributeNS(W_NS,'w:space','425'); sp.appendChild(cols);
  const dg = doc.createElementNS(W_NS,'w:docGrid');
  dg.setAttributeNS(W_NS,'w:type','lines'); dg.setAttributeNS(W_NS,'w:linePitch','360'); sp.appendChild(dg);
  return sp;
}

function updateMainSectPr(sectPr) {
  while (sectPr.firstChild) sectPr.removeChild(sectPr.firstChild);
  const doc = sectPr.ownerDocument;
  const pgSz = doc.createElementNS(W_NS,'w:pgSz');
  pgSz.setAttributeNS(W_NS,'w:w','11906'); pgSz.setAttributeNS(W_NS,'w:h','16838'); sectPr.appendChild(pgSz);
  const pgMar = doc.createElementNS(W_NS,'w:pgMar');
  pgMar.setAttributeNS(W_NS,'w:top','1440'); pgMar.setAttributeNS(W_NS,'w:right','1797');
  pgMar.setAttributeNS(W_NS,'w:bottom','1440'); pgMar.setAttributeNS(W_NS,'w:left','1797');
  pgMar.setAttributeNS(W_NS,'w:header','851'); pgMar.setAttributeNS(W_NS,'w:footer','992');
  pgMar.setAttributeNS(W_NS,'w:gutter','0'); sectPr.appendChild(pgMar);
  const pgNum = doc.createElementNS(W_NS,'w:pgNumType'); pgNum.setAttributeNS(W_NS,'w:start','1'); sectPr.appendChild(pgNum);
  const cols = doc.createElementNS(W_NS,'w:cols'); cols.setAttributeNS(W_NS,'w:space','720'); sectPr.appendChild(cols);
  const dg = doc.createElementNS(W_NS,'w:docGrid');
  dg.setAttributeNS(W_NS,'w:type','lines'); dg.setAttributeNS(W_NS,'w:linePitch','360'); sectPr.appendChild(dg);
}

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:p><w:pPr><w:jc w:val="center"/></w:pPr>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:p></w:ftr>`;

async function convertDocx(file, onLog) {
  onLog('讀取 DOCX...');
  const ab = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(ab);
  onLog('解析文件結構...');
  const docXml = await zip.file('word/document.xml').async('string');
  const doc = parseXML(docXml);
  const body = doc.getElementsByTagNameNS(W_NS,'body')[0];
  const children = [...body.childNodes].filter(n=>n.nodeType===1);

  onLog('套用格式...');
  let paraIdx=0, firstH1=null;
  const stats={cover:0,h1:0,h2:0,body:0,toc:0,empty:0,table:0};

  for (const el of children) {
    if (el.localName==='p'&&el.namespaceURI===W_NS) {
      const style=getStyle(el), text=getText(el);
      if (!text) { applyEmpty(el); stats.empty++; }
      else if (style==='1'||style==='Heading1'||style==='Heading 1') {
        if (!firstH1) firstH1=el;
        applyH1(el); stats.h1++;
      }
      else if (style==='2'||style==='Heading2'||style==='Heading 2'||style==='3'||style==='Heading3') { applyH2(el); stats.h2++; }
      else if (paraIdx<17) { applyCover(el); stats.cover++; }
      else if (text.includes('……')) { applyTOC(el); stats.toc++; }
      else { applyBody(el); stats.body++; }
      paraIdx++;
    } else if (el.localName==='tbl'&&el.namespaceURI===W_NS) {
      applyTable(el); stats.table++;
    } else if (el.localName==='sectPr'&&el.namespaceURI===W_NS) {
      updateMainSectPr(el);
    }
  }
  onLog(`H1=${stats.h1} H2=${stats.h2} 正文=${stats.body} 表格=${stats.table}`);

  // 修正5：插入封面 sectPr
  let footerRid = 'rId_f1';
  let relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
  const m = relsXml.match(/Id="(rId\d+)"[^>]+footer1\.xml/);
  if (m) footerRid = m[1];
  else {
    const nums=[...relsXml.matchAll(/rId(\d+)/g)].map(x=>parseInt(x[1]));
    footerRid='rId'+(nums.length?Math.max(...nums)+1:10);
  }
  if (firstH1) {
    const coverPara = doc.createElementNS(W_NS,'w:p');
    const cPPr = doc.createElementNS(W_NS,'w:pPr');
    cPPr.appendChild(buildCoverSectPr(doc, footerRid));
    coverPara.appendChild(cPPr);
    body.insertBefore(coverPara, firstH1);
    onLog('封面分節已插入');
  }

  // rels
  relsXml = relsXml.replace(/<Relationship[^>]+\/header\d*\.xml[^>]*\/>/g,'');
  if (!relsXml.includes('footer1.xml')) {
    relsXml = relsXml.replace('</Relationships>',
      `<Relationship Id="${footerRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`);
  }
  zip.file('word/_rels/document.xml.rels', relsXml);

  onLog('寫入檔案...');
  zip.file('word/document.xml', serXML(doc));
  zip.file('word/footer1.xml', FOOTER_XML);

  let ctXml = await zip.file('[Content_Types].xml').async('string');
  if (!ctXml.includes('footer1.xml')) {
    ctXml = ctXml.replace('</Types>',
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
    zip.file('[Content_Types].xml', ctXml);
  }
  const blob = await zip.generateAsync({type:'blob',compression:'DEFLATE'});
  return blob;
}
window.convertDocx = convertDocx;
