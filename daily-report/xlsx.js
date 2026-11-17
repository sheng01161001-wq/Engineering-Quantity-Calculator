/* 把算好的儲存格寫進範本 Excel（直接改 xlsx 裡的 XML，版面、框線、字型完全照範本） */
(function (root) {
  'use strict';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
  const T_SHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';

  const parse = s => new DOMParser().parseFromString(s, 'application/xml');
  const ser = d => new XMLSerializer().serializeToString(d);
  const colNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
  const refParts = ref => { const m = ref.match(/^([A-Z]+)(\d+)$/); return { col: m[1], row: Number(m[2]) }; };
  const kids = (el, name) => Array.prototype.filter.call(el.childNodes, n => n.nodeType === 1 && n.localName === name);

  /** 字型加粗：同一個樣式只做一次 */
  function boldStyler(stylesDoc) {
    const cache = {};
    const fonts = stylesDoc.getElementsByTagNameNS(NS, 'fonts')[0];
    const xfs = stylesDoc.getElementsByTagNameNS(NS, 'cellXfs')[0];
    return s => {
      s = Number(s || 0);
      if (s in cache) return cache[s];
      const xf = kids(xfs, 'xf')[s];
      const fontList = kids(fonts, 'font');
      const font = fontList[Number(xf.getAttribute('fontId') || 0)].cloneNode(true);
      if (!kids(font, 'b').length) font.insertBefore(stylesDoc.createElementNS(NS, 'b'), font.firstChild);
      fonts.appendChild(font);
      fonts.setAttribute('count', kids(fonts, 'font').length);
      const nx = xf.cloneNode(true);
      nx.setAttribute('fontId', kids(fonts, 'font').length - 1);
      nx.setAttribute('applyFont', '1');
      xfs.appendChild(nx);
      xfs.setAttribute('count', kids(xfs, 'xf').length);
      return (cache[s] = kids(xfs, 'xf').length - 1);
    };
  }

  const colStr = n => { let s = ''; while (n) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - 1 - m) / 26; } return s; };
  /** 公式平移：E10 往下 3 列 → E13（有 $ 的不動；引號裡的字不動） */
  function shiftFormula(f, dRow, dCol) {
    return f.split('"').map((part, i) => i % 2 ? part : part.replace(/(^|[^A-Za-z0-9_$!.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_])/g,
      (m0, pre, cAbs, col, rAbs, row) => pre + cAbs + (cAbs ? col : colStr(colNum(col) + dCol)) + rAbs + (rAbs ? row : String(Number(row) + dRow)))).join('"');
  }
  /** 共用公式（Excel 存檔時把一整欄相同公式只記一次）展開成每格各自的公式，不然改掉第一格後下面的格子會壞 */
  function expandSharedFormulas(doc) {
    const fs = Array.prototype.slice.call(doc.getElementsByTagNameNS(NS, 'f')).filter(f => f.getAttribute('t') === 'shared');
    const master = {};
    fs.forEach(f => { if (f.getAttribute('ref') && f.textContent) master[f.getAttribute('si')] = { text: f.textContent, at: refParts(f.parentNode.getAttribute('r')) }; });
    fs.forEach(f => {
      const m = master[f.getAttribute('si')];
      if (m && !f.textContent) {
        const at = refParts(f.parentNode.getAttribute('r'));
        f.textContent = shiftFormula(m.text, at.row - m.at.row, colNum(at.col) - colNum(m.at.col));
      }
      ['t', 'ref', 'si'].forEach(a => f.removeAttribute(a));
    });
  }

  /** 改一張工作表：cells={A1: 值}，值可以是數字、字串、''（清空）或 {date:'yyyy-mm-dd'} */
  function fillSheet(xml, cells, bold, toBold, excelSerial) {
    const doc = parse(xml);
    expandSharedFormulas(doc);
    const sheetData = doc.getElementsByTagNameNS(NS, 'sheetData')[0];
    const rowEl = r => {
      const rows = kids(sheetData, 'row');
      let el = rows.find(x => Number(x.getAttribute('r')) === r);
      if (el) return el;
      el = doc.createElementNS(NS, 'row'); el.setAttribute('r', r);
      sheetData.insertBefore(el, rows.find(x => Number(x.getAttribute('r')) > r) || null);
      return el;
    };
    const cellEl = ref => {
      const p = refParts(ref), row = rowEl(p.row), cs = kids(row, 'c');
      let el = cs.find(x => x.getAttribute('r') === ref);
      if (el) return el;
      el = doc.createElementNS(NS, 'c'); el.setAttribute('r', ref);
      row.insertBefore(el, cs.find(x => colNum(refParts(x.getAttribute('r')).col) > colNum(p.col)) || null);
      return el;
    };
    Object.keys(cells).forEach(ref => {
      let v = cells[ref];
      const el = cellEl(ref);
      while (el.firstChild) el.removeChild(el.firstChild);
      el.removeAttribute('t');
      if (v && typeof v === 'object' && v.date) v = excelSerial(v.date);
      if (typeof v === 'number') {
        const ve = doc.createElementNS(NS, 'v'); ve.textContent = String(v); el.appendChild(ve);
      } else if (v !== '' && v !== null && v !== undefined) {
        el.setAttribute('t', 'inlineStr');
        const is = doc.createElementNS(NS, 'is'), t = doc.createElementNS(NS, 't');
        t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve'); t.textContent = String(v);
        is.appendChild(t); el.appendChild(is);
      }
    });
    (bold || []).forEach(ref => { const el = cellEl(ref); el.setAttribute('s', toBold(el.getAttribute('s'))); });
    // 公式的舊計算結果拿掉，打開時 Excel 會重算
    Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS, 'f'), f => {
      kids(f.parentNode, 'v').forEach(v => f.parentNode.removeChild(v));
    });
    return doc;
  }

  /**
   * sheets：[{name, cells, bold}]（第一張會是打開時看到的那張）
   * 回傳 Blob（xlsx）
   */
  async function buildWorkbook(templateBuf, sheets, excelSerial) {
    const zip = await JSZip.loadAsync(templateBuf);
    const wbDoc = parse(await zip.file('xl/workbook.xml').async('string'));
    const relsDoc = parse(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
    const ctDoc = parse(await zip.file('[Content_Types].xml').async('string'));
    const stylesDoc = parse(await zip.file('xl/styles.xml').async('string'));
    const toBold = boldStyler(stylesDoc);

    // 範本第一張工作表
    const sheetEl = wbDoc.getElementsByTagNameNS(NS, 'sheet')[0];
    const rid = sheetEl.getAttributeNS(REL_NS, 'id') || sheetEl.getAttribute('r:id');
    const rel = Array.prototype.find.call(relsDoc.getElementsByTagNameNS(PKG_NS, 'Relationship'), r => r.getAttribute('Id') === rid);
    const sheetPath = 'xl/' + rel.getAttribute('Target').replace(/^\/?xl\//, '');
    const sheetXml = await zip.file(sheetPath).async('string');
    const sheetRelsPath = sheetPath.replace(/worksheets\//, 'worksheets/_rels/') + '.rels';
    const sheetRels = zip.file(sheetRelsPath) ? await zip.file(sheetRelsPath).async('string') : null;
    const printArea = Array.prototype.find.call(wbDoc.getElementsByTagNameNS(NS, 'definedName'), d => d.getAttribute('name') === '_xlnm.Print_Area');
    const areaRef = printArea ? printArea.textContent.replace(/^.*!/, '') : null;

    // 拿掉範本原本的工作表、計算鏈
    zip.remove(sheetPath); if (sheetRels) zip.remove(sheetRelsPath);
    zip.remove('xl/calcChain.xml');
    Array.prototype.slice.call(relsDoc.getElementsByTagNameNS(PKG_NS, 'Relationship')).forEach(r => {
      if (r.getAttribute('Type') === T_SHEET || /calcChain$/.test(r.getAttribute('Type'))) r.parentNode.removeChild(r);
    });
    Array.prototype.slice.call(ctDoc.getElementsByTagNameNS(CT_NS, 'Override')).forEach(o => {
      if (/worksheets\/|calcChain/.test(o.getAttribute('PartName'))) o.parentNode.removeChild(o);
    });
    const sheetsEl = wbDoc.getElementsByTagNameNS(NS, 'sheets')[0];
    while (sheetsEl.firstChild) sheetsEl.removeChild(sheetsEl.firstChild);
    let dnEl = wbDoc.getElementsByTagNameNS(NS, 'definedNames')[0];
    const oldName = sheetEl.getAttribute('name');
    if (dnEl) Array.prototype.slice.call(kids(dnEl, 'definedName')).forEach(d => { if (d.hasAttribute('localSheetId') || d.textContent.indexOf(oldName + '!') >= 0 || d.textContent.indexOf("'" + oldName + "'!") >= 0) dnEl.removeChild(d); });
    const binFile = sheetRels && (sheetRels.match(/Target="\.\.\/printerSettings\/([^"]+)"/) || [])[1];
    const bin = binFile ? await zip.file('xl/printerSettings/' + binFile).async('uint8array') : null;
    if (binFile) zip.remove('xl/printerSettings/' + binFile);

    sheets.forEach((sh, i) => {
      const n = i + 1;
      const doc = fillSheet(sheetXml, sh.cells, sh.bold, toBold, excelSerial);
      const ws = doc.documentElement;
      ws.removeAttribute('xr:uid');
      Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS, 'sheetView'), v => { if (i) v.removeAttribute('tabSelected'); else v.setAttribute('tabSelected', '1'); });
      zip.file('xl/worksheets/sheet' + n + '.xml', ser(doc));
      if (sheetRels) {
        let rx = sheetRels;
        if (bin) { zip.file('xl/printerSettings/printerSettings' + n + '.bin', bin); rx = rx.replace(/printerSettings\/[^"]+/, 'printerSettings/printerSettings' + n + '.bin'); }
        zip.file('xl/worksheets/_rels/sheet' + n + '.xml.rels', rx);
      }
      const r = relsDoc.createElementNS(PKG_NS, 'Relationship');
      r.setAttribute('Id', 'rIdS' + n); r.setAttribute('Type', T_SHEET); r.setAttribute('Target', 'worksheets/sheet' + n + '.xml');
      relsDoc.documentElement.appendChild(r);
      const o = ctDoc.createElementNS(CT_NS, 'Override');
      o.setAttribute('PartName', '/xl/worksheets/sheet' + n + '.xml');
      o.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
      ctDoc.documentElement.appendChild(o);
      const s = wbDoc.createElementNS(NS, 'sheet');
      s.setAttribute('name', sh.name); s.setAttribute('sheetId', n); s.setAttributeNS(REL_NS, 'r:id', 'rIdS' + n);
      sheetsEl.appendChild(s);
      if (areaRef) {
        if (!dnEl) { dnEl = wbDoc.createElementNS(NS, 'definedNames'); sheetsEl.parentNode.insertBefore(dnEl, sheetsEl.nextSibling); }
        const d = wbDoc.createElementNS(NS, 'definedName');
        d.setAttribute('name', '_xlnm.Print_Area'); d.setAttribute('localSheetId', i);
        d.textContent = "'" + sh.name.replace(/'/g, "''") + "'!" + areaRef;
        dnEl.appendChild(d);
      }
    });
    if (dnEl && !kids(dnEl, 'definedName').length) dnEl.parentNode.removeChild(dnEl);
    // 第一張是打開時的那張；打開時重算公式
    Array.prototype.forEach.call(wbDoc.getElementsByTagNameNS(NS, 'workbookView'), v => { v.removeAttribute('activeTab'); v.removeAttribute('firstSheet'); });
    let calc = wbDoc.getElementsByTagNameNS(NS, 'calcPr')[0];
    if (calc) calc.setAttribute('fullCalcOnLoad', '1');
    // 範本的檔案路徑不要帶出去
    Array.prototype.slice.call(wbDoc.getElementsByTagName('mc:AlternateContent')).forEach(e => e.parentNode.removeChild(e));

    zip.file('xl/workbook.xml', ser(wbDoc));
    zip.file('xl/_rels/workbook.xml.rels', ser(relsDoc));
    zip.file('[Content_Types].xml', ser(ctDoc));
    zip.file('xl/styles.xml', ser(stylesDoc));
    if (zip.file('docProps/app.xml')) zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>');
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
  }

  /** 讀 xlsx：回傳 [{name, cells: {A1: 值}}]（公式取存檔時算好的結果；日期是 Excel 序號數字） */
  async function readWorkbook(buf) {
    const zip = await JSZip.loadAsync(buf);
    const read = async p => zip.file(p) ? parse(await zip.file(p).async('string')) : null;
    const wb = await read('xl/workbook.xml');
    if (!wb) throw new Error('這不是 Excel（.xlsx）檔');
    const rels = await read('xl/_rels/workbook.xml.rels');
    const target = {};
    Array.prototype.forEach.call(rels.getElementsByTagNameNS(PKG_NS, 'Relationship'), r => {
      const t = r.getAttribute('Target');
      target[r.getAttribute('Id')] = t.charAt(0) === '/' ? t.slice(1) : 'xl/' + t;
    });
    const ssDoc = await read('xl/sharedStrings.xml');
    const shared = ssDoc ? Array.prototype.map.call(ssDoc.getElementsByTagNameNS(NS, 'si'),
      si => Array.prototype.map.call(si.getElementsByTagNameNS(NS, 't'), t => t.textContent).join('')) : [];
    const out = [];
    for (const s of Array.prototype.slice.call(wb.getElementsByTagNameNS(NS, 'sheet'))) {
      const rid = s.getAttributeNS(REL_NS, 'id') || s.getAttribute('r:id');
      const doc = await read(target[rid]);
      if (!doc) continue;
      const cells = {};
      Array.prototype.forEach.call(doc.getElementsByTagNameNS(NS, 'c'), c => {
        const t = c.getAttribute('t'), v = kids(c, 'v')[0];
        let val;
        if (t === 'inlineStr') val = Array.prototype.map.call(c.getElementsByTagNameNS(NS, 't'), x => x.textContent).join('');
        else if (!v) return;
        else if (t === 's') val = shared[Number(v.textContent)];
        else if (t === 'str' || t === 'e') val = v.textContent;
        else if (t === 'b') val = v.textContent === '1';
        else val = Number(v.textContent);
        if (val !== '' && val !== undefined) cells[c.getAttribute('r')] = val;
      });
      out.push({ name: s.getAttribute('name'), cells: cells });
    }
    return out;
  }

  root.DRXlsx = { buildWorkbook, readWorkbook };
})(this);
