/* 施工日報表產生器：解析日誌文字 → 算出每一格要填什麼（不碰 Excel 檔，瀏覽器和 Node 都能跑） */
(function (root) {
  'use strict';

  // ===== 小工具 =====
  const toHalf = s => String(s || '').replace(/[！-～]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
  const pad = n => ('0' + n).slice(-2);
  const ymd = (y, m, d) => y + '-' + pad(m) + '-' + pad(d);
  const num = s => { const n = Number(s); return isFinite(n) ? n : 0; };
  /** 拿掉項目編號「1、」「2.」「3)」（「2.5m」「1.2F」這種數字不動） */
  const stripNo = s => String(s).replace(/^\s*\d+\s*(?:、|[.)](?![\dF]))\s*/, '').trim();
  const round2 = n => Math.round(n * 100) / 100;

  /** 工種 → 日報表工項（和機器人的對照表一樣，取最長的關鍵字） */
  const TRADE_MAP = [
    ['水電消防', '水電工'], ['水電', '水電工'], ['消防', '消防工'], ['模板', '模板工'], ['鋼筋', '鋼筋工'],
    ['粉光', '整體粉光'], ['粗工', '粗工'], ['放樣', '放樣工'], ['開挖', '開挖工'], ['泥作', '泥作工'],
    ['鐵工', '鐵工'], ['弱電', '弱電工'], ['防水', '防水工'], ['油漆', '油漆工'], ['清潔', '清潔工'],
    ['鷹架', '鷹架工'], ['安全防護', '安全防護工'], ['支撐', '安全支撐工'], ['地錨', '地錨工'], ['打井', '打井工'],
    ['鋁窗', '鋁窗工'], ['輕隔間', '輕隔間工'], ['衛浴', '衛浴設備'], ['景觀', '景觀工'], ['石材', '石材工'],
    ['電梯', '電梯工'], ['玻璃', '玻璃工'], ['廚具', '廚具'], ['木工', '木工'], ['扶手', '樓梯扶手'],
    ['EPOXY', 'EPOXY'], ['鋁包板', '鋁包版'], ['鋁包版', '鋁包版'], ['地質改良', 'CCP地質改良工'], ['連續壁', '連續壁工'],
  ];
  function mapTrade(trade) {
    trade = String(trade || '').trim();
    let best = null;
    TRADE_MAP.forEach(p => { if (trade.indexOf(p[0]) >= 0 && (!best || p[0].length > best[0].length)) best = p; });
    return best ? best[1] : (trade ? (/工$/.test(trade) ? trade : trade + '工') : '其他');
  }
  const isRough = e => e.item === '粗工' || /粗工/.test(e.trade || '');

  // ===== 混凝土強度 =====
  const PSI_TO_KGF = { 2000: 140, 2500: 175, 3000: 210, 3500: 245, 4000: 280, 4500: 315, 5000: 350, 6000: 420 };
  const KGF_TO_PSI = {}; Object.keys(PSI_TO_KGF).forEach(p => { KGF_TO_PSI[PSI_TO_KGF[p]] = Number(p); });
  /** 「3000 PSI」「210kgf」「210」→ 範本寫法「混凝土210kgf/cm2/3000psi」 */
  function concreteName(gradeText) {
    const s = toHalf(gradeText);
    const m = s.match(/(\d{3,4})\s*(psi|磅|kgf\/cm2|kgf|kg)?/i);
    if (!m) return '混凝土' + s.trim();
    const n = Number(m[1]);
    const isPsi = m[1].length === 4 || /psi|磅/i.test(m[2] || '');
    const psi = isPsi ? n : KGF_TO_PSI[n], kgf = isPsi ? PSI_TO_KGF[n] : n;
    if (kgf && psi) return '混凝土' + kgf + 'kgf/cm2/' + psi + 'psi';
    return isPsi ? '混凝土' + n + 'psi' : '混凝土' + n + 'kgf/cm2';
  }
  function concreteShort(name) {
    const m = String(name).match(/\/(\d+)psi$/) || String(name).match(/^混凝土(\d+)psi$/);
    if (m) return m[1] + ' PSI';
    return String(name).replace(/^混凝土/, '');
  }

  // ===== 數量＋單位 =====
  const UNIT_NORM = { '立方公尺': 'm3', '立方米': 'm3', '立方': 'm3', '方': 'm3', 'M3': 'm3', 'm³': 'm3', '公斤': 'kg', 'KG': 'kg', '平方米': 'm2', '部': '台', '輛': '台', '桶裝': '桶', '卷': '捲' };
  /** 「黏著劑820 20包」「挖土機 1台」「山貓2」→ {name, qty, unit} */
  function parseQtyItems(text, defUnit) {
    const out = [];
    toHalf(text).split(/[\n,，、;；。]+/).forEach(piece => {
      piece = stripNo(piece);
      if (!piece) return;
      const m = piece.match(/^(.+?)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*([^\d\s]{0,4})$/);
      if (!m || !/[一-鿿A-Za-z]/.test(m[1])) return;
      const unit = m[3] ? (UNIT_NORM[m[3]] || m[3]) : defUnit;
      if (unit === '車') return;   // 車數不是材料數量
      out.push({ name: m[1].trim(), qty: Number(m[2]), unit: unit });
    });
    return out;
  }

  // ===== 解析日誌 =====
  /**
   * 吃機器人產生的日誌：
   *   【案名】2026/11/17（二）施工日誌
   *   定晨_水電*5工：      ← 廠商_工種*人數工：（也接受 x、×）
   *   1、B2F牆面配管
   *   國產_混凝土：  1、3000 PSI 30m3
   *   進料：  1、黏著劑820 20包
   *   機具：  1、挖土機 1台
   *   天氣：上午晴下午雨
   *   重要記事：  1、xxx
   *   今日合計 / 明日施工 → 略過（範本沒有這欄）
   */
  function parseLog(raw, fallbackDate) {
    const lines = toHalf(raw).replace(/\r/g, '').split('\n');
    const r = { date: '', weather: null, entries: [], concrete: [], materials: [], machines: [], notes: [], warnings: [] };
    let block = null;   // {type, ref}
    const today = fallbackDate || '';

    // 日期：第一個看到的 yyyy/m/d 或 m/d
    for (const ln of lines.filter(x => x.trim()).slice(0, 1)) {   // 只看第一行（標題行），避免把內容裡的 3/4 當日期
      let m = ln.match(/(\d{3,4})\s*[\/.\-年]\s*(\d{1,2})\s*[\/.\-月]\s*(\d{1,2})/);
      if (m) { const y = Number(m[1]) < 1000 ? Number(m[1]) + 1911 : Number(m[1]); r.date = ymd(y, Number(m[2]), Number(m[3])); break; }
      m = ln.match(/(?:^|[^\d])(\d{1,2})\s*[\/月]\s*(\d{1,2})\s*日?/);
      if (m && today) { r.date = ymd(today.slice(0, 4), Number(m[1]), Number(m[2])); break; }
    }

    lines.forEach(line => {
      const t = line.trim();
      if (!t) return;
      // 標題行
      if (/施工日誌|^【.*】\s*$/.test(t)) { block = null; return; }
      let m;
      if ((m = t.match(/^天氣\s*[:：]?\s*(.*)$/))) { r.weather = parseWeather(m[1]); block = null; return; }
      if (/^今日合計/.test(t)) { block = null; return; }
      if (/^明日(施工)?.*[:：]$/.test(t)) { block = { type: 'skip' }; return; }
      if (/^(進料|材料)\s*[:：]/.test(t)) { block = { type: 'mat' }; const rest = t.replace(/^(進料|材料)\s*[:：]/, ''); if (rest.trim()) r.materials.push.apply(r.materials, parseQtyItems(rest, '')); return; }
      if (/^(機具|施工機具)\s*[:：]/.test(t)) { block = { type: 'mach' }; const rest = t.replace(/^(施工)?機具\s*[:：]/, ''); if (rest.trim()) r.machines.push.apply(r.machines, parseQtyItems(rest, '台')); return; }
      if (/^(重要記事|記事)\s*[:：]/.test(t)) { block = { type: 'note' }; const rest = t.replace(/^(重要)?記事\s*[:：]/, '').trim(); if (rest) r.notes.push(rest); return; }
      // 混凝土：「國產_混凝土：」
      if ((m = t.match(/^(?:(.*?)_)?(?:預拌)?混凝土\s*[:：]\s*(.*)$/))) {
        block = { type: 'conc', vendor: (m[1] || '').trim() };
        if (m[2].trim()) addConcrete(m[2], block.vendor);
        return;
      }
      // 廠商標題：「定晨_水電*5工：」「定晨_水電x2.5工：」
      if ((m = t.match(/^(.+?)_(.*?)\s*[*＊xX×]\s*(\d+(?:\.\d+)?)\s*工\s*[:：]\s*(.*)$/))) {
        const e = { vendor: m[1].trim(), trade: m[2].trim(), n: Number(m[3]), items: [] };
        e.item = mapTrade(e.trade);
        r.entries.push(e);
        block = { type: 'entry', ref: e };
        if (m[4].trim()) e.items.push(m[4].trim());
        return;
      }
      // 項目行「1、xxx」或接在區塊後的文字
      const item = stripNo(t);
      if (!block) { r.warnings.push('看不懂這行，已略過：' + t); return; }
      if (block.type === 'entry') block.ref.items.push(item);
      else if (block.type === 'conc') addConcrete(item, block.vendor);
      else if (block.type === 'mat') r.materials.push.apply(r.materials, parseQtyItems(item, ''));
      else if (block.type === 'mach') r.machines.push.apply(r.machines, parseQtyItems(item, '台'));
      else if (block.type === 'note') r.notes.push(item);
    });

    function addConcrete(text, vendor) {
      const s = toHalf(text);
      const m = s.match(/^(.*?)\s*(\d+(?:\.\d+)?)\s*(m3|M3|立方公尺|立方米|立方|方)\s*$/);
      if (!m || !/\d{3,4}/.test(m[1])) { r.warnings.push('混凝土看不懂：' + text); return; }
      r.concrete.push({ vendor: vendor || '', name: concreteName(m[1]), qty: Number(m[2]), unit: 'm3' });
    }
    // 同一家廠商＋工種出現兩次：人數取後面的、內容合併
    const merged = [], idx = {};
    r.entries.forEach(e => {
      const k = e.vendor + '|' + e.trade;
      if (k in idx) { const o = merged[idx[k]]; o.n = e.n; e.items.forEach(x => { if (o.items.indexOf(x) < 0) o.items.push(x); }); }
      else { idx[k] = merged.length; merged.push(e); }
    });
    r.entries = merged;
    return r;
  }

  function parseWeather(s) {
    s = toHalf(s);
    const pick = re => { const m = s.match(re); return m ? m[1] : ''; };
    let am = pick(/上午\s*([晴陰雨])/), pm = pick(/下午\s*([晴陰雨])/);
    const all = s.match(/[晴陰雨]/g) || [];
    if (!am && !pm) { am = all[0] || ''; pm = all[1] || am; }
    return { am: am || pm, pm: pm || am };
  }

  // ===== 日期 =====
  const dUTC = s => { const p = String(s).split('-').map(Number); return Date.UTC(p[0], p[1] - 1, p[2]); };
  const excelSerial = s => Math.round((dUTC(s) - Date.UTC(1899, 11, 30)) / 86400000);
  const daysBetween = (a, b) => Math.round((dUTC(b) - dUTC(a)) / 86400000);
  const slashDate = d => { const p = String(d).split('-'); return Number(p[0]) + '/' + Number(p[1]) + '/' + Number(p[2]); };
  const tabName = d => Number(d.slice(5, 7)) + '月' + Number(d.slice(8, 10)) + '日';

  // ===== 填表 =====
  /** 版面（施工日報表_空白.xlsx） */
  const LAYOUT = {
    trades: { col: 'B', start: 9, rows: 42 },        // 工項 B9:B50，C 上午 D 下午 E 平均(公式) F 累計
    mats: { start: 9, rows: 16 },                   // 進場材料 H9:H24，J 單位 K 本日 M 累計
    machs: { start: 26, rows: 8 },                  // 施工機具 H26:H33
    summary: { col: 'I', start: 34, rows: 30 },     // 本日施工概要 I34:I63
    rough: { col: 'I', start: 64, rows: 6 },        // 粗工報表 I64:I69
    notes: { col: 'B', start: 51, rows: 19 },       // 重要記事 B51:B69
  };

  /** 行數不夠時，從最舊的開始兩筆併一列 */
  function packLines(lines, n) {
    const L = lines.slice();
    let i = 0;
    while (L.length > n && i < L.length - 1) { L.splice(i, 2, L[i] + '；' + L[i + 1]); i++; }
    return L;
  }
  function fillRows(c, col, start, n, lines) {
    const L = lines.slice(0, n);
    if (lines.length > n) L[n - 1] = lines.slice(n - 1).join('；');
    for (let i = 0; i < n; i++) c[col + (start + i)] = L[i] || '';
  }

  /**
   * days：{ 'yyyy-mm-dd': { log, weatherAm, weatherPm, machines, notes } }（全部存的日子，算累計用）
   * 回傳：{ cells: {A1: 值}, bold: [儲存格...], warnings }
   *   值：數字、字串，或 {date: 'yyyy-mm-dd'}（Excel 日期）
   */
  function buildDay(day, days, settings) {
    settings = settings || {};
    // 匯入的舊日報表：那天以前的累計當起點，只接著算那天之後存的日子
    const base = settings.base && settings.base.date < day ? settings.base : null;
    const dates = Object.keys(days).filter(d => d <= day && (!base || d > base.date)).sort();
    const parsed = {};
    dates.forEach(d => { parsed[d] = parseDayCached(days[d], d); });
    const p = parsed[day] || parseDay({}, day);
    const c = {}, bold = [];

    // 表頭
    if (settings.projectName) c.C5 = settings.projectName;
    if (settings.startDate) {
      c.C6 = { date: settings.startDate };
      c.P5 = daysBetween(settings.startDate, day) + 1;   // 工作累計（開工第幾天）
    }
    if (settings.plannedDays) c.O5 = Number(settings.plannedDays);
    c.L6 = { date: day };

    // 天氣 L=晴 M=陰 N=雨；第 4 列上午、第 5 列下午
    const col = { '晴': 'L', '陰': 'M', '雨': 'N' };
    ['L4', 'M4', 'N4', 'L5', 'M5', 'N5'].forEach(a => { c[a] = ''; });
    const w = p.weather || { am: settings.weather || '晴', pm: settings.weather || '晴' };
    if (col[w.am]) c[col[w.am] + '4'] = '●';
    if (col[w.pm]) c[col[w.pm] + '5'] = '●';

    // 出工：工項依「第一次出現的日期」由上往下排，每天位置固定；累計＝到今天為止每日人數總和
    const order = [], cum = {};
    if (base) base.trades.forEach(t => { if (order.indexOf(t.name) < 0) order.push(t.name); cum[t.name] = (cum[t.name] || 0) + t.cum; });
    dates.forEach(d => parsed[d].entries.forEach(e => {
      if (order.indexOf(e.item) < 0) order.push(e.item);
      cum[e.item] = (cum[e.item] || 0) + e.n;
    }));
    const todayN = {};
    p.entries.forEach(e => { todayN[e.item] = (todayN[e.item] || 0) + e.n; });
    const T = LAYOUT.trades;
    order.slice(0, T.rows).forEach((name, i) => {
      const r = T.start + i;
      c['B' + r] = name;
      c['C' + r] = name in todayN ? round2(todayN[name]) : '';
      c['D' + r] = name in todayN ? round2(todayN[name]) : '';
      c['F' + r] = round2(cum[name] || 0);
    });
    if (order.length > T.rows) p.warnings.push('工項超過 ' + T.rows + ' 種，多的沒有列出');

    // 材料（含混凝土）、機具：本日＋累計
    const fillQty = (start, rows, key, label, baseList) => {
      const ord = [], unit = {}, sum = {}, td = {};
      (baseList || []).forEach(x => { if (ord.indexOf(x.name) < 0) { ord.push(x.name); unit[x.name] = x.unit; } sum[x.name] = (sum[x.name] || 0) + x.cum; });
      dates.forEach(d => parsed[d][key].forEach(x => {
        if (ord.indexOf(x.name) < 0) ord.push(x.name);
        if (!unit[x.name]) unit[x.name] = x.unit;
        sum[x.name] = (sum[x.name] || 0) + x.qty;
        if (d === day) td[x.name] = (td[x.name] || 0) + x.qty;
      }));
      ord.slice(0, rows).forEach((name, i) => {
        const r = start + i;
        c['H' + r] = name; c['J' + r] = unit[name] || '';
        c['K' + r] = name in td ? round2(td[name]) : '';
        c['M' + r] = round2(sum[name]);
      });
      if (ord.length > rows) p.warnings.push(label + '超過 ' + rows + ' 種，多的沒有列出');
    };
    dates.forEach(d => { parsed[d].allMats = parsed[d].concrete.concat(parsed[d].materials); });
    fillQty(LAYOUT.mats.start, LAYOUT.mats.rows, 'allMats', '材料', base && base.mats);
    fillQty(LAYOUT.machs.start, LAYOUT.machs.rows, 'machines', '機具', base && base.machs);

    // 本日施工概要：每家廠商一段，標題粗體「定晨_水電x5工：」，底下「1、…」
    const blocks = p.entries.filter(e => !isRough(e)).map(e => ({ head: e.vendor + '_' + (e.trade || '') + 'x' + fmtN(e.n) + '工：', items: e.items }));
    const cv = {}, cvOrder = [];
    p.concrete.forEach(x => { if (!cv[x.vendor]) { cv[x.vendor] = []; cvOrder.push(x.vendor); } cv[x.vendor].push(concreteShort(x.name) + ' ' + fmtN(x.qty) + 'm3'); });
    cvOrder.forEach(v => blocks.push({ head: (v ? v + '_' : '') + '混凝土：', items: cv[v] }));
    const S = LAYOUT.summary;
    const need = blocks.reduce((a, b) => a + 1 + b.items.length, 0);
    let lines = [], boldIdx = [];
    blocks.forEach(b => {
      boldIdx.push(lines.length);
      if (need <= S.rows) { lines.push(b.head); b.items.forEach((x, k) => lines.push((k + 1) + '、' + x)); }
      else lines.push(b.head + b.items.map((x, k) => (k + 1) + '、' + x).join(' '));   // 列數不夠：每家壓成一列
    });
    if (lines.length > S.rows) { lines[S.rows - 1] = lines.slice(S.rows - 1).join('；'); lines = lines.slice(0, S.rows); boldIdx = boldIdx.filter(i => i < S.rows); }
    fillRows(c, S.col, S.start, S.rows, lines);
    boldIdx.forEach(i => bold.push(S.col + (S.start + i)));

    // 粗工報表：「邦丞 2 工：環境清潔、搬運」
    const rough = {}, roughWork = {}, roughOrder = [];
    p.entries.filter(isRough).forEach(e => {
      if (!(e.vendor in rough)) { rough[e.vendor] = 0; roughWork[e.vendor] = []; roughOrder.push(e.vendor); }
      rough[e.vendor] += e.n; roughWork[e.vendor] = roughWork[e.vendor].concat(e.items);
    });
    fillRows(c, LAYOUT.rough.col, LAYOUT.rough.start, LAYOUT.rough.rows,
      roughOrder.map(v => v + ' ' + fmtN(rough[v]) + ' 工：' + roughWork[v].join('、')));

    // 重要記事：到今天為止全部保留（日期＋內容）
    const notes = base ? base.notes.slice() : [];
    dates.forEach(d => parsed[d].notes.forEach(n => notes.push(slashDate(d) + ' ' + n)));
    fillRows(c, LAYOUT.notes.col, LAYOUT.notes.start, LAYOUT.notes.rows, packLines(notes, LAYOUT.notes.rows));

    return { cells: c, bold: bold, warnings: p.warnings, parsed: p };
  }
  const fmtN = n => String(round2(n));

  const PCACHE = new Map();
  function parseDayCached(rec, day) {
    const k = day + '|' + JSON.stringify(rec || {});
    let p = PCACHE.get(k);
    if (!p) { if (PCACHE.size > 2000) PCACHE.clear(); p = parseDay(rec, day); PCACHE.set(k, p); }
    // buildDay 會加警告、allMats，複製一份避免改到快取
    return Object.assign({}, p, { warnings: p.warnings.slice() });
  }
  /** 一天存的資料 → 解析結果（日誌文字＋另外填的天氣、機具、記事） */
  function parseDay(rec, day) {
    rec = rec || {};
    const p = parseLog(rec.log || '', day);
    if (rec.weatherAm || rec.weatherPm) p.weather = { am: rec.weatherAm || rec.weatherPm, pm: rec.weatherPm || rec.weatherAm };
    if (rec.machines) p.machines = p.machines.concat(parseQtyItems(rec.machines, '台'));
    if (rec.materials) p.materials = p.materials.concat(parseQtyItems(rec.materials, ''));
    if (rec.notes) toHalf(rec.notes).split('\n').map(s => s.replace(/^\s*\d+\s*[、.)]\s*/, '').trim()).filter(String).forEach(n => p.notes.push(n));
    return p;
  }

  // ===== 匯入舊日報表 =====
  const fromSerial = n => { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 86400000); return ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); };
  /** 儲存格的日期：Excel 序號、「2026年11月17日」、「2026/11/17」、民國「115/11/17」 */
  function anyDate(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) return fromSerial(v);
    const m = toHalf(v).match(/(\d{2,4})\s*[\/.\-年]\s*(\d{1,2})\s*[\/.\-月]\s*(\d{1,2})/);
    if (!m) return '';
    const y = Number(m[1]) < 1000 ? Number(m[1]) + 1911 : Number(m[1]);
    return ymd(y, Number(m[2]), Number(m[3]));
  }
  /**
   * sheets：readWorkbook 讀出來的 [{name, cells}]
   * 找日期最新的那一頁，取出累計當起點。回傳 {base, settings, sheetName} 或丟出錯誤
   */
  function parseOldReport(sheets, fileName) {
    const yearHint = (String(fileName || '').match(/(20\d\d)/) || [])[1] || String(new Date().getFullYear());
    let best = null;
    sheets.forEach(sh => {
      let d = anyDate(sh.cells.L6);
      if (!d) { const m = String(sh.name).match(/(\d{1,2})\s*月\s*(\d{1,2})\s*日/); if (m) d = ymd(yearHint, Number(m[1]), Number(m[2])); }
      if (d && (!best || d > best.date)) best = { date: d, sh: sh };
    });
    if (!best) throw new Error('找不到日報的日期（「日期」那格是空的，分頁名稱也不是「11月17日」這種）。請選有填過資料的日報表。');
    const c = best.sh.cells;
    const str = a => c[a] === undefined || c[a] === null ? '' : String(c[a]).trim();
    const n = a => { const v = Number(c[a]); return isFinite(v) ? round2(v) : 0; };
    const base = { date: best.date, trades: [], mats: [], machs: [], notes: [] };
    const T = LAYOUT.trades;
    for (let r = T.start; r < T.start + T.rows; r++) if (str('B' + r)) base.trades.push({ name: str('B' + r), cum: n('F' + r) });
    for (let r = LAYOUT.mats.start; r < LAYOUT.mats.start + LAYOUT.mats.rows; r++) if (str('H' + r)) base.mats.push({ name: str('H' + r), unit: str('J' + r), cum: n('M' + r) });
    for (let r = LAYOUT.machs.start; r < LAYOUT.machs.start + LAYOUT.machs.rows; r++) if (str('H' + r)) base.machs.push({ name: str('H' + r), unit: str('J' + r), cum: n('M' + r) });
    const noteEnd = /交辦/.test(str('A64')) ? 63 : LAYOUT.notes.start + LAYOUT.notes.rows - 1;   // 舊版範本 64 列以後是交辦事項
    for (let r = LAYOUT.notes.start; r <= noteEnd; r++) if (str('B' + r)) base.notes.push(str('B' + r));
    if (!base.trades.some(t => t.cum) && !base.mats.some(x => x.cum) && !base.machs.some(x => x.cum) && !base.notes.length)
      throw new Error('這份日報表（' + best.sh.name + '）沒有任何累計數字，看起來是空白表。請選有填過資料的日報表。');
    const settings = {};
    if (str('C5') && str('C5') !== '新建工程') settings.projectName = str('C5');
    if (anyDate(c.C6)) settings.startDate = anyDate(c.C6);
    if (n('O5') > 0) settings.plannedDays = String(n('O5'));
    return { base: base, settings: settings, sheetName: best.sh.name };
  }

  const api = { parseOldReport, anyDate, parseLog, parseDay, parseQtyItems, parseWeather, buildDay, mapTrade, concreteName, excelSerial, tabName, LAYOUT, toHalf };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.DRCore = api;
})(this);
