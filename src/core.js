// 매출현황 워크북 최신화의 순수 로직. 브라우저와 node 양쪽에서 쓴다.
// 외부 라이브러리(XLSX=SheetJS, JSZip)는 인자로 받는다.

import {
  parseSheet, serializeSheet, parseSharedStrings, cellText, cellNumber, getCell, ensureCell,
  setFormula, setNumber, setText, clearCell, expandSharedFormulas, dropFormulaCache,
  shiftFormula, numToCol, getAttr, setAttr, parseAttrs, attrsToStr, lastRowWithValue, esc,
} from './xml.js';

export const SALES_SHEET = '판매현황';
export const STOCK_SHEET = '재고';
export const DAILY_SHEET = '일별';

// 판매현황 A~U 머리글. 위치가 바뀌면 쓰기 전에 멈춘다.
export const SALES_HEADERS = ['모델NO', '제품명', '결제금액', '수량', '판매총액', '판매총액(미포함)', '운송비',
  '정산', '원가', '이익', '%', '구매자', '판매사이트', '마켓주문번호', '출고지시일', '주문번호', '상태',
  '카테고리', '모델명', '년월', '주'];

// 판매사이트 → 가격세팅 정산 블록 이름. 없는 마켓은 G마켓 기준 정산가를 쓴다.
export const DEFAULT_MARKET_BLOCK = {
  '지마켓': 'G마켓·옥션', 'G마켓': 'G마켓·옥션', '옥션': 'G마켓·옥션',
  '하이마트': '하이마트',
  'N페이': '스마트스토어', '네이버삽N': '스마트스토어', '스마트스토어': '스마트스토어',
  '쿠팡': '쿠팡',
  '컴퓨존': '컴퓨존',
  '자사몰': '자사몰',
};

const CATEGORY_ORDER = ['RAM', 'SSD', 'VGA'];
const DOW = ['일', '월', '화', '수', '목', '금', '토'];

// ---------------------------------------------------------------- 날짜

export const serialFromYMD = (y, m, d) => Date.UTC(y, m - 1, d) / 86400000 + 25569;
export function ymdFromSerial(s) {
  const dt = new Date((s - 25569) * 86400000);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), dow: dt.getUTCDay() };
}
const daysInMonth = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
const yymm = (y, m) => (y % 100) * 100 + m;
const fromYymm = (n) => ({ y: 2000 + Math.floor(n / 100), m: n % 100 });
const prevYymm = (n) => { const { y, m } = fromYymm(n); return m === 1 ? yymm(y - 1, 12) : yymm(y, m - 1); };

// 분기 첫날이 속한 주(월요일 시작)가 W1. 판매현황 U열 표기와 같다.
export function weekCode(y, m, d) {
  const q = Math.floor((m - 1) / 3);
  const qs = Date.UTC(y, q * 3, 1);
  const qsDow = (new Date(qs).getUTCDay() + 6) % 7;
  const weekStart = qs - qsDow * 86400000;
  const w = Math.floor((Date.UTC(y, m - 1, d) - weekStart) / (7 * 86400000)) + 1;
  return `${y % 100}Q${q + 1}W${w}`;
}

function parseShipDate(v) {
  if (v == null || v === '') return null;
  if (typeof v === 'number') {
    if (v > 19000000 && v < 30000000) {
      const y = Math.floor(v / 10000), m = Math.floor(v / 100) % 100, d = v % 100;
      return { y, m, d };
    }
    if (v > 30000 && v < 80000) { const r = ymdFromSerial(Math.floor(v)); return { y: r.y, m: r.m, d: r.d }; }
    return null;
  }
  if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  const s = String(v).trim();
  let m = /^(\d{4})[-./]?(\d{1,2})[-./]?(\d{1,2})/.exec(s);
  if (m) return { y: +m[1], m: +m[2], d: +m[3] };
  return null;
}

export const normId = (v) => {
  if (v == null) return '';
  if (typeof v === 'number') return String(Math.trunc(v));
  return String(v).trim().replace(/\.0+$/, '');
};

// ---------------------------------------------------------------- 상품명 정규화

// 가격세팅 모델명과 판매현황 모델명은 표기가 다르다(지포스/GeForce, Triple Fan/TF 등).
// 표기 차이만 지우고 토큰 집합이 정확히 같을 때만 같은 상품으로 본다. 부분 일치는 쓰지 않는다.
const NOISE = new Set(['PNY', 'ACER', '에이서', '지포스', 'GEFORCE', '라데온', 'RADEON', 'FAN', 'M.2', 'NVME',
  'UDIMM', 'M', 'BIFROST', 'RGB', 'VGA', 'SSD', 'RAM', 'DESK', 'DSEK', '#']);

export function nameKey(s) {
  if (!s) return '';
  const t = String(s).toUpperCase()
    .replace(/TRIPLE\s*FAN/g, ' TF ').replace(/\bTRIPLE\b/g, ' TF ')
    .replace(/DUAL\s*FAN/g, ' DUAL ').replace(/\bDF\b/g, ' DUAL ')
    .replace(/[()[\],#]/g, ' ');
  const toks = t.split(/\s+/).filter((x) => x && !NOISE.has(x));
  return [...new Set(toks)].sort().join(' ');
}

// ---------------------------------------------------------------- 패키지

async function readText(zip, path) {
  const f = zip.file(path);
  return f ? f.async('string') : null;
}

function resolveTarget(target) {
  if (target.startsWith('/')) return target.slice(1);
  return 'xl/' + target.replace(/^\.\//, '');
}

export async function openWorkbook(JSZip, buf) {
  const zip = await JSZip.loadAsync(buf);
  const wbXml = await readText(zip, 'xl/workbook.xml');
  const relsXml = await readText(zip, 'xl/_rels/workbook.xml.rels');
  if (!wbXml || !relsXml) throw new Error('xlsx 형식이 아닙니다.');
  const rels = new Map();
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const a = parseAttrs(m[0]);
    rels.set(getAttr(a, 'Id'), { target: getAttr(a, 'Target'), type: getAttr(a, 'Type') });
  }
  const sheets = [];
  for (const m of wbXml.matchAll(/<sheet\b[^>]*\/>/g)) {
    const a = parseAttrs(m[0]);
    const rid = getAttr(a, 'r:id');
    sheets.push({
      name: unescAttr(getAttr(a, 'name')),
      sheetId: +getAttr(a, 'sheetId'),
      rid,
      state: getAttr(a, 'state') || 'visible',
      path: resolveTarget(rels.get(rid).target),
    });
  }
  const sst = parseSharedStrings(await readText(zip, 'xl/sharedStrings.xml'));
  return { zip, wbXml, relsXml, rels, sheets, sst, cache: new Map() };
}

const unescAttr = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&');

async function sheetXml(wb, name) {
  const s = wb.sheets.find((x) => x.name === name);
  if (!s) return null;
  if (!wb.cache.has(s.path)) wb.cache.set(s.path, await readText(wb.zip, s.path));
  return wb.cache.get(s.path);
}

const isMonthSheet = (name) => /^\d{4}$/.test(name) && +name % 100 >= 1 && +name % 100 <= 12;

// ---------------------------------------------------------------- 입력 판별

export function detectKind(XLSX, buf) {
  const wb = XLSX.read(buf, { type: 'array', sheetRows: 4, bookVBA: false });
  const names = wb.SheetNames;
  if (names.includes(SALES_SHEET) && names.includes(STOCK_SHEET)) return 'base';
  if (names.includes('정산가세팅')) return 'price';
  for (const n of names) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    const flat = rows.flat().map((x) => (x == null ? '' : String(x).trim()));
    if (flat.includes('모델NO') && flat.includes('주문번호') && flat.includes('출고일')) return 'sales';
    if (flat.includes('상품코드') && flat.includes('가용수량')) return 'stock';
  }
  return null;
}

// ---------------------------------------------------------------- 입력 읽기

function headerIndex(rows, required) {
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const r = (rows[i] || []).map((x) => (x == null ? '' : String(x).trim()));
    if (required.every((h) => r.includes(h))) {
      const idx = {};
      r.forEach((h, j) => { if (h && !(h in idx)) idx[h] = j; });
      return { row: i, idx };
    }
  }
  return null;
}

export function readSalesFile(XLSX, buf, fileName) {
  const wb = XLSX.read(buf, { type: 'array', raw: true });
  const req = ['모델NO', '제품명', '결제금액', '수량', '판매사이트', '출고일', '주문번호'];
  for (const n of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    const h = headerIndex(rows, req);
    if (!h) continue;
    const g = (r, k) => (h.idx[k] == null ? null : r[h.idx[k]]);
    const out = [];
    for (let i = h.row + 1; i < rows.length; i++) {
      const r = rows[i];
      const code = g(r, '모델NO');
      if (code == null || String(code).trim() === '') continue;
      const date = parseShipDate(g(r, '출고일'));
      out.push({
        file: fileName,
        line: i + 1,
        code: String(code).trim(),
        name: g(r, '제품명') == null ? '' : String(g(r, '제품명')).trim(),
        price: Number(g(r, '결제금액')) || 0,
        qty: Number(g(r, '수량')) || 0,
        buyer: g(r, '구매자') == null ? '' : String(g(r, '구매자')),
        market: g(r, '판매사이트') == null ? '' : String(g(r, '판매사이트')).trim(),
        marketOrder: normId(g(r, '마켓주문번호')),
        orderId: normId(g(r, '주문번호')),
        status: g(r, '상태') == null ? '' : String(g(r, '상태')),
        date,
      });
    }
    return out;
  }
  throw new Error(`${fileName}: 판매 내역 머리글(모델NO·주문번호·출고일)을 찾지 못했습니다.`);
}

export function readStockFile(XLSX, buf, fileName) {
  const wb = XLSX.read(buf, { type: 'array', raw: true });
  for (const n of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null });
    const h = headerIndex(rows, ['상품코드', '상품명', '가용수량']);
    if (!h) continue;
    const meta = rows.slice(0, h.row).flat().filter((x) => x != null && x !== '');
    const data = [];
    for (let i = h.row + 1; i < rows.length; i++) {
      const r = rows[i];
      const blank = (k) => r[h.idx[k]] == null || String(r[h.idx[k]]).trim() === '';
      if (blank('상품코드') && blank('상품명')) continue;
      data.push(r);
    }
    return { file: fileName, meta, idx: h.idx, rows: data };
  }
  throw new Error(`${fileName}: 재고 머리글(상품코드·가용수량)을 찾지 못했습니다.`);
}

export function readPriceFile(XLSX, buf, fileName) {
  const wb = XLSX.read(buf, { type: 'array', raw: true });
  const ws = wb.Sheets['정산가세팅'];
  if (!ws) throw new Error(`${fileName}: '정산가세팅' 시트가 없습니다.`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const h = headerIndex(rows, ['모델명', '정산가', '원가']);
  if (!h) throw new Error(`${fileName}: 정산가세팅의 머리글(모델명·정산가·원가)을 찾지 못했습니다.`);
  const head = rows[h.row].map((x) => (x == null ? '' : String(x).trim()));
  const group = (rows[h.row - 1] || []).map((x) => (x == null ? '' : String(x).trim()));
  const blocks = {};
  head.forEach((v, j) => {
    if (v !== '정산액') return;
    for (let k = j; k >= 0; k--) {
      if (group[k]) { blocks[group[k]] = j; break; }
    }
  });
  const items = [];
  for (let i = h.row + 1; i < rows.length; i++) {
    const r = rows[i];
    const name = r[h.idx['모델명']];
    if (name == null || String(name).trim() === '') break; // 첫 표만 읽는다
    const num = (j) => (j == null || typeof r[j] !== 'number' ? null : r[j]);
    const market = {};
    for (const [label, j] of Object.entries(blocks)) market[label] = num(j);
    items.push({
      name: String(name).trim(),
      category: r[h.idx['카테고리']] == null ? '' : String(r[h.idx['카테고리']]).trim(),
      settle: num(h.idx['정산가']),
      cost: num(h.idx['원가']),
      market,
      key: nameKey(name),
    });
  }
  if (!items.length) throw new Error(`${fileName}: 정산가세팅에 상품 행이 없습니다.`);
  return { file: fileName, items, blocks: Object.keys(blocks) };
}

// ---------------------------------------------------------------- 기준 워크북 읽기

export async function readBase(JSZip, buf, fileName) {
  const wb = await openWorkbook(JSZip, buf);
  const xml = await sheetXml(wb, SALES_SHEET);
  if (!xml) throw new Error(`'${SALES_SHEET}' 시트가 없습니다.`);
  const sh = parseSheet(xml);
  const header = [];
  for (let c = 1; c <= SALES_HEADERS.length; c++) header.push((cellText(getCell(sh, 2, c), wb.sst) || '').trim());
  const bad = SALES_HEADERS.map((h, i) => (header[i] === h ? null : `${numToCol(i + 1)}열 '${header[i]}' (기대: '${h}')`)).filter(Boolean);
  if (bad.length) throw new Error(`판매현황 머리글이 예상과 다릅니다: ${bad.join(', ')}`);

  const lastRow = lastRowWithValue(sh, 1);
  const orderIds = new Set();
  const history = new Map();
  const byDate = new Map(); // serial → [{code, model, cat, market, qty}]
  for (let r = 3; r <= lastRow; r++) {
    const row = sh.rows.get(r);
    if (!row) continue;
    const t = (c) => cellText(row.cells.get(c), wb.sst);
    const n = (c) => cellNumber(row.cells.get(c));
    const code = t(1);
    if (!code) continue;
    const oid = normId(t(16));
    if (oid) orderIds.add(oid);
    const h = history.get(code) || {};
    const cat = t(18), model = t(19), cost = n(9);
    if (cat) h.cat = cat;
    if (model) h.model = model;
    if (cost != null) h.cost = cost;
    h.name = t(2) || h.name;
    history.set(code, h);
    const o = n(15);
    if (o != null) {
      const list = byDate.get(o) || [];
      list.push({ code, model: model || '', cat: cat || '', market: t(13) || '', qty: n(4) || 0 });
      byDate.set(o, list);
    }
  }
  // 마지막 데이터 행의 셀 서식을 새 행에 그대로 쓴다
  const styleRow = sh.rows.get(lastRow);
  const styles = [];
  for (let c = 1; c <= SALES_HEADERS.length; c++) {
    const cell = styleRow && styleRow.cells.get(c);
    styles.push(cell ? getAttr(cell.attrs, 's') : null);
  }

  // 월 시트(YYMM)의 코드→카테고리·모델명
  const months = [];
  for (const s of wb.sheets) {
    if (!isMonthSheet(s.name)) continue;
    months.push(+s.name);
  }
  months.sort((a, b) => a - b);
  const monthCodes = new Map();
  const monthSheetCodes = new Map(); // yymm → Set(code)
  for (const m of [...months].reverse()) {
    const msh = parseSheet(await sheetXml(wb, String(m)));
    const codes = new Set();
    let cat = '';
    const rowNums = [...msh.rows.keys()].sort((a, b) => a - b);
    for (const r of rowNums) {
      const a = cellText(getCell(msh, r, 1), wb.sst);
      if (a && !/요약|TOTAL|카테고리/.test(a)) cat = a.trim();
      const b = cellText(getCell(msh, r, 2), wb.sst);
      for (const col of [42, 43]) { // AP, AQ
        const code = cellText(getCell(msh, r, col), wb.sst);
        if (!code || code === '코드') continue;
        codes.add(code);
        if (!monthCodes.has(code) && cat && b) monthCodes.set(code, { cat, model: b.trim() });
      }
    }
    monthSheetCodes.set(m, codes);
  }

  return { fileName, wb, lastRow, orderIds, history, byDate, styles, months, monthCodes, monthSheetCodes };
}

// ---------------------------------------------------------------- 계획

// 상품 코드 → 가격세팅 상품. 이름 토큰이 정확히 같거나, 원가가 유일하게 같을 때만 잇는다.
export function matchPrice(code, names, histCost, price) {
  if (!price) return null;
  const keys = new Set(names.map(nameKey).filter(Boolean));
  const hits = price.items.filter((it) => keys.has(it.key));
  const uniq = [...new Set(hits)];
  if (uniq.length === 1) return { item: uniq[0], how: '이름' };
  if (uniq.length > 1) return null;
  if (histCost != null) {
    const byCost = price.items.filter((it) => it.cost === histCost);
    if (byCost.length === 1) return { item: byCost[0], how: '원가' };
  }
  return null;
}

const blockFor = (market, price, marketMap) => {
  const label = (marketMap && marketMap[market]) || DEFAULT_MARKET_BLOCK[market];
  if (!label || !price) return null;
  const norm = (s) => s.replace(/[\s·.]/g, '');
  return price.blocks.find((b) => norm(b) === norm(label)) || null;
};

// overrides: { products: {code: {cat, model, include}}, priceMap: {code: 가격세팅 모델명 | ''} }
export function plan(base, salesRows, price, overrides = {}) {
  const ov = overrides.products || {};
  const pm = overrides.priceMap || {};
  const seen = new Set();
  const skipped = { dup: [], dupInUpload: [], noDate: [], excluded: [] };
  const products = new Map();
  const rows = [];

  for (const s of salesRows) {
    if (!s.date) { skipped.noDate.push(s); continue; }
    const key = s.orderId || `${s.file}#${s.line}`;
    if (s.orderId && base.orderIds.has(s.orderId)) { skipped.dup.push(s); continue; }
    if (seen.has(key)) { skipped.dupInUpload.push(s); continue; }
    seen.add(key);

    let p = products.get(s.code);
    if (!p) {
      const h = base.history.get(s.code) || {};
      const mc = base.monthCodes.get(s.code) || {};
      const o = ov[s.code] || {};
      const cat = o.cat || h.cat || mc.cat || '';
      const model = o.model || h.model || mc.model || '';
      const known = Boolean((h.cat || mc.cat) && (h.model || mc.model));
      let match = null;
      if (pm[s.code] !== undefined) {
        const it = pm[s.code] && price ? price.items.find((x) => x.name === pm[s.code]) : null;
        match = it ? { item: it, how: '직접 지정' } : null;
      } else {
        match = matchPrice(s.code, [model, h.model, mc.model, s.name], h.cost, price);
      }
      p = {
        code: s.code, name: s.name, cat, model, known,
        include: o.include != null ? o.include : known,
        histCost: h.cost ?? null,
        match,
        count: 0, qty: 0,
      };
      products.set(s.code, p);
    }
    p.count++;
    p.qty += s.qty;
    if (!p.include || !p.cat || !p.model) { skipped.excluded.push(s); continue; }

    const blk = blockFor(s.market, price, overrides.marketMap);
    let unit = null, settleFrom = null;
    if (p.match) {
      const it = p.match.item;
      if (blk && it.market[blk] != null) { unit = it.market[blk]; settleFrom = blk; }
      else if (it.settle != null) { unit = it.settle; settleFrom = '정산가(G마켓 기준)'; }
    }
    const cost = p.match && p.match.item.cost != null ? p.match.item.cost : p.histCost;
    const { y, m, d } = s.date;
    rows.push({
      ...s,
      cat: p.cat, model: p.model,
      serial: serialFromYMD(y, m, d),
      yymm: yymm(y, m),
      week: weekCode(y, m, d),
      settle: unit == null ? null : Math.round(unit) * s.qty,
      settleUnit: unit == null ? null : Math.round(unit),
      settleFrom,
      cost,
      costFrom: p.match && p.match.item.cost != null ? '가격세팅' : (p.histCost != null ? '이전 판매행' : null),
    });
  }
  rows.sort((a, b) => a.serial - b.serial || a.market.localeCompare(b.market) || a.code.localeCompare(b.code));
  return { rows, skipped, products: [...products.values()] };
}

// ---------------------------------------------------------------- 쓰기

const RANGE_RE = (sheet) => new RegExp(`((?:'${sheet}'|${sheet})!\\$?[A-Z]{1,3}\\$?3:\\$?[A-Z]{1,3}\\$?)(\\d+)`, 'g');

function extendRanges(xml, sheet, test, to) {
  return xml.replace(RANGE_RE(sheet), (m, pre, n) => (test(+n) ? pre + to : m));
}

function salesRowXml(r, row, styles) {
  const s = (i) => (styles[i] != null ? ` s="${styles[i]}"` : '');
  const c = (i) => numToCol(i + 1) + r;
  const str = (i, v) => (v == null || v === '' ? `<c r="${c(i)}"${s(i)}/>`
    : `<c r="${c(i)}"${s(i)} t="inlineStr"><is><t>${esc(v)}</t></is></c>`);
  const num = (i, v) => (v == null ? `<c r="${c(i)}"${s(i)}/>` : `<c r="${c(i)}"${s(i)}><v>${v}</v></c>`);
  const fx = (i, f) => `<c r="${c(i)}"${s(i)}><f>${esc(f)}</f></c>`;
  return `<row r="${r}" spans="1:21">`
    + str(0, row.code) + str(1, row.name) + num(2, row.price) + num(3, row.qty)
    + fx(4, `C${r}*D${r}`) + fx(5, `E${r}/1.1`) + num(6, 0)
    + num(7, row.settle) + num(8, row.cost)
    + fx(9, `H${r}-(I${r}*D${r})-G${r}`) + fx(10, `IF(I${r}*D${r}=0,"",J${r}/(I${r}*D${r}))`)
    + str(11, row.buyer) + str(12, row.market) + str(13, row.marketOrder) + num(14, row.serial)
    + str(15, row.orderId) + str(16, row.status) + str(17, row.cat) + str(18, row.model)
    + num(19, row.yymm) + str(20, row.week)
    + '</row>';
}

function setDimension(xml, ref) {
  return xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="${ref}"/>`);
}

function fixDates(text, fn) {
  return text.replace(/DATE\((\d{4}),(\d{1,2}),(\d{1,2})\)/g, (m, y, mo, d) => {
    const r = fn(+y, +mo, +d);
    return r ? `DATE(${r.y},${r.m},${r.d})` : m;
  });
}

// 월 시트에서 '카테고리' 머리글 행과 C~AG 날짜 열을 찾는다
function monthLayout(sh, sst) {
  const headerRows = [];
  for (const [r, row] of sh.rows) {
    const a = cellText(row.cells.get(1), sst);
    if (a && a.trim() === '카테고리' && cellNumber(row.cells.get(3)) != null) headerRows.push(r);
  }
  headerRows.sort((a, b) => a - b);
  return { headerRows, dateCols: Array.from({ length: 31 }, (_, i) => 3 + i) }; // C..AG
}

function rowsWithDailyFormula(sh) {
  const out = [];
  for (const [r, row] of sh.rows) {
    const c = row.cells.get(3);
    if (c && c.f && c.f.text && c.f.text !== '0') out.push(r);
  }
  return out;
}

function updateCosts(sh, sst, costByCode, report, sheetName) {
  for (const [r, row] of sh.rows) {
    const code = cellText(row.cells.get(42), sst);
    const cell = row.cells.get(47); // AU 원가
    if (!code || !cell || cell.f) continue;
    const cost = costByCode.get(code);
    const old = cellNumber(cell);
    if (cost == null || old == null || old === cost) continue;
    setNumber(cell, cost);
    report.push(`${sheetName} AU${r} ${code} 원가 ${old.toLocaleString()} → ${cost.toLocaleString()}`);
  }
}

// 기존 월 시트: 이달 날짜 열 중 비어 있거나 '=0'인 칸에 C열 수식을 옮겨 채우고 총계 날짜 범위를 월말로 맞춘다
function refreshMonthSheet(sh, sst, ym) {
  const { y, m } = fromYymm(ym);
  const last = daysInMonth(y, m);
  const lo = serialFromYMD(y, m, 1), hi = serialFromYMD(y, m, last);
  const { headerRows, dateCols } = monthLayout(sh, sst);
  if (!headerRows.length) return { filled: 0 };
  expandSharedFormulas(sh);
  const cols = dateCols.filter((c) => {
    const v = cellNumber(getCell(sh, headerRows[0], c));
    return c !== 3 && v != null && v >= lo && v <= hi;
  });
  let filled = 0;
  for (const r of rowsWithDailyFormula(sh)) {
    const base = getCell(sh, r, 3).f.text;
    for (const c of cols) {
      const cell = ensureCell(sh, r, c, getAttr(getCell(sh, r, 3).attrs, 's'));
      if (cell.f && cell.f.text && cell.f.text !== '0') continue;
      setFormula(cell, shiftFormula(base, c - 3));
      filled++;
    }
  }
  for (const row of sh.rows.values()) {
    for (const cell of row.cells.values()) {
      if (!cell.f || !cell.f.text) continue;
      cell.f.text = fixDates(cell.f.text, (yy, mm, dd) => (yy === y && mm === m && dd !== 1 && dd !== last ? { y, m, d: last } : null));
    }
  }
  dropFormulaCache(sh);
  return { filled };
}

// 새 월 시트: 직전 월 시트를 복제해 날짜·수식·전월 비교를 새 달로 바꾼다
function buildMonthSheet(templateXml, sst, fromYm, toYm) {
  const sh = parseSheet(templateXml);
  expandSharedFormulas(sh);
  const { y, m } = fromYymm(toYm);
  const last = daysInMonth(y, m);
  const t = fromYymm(fromYm);
  const tLast = daysInMonth(t.y, t.m);
  const p = fromYymm(prevYymm(toYm));
  const pLast = daysInMonth(p.y, p.m);
  const { headerRows, dateCols } = monthLayout(sh, sst);
  if (!headerRows.length) throw new Error(`${fromYm} 시트에서 날짜 머리글 행을 찾지 못했습니다.`);

  const a1 = ensureCell(sh, 1, 1);
  setNumber(a1, toYm);

  for (const hr of headerRows) {
    const styleC = getAttr(getCell(sh, hr, 3).attrs, 's');
    for (const c of dateCols) {
      const cell = ensureCell(sh, hr, c, styleC);
      const day = c - 2;
      if (day <= last) setNumber(cell, serialFromYMD(y, m, day)); else clearCell(cell);
    }
  }
  for (const r of rowsWithDailyFormula(sh)) {
    const cC = getCell(sh, r, 3);
    const base = cC.f.text;
    for (const c of dateCols.slice(1)) {
      const cell = ensureCell(sh, r, c, getAttr(cC.attrs, 's'));
      if (c - 2 <= last) setFormula(cell, shiftFormula(base, c - 3)); else clearCell(cell);
    }
  }
  const toNew = (yy, mm, dd) => {
    if (yy !== t.y || mm !== t.m) return null;
    return dd === 1 ? { y, m, d: 1 } : { y, m, d: last };
  };
  for (const row of sh.rows.values()) {
    for (const cell of row.cells.values()) {
      if (cell.f && cell.f.text) cell.f.text = fixDates(cell.f.text, toNew);
    }
  }
  void tLast;
  // AJ·AK(전월 수량·매출액): AH·AI 수식을 전월 범위로 바꿔 넣는다
  const toPrev = (yy, mm, dd) => (yy === y && mm === m ? { y: p.y, m: p.m, d: dd === 1 ? 1 : pLast } : null);
  for (const [r, row] of sh.rows) {
    for (const [src, dst] of [[34, 36], [35, 37]]) { // AH→AJ, AI→AK
      const s = row.cells.get(src);
      const d = row.cells.get(dst);
      if (!s || !s.f || !s.f.text || (d && d.f)) continue;
      const text = /DATE\(/.test(s.f.text) ? fixDates(s.f.text, toPrev) : shiftFormula(s.f.text, dst - src);
      setFormula(ensureCell(sh, r, dst, getAttr(s.attrs, 's')), text);
    }
  }
  dropFormulaCache(sh);

  let xml = serializeSheet(sh);
  // 날짜가 있는 날짜 열은 보이게, 없는 열은 숨긴다(C~AG 범위 안의 col 정의만)
  xml = xml.replace(/<cols>([\s\S]*?)<\/cols>/, (all, inner) => {
    const fixed = inner.replace(/<col\b([^>]*)\/>/g, (cm, attr) => {
      const a = parseAttrs(attr);
      const mn = +getAttr(a, 'min'), mx = +getAttr(a, 'max');
      if (mn < 3 || mx > 33) return cm;
      const anyMissing = mx - 2 > last;
      const allMissing = mn - 2 > last;
      if (allMissing) setAttr(a, 'hidden', '1');
      else if (!anyMissing) setAttr(a, 'hidden', null);
      return `<col${attrsToStr(a)}/>`;
    });
    return `<cols>${fixed}</cols>`;
  });
  // 복제본에는 관계(rels)를 붙이지 않는다
  xml = xml.replace(/\s+xr:uid="[^"]*"/, '')
    .replace(/<drawing\b[^>]*\/>/g, '').replace(/<legacyDrawing\b[^>]*\/>/g, '')
    .replace(/<tableParts\b[\s\S]*?<\/tableParts>/g, '').replace(/<tableParts\b[^>]*\/>/g, '')
    .replace(/<hyperlinks>[\s\S]*?<\/hyperlinks>/g, '')
    .replace(/(<pageSetup\b[^>]*?)\s+r:id="[^"]*"/, '$1');
  return xml;
}

function buildStockSheet(xml, sst, stock) {
  const sh = parseSheet(xml);
  const headers = [];
  const row2 = sh.rows.get(2);
  const maxCol = row2 ? Math.max(...row2.cells.keys()) : 0;
  for (let c = 1; c <= maxCol; c++) headers.push((cellText(getCell(sh, 2, c), sst) || '').trim());
  const missing = headers.filter((h) => h && stock.idx[h] == null);
  if (missing.length) throw new Error(`재고 파일에 '${missing.join("', '")}' 열이 없습니다.`);
  const style = (r, c) => { const x = getCell(sh, r, c); return x ? getAttr(x.attrs, 's') : null; };
  const rowAttr = (r) => {
    const row = sh.rows.get(r);
    return row ? attrsToStr(row.attrs.filter(([k]) => k !== 'r')) : '';
  };
  const s1 = headers.map((_, i) => style(1, i + 1) ?? style(3, i + 1));
  const s2 = headers.map((_, i) => style(2, i + 1));
  const s3 = headers.map((_, i) => style(3, i + 1));
  const sAttr = (v) => (v != null ? ` s="${v}"` : '');
  const cell = (ref, s, v) => {
    if (v == null || v === '') return `<c r="${ref}"${sAttr(s)}/>`;
    if (typeof v === 'number') return `<c r="${ref}"${sAttr(s)}><v>${v}</v></c>`;
    return `<c r="${ref}"${sAttr(s)} t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  };
  let body = `<row r="1"${rowAttr(1)}>`;
  stock.meta.slice(0, 4).forEach((v, i) => { body += cell(numToCol(i + 1) + 1, s1[i], v); });
  body += `</row><row r="2"${rowAttr(2)}>`;
  headers.forEach((h, i) => { body += cell(numToCol(i + 1) + 2, s2[i], h); });
  body += '</row>';
  const ra = rowAttr(3);
  stock.rows.forEach((src, k) => {
    const r = k + 3;
    body += `<row r="${r}"${ra}>`;
    headers.forEach((h, i) => {
      let v = src[stock.idx[h]];
      if (typeof v === 'string') v = v.trim();
      body += cell(numToCol(i + 1) + r, s3[i], v);
    });
    body += '</row>';
  });
  const last = stock.rows.length + 2;
  let out = sh.head + body + sh.tail;
  out = setDimension(out, `A1:${numToCol(headers.length)}${last}`);
  out = out.replace(/<autoFilter ref="([A-Z]+)2:([A-Z]+)\d+"/, `<autoFilter ref="$12:$2${last}"`);
  return { xml: out, last };
}

// 일별 시트: 이번에 들어온 출고일마다 카테고리·모델 × 판매사이트 수량표를 만든다(최신 날짜가 위)
// 일별 시트의 서식은 행 위치가 아니라 내용으로 찾는다(이 도구가 만든 결과를 다시 넣어도 같게)
function dailyStyles(sh, sst) {
  const st = (r, c) => { const x = getCell(sh, r, c); return x ? getAttr(x.attrs, 's') : null; };
  const rows = [...sh.rows.keys()].sort((a, b) => a - b);
  const find = (re) => rows.find((r) => re.test((cellText(getCell(sh, r, 1), sst) || '').trim()));
  const headRow = find(/^카테고리$/);
  const sumRow = find(/요약$/);
  let sumNum = null;
  if (sumRow) for (let c = 3; c <= 30 && sumNum == null; c++) sumNum = st(sumRow, c);
  return {
    title: st(rows[0] || 1, 1),
    head: headRow ? st(headRow, 1) : null,
    headRow,
    sumLabel: sumRow ? st(sumRow, 1) : null,
    sumNum,
  };
}

// 요약 행 서식에서 굵기·배경만 뺀 모델 행 서식을 styles.xml에 더한다. 같은 서식이 이미 있으면 그것을 쓴다.
export function plainDailyStyles(stylesXml, labelXf, numXf) {
  const cx = /<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
  const fo = /<fonts count="(\d+)"([^>]*)>([\s\S]*?)<\/fonts>/.exec(stylesXml);
  if (!cx || !fo || labelXf == null) return { xml: stylesXml, label: labelXf, num: numXf };
  const xfs = cx[2].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g);
  const fonts = fo[3].match(/<font\b[\s\S]*?<\/font>|<font\/>/g);
  let xml = stylesXml;
  let fontList = [...fonts];
  const plainFont = (fid) => {
    const f = fonts[fid];
    if (!f || !f.includes('<b/>')) return fid;
    const nb = f.replace('<b/>', '');
    let i = fontList.indexOf(nb);
    if (i < 0) { fontList.push(nb); i = fontList.length - 1; }
    return i;
  };
  const strip = (xf) => {
    const fid = +(/fontId="(\d+)"/.exec(xf) || [0, 0])[1];
    return xf.replace(/fontId="\d+"/, `fontId="${plainFont(fid)}"`).replace(/fillId="\d+"/, 'fillId="0"').replace(/\s+applyFill="1"/, '');
  };
  let xfList = [...xfs];
  const addXf = (x) => { let i = xfList.indexOf(x); if (i < 0) { xfList.push(x); i = xfList.length - 1; } return i; };
  const label = addXf(strip(xfs[+labelXf]));
  const num = numXf != null ? addXf(strip(xfs[+numXf])) : label;
  if (fontList.length !== fonts.length) {
    xml = xml.replace(fo[0], () => `<fonts count="${fontList.length}"${fo[2]}>${fontList.join('')}</fonts>`);
  }
  if (xfList.length !== xfs.length) {
    xml = xml.replace(/<cellXfs count="\d+">[\s\S]*?<\/cellXfs>/, () => `<cellXfs count="${xfList.length}">${xfList.join('')}</cellXfs>`);
  }
  return { xml, label: String(label), num: String(num) };
}

function buildDailySheet(xml, sst, dates, byDate, modelOrder, stylesRef) {
  const sh = parseSheet(xml);
  const S = dailyStyles(sh, sst);
  const P = stylesRef ? stylesRef(S) : { label: null, num: null };
  const baseMarkets = [];
  for (let c = 3; c <= 30 && S.headRow; c++) {
    const v = cellText(getCell(sh, S.headRow, c), sst);
    if (v && v.trim()) baseMarkets.push(v.trim());
  }
  const sAttr = (v) => (v != null ? ` s="${v}"` : '');
  const c = (ref, s, v) => {
    if (v == null || v === '') return `<c r="${ref}"${sAttr(s)}/>`;
    if (typeof v === 'number') return `<c r="${ref}"${sAttr(s)}><v>${v}</v></c>`;
    return `<c r="${ref}"${sAttr(s)} t="inlineStr"><is><t>${esc(v)}</t></is></c>`;
  };
  let body = '';
  let r = 1;
  let maxCol = 6;
  let lastWritten = 1;
  for (const serial of [...dates].sort((a, b) => b - a)) {
    const list = byDate.get(serial) || [];
    const markets = [...baseMarkets];
    for (const x of list) if (x.market && !markets.includes(x.market)) markets.push(x.market);
    const totalCol = markets.length + 3;
    maxCol = Math.max(maxCol, totalCol);
    const { m, d, dow } = ymdFromSerial(serial);
    body += `<row r="${r}">${c('A' + r, S.title, `${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}(${DOW[dow]})`)}</row>`;
    r++;
    let h = `<row r="${r}">${c('A' + r, S.head, '')}${c('B' + r, S.head, '')}`;
    markets.forEach((_, i) => { h += c(numToCol(i + 3) + r, S.head, i === 0 ? '판매사이트' : ''); });
    h += c(numToCol(totalCol) + r, S.head, '합계') + '</row>';
    body += h; r++;
    let blank = `<row r="${r}">`;
    for (let k = 1; k <= totalCol; k++) blank += c(numToCol(k) + r, S.head, '');
    body += blank + '</row>'; r++;
    let hh = `<row r="${r}">${c('A' + r, S.head, '카테고리')}${c('B' + r, S.head, '모델명')}`;
    markets.forEach((mk, i) => { hh += c(numToCol(i + 3) + r, S.head, mk); });
    hh += c(numToCol(totalCol) + r, S.head, '') + '</row>';
    body += hh; r++;

    const cats = [...new Set([...CATEGORY_ORDER, ...list.map((x) => x.cat).filter(Boolean)])];
    for (const cat of cats) {
      const items = list.filter((x) => x.cat === cat);
      const models = [...new Set(items.map((x) => x.model))].sort((a, b) => {
        const ia = modelOrder.indexOf(a), ib = modelOrder.indexOf(b);
        return (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib) || a.localeCompare(b);
      });
      const sums = markets.map(() => 0);
      models.forEach((model, k) => {
        let line = `<row r="${r}">${c('A' + r, P.label, k === 0 ? cat : '')}${c('B' + r, P.label, model)}`;
        let tot = 0;
        markets.forEach((mk, i) => {
          const q = items.filter((x) => x.model === model && x.market === mk).reduce((s, x) => s + x.qty, 0);
          sums[i] += q; tot += q;
          line += c(numToCol(i + 3) + r, P.num, q || '');
        });
        line += c(numToCol(totalCol) + r, P.num, tot) + '</row>';
        body += line; r++;
      });
      let sl = `<row r="${r}">${c('A' + r, S.sumLabel, `${cat} 요약`)}${c('B' + r, S.sumLabel, '')}`;
      markets.forEach((_, i) => { sl += c(numToCol(i + 3) + r, S.sumNum, sums[i] || ''); });
      sl += c(numToCol(totalCol) + r, S.sumNum, sums.reduce((a, b) => a + b, 0)) + '</row>';
      body += sl; lastWritten = r; r++;
    }
    r++;
  }
  let out = sh.head + body + sh.tail;
  out = setDimension(out, `A1:${numToCol(maxCol)}${lastWritten}`);
  out = out.replace(/<selection\b[^>]*\/>/, '<selection activeCell="A1" sqref="A1"/>');
  return out;
}

function setContentTypeOverride(xml, part) {
  if (xml.includes(`PartName="/${part}"`)) return xml;
  return xml.replace('</Types>',
    `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
}

/**
 * 워크북을 고친다.
 * @returns {Promise<{blob: Uint8Array, report: object}>}
 */
export async function apply(JSZip, base, planned, opts = {}) {
  const { stock = null, price = null } = opts;
  const wb = base.wb;
  const zip = wb.zip;
  const report = { added: planned.rows.length, sheetsCreated: [], sheetsUpdated: [], costChanges: [], warnings: [], stockRows: 0 };
  const oldLast = base.lastRow;
  const newLast = oldLast + planned.rows.length;

  // 1. 모든 시트 XML 로드 + 고정 범위 확장 + 그룹 선택 해제
  const stockLast = stock ? stock.rows.length + 2 : null;
  const texts = new Map();
  const original = new Map();
  for (const s of wb.sheets) {
    let x = wb.cache.get(s.path) ?? await readText(zip, s.path);
    if (x == null) continue;
    original.set(s.name, x);
    if (newLast !== oldLast) x = extendRanges(x, SALES_SHEET, (n) => n === oldLast, newLast);
    if (stockLast) x = extendRanges(x, STOCK_SHEET, (n) => n < stockLast, stockLast);
    texts.set(s.name, x);
  }

  // 2. 판매현황 행 추가
  if (planned.rows.length) {
    let x = texts.get(SALES_SHEET);
    // 판매현황 1행의 SUBTOTAL 범위(시트명 없는 참조)
    x = x.replace(/<row r="1"[\s\S]*?<\/row>/, (row) =>
      row.replace(/(\$?[A-Z]{1,3}\$?3:\$?[A-Z]{1,3}\$?)(\d+)/g, (m, pre, n) => (+n === oldLast ? pre + newLast : m)));
    const add = planned.rows.map((row, i) => salesRowXml(oldLast + 1 + i, row, base.styles)).join('');
    x = x.replace('</sheetData>', add + '</sheetData>');
    x = setDimension(x, `A1:U${newLast}`);
    x = x.replace(/<autoFilter ref="A2:U\d+"/, `<autoFilter ref="A2:U${newLast}"`);
    texts.set(SALES_SHEET, x);
  }

  // 3. 재고 교체
  if (stock) {
    const r = buildStockSheet(texts.get(STOCK_SHEET), wb.sst, stock);
    texts.set(STOCK_SHEET, r.xml);
    report.stockRows = stock.rows.length;
  }

  // 4. 월 시트
  const costByCode = new Map();
  for (const p of planned.products) if (p.match && p.match.item.cost != null) costByCode.set(p.code, p.match.item.cost);
  const monthsNeeded = [...new Set(planned.rows.map((r) => r.yymm))].sort((a, b) => a - b);
  let wbXml = wb.wbXml;
  let relsXml = wb.relsXml;
  let ctXml = await readText(zip, '[Content_Types].xml');
  const sheetOrder = wb.sheets.map((s) => s.name);
  const newSheets = [];
  let months = [...base.months];
  for (const ym of monthsNeeded) {
    const name = String(ym);
    if (texts.has(name)) {
      const sh = parseSheet(texts.get(name));
      const { filled } = refreshMonthSheet(sh, wb.sst, ym);
      updateCosts(sh, wb.sst, costByCode, report.costChanges, name);
      texts.set(name, serializeSheet(sh));
      report.sheetsUpdated.push(`${name} (날짜 칸 ${filled}개 채움)`);
    } else {
      const tpl = months.filter((m) => m < ym).pop() ?? months[months.length - 1];
      if (tpl == null) { report.warnings.push(`${name} 월 시트를 만들 원본 월 시트가 없습니다.`); continue; }
      const xml = buildMonthSheet(texts.get(String(tpl)), wb.sst, tpl, ym);
      const sh = parseSheet(xml);
      updateCosts(sh, wb.sst, costByCode, report.costChanges, name);
      texts.set(name, serializeSheet(sh));
      newSheets.push({ name, after: String(tpl) });
      months = [...months, ym].sort((a, b) => a - b);
      report.sheetsCreated.push(`${name} (${tpl} 시트를 복제)`);
    }
    const listed = base.monthSheetCodes.get(ym) || base.monthSheetCodes.get(months.filter((m) => m < ym).pop()) || new Set();
    const miss = [...new Set(planned.rows.filter((r) => r.yymm === ym && !listed.has(r.code)).map((r) => `${r.code} ${r.model}`))];
    if (miss.length) report.warnings.push(`${name} 시트 상품 목록에 없는 코드(월 시트 합계에서 빠짐): ${miss.join(', ')}`);
  }

  // 5. 일별
  if (planned.rows.length && texts.has(DAILY_SHEET)) {
    const byDate = new Map();
    for (const [k, v] of base.byDate) byDate.set(k, [...v]);
    for (const r of planned.rows) {
      const l = byDate.get(r.serial) || [];
      l.push({ code: r.code, model: r.model, cat: r.cat, market: r.market, qty: r.qty });
      byDate.set(r.serial, l);
    }
    const modelOrder = [];
    for (const v of base.monthCodes.values()) if (!modelOrder.includes(v.model)) modelOrder.push(v.model);
    const dates = [...new Set(planned.rows.map((r) => r.serial))];
    let stylesXml = await readText(zip, 'xl/styles.xml');
    const daily = buildDailySheet(texts.get(DAILY_SHEET), wb.sst, dates, byDate, modelOrder, (S) => {
      const r = plainDailyStyles(stylesXml, S.sumLabel, S.sumNum);
      if (r.xml !== stylesXml) { stylesXml = r.xml; zip.file('xl/styles.xml', stylesXml); }
      return r;
    });
    texts.set(DAILY_SHEET, daily);
    report.daily = dates.length;
  }

  // 6. 새 시트 등록
  let maxSheetFile = Math.max(...wb.sheets.map((s) => +(/sheet(\d+)\.xml$/.exec(s.path) || [0, 0])[1]));
  let maxRid = Math.max(...[...wb.rels.keys()].map((k) => +(/(\d+)$/.exec(k) || [0, 0])[1]));
  let maxSheetId = Math.max(...wb.sheets.map((s) => s.sheetId));
  const paths = new Map(wb.sheets.map((s) => [s.name, s.path]));
  let activeName = null;
  for (const ns of newSheets) {
    const file = `xl/worksheets/sheet${++maxSheetFile}.xml`;
    const rid = `rId${++maxRid}`;
    const sid = ++maxSheetId;
    paths.set(ns.name, file);
    relsXml = relsXml.replace('</Relationships>',
      `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${maxSheetFile}.xml"/></Relationships>`);
    ctXml = setContentTypeOverride(ctXml, file);
    const pos = sheetOrder.indexOf(ns.after) + 1;
    // 원본 월 시트는 숨긴다(지난 달은 숨김으로 두는 기존 관례)
    wbXml = wbXml.replace(new RegExp(`<sheet\\b[^>]*name="${ns.after}"[^>]*/>`), (m) => {
      const a = parseAttrs(m.slice(6, -2));
      setAttr(a, 'state', 'hidden');
      return `<sheet${attrsToStr(a)}/>${'\u0000'}`;
    }).replace('\u0000', `<sheet name="${ns.name}" sheetId="${sid}" r:id="${rid}"/>`);
    // 시트 인덱스를 가리키는 localSheetId를 한 칸씩 민다
    wbXml = wbXml.replace(/localSheetId="(\d+)"/g, (m, n) => (+n >= pos ? `localSheetId="${+n + 1}"` : m));
    sheetOrder.splice(pos, 0, ns.name);
    activeName = ns.name;
  }

  // 7. definedNames(필터 범위)·탭 선택·재계산
  const idx = (n) => sheetOrder.indexOf(n);
  const fixFilter = (sheet, ref) => {
    // ref에 '$2' 같은 문자열이 들어 있어 치환 문자열 대신 함수를 쓴다
    wbXml = wbXml.replace(new RegExp(`(<definedName name="_xlnm._FilterDatabase" localSheetId="${idx(sheet)}"[^>]*>)[^<]*(</definedName>)`),
      (m, open, close) => `${open}${sheet}!${ref}${close}`);
  };
  if (planned.rows.length) fixFilter(SALES_SHEET, `$A$2:$U$${newLast}`);
  if (stock) fixFilter(STOCK_SHEET, `$B$2:$K$${stockLast}`);
  if (activeName) {
    for (const [n, x] of texts) texts.set(n, x.replace(/\s+tabSelected="1"/g, ''));
    const x = texts.get(activeName).replace(/<sheetView\b/, '<sheetView tabSelected="1"');
    texts.set(activeName, x);
    wbXml = wbXml.replace(/(<workbookView\b[^>]*?)\s+activeTab="\d+"/, '$1').replace(/<workbookView\b/, `<workbookView activeTab="${idx(activeName)}"`);
    wbXml = wbXml.replace(/(<workbookView\b[^>]*?)\s+firstSheet="(\d+)"/, (m, pre, n) => (+n > idx(activeName) ? pre : m));
  }
  wbXml = wbXml.replace(/<calcPr\b([^>]*?)\/>/, (m, a) => {
    const at = parseAttrs(a);
    setAttr(at, 'fullCalcOnLoad', '1');
    return `<calcPr${attrsToStr(at)}/>`;
  });

  // calcChain은 셀 목록이 바뀌면 손상 경고를 낸다 — 지우고 Excel이 다시 만들게 한다
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
    relsXml = relsXml.replace(/<Relationship\b[^>]*calcChain[^>]*\/>/, '');
    ctXml = ctXml.replace(/<Override\b[^>]*calcChain[^>]*\/>/, '');
  }

  for (const [n, x] of texts) if (x !== original.get(n)) zip.file(paths.get(n), x);
  zip.file('xl/workbook.xml', wbXml);
  zip.file('xl/_rels/workbook.xml.rels', relsXml);
  zip.file('[Content_Types].xml', ctXml);
  void price;

  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  report.newLast = newLast;
  return { bytes: out, report };
}

export function outputName(baseName, now = new Date()) {
  const stamp = `${String(now.getFullYear() % 100).padStart(2, '0')}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  const stem = baseName.replace(/\.xlsx$/i, '');
  return (/^\d{6}/.test(stem) ? stamp + stem.slice(6) : `${stamp}_${stem}`) + '.xlsx';
}
