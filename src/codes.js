// 같은 상품의 다른 ERP 코드를 최신 월 시트의 행(AP 단품·AQ 부품 코드)에 잇는다.
// ERP는 같은 상품을 자재 코드(WD05-1072, 'DESK RAM … #')와 완성품 코드(BL.9BWWA.420, 'RAM …')로
// 나눠 들고 있고, 출고 코드가 바뀌거나 코드에 점이 붙기도 한다. 월 시트 SUMIFS/SUMIF는 코드 글자가
// 정확히 같아야 세므로, 판매 코드는 월 시트 AP 코드로 바꿔 쓰고 재고는 바뀐 코드까지 더하게 한다.
//
// 같은 상품으로 보는 근거(하나라도 맞고, 맞는 월 시트 행이 하나뿐일 때만):
//   1. 코드에서 점·공백만 다름              BL.9BWWR.130 ↔ BL9BWWR.130
//   2. 과거 월 시트에서 같은 모델명으로 쓰임   BL.9BWWA.420 = 2606 시트 'ACER D5-4800 16GB'
//   3. ERP 상품명이 같음(DESK·RAM·SSD·VGA 머리말과 '#'만 다름)

import { nameKey } from './core.js';

export const normCode = (c) => String(c || '').toUpperCase().replace(/[.\s]/g, '');

export function normErpName(n) {
  let s = String(n || '').toUpperCase().replace(/#/g, ' ').trim();
  for (;;) {
    const t = s.replace(/^(DESK|DSEK|RAM|SSD|VGA)\s+/, '');
    if (t === s) break;
    s = t;
  }
  return s.replace(/\s+/g, '');
}

/**
 * @param base   readBase 결과(monthRows, allMonthRows, history)
 * @param stock  재고(readStockFile 결과 또는 base.sheetStock) — 상품명 출처
 * @param sales  판매 행 — 상품명 출처
 */
export function buildCodeResolver(base, stock, sales = []) {
  const rows = base.monthRows.map((r, i) => ({ ...r, i }));
  const own = new Map(); // 월 시트에 그대로 있는 코드 → 행
  for (const r of rows) {
    own.set(r.ap, r);
    if (r.aq) own.set(r.aq, r);
  }

  // 코드 → ERP 상품명들
  const names = new Map();
  const addName = (code, name) => {
    if (!code || !name) return;
    const k = normErpName(name);
    if (!k) return;
    if (!names.has(code)) names.set(code, new Set());
    names.get(code).add(k);
  };
  if (stock) {
    const ci = stock.idx['상품코드'], ni = stock.idx['상품명'];
    if (ci != null && ni != null) for (const r of stock.rows) addName(String(r[ci] ?? '').trim(), r[ni]);
  }
  for (const [code, h] of base.history) addName(code, h.name);
  for (const s of sales) addName(s.code, s.name);

  // 월 시트 행마다 이름 지문(AP·AQ 코드의 ERP 상품명)
  const rowByName = new Map();
  for (const r of rows) {
    for (const c of [r.ap, r.aq]) {
      for (const k of names.get(c) || []) {
        if (!rowByName.has(k)) rowByName.set(k, new Set());
        rowByName.get(k).add(r);
      }
    }
  }
  const rowByNorm = new Map();
  for (const r of rows) {
    for (const c of [r.ap, r.aq]) {
      if (!c) continue;
      const k = normCode(c);
      if (!rowByNorm.has(k)) rowByNorm.set(k, new Set());
      rowByNorm.get(k).add(r);
    }
  }
  const rowByModel = new Map();
  for (const r of rows) {
    const k = nameKey(r.model);
    if (!rowByModel.has(k)) rowByModel.set(k, new Set());
    rowByModel.get(k).add(r);
  }
  const pastModel = new Map(); // 코드 → 과거 월 시트 모델명들
  for (const x of base.allMonthRows || []) {
    if (!pastModel.has(x.code)) pastModel.set(x.code, new Set());
    pastModel.get(x.code).add(x.model);
  }

  const one = (set) => (set && set.size === 1 ? [...set][0] : null);
  const cache = new Map();

  /** 코드가 속한 월 시트 행. 그대로 있으면 how='같은 코드'. */
  function resolve(code) {
    code = String(code || '').trim();
    if (!code) return null;
    if (cache.has(code)) return cache.get(code);
    let out = null;
    if (own.has(code)) out = { row: own.get(code), how: '같은 코드' };
    if (!out) { const r = one(rowByNorm.get(normCode(code))); if (r) out = { row: r, how: '코드 표기(점) 차이' }; }
    if (!out && pastModel.has(code)) {
      const cand = new Set();
      for (const m of pastModel.get(code)) for (const r of rowByModel.get(nameKey(m)) || []) cand.add(r);
      const r = one(cand);
      if (r) out = { row: r, how: '과거 월 시트에서 같은 모델' };
    }
    if (!out && names.has(code)) {
      const cand = new Set();
      for (const k of names.get(code)) for (const r of rowByName.get(k) || []) cand.add(r);
      const r = one(cand);
      if (r) out = { row: r, how: 'ERP 상품명 같음' };
    }
    cache.set(code, out);
    return out;
  }

  /** 재고 파일에서 이 행의 상품인데 AP·AQ와 글자가 다른 코드들 */
  function extraStockCodes(row) {
    if (!stock) return [];
    const ci = stock.idx['상품코드'];
    const seen = new Set();
    for (const r of stock.rows) {
      const c = String(r[ci] ?? '').trim();
      if (!c || c === row.ap || c === row.aq || seen.has(c)) continue;
      const res = resolve(c);
      if (res && res.row.i === row.i) seen.add(c);
    }
    return [...seen];
  }

  return { rows, resolve, extraStockCodes };
}
