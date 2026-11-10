// 같은 상품의 다른 ERP 코드를 월 시트 행에 잇는 규칙. 실제 데이터 없이 가짜 값으로 시험한다.
import test from 'node:test';
import assert from 'node:assert/strict';
import { buildCodeResolver, normCode, normErpName } from '../src/codes.js';
import * as core from '../src/core.js';

const base = () => ({
  monthRows: [
    { cat: 'RAM', model: 'ACME D5-4800 16GB', ap: 'WD05-0001', aq: null },
    { cat: 'RAM', model: 'ACME D5-5600 16GB', ap: 'WD05-0002', aq: null },
    { cat: 'SSD', model: 'ACME SSD 2TB', ap: 'BL9XX.130', aq: 'WN08-0003' },
  ],
  allMonthRows: [
    { ym: 2606, model: 'ACME D5-4800 16GB', code: 'BL.9AA.420' },
    { ym: 2610, model: 'ACME D5-4800 16GB', code: 'WD05-0001' },
  ],
  history: new Map([
    ['WD05-0001', { cat: 'RAM', model: 'ACME D5-4800 16GB', name: 'DESK RAM ACME DDR5-4800 16GB #', cost: 100 }],
    ['WD05-0002', { cat: 'RAM', model: 'ACME D5-5600 16GB', name: 'DSEK RAM ACME DDR5-5600 16GB #', cost: 120 }],
  ]),
  monthCodes: new Map(),
  orderIds: new Set(),
  sheetStock: null,
});
const stock = {
  idx: { 상품코드: 0, 상품명: 1, 가용수량: 2, 창고명: 3 },
  rows: [
    ['WD05-0001', 'DESK RAM ACME DDR5-4800 16GB #', 50, '자재'],
    ['BL.9AA.420', 'RAM ACME DDR5-4800 16GB', 7, '완성품'],
    ['K5.NEW.5600', 'RAM ACME DDR5-5600 16GB', 3, '완성품'],
    ['WN08-0003', 'SSD ACME SSD M.2 2TB #', 20, '자재'],
    ['BL.9XX.130', 'SSD ACME SSD M.2 2TB #', 4, '완성품'],
    ['ZZ-9999', 'RAM 다른 상품', 1, '완성품'],
  ],
};

test('코드·상품명 정규화', () => {
  assert.equal(normCode('BL.9BWWR.130'), normCode('BL9BWWR.130'));
  assert.notEqual(normCode('VCG50608DFXPB1-O'), normCode('VCG50608DFXPB1'));
  assert.equal(normErpName('DESK RAM ACER DDR5-4800 16GB #'), normErpName('RAM ACER DDR5-4800 16GB'));
  assert.equal(normErpName('DSEK RAM X 16GB #'), normErpName('RAM X 16GB'));
});

test('바뀐 코드를 월 시트 행에 잇는다 — 점 차이·과거 월 시트·ERP 상품명', () => {
  const r = buildCodeResolver(base(), stock);
  assert.equal(r.resolve('WD05-0001').how, '같은 코드');
  assert.equal(r.resolve('BL.9XX.130').row.ap, 'BL9XX.130');
  assert.equal(r.resolve('BL.9XX.130').how, '코드 표기(점) 차이');
  assert.equal(r.resolve('BL.9AA.420').row.ap, 'WD05-0001');
  assert.equal(r.resolve('BL.9AA.420').how, '과거 월 시트에서 같은 모델');
  assert.equal(r.resolve('K5.NEW.5600').row.ap, 'WD05-0002');
  assert.equal(r.resolve('K5.NEW.5600').how, 'ERP 상품명 같음');
  assert.equal(r.resolve('ZZ-9999'), null);
});

test('상품명이 두 행에 걸치면 잇지 않는다', () => {
  const b = base();
  b.history.set('WD05-0002', { ...b.history.get('WD05-0002'), name: 'DESK RAM ACME DDR5-4800 16GB #' });
  const r = buildCodeResolver(b, { idx: stock.idx, rows: [['K5.X', 'RAM ACME DDR5-4800 16GB', 1, '완성품']] });
  assert.equal(r.resolve('K5.X'), null);
});

test('재고: AP·AQ와 글자가 다른 같은 상품 코드를 찾는다', () => {
  const r = buildCodeResolver(base(), stock);
  assert.deepEqual(r.extraStockCodes(r.rows[0]), ['BL.9AA.420']);
  assert.deepEqual(r.extraStockCodes(r.rows[1]), ['K5.NEW.5600']);
  assert.deepEqual(r.extraStockCodes(r.rows[2]), ['BL.9XX.130']);
});

test('판매 계획: 바뀐 코드는 월 시트 AP 코드로 바꿔 넣고 내역을 남긴다', () => {
  const b = base();
  const sale = (code, qty, name) => ({ file: 'f', line: 1, code, name, price: 1000, qty, market: '지마켓', orderId: `${code}-${qty}`, date: { y: 2026, m: 10, d: 8 } });
  const p = core.plan(b, [sale('BL.9AA.420', 2, 'DESK RAM ACME DDR5-4800 16GB #'), sale('BL.9XX.130', 1, 'SSD ACME SSD M.2 2TB #'), sale('WN08-0003', 3, 'SSD ACME SSD M.2 2TB #')], null, {}, stock);
  assert.equal(p.skipped.excluded.length, 0);
  assert.deepEqual(p.rows.map((r) => r.code), ['WD05-0001', 'BL9XX.130', 'BL9XX.130']);
  assert.deepEqual(p.rows.map((r) => r.origCode), ['BL.9AA.420', 'BL.9XX.130', 'WN08-0003']);
  assert.equal(p.aliases.find((a) => a.from === 'WN08-0003').how, '월 시트 부품(AQ) 코드');
  assert.equal(p.rows[0].model, 'ACME D5-4800 16GB');
});

test('판매 중복: 같은 주문번호라도 마켓주문번호·상품이 다르면 다른 줄', () => {
  const b = { ...base(), saleKeys: new Set([core.saleKey('4273564', '4482422272', 'WD05-0001')]) };
  const sale = (orderId, marketOrder, code, line) => ({ file: 'f', line, code, name: '', price: 1000, qty: 1, market: '지마켓', orderId, marketOrder, date: { y: 2026, m: 10, d: 8 } });
  const p = core.plan(b, [
    sale('4273564', '4482422272', 'WD05-0001', 1), // 워크북에 이미 있음
    sale('4273564', '4482422270', 'WD05-0002', 2), // 같은 주문의 다른 상품 → 새 줄
    sale('4273564', '4482422270', 'WD05-0002', 3), // 같은 파일 안 같은 줄 → 중복
    sale('4273564', '4482422272', 'BL.9AA.420', 4), // 바뀐 코드지만 월 시트 코드(WD05-0001)로는 이미 있음
  ], null, {}, stock);
  assert.equal(p.skipped.dup.length, 2);
  assert.equal(p.skipped.dupInUpload.length, 1);
  assert.deepEqual(p.rows.map((r) => r.code), ['WD05-0002']);
});
