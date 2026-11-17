// 레포에는 실제 데이터를 두지 않는다. 순수 함수는 가짜 값으로 시험하고,
// 실제 파일 통합 시험은 SRU_FIXTURES(레포 밖 폴더)를 지정했을 때만 돈다.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import * as core from '../src/core.js';
import { shiftFormula, parseSheet, serializeSheet, expandSharedFormulas, getCell } from '../src/xml.js';

const require = createRequire(import.meta.url);

test('주 코드: 분기 첫날이 속한 월요일 시작 주가 W1', () => {
  assert.equal(core.weekCode(2026, 9, 28), '26Q3W14');
  assert.equal(core.weekCode(2026, 10, 1), '26Q4W1');
  assert.equal(core.weekCode(2026, 10, 5), '26Q4W2');
  assert.equal(core.weekCode(2026, 1, 2), '26Q1W1');
});

test('엑셀 날짜 serial', () => {
  assert.equal(core.serialFromYMD(2026, 9, 28), 46293);
  assert.deepEqual(core.ymdFromSerial(46296), { y: 2026, m: 10, d: 1, dow: 4 });
});

test('상품명 키: 표기 차이만 지우고 모델 구분 토큰은 남긴다', () => {
  const k = core.nameKey;
  assert.equal(k('RTX 5070 Ti OC D7 16GB TF'), k('PNY 지포스 RTX 5070 Ti OC D7 16GB Triple Fan'));
  assert.equal(k('삼성전자 990 PRO M.2 NVMe 1TB'), k('삼성전자 990 PRO M.2 NVMe (1TB)'));
  assert.equal(k('Predator RX 9070 XT WHITE OC D6 16GB'), k('에이서 PREDATOR 라데온 RX 9070 XT BiFrost OC WHITE D6 16GB'));
  assert.notEqual(k('RTX 5060 OC D7 8GB Dual'), k('PNY 지포스 RTX 5060 Ti OC D7 8GB Dual Fan'));
  assert.notEqual(k('KingBank KRRB DDR5-6000 CL30 32GB(16GBx2)'), k('KingBank KRXB DDR5-6000 CL30 32GB(16GBx2)'));
});

test('가격세팅 매칭: 이름이 둘 이상 겹치면 잇지 않는다', () => {
  const price = { items: [
    { name: 'A 1TB', key: core.nameKey('A 1TB'), cost: 10 },
    { name: 'A (1TB)', key: core.nameKey('A (1TB)'), cost: 11 },
    { name: 'B 2TB', key: core.nameKey('B 2TB'), cost: 20 },
  ] };
  assert.equal(core.matchPrice('x', ['A 1TB'], null, price), null);
  assert.equal(core.matchPrice('x', ['B 2TB'], null, price).item.name, 'B 2TB');
  assert.equal(core.matchPrice('x', ['모름'], 11, price).how, '원가');
});

test('수식 이동: 상대 참조만, 문자열·절대 참조는 그대로', () => {
  assert.equal(
    shiftFormula('SUMIFS(판매현황!$D$3:$D$99,판매현황!$O$3:$O$99,C$3,판매현황!$A$3:$A$99,$AP4)', 5),
    'SUMIFS(판매현황!$D$3:$D$99,판매현황!$O$3:$O$99,H$3,판매현황!$A$3:$A$99,$AP4)');
  assert.equal(shiftFormula('SUM(C4,C5,"C6")', 1), 'SUM(D4,D5,"C6")');
  assert.equal(shiftFormula("'2607'!C$2+LOG10(4)", 2), "'2607'!E$2+LOG10(4)");
  assert.equal(shiftFormula('">="&DATE(2026,9,1)', 3), '">="&DATE(2026,9,1)');
});

test('공유 수식 풀기 후 직렬화가 셀을 보존한다', () => {
  const xml = '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1" s="3"><f t="shared" ref="B1:D1" si="0">A1*2</f><v>2</v></c>'
    + '<c r="C1" s="3"><f t="shared" si="0"/><v>4</v></c></row></sheetData></worksheet>';
  const sh = parseSheet(xml);
  expandSharedFormulas(sh);
  assert.equal(getCell(sh, 1, 3).f.text, 'B1*2');
  const out = serializeSheet(sh);
  assert.match(out, /<c r="C1" s="3"><f>B1\*2<\/f><v>4<\/v><\/c>/);
});

test('결과 파일명은 기준일로 앞 6자리를 바꾼다', () => {
  assert.equal(core.outputName('261001_매출현황_컴포넌트.xlsx', { y: 2026, m: 10, d: 2 }), '261002_매출현황_컴포넌트.xlsx');
  assert.equal(core.outputName('매출.xlsx', { y: 2026, m: 10, d: 2 }), '261002_매출.xlsx');
});

test('1영업일 전: 주말·대체공휴일을 건너뛴다', () => {
  const p = (y, m, d) => core.prevBusinessDay(new Date(y, m - 1, d));
  assert.deepEqual(p(2026, 10, 6), { y: 2026, m: 10, d: 2 }); // 10/5 개천절 대체공휴일
  assert.deepEqual(p(2026, 10, 5), { y: 2026, m: 10, d: 2 });
  assert.deepEqual(p(2026, 10, 2), { y: 2026, m: 10, d: 1 });
  assert.deepEqual(p(2026, 10, 12), { y: 2026, m: 10, d: 8 }); // 10/9 한글날
});

test("일별 날짜: 기준일과 그 전 영업일, 최신이 먼저", () => {
  const s = core.serialFromYMD;
  const got = core.dailyDates([s(2026, 10, 1)], { y: 2026, m: 10, d: 2 }, new Map());
  assert.deepEqual(got, [s(2026, 10, 2), s(2026, 10, 1)]);
});

const FIX = process.env.SRU_FIXTURES;
test('실제 파일 통합 시험 (SRU_FIXTURES)', { skip: !FIX && 'SRU_FIXTURES 미지정' }, async () => {
  const XLSX = require('xlsx');
  const JSZip = require('jszip');
  const files = fs.readdirSync(FIX).filter((f) => /\.xlsx?$/i.test(f)).map((f) => path.join(FIX, f));
  const inp = { sales: [] };
  for (const f of files) {
    const buf = new Uint8Array(fs.readFileSync(f));
    const kind = core.detectKind(XLSX, buf);
    if (kind === 'base') inp.base = await core.readBase(JSZip, buf, path.basename(f));
    if (kind === 'sales') inp.sales.push(...core.readSalesFile(XLSX, buf, path.basename(f)));
    if (kind === 'stock') inp.stock = core.readStockFile(XLSX, buf, path.basename(f));
    if (kind === 'price') inp.price = core.readPriceFile(XLSX, buf, path.basename(f));
  }
  assert.ok(inp.base, '기준 워크북 없음');
  const p = core.plan(inp.base, inp.sales, inp.price);
  const { bytes } = await core.apply(JSZip, inp.base, p, { stock: inp.stock, price: inp.price });
  // 결과를 다시 읽으면 같은 판매가 전부 중복으로 걸러져야 한다(멱등)
  const again = await core.readBase(JSZip, bytes, 'out.xlsx');
  assert.equal(again.lastRow, inp.base.lastRow + p.rows.length);
  const p2 = core.plan(again, inp.sales, inp.price);
  assert.equal(p2.rows.length, 0);
  // SheetJS로도 열려야 한다
  const wb = XLSX.read(bytes, { type: 'array', sheetRows: 2 });
  for (const ym of new Set(p.rows.map((r) => String(r.yymm)))) assert.ok(wb.SheetNames.includes(ym));
});

test('전월 수식: 원본 시트의 두 달 전 범위를 직전 달로 옮긴다', () => {
  const f = 'SUMIFS(판매현황!$D$3:$D$9,판매현황!$O$3:$O$9,">="&DATE(2026,9,1),판매현황!$O$3:$O$9,"<="&DATE(2026,9,30),판매현황!$A$3:$A$9,$AP4)';
  assert.equal(core.fixPrevMonthDates(f, 2611), f.replace('DATE(2026,9,1)', 'DATE(2026,10,1)').replace('DATE(2026,9,30)', 'DATE(2026,10,31)'));
  assert.equal(core.fixPrevMonthDates(f, 2610), f); // 이미 직전 달
  // 해 넘김: 2601의 전월은 2512
  const nov = f.replace('DATE(2026,9,1)', 'DATE(2025,11,1)').replace('DATE(2026,9,30)', 'DATE(2025,11,30)');
  assert.equal(core.fixPrevMonthDates(nov, 2601), f.replace('DATE(2026,9,1)', 'DATE(2025,12,1)').replace('DATE(2026,9,30)', 'DATE(2025,12,31)'));
  assert.equal(core.fixPrevMonthDates('SUM(AJ4:AJ9)', 2611), 'SUM(AJ4:AJ9)');
});

test('요약 SUM: 같은 묶음에서 빠진 줄을 더하고 다른 꼴은 그대로', () => {
  const s = (r, t) => `<c r="${r}" t="inlineStr"><is><t>${t}</t></is></c>`;
  const v = (r, n) => `<c r="${r}"><v>${n}</v></c>`;
  const f = (r, t) => `<c r="${r}"><f>${t}</f></c>`;
  const row = (r, cells) => `<row r="${r}">${cells.join('')}</row>`;
  const xml = '<worksheet><sheetData>'
    + row(1, [s('A1', '카테고리'), s('B1', '판매사이트')])
    + row(2, [s('A2', 'RAM'), s('B2', '자사몰'), v('C2', 1), v('D2', 1), v('E2', 1), v('F2', 1)])
    + row(3, [s('B3', '지마켓'), v('C3', 2), v('D3', 2), v('E3', 2), v('F3', 2)])
    + row(4, [s('B4', '11번가'), v('C4', 4), v('D4', 4), v('E4', 4)])
    + row(5, [s('B5', 'B2B'), v('C5', 8), v('D5', 8), v('E5', 8)])
    + row(6, [s('A6', 'RAM 요약'), f('C6', 'SUM(C2,C3,C5)'), f('D6', 'SUM(D2:D5)'), f('E6', 'E2*3'), f('F6', 'SUM(F2,F3)'), f('G6', 'SUM(G2,G1003)')])
    + row(7, [s('A7', 'SSD'), s('B7', '자사몰'), v('C7', 16)])
    + row(8, [s('B8', '11번가'), v('C8', 32)])
    + row(9, [s('A9', 'SSD 요약'), f('C9', 'SUM(C7)'), f('D9', 'SUM(C2:C5)')])
    + '</sheetData></worksheet>';
  const sh = parseSheet(xml);
  const r = core.fixSummaryRows(sh, []);
  assert.equal(getCell(sh, 6, 3).f.text, 'SUM(C2,C3,C5,C4)');
  assert.equal(getCell(sh, 6, 4).f.text, 'SUM(D2:D5)'); // 이미 다 들어 있음
  assert.equal(getCell(sh, 6, 5).f.text, 'E2*3'); // SUM 꼴이 아님
  assert.equal(getCell(sh, 6, 6).f.text, 'SUM(F2,F3,F4,F5)'); // 아직 빈 칸도 묶음이면 더함
  assert.equal(getCell(sh, 6, 7).f.text, 'SUM(G2,G1003)'); // 묶음 밖 참조
  assert.equal(getCell(sh, 9, 3).f.text, 'SUM(C7,C8)');
  assert.equal(getCell(sh, 9, 4).f.text, 'SUM(C2:C5)'); // 다른 열·다른 묶음
  assert.equal(r.cells, 3);
  assert.deepEqual(r.rows, ['RAM 요약(6행)에 11번가·B2B', 'SSD 요약(9행)에 11번가']);
});

test('정산 = 가격세팅 정산액 × 수량 × 0.98', () => {
  const base = { history: new Map(), monthCodes: new Map([['X1', { cat: 'RAM', model: 'A 16GB' }]]), saleKeys: new Set(), monthRows: [], allMonthRows: [], sheetStock: null };
  const price = { blocks: ['G마켓·옥션'], items: [{ name: 'A 16GB', key: core.nameKey('A 16GB'), cost: 100, settle: 278400, market: { 'G마켓·옥션': 278378 } }] };
  const sale = { file: 'f', line: 2, code: 'X1', name: 'A 16GB', price: 300000, qty: 4, market: '지마켓', marketOrder: 'm1', orderId: 'o1', date: { y: 2026, m: 11, d: 13 } };
  const p = core.plan(base, [sale], price);
  assert.equal(core.SETTLE_RATE, 0.98);
  assert.equal(p.rows[0].settle, Math.round(278378 * 4 * 0.98)); // 1091242
  assert.equal(p.rows[0].settleFrom, 'G마켓·옥션');
});

test('B2B 정산은 0.98을 곱하지 않는다', () => {
  const base = { history: new Map(), monthCodes: new Map([['X1', { cat: 'RAM', model: 'A 16GB' }]]), saleKeys: new Set(), monthRows: [], allMonthRows: [], sheetStock: null };
  const price = { blocks: ['G마켓·옥션'], items: [{ name: 'A 16GB', key: core.nameKey('A 16GB'), cost: 100, settle: 278400, market: { 'G마켓·옥션': 278378 } }] };
  const sale = (market, n) => ({ file: 'f', line: n, code: 'X1', name: 'A 16GB', price: 300000, qty: 4, market, marketOrder: 'm' + n, orderId: 'o' + n, date: { y: 2026, m: 11, d: 13 } });
  const p = core.plan(base, [sale('B2B', 2), sale('지마켓', 3)], price);
  assert.equal(p.rows[0].settle, 278400 * 4); // 블록 없음 → 정산가, 100%
  assert.equal(p.rows[1].settle, Math.round(278378 * 4 * 0.98));
});
