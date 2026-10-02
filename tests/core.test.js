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

test('결과 파일명은 오늘 날짜로 앞 6자리를 바꾼다', () => {
  assert.equal(core.outputName('261001_매출현황_컴포넌트.xlsx', new Date(2026, 9, 2)), '261002_매출현황_컴포넌트.xlsx');
  assert.equal(core.outputName('매출.xlsx', new Date(2026, 9, 2)), '261002_매출.xlsx');
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
