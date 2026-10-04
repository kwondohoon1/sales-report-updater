import test from 'node:test';
import assert from 'node:assert/strict';
import { parse, evaluate, r1c1, sameValue, recalcWorkbook } from '../src/formula.js';

const cells = {
  A1: 10, A2: 3, B1: '', B2: '11%', C1: 'abc', D1: null,
};
const get = (sheet, c, r) => {
  const addr = String.fromCharCode(64 + c) + r;
  return addr in cells ? cells[addr] : null;
};
const ev = (f) => evaluate(parse(f), get, 'S');

test('사칙연산·비교·퍼센트·문자열', () => {
  assert.equal(ev('=A1*2+A2'), 23);
  assert.equal(ev('=A1/(1-0.11-0.08)'), 10 / 0.81);
  assert.equal(ev('=A1>=10'), true);
  assert.equal(ev('=B1=""'), true);
  assert.equal(ev('=D1=""'), true); // 빈 셀은 ""와 같다
  assert.equal(ev('=A1*2%'), 0.2);
  assert.equal(ev('="a"&A1'), 'a10');
  assert.equal(ev('=B2*100'), 11); // "11%" 문자열도 숫자로
});

test('IF·OR·MIN·ROUNDDOWN·ROUNDUP·ROUND', () => {
  assert.equal(ev('=IF(A1>5,"big","small")'), 'big');
  assert.equal(ev('=IF(OR(B1="",A2=0),"",A1/A2)'), '');
  assert.equal(ev('=MIN(A1,A2,100)'), 3);
  assert.equal(ev('=ROUNDDOWN(2561430*0.11,-1)'), 281750);
  assert.equal(ev('=ROUNDUP(1234.1,-1)'), 1240);
  assert.equal(ev('=ROUND(2.5,0)'), 3);
  assert.equal(ev('=ROUNDDOWN(-17,-1)'), -10);
  assert.equal(ev('=SUM(A1:A2)'), 13);
});

test('오류 전파와 지원 안 하는 함수', () => {
  assert.deepEqual(ev('=A1/0'), { err: '#DIV/0!' });
  assert.deepEqual(ev('=C1*2'), { err: '#VALUE!' });
  assert.equal(ev('=IFERROR(A1/0,-1)'), -1);
  assert.equal(ev('=VLOOKUP(1,A1:A2,1)').unsupported, true);
});

test('R1C1 변환은 같은 열의 행끼리 패턴을 같게 만든다', () => {
  const a = r1c1('IF(L17="","",L17-ROUNDDOWN(L17*$D$3,-1))', 17, 6);
  const b = r1c1('IF(L18="","",L18-ROUNDDOWN(L18*$D$3,-1))', 18, 6);
  const c = r1c1('IF(L19="","",L19-ROUNDDOWN(L19*$D$3,-1))', 20, 6); // 한 줄 어긋남
  assert.equal(a, b);
  assert.notEqual(a, c);
  assert.equal(r1c1("'2607'!C$2+D5", 5, 3), "'2607'!C[0]R2+C[1]R[0]");
});

test('값 비교는 반올림 오차와 빈 값을 너그럽게 본다', () => {
  assert.ok(sameValue(0.1 + 0.2, 0.3));
  assert.ok(sameValue('', null));
  assert.ok(!sameValue(10, 11));
  assert.ok(sameValue({ err: '#DIV/0!' }, { err: '#DIV/0!' }));
});

test('워크북 재계산: 저장값과 다르면 mismatch, 의존 관계를 기록', () => {
  const wb = { SheetNames: ['S'], Sheets: { S: {
    '!ref': 'A1:C1',
    A1: { t: 'n', v: 5 },
    B1: { t: 'n', v: 10, f: 'A1*2' },
    C1: { t: 'n', v: 999, f: 'B1+1' }, // 저장값이 틀림
  } } };
  const { results, deps } = recalcWorkbook(wb);
  const byAddr = Object.fromEntries(results.map((r) => [r.addr, r]));
  assert.equal(byAddr.B1.status, 'ok');
  assert.equal(byAddr.C1.status, 'mismatch');
  assert.equal(byAddr.C1.computed, 11);
  assert.ok(deps.get('S!3,1').has('S!2,1'));
});
