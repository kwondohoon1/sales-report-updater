// 가격세팅 워크북의 수식을 다시 계산하는 작은 엑셀 수식 계산기.
// 저장된 값과 다른 셀(재계산 안 됨·값 덮어씀), 오류값, 순환참조를 찾는 데 쓴다.
// 지원: 사칙연산·^·&·비교·%, 셀/범위 참조(시트 포함), IF IFERROR AND OR NOT MIN MAX SUM AVERAGE COUNT
//       ROUND ROUNDDOWN ROUNDUP INT ABS ISBLANK ISNUMBER ISERROR TRUE FALSE.
// 모르는 함수는 {unsupported} 를 돌려주고, 호출한 쪽에서 비교를 건너뛴다.

import { colToNum, numToCol } from './xml.js';

export const isErr = (v) => v != null && typeof v === 'object' && 'err' in v;
export const isUnsupported = (v) => v != null && typeof v === 'object' && v.unsupported === true;
const ERR = (code) => ({ err: code });

// ---------------------------------------------------------------- 토큰

const REF_RE = /^(?:(?:'((?:[^']|'')+)'|([^\s!'"()+\-*/^&=<>,;:%]+))!)?(\$?)([A-Z]{1,3})(\$?)(\d+)/;

export function tokenize(src) {
  const out = [];
  let i = 0;
  const s = src.startsWith('=') ? src.slice(1) : src;
  while (i < s.length) {
    const ch = s[i];
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') { i++; continue; }
    if (ch === '"') {
      let j = i + 1, text = '';
      while (j < s.length) {
        if (s[j] === '"') { if (s[j + 1] === '"') { text += '"'; j += 2; continue; } break; }
        text += s[j++];
      }
      out.push({ t: 'str', v: text });
      i = j + 1;
      continue;
    }
    if (/[0-9.]/.test(ch) && !(ch === '.' && !/[0-9]/.test(s[i + 1] || ''))) {
      const m = /^\d*\.?\d+(?:[eE][+-]?\d+)?|^\d+\.?/.exec(s.slice(i));
      out.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    const two = s.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') { out.push({ t: 'op', v: two }); i += 2; continue; }
    if ('+-*/^&=<>%(),;:'.includes(ch)) { out.push({ t: 'op', v: ch }); i++; continue; }
    const rm = REF_RE.exec(s.slice(i));
    if (rm) {
      const sheet = rm[1] != null ? rm[1].replace(/''/g, "'") : rm[2];
      out.push({ t: 'ref', sheet: sheet || null, col: colToNum(rm[4]), row: +rm[6], absC: !!rm[3], absR: !!rm[5] });
      i += rm[0].length;
      continue;
    }
    const im = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(s.slice(i));
    if (im) {
      const name = im[0].toUpperCase();
      i += im[0].length;
      if (s[i] === '(') out.push({ t: 'fn', v: name });
      else if (name === 'TRUE' || name === 'FALSE') out.push({ t: 'bool', v: name === 'TRUE' });
      else out.push({ t: 'name', v: name });
      continue;
    }
    throw new Error(`수식을 읽을 수 없습니다: ${s.slice(i, i + 12)}`);
  }
  return out;
}

// ---------------------------------------------------------------- 파서 → AST

export function parse(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const take = () => toks[p++];
  const isOp = (v) => peek() && peek().t === 'op' && peek().v === v;

  function comparison() {
    let left = concat();
    while (peek() && peek().t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(peek().v)) {
      const op = take().v;
      left = { k: 'cmp', op, l: left, r: concat() };
    }
    return left;
  }
  function concat() {
    let left = additive();
    while (isOp('&')) { take(); left = { k: 'bin', op: '&', l: left, r: additive() }; }
    return left;
  }
  function additive() {
    let left = multiplicative();
    while (isOp('+') || isOp('-')) { const op = take().v; left = { k: 'bin', op, l: left, r: multiplicative() }; }
    return left;
  }
  function multiplicative() {
    let left = power();
    while (isOp('*') || isOp('/')) { const op = take().v; left = { k: 'bin', op, l: left, r: power() }; }
    return left;
  }
  function power() {
    let left = unary();
    while (isOp('^')) { take(); left = { k: 'bin', op: '^', l: left, r: unary() }; }
    return left;
  }
  function unary() {
    if (isOp('-')) { take(); return { k: 'neg', e: unary() }; }
    if (isOp('+')) { take(); return unary(); }
    return postfix();
  }
  function postfix() {
    let e = primary();
    while (isOp('%')) { take(); e = { k: 'pct', e }; }
    return e;
  }
  function primary() {
    const t = take();
    if (!t) throw new Error('수식이 중간에 끝났습니다');
    if (t.t === 'num') return { k: 'num', v: t.v };
    if (t.t === 'str') return { k: 'str', v: t.v };
    if (t.t === 'bool') return { k: 'bool', v: t.v };
    if (t.t === 'ref') {
      if (isOp(':') && toks[p + 1] && toks[p + 1].t === 'ref') { take(); const b = take(); return { k: 'range', a: t, b }; }
      return { k: 'ref', r: t };
    }
    if (t.t === 'fn') {
      take(); // (
      const args = [];
      if (!isOp(')')) {
        for (;;) {
          if (isOp(',') || isOp(';')) { args.push({ k: 'blank' }); take(); continue; }
          args.push(comparison());
          if (isOp(',') || isOp(';')) { take(); if (isOp(')')) { args.push({ k: 'blank' }); } continue; }
          break;
        }
      }
      if (!isOp(')')) throw new Error(`${t.v}( 의 닫는 괄호가 없습니다`);
      take();
      return { k: 'fn', name: t.v, args };
    }
    if (t.t === 'op' && t.v === '(') {
      const e = comparison();
      if (!isOp(')')) throw new Error('괄호가 맞지 않습니다');
      take();
      return e;
    }
    if (t.t === 'name') return { k: 'name', v: t.v };
    throw new Error(`예상하지 못한 토큰: ${t.v}`);
  }
  const ast = comparison();
  if (p !== toks.length) throw new Error('수식 끝에 남은 글자가 있습니다');
  return ast;
}

// ---------------------------------------------------------------- 값 변환

function toNum(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (isErr(v) || isUnsupported(v)) return v;
  if (typeof v === 'string') {
    const s = v.trim();
    if (s === '') return ERR('#VALUE!');
    if (/^-?\d[\d,]*\.?\d*%$/.test(s)) return parseFloat(s.replace(/,/g, '')) / 100;
    if (/^-?\d[\d,]*\.?\d*$/.test(s) || /^-?\.\d+$/.test(s)) return parseFloat(s.replace(/,/g, ''));
    return ERR('#VALUE!');
  }
  return ERR('#VALUE!');
}
const toStr = (v) => (v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));
const truthy = (v) => {
  if (isErr(v)) return v;
  if (v == null) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (typeof v === 'string') { const s = v.toUpperCase(); if (s === 'TRUE') return true; if (s === 'FALSE') return false; return ERR('#VALUE!'); }
  return false;
};

function compare(op, a, b) {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  // 빈 셀은 상대 타입에 맞춰 0 또는 ""로 본다
  if (a == null) a = typeof b === 'string' ? '' : 0;
  if (b == null) b = typeof a === 'string' ? '' : 0;
  const rank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  let c;
  if (rank(a) !== rank(b)) c = rank(a) - rank(b);
  else if (typeof a === 'string') { const x = a.toUpperCase(), y = b.toUpperCase(); c = x < y ? -1 : x > y ? 1 : 0; }
  else c = Number(a) - Number(b);
  switch (op) {
    case '=': return c === 0;
    case '<>': return c !== 0;
    case '<': return c < 0;
    case '>': return c > 0;
    case '<=': return c <= 0;
    default: return c >= 0;
  }
}

const EPS = 1e-9;
function roundDir(x, digits, mode) {
  const f = Math.pow(10, digits);
  const sign = x < 0 ? -1 : 1;
  const y = Math.abs(x) * f;
  let r;
  if (mode === 'down') r = Math.floor(y + y * 1e-13 + EPS);
  else if (mode === 'up') r = Math.ceil(y - y * 1e-13 - EPS);
  else r = Math.floor(y + 0.5 + EPS);
  return (sign * r) / f;
}

// ---------------------------------------------------------------- 평가

/**
 * @param {(sheet:string|null, col:number, row:number) => any} getCellValue  셀 값(수식이면 평가된 값)
 */
export function evaluate(ast, getCellValue, sheet) {
  const ev = (n) => {
    switch (n.k) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'blank': return null;
      case 'name': return ERR('#NAME?');
      case 'ref': return getCellValue(n.r.sheet || sheet, n.r.col, n.r.row);
      case 'range': return rangeValues(n);
      case 'pct': { const v = toNum(ev(n.e)); return isErr(v) ? v : v / 100; }
      case 'neg': { const v = toNum(ev(n.e)); return isErr(v) ? v : -v; }
      case 'cmp': return compare(n.op, ev(n.l), ev(n.r));
      case 'bin': {
        const a = ev(n.l), b = ev(n.r);
        if (isUnsupported(a) || isUnsupported(b)) return { unsupported: true };
        if (n.op === '&') { if (isErr(a)) return a; if (isErr(b)) return b; return toStr(a) + toStr(b); }
        const x = toNum(a), y = toNum(b);
        if (isErr(x)) return x;
        if (isErr(y)) return y;
        switch (n.op) {
          case '+': return x + y;
          case '-': return x - y;
          case '*': return x * y;
          case '/': return y === 0 ? ERR('#DIV/0!') : x / y;
          default: return Math.pow(x, y);
        }
      }
      case 'fn': return callFn(n);
      default: return ERR('#VALUE!');
    }
  };
  const rangeValues = (n) => {
    const sh = n.a.sheet || sheet;
    const out = [];
    for (let r = Math.min(n.a.row, n.b.row); r <= Math.max(n.a.row, n.b.row); r++) {
      for (let c = Math.min(n.a.col, n.b.col); c <= Math.max(n.a.col, n.b.col); c++) out.push(getCellValue(sh, c, r));
    }
    return { range: out };
  };
  const flat = (args) => {
    const out = [];
    for (const a of args) {
      const v = ev(a);
      if (v && typeof v === 'object' && 'range' in v) out.push(...v.range); else out.push(v);
    }
    return out;
  };
  const nums = (vals, { skipText = true } = {}) => {
    const out = [];
    for (const v of vals) {
      if (isUnsupported(v)) return v;
      if (isErr(v)) return v;
      if (v == null) continue;
      if (typeof v === 'number') out.push(v);
      else if (typeof v === 'boolean') { if (!skipText) out.push(v ? 1 : 0); }
      else if (typeof v === 'string') { if (!skipText) { const n = toNum(v); if (isErr(n)) return n; out.push(n); } }
    }
    return out;
  };
  const callFn = (n) => {
    const { name, args } = n;
    switch (name) {
      case 'IF': {
        const c = truthy(ev(args[0]));
        if (isErr(c)) return c;
        if (isUnsupported(c)) return c;
        if (c) return args[1] ? ev(args[1]) : true;
        return args[2] ? ev(args[2]) : false;
      }
      case 'IFERROR': { const v = ev(args[0]); return isErr(v) ? ev(args[1]) : v; }
      case 'AND': case 'OR': {
        const vals = flat(args);
        let acc = name === 'AND';
        for (const v of vals) {
          if (isUnsupported(v)) return v;
          if (v == null || typeof v === 'string') continue;
          const t = truthy(v);
          if (isErr(t)) return t;
          acc = name === 'AND' ? acc && t : acc || t;
        }
        return acc;
      }
      case 'NOT': { const t = truthy(ev(args[0])); return isErr(t) ? t : !t; }
      // INDEX(한 줄·한 칸짜리 범위, n) / MATCH(값, 범위, 0) — 가격세팅 마켓별상품코드의 '세팅 등재가'에 쓰임
      case 'INDEX': {
        const r = ev(args[0]);
        if (isErr(r) || isUnsupported(r)) return r;
        if (!(r && typeof r === 'object' && 'range' in r) || args.length > 2) return { unsupported: true, fn: 'INDEX' };
        const i = ev(args[1]);
        if (isErr(i) || isUnsupported(i)) return i;
        if (typeof i !== 'number' || i < 1 || i > r.range.length) return ERR('#REF!');
        const v = r.range[Math.trunc(i) - 1];
        return v == null ? 0 : v;
      }
      case 'MATCH': {
        const v = ev(args[0]);
        if (isErr(v) || isUnsupported(v)) return v;
        const r = ev(args[1]);
        if (isErr(r) || isUnsupported(r)) return r;
        const type = args[2] ? ev(args[2]) : 1;
        if (type !== 0 || !(r && typeof r === 'object' && 'range' in r)) return { unsupported: true, fn: 'MATCH' };
        const key = typeof v === 'string' ? v.toLowerCase() : v;
        const at = r.range.findIndex((x) => (typeof x === 'string' ? x.toLowerCase() : x) === key);
        return at < 0 ? ERR('#N/A') : at + 1;
      }
      case 'MIN': case 'MAX': case 'SUM': case 'AVERAGE': case 'COUNT': {
        const vals = flat(args);
        const direct = args.map((a) => a.k !== 'range' && a.k !== 'ref');
        // 직접 쓴 인자(숫자 문자열 등)는 숫자로 바꾸고, 범위·참조 안의 글자는 건너뛴다
        const list = [];
        let i = 0;
        for (const a of args) {
          const v = ev(a);
          if (v && typeof v === 'object' && 'range' in v) { const r = nums(v.range); if (!Array.isArray(r)) return r; list.push(...r); }
          else { const r = nums([v], { skipText: direct[i] }); if (!Array.isArray(r)) return r; list.push(...r); }
          i++;
        }
        void vals;
        if (name === 'COUNT') return list.length;
        if (!list.length) return name === 'AVERAGE' ? ERR('#DIV/0!') : 0;
        if (name === 'MIN') return Math.min(...list);
        if (name === 'MAX') return Math.max(...list);
        const s = list.reduce((x, y) => x + y, 0);
        return name === 'SUM' ? s : s / list.length;
      }
      case 'ROUND': case 'ROUNDDOWN': case 'ROUNDUP': {
        const x = toNum(ev(args[0]));
        if (isErr(x) || isUnsupported(x)) return x;
        const d = args[1] ? toNum(ev(args[1])) : 0;
        if (isErr(d)) return d;
        return roundDir(x, Math.trunc(d), name === 'ROUNDDOWN' ? 'down' : name === 'ROUNDUP' ? 'up' : 'half');
      }
      case 'INT': { const x = toNum(ev(args[0])); return isErr(x) ? x : Math.floor(x); }
      case 'ABS': { const x = toNum(ev(args[0])); return isErr(x) ? x : Math.abs(x); }
      case 'ISBLANK': return ev(args[0]) == null;
      case 'ISNUMBER': return typeof ev(args[0]) === 'number';
      case 'ISERROR': return isErr(ev(args[0]));
      case 'TRUE': return true;
      case 'FALSE': return false;
      default: return { unsupported: true, fn: name };
    }
  };
  return ev(ast);
}

// ---------------------------------------------------------------- 워크북 평가기

const SHEETJS_ERR = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!', 0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A' };

export function cachedValue(cell) {
  if (!cell) return null;
  if (cell.t === 'e') return ERR(cell.w || SHEETJS_ERR[cell.v] || '#ERR');
  if (cell.t === 'n') return typeof cell.v === 'number' ? cell.v : null;
  if (cell.t === 'b') return !!cell.v;
  if (cell.t === 's' || cell.t === 'str') return cell.v == null ? '' : String(cell.v);
  if (cell.t === 'd') return cell.v instanceof Date ? cell.v.getTime() / 86400000 + 25569 : null;
  return cell.v == null ? null : cell.v;
}

/**
 * SheetJS 워크북의 모든 수식을 다시 계산한다.
 * @returns {{ get(sheet, col, row), results: Array<{sheet, addr, formula, cached, computed, status}> }}
 *   status: 'ok' | 'mismatch' | 'error' | 'cycle' | 'unsupported' | 'parse'
 */
export function recalcWorkbook(wb) {
  const memo = new Map();
  const visiting = new Set();
  const results = [];
  const deps = new Map();   // 수식 셀 → 직접 참조한 셀 집합(연쇄 오류의 원인을 가리는 데 쓴다)
  const stack = [];
  const key = (s, c, r) => `${s}!${c},${r}`;
  const sheetOf = (name) => wb.Sheets[name] || (wb.SheetNames.find((n) => n.toUpperCase() === String(name).toUpperCase()) && wb.Sheets[wb.SheetNames.find((n) => n.toUpperCase() === String(name).toUpperCase())]);

  function get(sheet, col, row) {
    const ws = sheetOf(sheet);
    if (!ws) return ERR('#REF!');
    const addr = numToCol(col) + row;
    const k = key(sheet, col, row);
    if (stack.length) stack[stack.length - 1].add(k);
    const cell = ws[addr];
    if (!cell) return null;
    if (!cell.f) return cachedValue(cell);
    if (memo.has(k)) return memo.get(k);
    if (visiting.has(k)) return ERR('#CYCLE!');
    visiting.add(k);
    const mine = new Set();
    stack.push(mine);
    let v;
    try { v = evaluate(parse(cell.f), get, sheet); }
    catch (e) { v = { parseError: e.message }; }
    stack.pop();
    deps.set(k, mine);
    visiting.delete(k);
    memo.set(k, v);
    return v;
  }

  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const cell = ws[addr];
      if (!cell.f) continue;
      const m = /^([A-Z]+)(\d+)$/.exec(addr);
      const computed = get(name, colToNum(m[1]), +m[2]);
      const cached = cachedValue(cell);
      let status = 'ok';
      if (computed && computed.parseError) status = 'parse';
      else if (isUnsupported(computed)) status = 'unsupported';
      else if (isErr(computed) && computed.err === '#CYCLE!') status = 'cycle';
      else if (isErr(cached)) status = 'error';
      else if (!sameValue(computed, cached)) status = 'mismatch';
      results.push({ sheet: name, addr, formula: cell.f, cached, computed, status, key: key(name, colToNum(m[1]), +m[2]) });
    }
  }
  return { get, results, deps };
}

export function sameValue(a, b) {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.err === b.err;
  if (a == null || a === '') return b == null || b === '' || b === 0 && a == null;
  if (b == null || b === '') return a == null || a === '';
  if (typeof a === 'boolean' || typeof b === 'boolean') return a === b;
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= Math.max(0.005, Math.abs(b) * 1e-9);
  if (typeof a === 'number' && typeof b === 'string') { const n = toNum(b); return !isErr(n) && Math.abs(a - n) <= 0.005; }
  if (typeof a === 'string' && typeof b === 'number') { const n = toNum(a); return !isErr(n) && Math.abs(n - b) <= 0.005; }
  return String(a) === String(b);
}

// 수식을 R1C1 식으로 바꿔 같은 열의 다른 행과 패턴을 비교할 수 있게 한다
export function r1c1(formula, row, col) {
  let out = '';
  let i = 0;
  const f = formula.startsWith('=') ? formula.slice(1) : formula;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"') {
      let j = i + 1;
      while (j < f.length) { if (f[j] === '"') { if (f[j + 1] === '"') { j += 2; continue; } break; } j++; }
      out += f.slice(i, j + 1); i = j + 1; continue;
    }
    const m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/.exec(f.slice(i));
    if (m && !/[A-Za-z0-9_.]/.test(f[i - 1] || '')) {
      const c = colToNum(m[2]), r = +m[4];
      out += (m[1] ? `C${c}` : `C[${c - col}]`) + (m[3] ? `R${r}` : `R[${r - row}]`);
      i += m[0].length;
      continue;
    }
    out += ch; i++;
  }
  return out;
}
