// 가격세팅 워크북 검사. 엑셀 자체 오류(오류값·재계산 불일치·수식 패턴 이탈)와
// 가격 업무 규칙(입력가·노출가, 마진, 10원 단위, 시트 간 불일치, 상품코드, 다나와 최저가),
// 매출현황과의 대조(원가 변경, 최대혜택가 이하 판매, 정산 단가 변경)를 항목별로 돌려준다.

import { recalcWorkbook, cachedValue, r1c1, isErr, isUnsupported } from './formula.js';
import { nameKey, ymdFromSerial, serialFromYMD, DEFAULT_MARKET_BLOCK } from './core.js';
import { colToNum, numToCol } from './xml.js';

export const SEV = { error: '오류', warn: '주의', info: '정보' };

// 분류 순서(화면 표시 순)
export const CATEGORIES = ['엑셀 오류', '입력 누락', '가격 오류', '마진', '정산 불일치', '시트 간 불일치', '상품코드', '수수료·쿠폰 설정', '다나와 최저가', '매출현황 대조'];

const LOW_MARGIN = 0.05;      // 마진율 5% 미만이면 주의
const HIGH_MARGIN = 1.0;      // 마진율 100% 초과면 정보로 알림
const SETTLE_DIFF = 50;       // 마켓 정산액과 기준 정산가 차이 허용(원)
const LIST_DIFF = 10;         // 등재가 비교 허용(원)
const COST_JUMP = 0.2;        // 원가 변동 20% 넘으면 주의
const UNDER_BENEFIT = 0.03;   // 최대혜택가보다 3% 넘게 싸게 팔렸으면 주의(최근 7일)
const UNDER_BIG = 0.15;       // 15% 넘게 싸게 팔린 건은 30일까지 본다
const RECENT_DAYS = 30;
const FRESH_DAYS = 7;

const norm = (s) => (s == null ? '' : String(s).replace(/\s+/g, '').replace(/[()]/g, (m) => m));
const fmt = (n) => (n == null || n === '' ? '' : typeof n === 'number' ? Math.round(n).toLocaleString('ko-KR') : String(n));
const pct = (x) => `${(x * 100).toFixed(1)}%`;
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const notTens = (v) => isNum(v) && Math.abs(v - Math.round(v / 10) * 10) > 1e-6;

export function checkPriceFile(XLSX, buf, fileName, ctx = {}) {
  const wb = XLSX.read(buf, { type: 'array', cellNF: false });
  const items = [];
  const add = (sev, cat, sheet, addr, product, msg) => items.push({ sev, cat, sheet, addr, product: product || '', msg });
  const { get, results, deps } = recalcWorkbook(wb);
  // 업무 규칙은 사용자가 엑셀에서 본 값(저장값)으로 검사한다. 재계산 결과는 '저장값 불일치' 탐지에만 쓴다.
  const val = (sheet, c, r) => {
    const ws = wb.Sheets[sheet];
    const cell = ws && ws[numToCol(c) + r];
    if (!cell) return null;
    const v = cachedValue(cell);
    if (isErr(v)) return null;
    if (v == null && cell.f) { const w = get(sheet, c, r); return isErr(w) || isUnsupported(w) || (w && w.parseError) ? null : w; }
    return v;
  };
  const addrOf = (c, r) => numToCol(c) + r;

  // ------------------------------------------------------------ 1. 엑셀 오류
  const unsupported = new Set();
  // 저장값 불일치는 원인 셀만 보고한다: 참조하는 수식 셀 중에 불일치가 없는 셀이 원인이다.
  const bad = new Set(results.filter((r) => r.status === 'mismatch' || r.status === 'error').map((r) => r.key));
  const isRoot = (r) => ![...(deps.get(r.key) || [])].some((k) => bad.has(k));
  const roots = [], cascaded = [];
  for (const r of results) {
    if (r.status === 'cycle') add('error', '엑셀 오류', r.sheet, r.addr, '', `순환 참조 — 수식 ${shortF(r.formula)}`);
    else if (r.status === 'parse') add('warn', '엑셀 오류', r.sheet, r.addr, '', `수식을 해석하지 못함: ${r.computed.parseError}`);
    else if (r.status === 'unsupported' && r.computed.fn) unsupported.add(r.computed.fn);
    else if (r.status === 'error' || r.status === 'mismatch') (isRoot(r) ? roots : cascaded).push(r);
  }
  for (const grp of groupRuns(roots)) {
    const r = grp[0];
    const where = grp.length > 1 ? `${grp[0].addr}~${grp[grp.length - 1].addr} (${grp.length}개 셀)` : r.addr;
    if (r.status === 'error') add('error', '엑셀 오류', r.sheet, where, '', `오류값 ${r.cached.err} — 수식 ${shortF(r.formula)}`);
    else add('error', '엑셀 오류', r.sheet, where, '', `저장된 값(${show(r.cached)})이 수식 결과(${show(r.computed)})와 다름 — 입력값을 바꾼 뒤 재계산이 안 됐거나 값을 덮어쓴 것${grp.length > 1 ? ' (첫 셀 기준)' : ''}`);
  }
  if (cascaded.length) {
    const bySheet = {};
    for (const r of cascaded) bySheet[r.sheet] = (bySheet[r.sheet] || 0) + 1;
    add('info', '엑셀 오류', '', '', '', `위 원인 때문에 연쇄로 저장값이 달라진 셀 ${cascaded.length}개 (${Object.entries(bySheet).map(([s, n]) => `${s} ${n}`).join(', ')}) — 엑셀에서 열어 다시 계산하면 맞춰짐`);
  }
  for (const fn of unsupported) add('info', '엑셀 오류', '', '', '', `${fn} 함수는 검사기가 다시 계산하지 못해 그 수식은 값 비교를 건너뜀`);
  // 수식이 없는 오류값 셀
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    for (const addr in ws) if (addr[0] !== '!' && ws[addr].t === 'e' && !ws[addr].f) add('error', '엑셀 오류', name, addr, '', `오류값 ${ws[addr].w || ''}이 값으로 들어 있음`);
  }
  patternChecks(wb, add);

  // ------------------------------------------------------------ 2. 정산가세팅 읽기
  const PS = '정산가세팅';
  if (!wb.Sheets[PS]) { add('error', '입력 누락', '', '', '', `'${PS}' 시트가 없음`); return finish(items); }
  const grid = sheetGrid(wb.Sheets[PS], (c, r) => val(PS, c, r));
  const tables = findTables(grid);
  if (!tables.length) { add('error', '입력 누락', PS, '', '', '머리글(모델명·정산가·원가)을 찾지 못함'); return finish(items); }
  const params = readParams(grid, tables[0].headerRow);
  const main = tables[0];
  // 상품마다 실제 세팅한 표('상태'=세팅, 없으면 첫 표). 단종은 가격 규칙 검사에서 뺀다
  const basisOf = new Map(), goneKeys = new Set();
  for (const t of tables) for (const row of t.rows) {
    const k = nameKey(row.model);
    if (row.status === '단종') goneKeys.add(k);
    if (row.status === '세팅' && !basisOf.has(k)) basisOf.set(k, { t, row });
  }
  // '세팅' 표시가 없는 상품: 상태 열이 있는 파일이면 자사몰 기준 표(동일스펙 사이트와 같게), 없으면 첫 표
  const hasState = tables.some((t) => t.col['상태']);
  const fallback = (hasState && tables.find((t) => /자사몰/.test(t.label))) || main;
  for (const row of main.rows) {
    const k = nameKey(row.model);
    if (basisOf.has(k)) continue;
    const fr = fallback.rows.find((x) => nameKey(x.model) === k);
    basisOf.set(k, fr ? { t: fallback, row: fr } : { t: main, row });
  }
  const isBasis = (t, row) => { const b = basisOf.get(nameKey(row.model)); return !!b && b.t === t && !goneKeys.has(nameKey(row.model)); };

  // 파라미터 범위
  for (const [label, row] of Object.entries(params.rows)) {
    for (const [mk, v] of Object.entries(row.values)) {
      const a = addrOf(params.cols[mk], row.r);
      if (v == null || v === '') { add('warn', '수수료·쿠폰 설정', PS, a, '', `${mk} ${label} 값이 비어 있음`); continue; }
      if (!isNum(v)) { add('error', '수수료·쿠폰 설정', PS, a, '', `${mk} ${label}이 숫자가 아님: ${show(v)}`); continue; }
      const isRate = /쿠폰|수수료|할인|비율/.test(label) && !/한도|기준금액/.test(label);
      if (isRate && (v < 0 || v > 1)) add('error', '수수료·쿠폰 설정', PS, a, '', `${mk} ${label} ${v} — 비율은 0~1 사이여야 함`);
      if (!isRate && v < 0) add('error', '수수료·쿠폰 설정', PS, a, '', `${mk} ${label}이 음수`);
    }
  }
  const gm = params.marketOf('G마켓') || Object.keys(params.cols)[0];
  const lim = (label) => params.value(label, gm);
  if (isNum(lim('카드할인 한도 (고가)')) && isNum(lim('카드할인 한도 (중저가)')) && lim('카드할인 한도 (고가)') < lim('카드할인 한도 (중저가)')) add('warn', '수수료·쿠폰 설정', PS, '', '', '카드할인 고가 한도가 중저가 한도보다 작음');

  // ------------------------------------------------------------ 3. 표별 가격 규칙
  for (const t of tables) {
    for (const row of t.rows) {
      const p = row.model;
      // 세팅하지 않은 표(같은 정산가를 다른 마켓에 맞춘 참고용)와 단종 상품은 가격·마진 규칙을 보지 않는다
      if (!isBasis(t, row)) continue;
      const A = (name) => addrOf(t.col[name], row.r);
      const E = row.v('입력가'), F = row.v('노출가(2차혜택가)'), G = row.v('최대혜택가'), H = row.v('정산가'), I = row.v('원가');
      const J = row.v('마진'), K = row.v('마진율'), D = row.v('기준');
      if (E == null || E === '') add('error', '입력 누락', PS, A('입력가'), p, `${t.label} 입력가가 비어 있음`);
      if (I == null || I === '' || I === 0) add('error', '입력 누락', PS, A('원가'), p, `${t.label} 원가가 비어 있거나 0`);
      if (D !== '노출가' && D !== '최대혜택가') add('error', '입력 누락', PS, A('기준'), p, `기준이 '노출가'/'최대혜택가'가 아님: ${show(D)}`);
      if (!isNum(E)) continue;
      // 입력가 = 기준 가격
      const target = D === '최대혜택가' ? G : F;
      if (isNum(target) && Math.abs(target - E) > 0.5) add('error', '가격 오류', PS, A(D === '최대혜택가' ? '최대혜택가' : '노출가(2차혜택가)'), p, `${t.label} 입력가 ${fmt(E)}인데 ${D} ${fmt(target)} (차이 ${fmt(target - E)}원) — 등재가 역산이 어긋남`);
      // 10원 단위
      for (const name of ['입력가', '노출가(2차혜택가)', '최대혜택가']) { const v = row.v(name); if (notTens(v)) add('warn', '가격 오류', PS, A(name), p, `${name} ${fmt(v)}이 10원 단위가 아님`); }
      // 마진
      if (isNum(I) && I > E) add('error', '마진', PS, A('원가'), p, `원가 ${fmt(I)}가 입력가 ${fmt(E)}보다 큼`);
      else if (isNum(J) && J < 0) add('error', '마진', PS, A('마진'), p, `${t.label} 마진 ${fmt(J)}원 — 정산가 ${fmt(H)} < 원가 ${fmt(I)}`);
      else if (isNum(K) && K < LOW_MARGIN) add('warn', '마진', PS, A('마진율'), p, `${t.label} 마진율 ${pct(K)} — ${pct(LOW_MARGIN)} 미만`);
      if (isNum(K) && K > HIGH_MARGIN) add('info', '마진', PS, A('마진율'), p, `${t.label} 마진율 ${pct(K)} — 원가가 맞는지 한 번 확인`);
      // 마켓 블록
      for (const b of t.blocks) {
        const list = row.at(b.list), expo = row.at(b.expo), settle = row.at(b.settle), diff = row.at(b.diff);
        // 가격 순서는 그 표의 기준 마켓 묶음에서만 (G마켓 기준 표 → G마켓·옥션, 하이마트 기준 표 → 하이마트)
        if (b.name.replace(/[\s·]/g, '').includes(t.label.replace(/\s*기준$/, '').replace(/\s/g, ''))) {
          // 노출가(MD쿠폰 적용)와 2차혜택가(선택·중복쿠폰 적용)는 서로 다른 할인이라 둘 사이 순서는 보지 않는다
          if (isNum(list) && isNum(expo) && isNum(F) && isNum(G) && !(list >= expo - 0.5 && list >= F - 0.5 && expo >= G - 0.5 && F >= G - 0.5)) add('error', '가격 오류', PS, addrOf(b.list, row.r), p, `가격 순서가 깨짐: 등재가 ${fmt(list)} ≥ 노출가 ${fmt(expo)}·2차혜택가 ${fmt(F)} ≥ 최대혜택가 ${fmt(G)} 이어야 함`);
        }
        if (isNum(H) && (list == null || list === '')) add('error', '정산 불일치', PS, addrOf(b.list, row.r), p, `${b.name} 등재가가 비어 있음`);
        if (notTens(list)) add('warn', '가격 오류', PS, addrOf(b.list, row.r), p, `${b.name} 등재가 ${fmt(list)}이 10원 단위가 아님`);
        if (notTens(expo)) add('warn', '가격 오류', PS, addrOf(b.expo, row.r), p, `${b.name} 노출가 ${fmt(expo)}이 10원 단위가 아님`);
        if (isNum(diff) && Math.abs(diff) > SETTLE_DIFF) add('warn', '정산 불일치', PS, addrOf(b.diff, row.r), p, `${b.name} 정산액 ${fmt(settle)}이 기준 정산가 ${fmt(H)}와 ${fmt(diff)}원 차이`);
        if (isNum(settle) && isNum(I) && settle < I) add('error', '마진', PS, addrOf(b.settle, row.r), p, `${b.name} 정산액 ${fmt(settle)} < 원가 ${fmt(I)} — 그 마켓에서 손실`);
      }
    }
    // 모델명 중복
    const seen = new Map();
    for (const row of t.rows) {
      const k = nameKey(row.model);
      if (seen.has(k)) add('error', '입력 누락', PS, addrOf(t.col['모델명'], row.r), row.model, `${t.label} 표에 같은 상품이 두 번 (${seen.get(k)}행)`);
      else seen.set(k, row.r);
    }
  }
  // 표끼리 비교(G마켓 기준 vs 자사몰 기준 등)
  for (const t of tables.slice(1)) {
    const a = new Map(main.rows.map((r) => [nameKey(r.model), r]));
    for (const row of t.rows) {
      const m = a.get(nameKey(row.model));
      if (!m) { add('warn', '시트 간 불일치', PS, addrOf(t.col['모델명'], row.r), row.model, `${t.label} 표에만 있고 ${main.label} 표에는 없음`); continue; }
      const c1 = m.v('원가'), c2 = row.v('원가');
      if (isNum(c1) && isNum(c2) && Math.abs(c1 - c2) > 0.5) add('error', '시트 간 불일치', PS, addrOf(t.col['원가'], row.r), row.model, `원가가 표마다 다름: ${main.label} ${fmt(c1)} / ${t.label} ${fmt(c2)}`);
    }
    const b = new Set(t.rows.map((r) => nameKey(r.model)));
    for (const row of main.rows) if (!b.has(nameKey(row.model))) add('warn', '시트 간 불일치', PS, addrOf(main.col['모델명'], row.r), row.model, `${main.label} 표에만 있고 ${t.label} 표에는 없음`);
  }

  const byKey = new Map(main.rows.map((r) => [nameKey(r.model), r]));
  const findRow = (name) => byKey.get(nameKey(name));
  const gmBlock = main.blocks.find((b) => b.isBase) || main.blocks[0];

  // ------------------------------------------------------------ 4. 마켓별상품코드
  const MC = '마켓별상품코드';
  let codeRows = [];
  if (wb.Sheets[MC]) {
    const g = sheetGrid(wb.Sheets[MC], (c, r) => val(MC, c, r));
    const hdr = g.findRow((cells) => cells.includes('모델명') && cells.includes('지마켓'));
    if (hdr == null) add('warn', '시트 간 불일치', MC, '', '', '머리글(모델명·지마켓)을 찾지 못함');
    else {
      const col = g.headerMap(hdr);
      const codeCols = Object.keys(col).filter((h) => !['카테고리', '브랜드', '모델명', '노출가', '등재가', '세팅 등재가', '세팅등재가', '마켓', '등재가(%)', '원가', '정산가', '마진율', '다나와'].includes(h));
      const RULES = {
        '지마켓': { re: /^\d{10}$/, need: true }, '옥션': { re: /^F\d{9}$/, need: true }, '자사몰': { re: /^\d{4,8}$/, need: true },
        '하이마트': { re: /^MW\d+_\d+$/, need: true }, '하이마트온라인상품': { re: /^\d{6,12}$/, need: true },
        '스마트스토어': { re: /^\d{8,14}$/, need: true }, '쿠팡': { re: /^\d{8,14}$/, need: false }, '컴퓨존': { re: /^\d+$/, need: false },
      };
      const dup = {};
      const optionalMissing = {};
      for (let r = hdr + 1; r <= g.maxRow; r++) {
        const model = g.get(col['모델명'], r);
        if (model == null || model === '') { if (codeRows.length) break; continue; }
        const codes = {};
        for (const h of codeCols) codes[h] = g.get(col[h], r);
        codeRows.push({ r, model, codes, cost: g.get(col['원가'], r), settle: g.get(col['정산가'], r), list: g.get(col['등재가'], r), rate: g.get(col['등재가(%)'], r) });
        const m = findRow(model);
        if (!m) { add('error', '시트 간 불일치', MC, addrOf(col['모델명'], r), model, '정산가세팅에 없는 상품'); continue; }
        const pc = m.v('원가'), ps = m.v('정산가'), pl = m.at(gmBlock.list);
        const c = g.get(col['원가'], r);
        if (isNum(c) && isNum(pc) && Math.abs(c - pc) > 0.5) add('error', '시트 간 불일치', MC, addrOf(col['원가'], r), model, `원가 ${fmt(c)} ≠ 정산가세팅 ${fmt(pc)}`);
        const s = g.get(col['정산가'], r);
        if (isNum(s) && isNum(ps) && Math.abs(s - ps) > 0.5) add('error', '시트 간 불일치', MC, addrOf(col['정산가'], r), model, `정산가 ${fmt(s)} ≠ 정산가세팅 ${fmt(ps)}`);
        const l = g.get(col['등재가'], r);
        if (isNum(l) && isNum(pl) && Math.abs(l - pl) > LIST_DIFF) add('warn', '시트 간 불일치', MC, addrOf(col['등재가'], r), model, `등재가 ${fmt(l)} ≠ 정산가세팅 ${gmBlock.name} 등재가 ${fmt(pl)}`);
        const rate = g.get(col['등재가(%)'], r);
        const ownRate = params.value('Cat 수수료', params.marketOf('자사몰'));
        if (isNum(rate) && isNum(ownRate) && Math.abs(rate - ownRate) > 1e-6) add('warn', '수수료·쿠폰 설정', MC, addrOf(col['등재가(%)'], r), model, `등재가(%) ${pct(rate)} ≠ 정산가세팅 자사몰 Cat 수수료 ${pct(ownRate)}`);
        for (const h of codeCols) {
          const v = codes[h];
          const s2 = v == null ? '' : String(v).trim();
          const rule = RULES[h] || { re: null, need: false };
          if (!s2) { if (rule.need) add('warn', '상품코드', MC, addrOf(col[h], r), model, `${h} 상품코드 비어 있음`); else (optionalMissing[h] = optionalMissing[h] || []).push(model); continue; }
          if (rule.re && !rule.re.test(s2)) add('warn', '상품코드', MC, addrOf(col[h], r), model, `${h} 코드 형식이 다름: ${s2}`);
          const dk = `${h}|${s2}`;
          if (dup[dk]) add('error', '상품코드', MC, addrOf(col[h], r), model, `${h} 코드 ${s2}가 ${dup[dk]}행과 중복`); else dup[dk] = r;
        }
      }
      for (const [h, list] of Object.entries(optionalMissing)) add('info', '상품코드', MC, addrOf(col[h], hdr), '', `${h} 상품코드 없는 상품 ${list.length}개 (선택 항목): ${list.slice(0, 4).join(', ')}${list.length > 4 ? ' 외' : ''}`);
      const inCodes = new Set(codeRows.map((x) => nameKey(x.model)));
      for (const row of main.rows) if (!inCodes.has(nameKey(row.model))) add('warn', '시트 간 불일치', PS, addrOf(main.col['모델명'], row.r), row.model, `${MC} 시트에 없는 상품`);
    }
  } else add('info', '시트 간 불일치', '', '', '', `'${MC}' 시트 없음 — 상품코드 검사 건너뜀`);

  // ------------------------------------------------------------ 5. 최저가비교
  const LP = '최저가비교';
  if (wb.Sheets[LP]) {
    const g = sheetGrid(wb.Sheets[LP], (c, r) => val(LP, c, r));
    const top = String(g.get(1, 1) || '');
    const dm = /(\d{4})-(\d{2})-(\d{2})/.exec(top);
    if (dm && ctx.reportDate) {
      const got = serialFromYMD(+dm[1], +dm[2], +dm[3]);
      const want = serialFromYMD(ctx.reportDate.y, ctx.reportDate.m, ctx.reportDate.d);
      if (want - got > 1) add('warn', '다나와 최저가', LP, 'A1', '', `최저가 수집일 ${dm[0]} — 기준일보다 ${want - got}일 전 자료`);
    }
    const hdr = g.findRow((cells) => cells.includes('판정') && cells.includes('상품명'));
    if (hdr == null) add('warn', '다나와 최저가', LP, '', '', '머리글(상품명·판정)을 찾지 못함');
    else {
      // 같은 이름의 머리글이 세 번 나온다(현재세팅/제안세팅/가격비교) — 처음 나온 것이 현재세팅
      const cells = g.rowCells(hdr);
      const first = (h) => cells.indexOf(h) + 1;
      const nth = (h, n) => { let k = -1; for (let i = 0; i < n; i++) k = cells.indexOf(h, k + 1); return k + 1; };
      const cName = first('상품명'), cExpo = first('노출가'), cBen = first('최대혜택가'), cCost = first('원가');
      const cPropExpo = nth('노출가', 2), cDanawa = nth('노출가', 3), cGap = first('가격차이'), cVerdict = first('판정'), cCode = first('다나와코드'), cRival = nth('상품명', 2);
      const noDanawa = [];
      for (let r = hdr + 1; r <= g.maxRow; r++) {
        const name = g.get(cName, r);
        if (!name) continue;
        if (goneKeys.has(nameKey(name))) continue;   // 단종
        const m = (basisOf.get(nameKey(name)) || {}).row;
        if (!m) { add('warn', '시트 간 불일치', LP, addrOf(cName, r), name, '정산가세팅에 없는 상품'); continue; }
        const cur = [['노출가', cExpo, m.v('노출가(2차혜택가)')], ['최대혜택가', cBen, m.v('최대혜택가')], ['원가', cCost, m.v('원가')]];
        for (const [lab, c, want] of cur) {
          const v = g.get(c, r);
          if (isNum(v) && isNum(want) && Math.abs(v - want) > 0.5) add('error', '시트 간 불일치', LP, addrOf(c, r), name, `현재세팅 ${lab} ${fmt(v)} ≠ 정산가세팅 ${fmt(want)} — 최저가비교가 옛 세팅 기준`);
        }
        const verdict = String(g.get(cVerdict, r) || '');
        const danawa = g.get(cDanawa, r), gap = g.get(cGap, r), expo = g.get(cExpo, r), prop = g.get(cPropExpo, r);
        const rival = g.get(cRival, r);
        if (!g.get(cCode, r)) noDanawa.push(name);
        if (/손실|보류/.test(verdict)) add('warn', '다나와 최저가', LP, addrOf(cVerdict, r), name, `손실·보류 — 다나와 최저가 ${fmt(danawa)}(${rival || ''})보다 ${fmt(gap)}원 비싸고, 맞추면 손실`);
        else if (/인상/.test(verdict)) add('info', '다나와 최저가', LP, addrOf(cVerdict, r), name, `인상 제안 — 노출가 ${fmt(expo)} → ${fmt(prop)} (다나와 최저가 ${fmt(danawa)})`);
        else if (isNum(gap) && gap > 0) add('warn', '다나와 최저가', LP, addrOf(cGap, r), name, `다나와 최저가 ${fmt(danawa)}보다 ${fmt(gap)}원 비쌈 (판정 ${verdict || '없음'})`);
        if (isNum(prop) && isNum(expo) && !/인상/.test(verdict) && Math.abs(prop - expo) > 0.5) add('info', '다나와 최저가', LP, addrOf(cPropExpo, r), name, `제안 노출가 ${fmt(prop)}가 현재 ${fmt(expo)}와 다름`);
      }
      if (noDanawa.length) add('info', '다나와 최저가', LP, addrOf(cCode, hdr), '', `다나와 코드 미등록 ${noDanawa.length}개 — 유사 상품 최저가로 비교 중: ${noDanawa.slice(0, 4).join(', ')}${noDanawa.length > 4 ? ' 외' : ''}`);
    }
  }

  // ------------------------------------------------------------ 6. G마켓_프로모션
  const PR = wb.SheetNames.find((n) => /프로모션/.test(n));
  // 이름에 '종료'가 붙은 프로모션 시트는 지난 행사 기록이라 지금 세팅과 비교하지 않는다
  if (PR && /종료/.test(PR)) add('info', '시트 간 불일치', PR, '', '', `'${PR}'는 종료된 행사라 지금 가격과 비교하지 않음`);
  else if (PR) {
    const ws = wb.Sheets[PR];
    const g = sheetGrid(ws, (c, r) => val(PR, c, r));
    const hdr = g.findRow((cells) => cells.includes('모델명') && cells.includes('노출가') && cells.includes('등재가'));
    if (hdr != null) {
      const col = g.headerMap(hdr);
      const gmk = params.marketOf('G마켓');
      const want = {
        'MD쿠폰': params.value('MD쿠폰', gmk), '선택쿠폰할인률(%)': params.value('선택쿠폰', gmk), '중복쿠폰할인률(%)': params.value('중복쿠폰 (셀러부담)', gmk),
        '카드즉시할인률(%)': params.value('카드즉시할인', gmk), 'Cat수수료(%)': params.value('Cat 수수료', gmk),
      };
      const codeByKey = new Map(codeRows.map((x) => [nameKey(x.model), x]));
      const seen = new Set();
      for (let r = hdr + 1; r <= g.maxRow; r++) {
        const model = g.get(col['모델명'], r);
        if (!model) continue;
        seen.add(nameKey(model));
        const m = findRow(model);
        if (!m) { add('warn', '시트 간 불일치', PR, addrOf(col['모델명'], r), model, '정산가세팅에 없는 상품'); continue; }
        const expo = g.get(col['노출가'], r), list = g.get(col['등재가'], r);
        if (isNum(expo) && isNum(m.at(gmBlock.expo)) && Math.abs(expo - m.at(gmBlock.expo)) > 0.5) add('error', '시트 간 불일치', PR, addrOf(col['노출가'], r), model, `노출가 ${fmt(expo)} ≠ 정산가세팅 ${gmBlock.name} 노출가 ${fmt(m.at(gmBlock.expo))}`);
        if (isNum(list) && isNum(m.at(gmBlock.list)) && Math.abs(list - m.at(gmBlock.list)) > 0.5) add('error', '시트 간 불일치', PR, addrOf(col['등재가'], r), model, `등재가 ${fmt(list)} ≠ 정산가세팅 ${gmBlock.name} 등재가 ${fmt(m.at(gmBlock.list))}`);
        const est = col['정산예상금액'] ? g.get(col['정산예상금액'], r) : null;
        if (isNum(est) && isNum(m.at(gmBlock.settle)) && Math.abs(est - m.at(gmBlock.settle)) > 1) add('warn', '정산 불일치', PR, addrOf(col['정산예상금액'], r), model, `정산예상금액 ${fmt(est)} ≠ 정산가세팅 ${gmBlock.name} 정산액 ${fmt(m.at(gmBlock.settle))}`);
        const cr = codeByKey.get(nameKey(model));
        if (cr) {
          for (const [h, mk] of [['G_상품코드', '지마켓'], ['A_상품코드', '옥션']]) {
            if (!col[h] || cr.codes[mk] == null) continue;
            const a = String(g.get(col[h], r) ?? '').trim(), b = String(cr.codes[mk] ?? '').trim();
            if (a && b && a !== b) add('error', '상품코드', PR, addrOf(col[h], r), model, `${h} ${a} ≠ 마켓별상품코드 ${mk} ${b}`);
          }
        }
        for (const [h, w] of Object.entries(want)) {
          if (!col[h] || !isNum(w)) continue;
          const raw = g.get(col[h], r);
          const v = typeof raw === 'string' && /%$/.test(raw.trim()) ? parseFloat(raw) / 100 : raw;
          if (isNum(v) && Math.abs(v - w) > 1e-6) add('warn', '수수료·쿠폰 설정', PR, addrOf(col[h], r), model, `${h} ${pct(v)} ≠ 정산가세팅 ${pct(w)}`);
          if (typeof raw === 'string' && /ERROR/i.test(raw)) add('error', '엑셀 오류', PR, addrOf(col[h], r), model, `${h}이 'ERROR' — 스타배송 Y/N 값을 확인`);
        }
        // 카드할인 한도는 수식 상수로 들어 있다 → 파라미터와 비교
        const tcol = col['카드할인최대금액'];
        if (tcol) {
          const cell = ws[addrOf(tcol, r)];
          const nums = cell && cell.f ? (cell.f.match(/\d{4,}/g) || []).map(Number) : [];
          const need = [params.value('카드할인 고가 기준금액', gmk), params.value('카드할인 한도 (고가)', gmk), params.value('카드할인 한도 (중저가)', gmk)].filter(isNum);
          const miss = need.filter((n) => !nums.includes(n));
          if (nums.length && miss.length) add('warn', '수수료·쿠폰 설정', PR, addrOf(tcol, r), model, `카드할인 한도 수식 상수(${nums.join('/')})가 정산가세팅(${need.join('/')})과 다름`);
        }
        for (const h of Object.keys(col)) { const v = g.get(col[h], r); if (typeof v === 'string' && /ERROR/i.test(v) && !(h in want)) add('error', '엑셀 오류', PR, addrOf(col[h], r), model, `${h} 값이 'ERROR'`); }
      }
      for (const row of main.rows) if (!seen.has(nameKey(row.model))) add('info', '시트 간 불일치', PS, addrOf(main.col['모델명'], row.r), row.model, `${PR} 시트에 없는 상품`);
    }
  }

  // ------------------------------------------------------------ 7. 매출현황 대조
  if (ctx.base) baseChecks(ctx.base, main, gmBlock, params, add, PS, addrOf, basisOf, goneKeys);

  return finish(items, { tables: tables.map((t) => ({ label: t.label, rows: t.rows.length })), params });
}

function finish(rawItems, extra = {}) {
  const order = { error: 0, warn: 1, info: 2 };
  const items = collapseRepeats(rawItems);
  items.sort((a, b) => CATEGORIES.indexOf(a.cat) - CATEGORIES.indexOf(b.cat) || order[a.sev] - order[b.sev] || (a.sheet + a.addr).localeCompare(b.sheet + b.addr, 'ko'));
  const counts = { error: 0, warn: 0, info: 0 };
  for (const i of items) counts[i.sev]++;
  return { items, counts, ...extra };
}

const shortF = (f) => (f.length > 60 ? f.slice(0, 57) + '…' : f);

// 같은 시트·같은 열에서 상품만 다르고 내용이 같은 항목이 3개 이상 이어지면 한 줄로 묶는다
function collapseRepeats(items) {
  const sig = (i) => `${i.cat}|${i.sev}|${i.sheet}|${colOf(i.addr)}|${i.msg}`;
  const groups = new Map();
  for (const i of items) {
    if (!i.sheet || !/^[A-Z]+\d+$/.test(i.addr)) { groups.set(Symbol(), [i]); continue; }
    const k = sig(i);
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(i);
  }
  const out = [];
  for (const list of groups.values()) {
    if (list.length < 3) { out.push(...list); continue; }
    list.sort((a, b) => rowOf(a.addr) - rowOf(b.addr));
    const names = [...new Set(list.map((i) => i.product).filter(Boolean))];
    out.push({ ...list[0], addr: `${list[0].addr}~${list[list.length - 1].addr} (${list.length}개 셀)`, product: names.length ? `상품 ${names.length}개` : '', msg: list[0].msg });
  }
  return out;
}

// 같은 시트·같은 열에서 행이 이어지는 셀들을 한 묶음으로
function groupRuns(list) {
  const sorted = [...list].sort((a, b) => a.sheet.localeCompare(b.sheet) || colOf(a.addr).localeCompare(colOf(b.addr)) || rowOf(a.addr) - rowOf(b.addr));
  const out = [];
  for (const r of sorted) {
    const last = out[out.length - 1];
    if (last && last[0].sheet === r.sheet && colOf(last[0].addr) === colOf(r.addr) && rowOf(last[last.length - 1].addr) + 1 === rowOf(r.addr) && last[0].status === r.status) last.push(r);
    else out.push([r]);
  }
  return out;
}
const colOf = (a) => a.replace(/\d+$/, '');
const rowOf = (a) => +a.replace(/^[A-Z]+/, '');
const show = (v) => (v == null ? '빈 값' : isErr(v) ? v.err : typeof v === 'number' ? (Number.isInteger(v) ? v.toLocaleString('ko-KR') : v.toLocaleString('ko-KR', { maximumFractionDigits: 2 })) : typeof v === 'string' ? `'${v}'` : String(v));

// ---------------------------------------------------------------- 시트 격자

function sheetGrid(ws, getter) {
  const ref = ws['!ref'] ? XLSXRange(ws['!ref']) : { r2: 0, c2: 0 };
  const rows = new Map();
  for (const addr in ws) {
    if (addr[0] === '!') continue;
    const m = /^([A-Z]+)(\d+)$/.exec(addr);
    const r = +m[2];
    if (!rows.has(r)) rows.set(r, new Set());
    rows.get(r).add(colToNum(m[1]));
  }
  const g = {
    maxRow: ref.r2, maxCol: ref.c2,
    get: (c, r) => getter(c, r),
    rowCells(r) { const out = []; for (let c = 1; c <= ref.c2; c++) out.push(norm(getter(c, r))); return out; },
    findRow(pred) { for (let r = 1; r <= Math.min(ref.r2, 60); r++) if (rows.has(r) && pred(g.rowCells(r))) return r; return null; },
    headerMap(r) { const out = {}; g.rowCells(r).forEach((h, i) => { if (h && !(h in out)) out[h] = i + 1; }); return out; },
  };
  return g;
}

function XLSXRange(ref) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(ref);
  return { c1: colToNum(m[1]), r1: +m[2], c2: colToNum(m[3] || m[1]), r2: +(m[4] || m[2]) };
}

// 정산가세팅: 머리글 행(모델명·정산가·원가)마다 표 하나. 머리글 위 행에 표 이름(G마켓 기준 등)과 마켓 블록 이름.
function findTables(g) {
  const tables = [];
  for (let r = 1; r <= g.maxRow; r++) {
    const cells = g.rowCells(r);
    if (!(cells.includes('모델명') && cells.includes('정산가') && cells.includes('원가'))) continue;
    const col = {};
    cells.forEach((h, i) => { if (h && !(h in col)) col[h] = i + 1; });
    const above = g.rowCells(r - 1);
    const label = above.find((x) => /기준$/.test(x)) || `표(${r}행)`;
    // 마켓 블록: '정산액' 머리글마다 왼쪽에서 가장 가까운 윗줄 이름
    const blocks = [];
    cells.forEach((h, i) => {
      if (h !== '정산액') return;
      let name = '';
      for (let k = i; k >= 0; k--) if (above[k] && !/기준$|^정산$/.test(above[k])) { name = above[k]; break; }
      const settle = i + 1;
      // 묶음 = 등재가(세팅) · 노출가 · (최대혜택가) · 정산액 · 차이 — 머리글로 자리를 찾는다
      const back = (h) => { for (let k = i - 1; k >= Math.max(0, i - 4); k--) if (cells[k] === h) return k + 1; return null; };
      blocks.push({ name, list: back('등재가(세팅)') || settle - 2, expo: back('노출가') || settle - 1, max: back('최대혜택가'), settle, diff: settle + 1, isBase: /G마켓|지마켓/.test(name) });
    });
    const rows = [];
    for (let rr = r + 1; rr <= g.maxRow; rr++) {
      const model = g.get(col['모델명'], rr);
      if (model == null || model === '') break;
      rows.push({
        r: rr, model: String(model).trim(),
        status: col['상태'] ? String(g.get(col['상태'], rr) ?? '').trim() : '',
        v: (name) => (col[name] ? g.get(col[name], rr) : null),
        at: (c) => (c ? g.get(c, rr) : null),
      });
    }
    tables.push({ headerRow: r, label, col, blocks, rows, headerRowIndex: r });
  }
  return tables;
}

// 머리글 위쪽의 파라미터 표: A열 이름, 1행 마켓 이름, D~ 값
function readParams(g, firstHeaderRow) {
  const cols = {};
  // 항목 이름 열 = 1행의 '마켓 조건' 칸 (앞에 '상태' 열이 끼면 B열). 그 오른쪽 이름이 마켓
  const head = g.rowCells(1);
  const lc = Math.max(0, head.findIndex((h) => h && h.replace(/\s/g, '') === '마켓조건'));
  head.forEach((h, i) => { if (h && i > lc) cols[h] = i + 1; });
  const rows = {};
  for (let r = 2; r < firstHeaderRow - 1; r++) {
    const label = g.get(lc + 1, r);
    if (label == null || label === '') continue;
    const values = {};
    for (const [mk, c] of Object.entries(cols)) values[mk] = g.get(c, r);
    rows[String(label).trim()] = { r, values };
  }
  const marketOf = (part) => Object.keys(cols).find((k) => k.replace(/\s/g, '').includes(part)) || null;
  return {
    cols, rows, marketOf,
    value(label, mk) { const row = rows[label] || rows[Object.keys(rows).find((k) => k.replace(/\s/g, '') === label.replace(/\s/g, ''))]; return row && mk ? row.values[mk] : null; },
  };
}

// 같은 열에서 연속된 수식 셀들의 R1C1 패턴을 비교한다. 다수 패턴과 다른 수식, 수식 자리에 들어간 값을 잡는다.
function patternChecks(wb, add) {
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws['!ref']) continue;
    // 최저가비교는 동일스펙 사이트가 상품별 기준 마켓(G마켓·자사몰·하이마트) 조건으로 수식을 넣어 행마다 다르다
    if (name.replace(/\s/g, '') === '최저가비교') continue;
    const { c2, r2 } = XLSXRange(ws['!ref']);
    for (let c = 1; c <= c2; c++) {
      let run = [];
      const flush = () => {
        if (run.length >= 4) {
          const counts = new Map();
          for (const x of run) if (x.pat) counts.set(x.pat, (counts.get(x.pat) || 0) + 1);
          const [top, n] = [...counts.entries()].sort((a, b) => b[1] - a[1])[0] || [null, 0];
          if (top && n >= Math.max(3, run.length * 0.6)) {
            for (const x of run) {
              if (!x.pat && x.isNum) add('warn', '엑셀 오류', name, x.addr, '', `수식 열에 값 ${fmt(x.v)}이 직접 들어 있음 (주변은 ${shortF(x.sample)})`);
              else if (x.pat && x.pat !== top) add('error', '엑셀 오류', name, x.addr, '', `같은 열의 다른 행과 수식이 다름: ${shortF(x.f)}`);
            }
          }
        }
        run = [];
      };
      for (let r = 1; r <= r2; r++) {
        const addr = numToCol(c) + r;
        const cell = ws[addr];
        if (!cell) { flush(); continue; }
        if (cell.f) run.push({ addr, pat: r1c1(cell.f, r, c), f: cell.f, sample: cell.f });
        else if (cell.t === 'n') run.push({ addr, pat: null, isNum: true, v: cell.v, sample: run.length ? run[run.length - 1].sample : '' });
        else flush();
      }
      flush();
    }
  }
}

// ---------------------------------------------------------------- 매출현황 대조

function baseChecks(base, main0, gmBlock0, params, add, PS, addrOf, basisOf, goneKeys) {
  const dates = [...base.byDate.keys()];
  if (!dates.length) return;
  const latest = Math.max(...dates);
  const since = latest - RECENT_DAYS;
  for (const row0 of main0.rows) {
    const key = nameKey(row0.model);
    if (goneKeys && goneKeys.has(key)) continue;
    // 그 상품을 세팅한 표로 비교 (없으면 첫 표)
    const bs = (basisOf && basisOf.get(key)) || { t: main0, row: row0 };
    const main = bs.t, row = bs.row;
    const gmBlock = main.blocks.find((b) => b.isBase) || gmBlock0;
    const codes = new Set();
    for (const r of base.monthRows) if (nameKey(r.model) === key) codes.add(r.ap);
    if (!codes.size) for (const [code, h] of base.history) if (h.model && nameKey(h.model) === key) codes.add(code);
    if (!codes.size) { add('warn', '매출현황 대조', PS, addrOf(main.col['모델명'], row.r), row.model, '매출현황 월 시트·판매이력에서 코드를 못 찾음 — 정산·원가 자동 입력이 안 됨'); continue; }
    const cost = row.v('원가');
    for (const code of codes) {
      const h = base.history.get(code);
      if (h && isNum(h.cost) && isNum(cost) && Math.abs(h.cost - cost) > 0.5) {
        const ratio = Math.abs(cost - h.cost) / h.cost;
        add(ratio > COST_JUMP ? 'warn' : 'info', '매출현황 대조', PS, addrOf(main.col['원가'], row.r), row.model, `원가 변경 ${fmt(h.cost)} → ${fmt(cost)} (${cost > h.cost ? '+' : ''}${pct((cost - h.cost) / h.cost)}) — 마지막 판매행 기준`);
      }
    }
    // 최근 판매: 최대혜택가 이하 판매, 정산 단가 변경
    const ben = row.v('최대혜택가');
    const under = [];
    const lastSettle = new Map(); // market → {serial, unit}
    for (const [serial, list] of base.byDate) {
      if (serial < since) continue;
      for (const x of list) {
        if (!codes.has(x.code) || !x.qty) continue;
        if (isNum(ben) && isNum(x.price)) {
          const fresh = latest - serial < FRESH_DAYS;
          if (x.price < ben * (1 - (fresh ? UNDER_BENEFIT : UNDER_BIG))) under.push({ serial, market: x.market, price: x.price });
        }
        if (isNum(x.settle) && x.qty) {
          const prev = lastSettle.get(x.market);
          if (!prev || prev.serial < serial) lastSettle.set(x.market, { serial, unit: x.settle / x.qty });
        }
      }
    }
    if (under.length) {
      const min = under.reduce((a, b) => (b.price < a.price ? b : a));
      const d = ymdFromSerial(min.serial);
      add('warn', '매출현황 대조', PS, addrOf(main.col['최대혜택가'], row.r), row.model, `최대혜택가 ${fmt(ben)}보다 싸게 팔린 건 ${under.length}건(최근 ${FRESH_DAYS}일 ${pct(UNDER_BENEFIT)}↑ 또는 ${RECENT_DAYS}일 ${pct(UNDER_BIG)}↑) — 최저 ${d.m}/${d.d} ${min.market} ${fmt(min.price)}원 (${pct((ben - min.price) / ben)} 쌈)`);
    }
    const changes = [];
    for (const [market, s] of lastSettle) {
      const blockName = DEFAULT_MARKET_BLOCK[market];
      const block = main.blocks.find((b) => blockName && b.name.replace(/[\s·.]/g, '') === blockName.replace(/[\s·.]/g, '')) || gmBlock;
      const now = row.at(block.settle);
      if (isNum(now) && Math.abs(now - s.unit) > 1) {
        const d = ymdFromSerial(s.serial);
        changes.push(`${market} ${fmt(s.unit)}(${d.m}/${d.d}) → ${fmt(now)}`);
      }
    }
    if (changes.length) add('info', '매출현황 대조', PS, addrOf(gmBlock.settle, row.r), row.model, `정산 단가가 마지막 판매행과 다름: ${changes.join(' · ')}`);
  }
}
