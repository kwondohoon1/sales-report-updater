// 화면 동작. 계산은 전부 core.js가 하고, 여기서는 파일을 모으고 결과를 보여준다.
import * as core from './core.js';
import * as PC from './pricecheck.js';

const XLSX = window.XLSX;
const JSZip = window.JSZip;
const $ = (s) => document.querySelector(s);

// 기억하는 것은 상품 분류·매칭 설정뿐이다. 판매 데이터는 저장하지 않는다.
const MEM_KEY = 'sru.settings.v1';
function loadMem() {
  try { return JSON.parse(localStorage.getItem(MEM_KEY)) || { products: {}, priceMap: {} }; }
  catch { return { products: {}, priceMap: {} }; }
}
function saveMem() {
  try { localStorage.setItem(MEM_KEY, JSON.stringify(mem)); } catch { /* 저장 불가 환경 */ }
}
let mem = loadMem();

const state = {
  files: { base: null, sales: [], stock: null, price: null }, // {name, buf}
  base: null, salesRows: [], stock: null, price: null,
  planned: null,
  downloadUrl: null,
};

const fmt = (n) => (n == null || n === '' ? '' : Math.round(n).toLocaleString('ko-KR'));
const el = (tag, attrs = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') e.className = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v === true) e.setAttribute(k, '');
    else if (v !== false && v != null) e.setAttribute(k, v);
  }
  for (const k of kids.flat()) if (k != null) e.append(k.nodeType ? k : document.createTextNode(String(k)));
  return e;
};
const dateText = (d) => (d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : '');

function setMsg(text, cls = '') {
  const m = $('#file-msg');
  m.textContent = text;
  m.className = 'msg ' + cls;
}

// ---------------------------------------------------------------- 파일 받기

async function addFiles(list) {
  setMsg('파일을 읽는 중…', 'busy');
  const problems = [];
  for (const f of list) {
    if (!/\.(xlsx|xls)$/i.test(f.name)) { problems.push(`${f.name}: 엑셀 파일이 아닙니다`); continue; }
    const buf = new Uint8Array(await f.arrayBuffer());
    let kind = null;
    try { kind = core.detectKind(XLSX, buf); } catch (e) { problems.push(`${f.name}: 읽지 못했습니다 (${e.message})`); continue; }
    if (!kind) { problems.push(`${f.name}: 어떤 파일인지 알 수 없습니다`); continue; }
    if (kind === 'sales') {
      if (!state.files.sales.some((x) => x.name === f.name && x.buf.length === buf.length)) state.files.sales.push({ name: f.name, buf });
    } else {
      state.files[kind] = { name: f.name, buf };
    }
  }
  renderSlots();
  await analyze(problems);
}

function removeFile(kind, i) {
  if (kind === 'sales') state.files.sales.splice(i, 1);
  else state.files[kind] = null;
  renderSlots();
  analyze([]);
}

function renderSlots() {
  for (const li of document.querySelectorAll('#slots li')) {
    const kind = li.dataset.kind;
    const box = li.querySelector('.files');
    box.replaceChildren();
    const list = kind === 'sales' ? state.files.sales : (state.files[kind] ? [state.files[kind]] : []);
    list.forEach((f, i) => box.append(el('div', { class: 'file' },
      el('span', { title: f.name }, f.name),
      el('button', { type: 'button', title: '빼기', 'aria-label': `${f.name} 빼기`, onclick: () => removeFile(kind, i) }, '×'))));
    li.classList.toggle('filled', list.length > 0);
  }
}

// ---------------------------------------------------------------- 분석

async function readInputs() {
  const f = state.files;
  state.base = f.base ? await core.readBase(JSZip, f.base.buf, f.base.name) : null;
  state.salesRows = [];
  for (const s of f.sales) state.salesRows.push(...core.readSalesFile(XLSX, s.buf, s.name));
  state.stock = f.stock ? core.readStockFile(XLSX, f.stock.buf, f.stock.name) : null;
  state.price = f.price ? core.readPriceFile(XLSX, f.price.buf, f.price.name) : null;
  state.priceCheck = null;
  if (f.price) {
    try { state.priceCheck = PC.checkPriceFile(XLSX, f.price.buf, f.price.name, { base: state.base, reportDate: readReportDate() }); }
    catch (e) { console.error(e); state.priceCheck = { failed: e.message }; }
  }
}

async function analyze(problems) {
  clearDownload();
  try {
    await readInputs();
  } catch (e) {
    problems.push(e.message);
    state.base = null;
  }
  renderPriceCheck();
  if (!state.base) {
    $('#step-check').hidden = true;
    $('#step-build').hidden = true;
    setMsg(problems.length ? problems.join(' · ') : '매출현황 워크북을 넣어 주세요.', problems.length ? 'bad' : '');
    return;
  }
  setMsg(problems.join(' · '), problems.length ? 'bad' : '');
  replan();
}

function replan() {
  state.planned = core.plan(state.base, state.salesRows, state.price, { products: mem.products, priceMap: mem.priceMap }, state.stock);
  renderStock();
  renderCheck();
}

// 가격세팅 검사 결과. 심각도 칩을 눌러 숨기거나 보일 수 있다.
const sevHidden = new Set();
function renderPriceCheck() {
  const card = $('#price-card');
  const pc = state.priceCheck;
  if (!pc) { card.hidden = true; return; }
  card.hidden = false;
  const body = $('#price-body');
  const chips = $('#price-chips');
  if (pc.failed) {
    $('#price-src').textContent = state.files.price.name;
    chips.replaceChildren();
    body.replaceChildren(el('div', { class: 'note bad' }, `검사하지 못했습니다: ${pc.failed}`));
    return;
  }
  const tables = (pc.tables || []).map((t) => `${t.label} ${t.rows}개`).join(' · ');
  $('#price-src').textContent = `${state.files.price.name} · ${tables}${state.base ? '' : ' · 매출현황이 없어 대조 항목은 건너뜀'}`;
  const total = pc.items.length;
  chips.replaceChildren(
    ...(total ? ['error', 'warn', 'info'].map((sev) => el('button', {
      type: 'button', class: `chip ${sev} ${sevHidden.has(sev) ? 'off' : ''}`,
      onclick: () => { if (sevHidden.has(sev)) sevHidden.delete(sev); else sevHidden.add(sev); renderPriceCheck(); },
    }, PC.SEV[sev], el('span', { class: 'n' }, String(pc.counts[sev])))) : [el('span', { class: 'chip ok' }, '문제 없음')]),
  );
  const shown = pc.items.filter((i) => !sevHidden.has(i.sev));
  const groups = PC.CATEGORIES.map((cat) => [cat, shown.filter((i) => i.cat === cat)]).filter(([, l]) => l.length);
  body.replaceChildren(...groups.map(([cat, list]) => {
    const c = { error: 0, warn: 0, info: 0 };
    for (const i of list) c[i.sev]++;
    const head = el('tr', {}, ['구분', '위치', '내용', '상품'].map((h) => el('th', {}, h)));
    const rows = list.map((i) => el('tr', {},
      el('td', {}, el('span', { class: `sev ${i.sev}` }, PC.SEV[i.sev])),
      el('td', { class: 'loc' }, i.sheet ? `${i.sheet}!${i.addr}` : ''),
      el('td', { class: 'msg' }, i.msg),
      el('td', { class: 'wrap dim' }, i.product)));
    return el('details', { class: 'cat', open: c.error + c.warn > 0 },
      el('summary', {}, el('b', {}, cat),
        el('span', { class: 'cnt' }, ...['error', 'warn', 'info'].filter((s) => c[s]).map((s) => el('span', { class: `sev ${s}` }, `${PC.SEV[s]} ${c[s]}`)))),
      el('div', { class: 'table-wrap' }, el('table', { class: 'grid' }, el('thead', {}, head), el('tbody', {}, rows))));
  }));
}

function priceCheckText() {
  const pc = state.priceCheck;
  if (!pc || pc.failed) return '';
  const lines = [`가격세팅 검사 — ${state.files.price.name} — 오류 ${pc.counts.error} · 주의 ${pc.counts.warn} · 정보 ${pc.counts.info}`];
  for (const cat of PC.CATEGORIES) {
    const list = pc.items.filter((i) => i.cat === cat);
    if (!list.length) continue;
    lines.push('', `[${cat}]`);
    for (const i of list) lines.push(`- (${PC.SEV[i.sev]}) ${i.sheet ? `${i.sheet}!${i.addr} ` : ''}${i.product ? `${i.product} — ` : ''}${i.msg}`);
  }
  return lines.join('\n');
}

// 가격세팅 순서대로 재고. 재고 파일이 있으면 그것, 없으면 매출현황의 재고 시트.
function renderStock() {
  const card = $('#stock-card');
  const stock = state.stock || (state.base && state.base.sheetStock);
  if (!state.price || !state.base || !stock) { card.hidden = true; return; }
  card.hidden = false;
  const rows = core.stockByPrice(state.base, state.price, stock, mem.priceMap);
  const when = stock.meta && stock.meta.length >= 4 ? ` · ${stock.meta[3]} 출력` : '';
  $('#stock-src').textContent = `${state.stock ? '재고 파일' : '매출현황 재고 시트(재고 파일 없음)'}${when}`;
  const whs = [...new Set(rows.flatMap((r) => Object.keys(r.byWh)))].sort();
  const head = el('tr', {}, ['카테고리', '모델명', '총재고', '단품', '부품', ...whs, '원가', '재고금액', '코드'].map((h, i) =>
    el('th', { class: i >= 2 && i < 5 + whs.length + 2 ? 'n' : '' }, h)));
  const line = (r, cls) => el('tr', { class: cls },
    el('td', {}, r.category || ''), el('td', { class: 'wrap' }, r.name),
    el('td', { class: 'n strong' }, fmt(r.total)),
    el('td', { class: 'n' }, fmt(r.single)), el('td', { class: 'n' }, fmt(r.part)),
    ...whs.map((w) => el('td', { class: 'n' }, fmt(r.byWh[w] || 0))),
    el('td', { class: 'n' }, fmt(r.cost)), el('td', { class: 'n' }, fmt(r.value)),
    el('td', { class: 'muted' }, r.found === false ? '코드 못 찾음' : [...(r.codes || []), ...(r.partCodes || [])].join(', ')));
  const body = [];
  const sumOf = (list, label) => {
    const t = { category: label, name: '', single: 0, part: 0, total: 0, byWh: {}, cost: null, value: 0 };
    for (const r of list) {
      t.single += r.single; t.part += r.part; t.total += r.total; t.value += r.value || 0;
      for (const [w, v] of Object.entries(r.byWh)) t.byWh[w] = (t.byWh[w] || 0) + v;
    }
    return t;
  };
  const cats = [...new Set(rows.map((r) => r.category))];
  for (const cat of cats) {
    const list = rows.filter((r) => r.category === cat);
    for (const r of list) body.push(line(r, [r.total === 0 ? 'zero' : '', r.found ? '' : 'need'].join(' ')));
    body.push(line(sumOf(list, `${cat} 합계`), 'sum'));
  }
  body.push(line(sumOf(rows, '전체'), 'sum total'));
  $('#stock-table').replaceChildren(el('thead', {}, head), el('tbody', {}, body));
}

function renderCheck() {
  const p = state.planned;
  const b = state.base;
  $('#step-check').hidden = false;
  $('#step-build').hidden = false;

  const qty = p.rows.reduce((s, r) => s + r.qty, 0);
  const amount = p.rows.reduce((s, r) => s + r.price * r.qty, 0);
  const months = [...new Set(p.rows.map((r) => r.yymm))].sort();
  const monthText = months.map((m) => (b.months.includes(m) ? `${m} 갱신` : `${m} 새로 만듦`)).join(', ');
  const dates = [...new Set(p.rows.map((r) => r.serial))].sort();
  const stat = (k, v, s) => el('div', { class: 'stat' }, el('div', { class: 'k' }, k), el('div', { class: String(v).length > 6 ? 'v long' : 'v' }, v), s ? el('div', { class: 's' }, s) : null);
  $('#stats').replaceChildren(
    stat('추가될 판매행', fmt(p.rows.length), `수량 ${fmt(qty)} · 판매총액 ${fmt(amount)}`),
    stat('이미 있는 주문(제외)', fmt(p.skipped.dup.length + p.skipped.dupInUpload.length), '주문번호 기준'),
    stat('미등록 상품(제외)', fmt(p.skipped.excluded.length), '아래 표에서 포함 가능'),
    stat('월 시트', months.length ? monthText : '변경 없음', dates.length ? `출고일 ${dates.map((s) => { const d = core.ymdFromSerial(s); return `${d.m}/${d.d}`; }).join(', ')}` : ''),
    stat('재고', state.stock ? fmt(state.stock.rows.length) + '줄' : '그대로', state.stock ? '재고 시트를 이 파일로 교체' : '재고 파일 없음'),
  );

  const warns = [];
  if (p.aliases && p.aliases.length) {
    warns.push(`같은 상품인데 ERP 코드가 달라 월 시트 코드로 바꿔 넣습니다: ${p.aliases.map((a) => `${a.from} → ${a.to} (${a.model}, ${a.how}, ${a.qty}개)`).join(' · ')}`);
  }
  if (!state.salesRows.length) warns.push('판매 내역 파일이 없습니다. 재고만 바꾸려면 그대로 진행하세요.');
  if (!state.price && p.rows.length) warns.push('가격세팅 파일이 없어 정산이 비고, 원가는 이전 판매행 값을 씁니다.');
  const noSettle = p.rows.filter((r) => r.settle == null);
  if (noSettle.length) warns.push(`정산을 못 채운 행 ${noSettle.length}개 — 상품 표에서 가격세팅 매칭을 골라 주세요.`);
  const fallback = [...new Set(p.rows.filter((r) => r.settleFrom && r.settleFrom.startsWith('정산가')).map((r) => r.market))];
  if (fallback.length) warns.push(`가격세팅에 정산 블록이 없는 판매사이트(${fallback.join(', ')})는 그 상품을 세팅한 표의 정산가를 썼습니다.`);
  if (p.skipped.noDate.length) warns.push(`출고일이 비어 있어 뺀 행 ${p.skipped.noDate.length}개`);
  $('#warnings').replaceChildren(...(warns.length ? [el('div', { class: 'note warn' }, el('ul', {}, warns.map((w) => el('li', {}, w))))] : []));

  renderProducts();
  renderRows();
  renderSkipped();
  $('#build-hint').textContent = p.rows.length || state.stock
    ? `원본 ${b.fileName}에 판매행 ${p.rows.length}개${state.stock ? '와 새 재고' : ''}를 반영한 사본을 만듭니다. 원본 파일은 바뀌지 않습니다.`
    : '반영할 내용이 없습니다.';
  $('#build').disabled = !(p.rows.length || state.stock);
  updateNamePreview();
}

function renderProducts() {
  const p = state.planned;
  const price = state.price;
  const head = el('tr', {}, ['포함', '코드', '카테고리', '모델명', '가격세팅 매칭', '원가', '건수', '수량', '판매 제품명'].map((h) => el('th', {}, h)));
  const body = p.products.map((pr) => {
    const o = mem.products[pr.code] || {};
    const save = (patch) => { mem.products[pr.code] = { ...mem.products[pr.code], ...patch }; saveMem(); replan(); };
    const inc = el('input', { type: 'checkbox', checked: pr.include, onchange: (e) => save({ include: e.target.checked }) });
    const text = (key, val) => (pr.known && !o[key]
      ? el('span', {}, val)
      : el('input', { type: 'text', value: val || '', placeholder: key === 'cat' ? 'RAM/SSD/VGA' : '월 시트 모델명', onchange: (e) => save({ [key]: e.target.value.trim() }) }));
    let match;
    if (!price) match = el('span', { class: 'tag' }, '가격세팅 없음');
    else {
      const auto = pr.code in mem.priceMap ? null : pr.match;
      const sel = el('select', {
        onchange: (e) => {
          const v = e.target.value;
          if (v === '__auto') delete mem.priceMap[pr.code]; else mem.priceMap[pr.code] = v === '__none' ? '' : v;
          saveMem(); replan();
        },
      },
      el('option', { value: '__auto', selected: !(pr.code in mem.priceMap) }, auto ? `자동: ${auto.item.name}` : '자동: 못 찾음'),
      el('option', { value: '__none', selected: mem.priceMap[pr.code] === '' }, '매칭 안 함'),
      price.items.map((it) => el('option', { value: it.name, selected: mem.priceMap[pr.code] === it.name }, `${it.category} · ${it.name}`)));
      match = el('div', {}, sel, ' ', pr.match ? el('span', { class: 'tag ok' }, pr.match.how) : el('span', { class: 'tag warn' }, '없음'));
    }
    const cost = pr.match && pr.match.item.cost != null ? pr.match.item.cost : pr.histCost;
    const need = pr.include && (!pr.cat || !pr.model || (price && !pr.match));
    return el('tr', { class: [pr.include ? '' : 'off', need ? 'need' : ''].join(' ') },
      el('td', {}, inc), el('td', {}, pr.code),
      el('td', {}, text('cat', pr.cat)), el('td', {}, text('model', pr.model)),
      el('td', {}, match), el('td', { class: 'n' }, fmt(cost)),
      el('td', { class: 'n' }, pr.count), el('td', { class: 'n' }, pr.qty),
      el('td', { class: 'muted' }, pr.name));
  });
  $('#products').replaceChildren(el('thead', {}, head), el('tbody', {}, body));
}

function renderRows() {
  const rows = state.planned.rows;
  $('#rows-box').querySelector('summary').textContent = `추가될 판매행 보기 (${rows.length})`;
  const head = el('tr', {}, ['출고일', '판매사이트', '모델명', '수량', '결제금액', '정산', '정산 기준', '원가', '이익', '주', '주문번호'].map((h) => el('th', {}, h)));
  const body = rows.map((r) => {
    const profit = r.settle == null || r.cost == null ? null : r.settle - r.cost * r.qty;
    return el('tr', { class: r.settle == null ? 'need' : '' },
      el('td', {}, dateText(r.date)), el('td', {}, r.market), el('td', {}, r.model),
      el('td', { class: 'n' }, r.qty), el('td', { class: 'n' }, fmt(r.price)),
      el('td', { class: 'n' }, fmt(r.settle)), el('td', {}, r.settleFrom || '—'),
      el('td', { class: 'n' }, fmt(r.cost)), el('td', { class: 'n' }, fmt(profit)),
      el('td', {}, r.week), el('td', {}, r.orderId));
  });
  $('#rows').replaceChildren(el('thead', {}, head), el('tbody', {}, body));
}

function renderSkipped() {
  const s = state.planned.skipped;
  const all = [
    ...s.dup.map((r) => ['이미 판매현황에 있음', r]),
    ...s.dupInUpload.map((r) => ['업로드 파일끼리 중복', r]),
    ...s.excluded.map((r) => ['미등록·제외 상품', r]),
    ...s.noDate.map((r) => ['출고일 없음', r]),
  ];
  $('#skip-box').hidden = !all.length;
  $('#skip-box').querySelector('summary').textContent = `제외된 행 보기 (${all.length})`;
  const head = el('tr', {}, ['사유', '파일', '출고일', '판매사이트', '코드', '제품명', '수량', '주문번호'].map((h) => el('th', {}, h)));
  const body = all.map(([why, r]) => el('tr', {}, el('td', {}, why), el('td', {}, r.file), el('td', {}, dateText(r.date)),
    el('td', {}, r.market), el('td', {}, r.code), el('td', {}, r.name), el('td', { class: 'n' }, r.qty), el('td', {}, r.orderId)));
  $('#skipped').replaceChildren(el('thead', {}, head), el('tbody', {}, body));
}

// ---------------------------------------------------------------- 만들기

function clearDownload() {
  if (state.downloadUrl) URL.revokeObjectURL(state.downloadUrl);
  state.downloadUrl = null;
  $('#download').hidden = true;
  $('#report').replaceChildren();
}

async function build() {
  const btn = $('#build');
  btn.disabled = true;
  btn.textContent = '만드는 중…';
  clearDownload();
  try {
    // apply가 패키지를 직접 고치므로 원본 버퍼에서 다시 읽는다
    const f = state.files.base;
    const base = await core.readBase(JSZip, f.buf, f.name);
    const planned = core.plan(base, state.salesRows, state.price, { products: mem.products, priceMap: mem.priceMap }, state.stock);
    const reportDate = readReportDate();
    const { bytes, report } = await core.apply(JSZip, base, planned, { stock: state.stock, price: state.price, reportDate });
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    state.downloadUrl = URL.createObjectURL(blob);
    const a = $('#download');
    a.href = state.downloadUrl;
    a.download = core.outputName(f.name, reportDate);
    a.textContent = `내려받기 · ${a.download}`;
    a.hidden = false;
    const items = [
      `판매현황 ${report.added}행 추가 (마지막 행 ${report.newLast})`,
      ...report.sheetsCreated.map((s) => `월 시트 생성: ${s}`),
      ...report.sheetsUpdated.map((s) => `월 시트 갱신: ${s}`),
      report.stockRows ? `재고 시트 교체: ${report.stockRows.toLocaleString()}줄` : null,
      report.daily ? `일별 시트: ${report.daily}일치 표(기준일·전 영업일, 수식으로)` : null,
      report.salePrice && report.salePrice.sheet ? `정산가(AX열): ${report.salePrice.sheet} 시트 ${report.salePrice.filled}개 입력${report.salePrice.missing.length ? `, 가격세팅에 없어 '-' ${report.salePrice.missing.length}개` : ''}` : null,
      ...report.costChanges.map((s) => `원가 갱신: ${s}`),
      ...(report.codeAliases || []).map((a) => `코드 변환: ${a.from} → ${a.to} (${a.model}) ${a.qty}개`),
      ...(report.stockAliases || []).map((s) => `재고 코드 추가: ${s}`),
    ].filter(Boolean);
    $('#report').replaceChildren(
      el('ul', {}, items.map((t) => el('li', {}, t))),
      ...(report.warnings.length ? [el('div', { class: 'note warn' }, el('ul', {}, report.warnings.map((w) => el('li', {}, w))))] : []),
      el('p', { class: 'hint' }, '엑셀에서 열면 수식이 전부 다시 계산됩니다. 저장할 때 한 번 더 저장 확인이 뜰 수 있습니다.'),
    );
  } catch (e) {
    console.error(e);
    $('#report').replaceChildren(el('div', { class: 'note bad' }, `만들지 못했습니다: ${e.message}`));
  } finally {
    btn.disabled = false;
    btn.textContent = '최신화한 엑셀 만들기';
  }
}

// 파일명·일별 기준일: 기본은 오늘의 1영업일 전(주말·공휴일 제외), 직접 바꿀 수 있다
function readReportDate() {
  const v = $('#report-date').value;
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v);
  return m ? { y: +m[1], m: +m[2], d: +m[3] } : core.prevBusinessDay(new Date());
}
function initReportDate() {
  const d = core.prevBusinessDay(new Date());
  $('#report-date').value = `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  $('#report-date').addEventListener('change', () => { clearDownload(); updateNamePreview(); if (state.files.price) analyze([]); });
}
function updateNamePreview() {
  const f = state.files.base;
  $('#name-preview').textContent = f ? core.outputName(f.name, readReportDate()) : '';
}

// ---------------------------------------------------------------- 연결
initReportDate();

const drop = $('#drop');
$('#picker').addEventListener('change', (e) => { addFiles([...e.target.files]); e.target.value = ''; });
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => addFiles([...e.dataTransfer.files]));
// 창 아무 데나 떨어뜨려도 브라우저가 파일을 열지 않게
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => { e.preventDefault(); if (!drop.contains(e.target)) addFiles([...e.dataTransfer.files]); });
$('#build').addEventListener('click', build);
$('#price-copy').addEventListener('click', async () => {
  const text = priceCheckText();
  if (!text) return;
  try { await navigator.clipboard.writeText(text); $('#price-copy').textContent = '복사했습니다'; }
  catch { $('#price-copy').textContent = '복사 실패 — 브라우저가 막음'; }
  setTimeout(() => { $('#price-copy').textContent = '검사 결과 복사'; }, 1500);
});
$('#reset-memory').addEventListener('click', () => {
  mem = { products: {}, priceMap: {} };
  saveMem();
  if (state.base) replan();
});

if (!XLSX || !JSZip) setMsg('라이브러리를 불러오지 못했습니다. 페이지를 새로 고쳐 주세요.', 'bad');
else setMsg('매출현황 워크북을 넣어 주세요.');
