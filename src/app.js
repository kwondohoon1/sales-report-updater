// 화면 동작. 계산은 전부 core.js가 하고, 여기서는 파일을 모으고 결과를 보여준다.
import * as core from './core.js';

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
}

async function analyze(problems) {
  clearDownload();
  try {
    await readInputs();
  } catch (e) {
    problems.push(e.message);
    state.base = null;
  }
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
  state.planned = core.plan(state.base, state.salesRows, state.price, { products: mem.products, priceMap: mem.priceMap });
  renderCheck();
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
  if (!state.salesRows.length) warns.push('판매 내역 파일이 없습니다. 재고만 바꾸려면 그대로 진행하세요.');
  if (!state.price && p.rows.length) warns.push('가격세팅 파일이 없어 정산이 비고, 원가는 이전 판매행 값을 씁니다.');
  const noSettle = p.rows.filter((r) => r.settle == null);
  if (noSettle.length) warns.push(`정산을 못 채운 행 ${noSettle.length}개 — 상품 표에서 가격세팅 매칭을 골라 주세요.`);
  const fallback = [...new Set(p.rows.filter((r) => r.settleFrom && r.settleFrom.startsWith('정산가')).map((r) => r.market))];
  if (fallback.length) warns.push(`가격세팅에 정산 블록이 없는 판매사이트(${fallback.join(', ')})는 G마켓 기준 정산가를 썼습니다.`);
  if (p.skipped.noDate.length) warns.push(`출고일이 비어 있어 뺀 행 ${p.skipped.noDate.length}개`);
  $('#warnings').replaceChildren(...(warns.length ? [el('div', { class: 'note warn' }, el('ul', {}, warns.map((w) => el('li', {}, w))))] : []));

  renderProducts();
  renderRows();
  renderSkipped();
  $('#build-hint').textContent = p.rows.length || state.stock
    ? `원본 ${b.fileName}에 판매행 ${p.rows.length}개${state.stock ? '와 새 재고' : ''}를 반영한 사본을 만듭니다. 원본 파일은 바뀌지 않습니다.`
    : '반영할 내용이 없습니다.';
  $('#build').disabled = !(p.rows.length || state.stock);
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
    const planned = core.plan(base, state.salesRows, state.price, { products: mem.products, priceMap: mem.priceMap });
    const { bytes, report } = await core.apply(JSZip, base, planned, { stock: state.stock, price: state.price });
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    state.downloadUrl = URL.createObjectURL(blob);
    const a = $('#download');
    a.href = state.downloadUrl;
    a.download = core.outputName(f.name);
    a.textContent = `내려받기 · ${a.download}`;
    a.hidden = false;
    const items = [
      `판매현황 ${report.added}행 추가 (마지막 행 ${report.newLast})`,
      ...report.sheetsCreated.map((s) => `월 시트 생성: ${s}`),
      ...report.sheetsUpdated.map((s) => `월 시트 갱신: ${s}`),
      report.stockRows ? `재고 시트 교체: ${report.stockRows.toLocaleString()}줄` : null,
      report.daily ? `일별 시트: 출고일 ${report.daily}일치 표` : null,
      ...report.costChanges.map((s) => `원가 갱신: ${s}`),
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

// ---------------------------------------------------------------- 연결

const drop = $('#drop');
$('#picker').addEventListener('change', (e) => { addFiles([...e.target.files]); e.target.value = ''; });
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => addFiles([...e.dataTransfer.files]));
// 창 아무 데나 떨어뜨려도 브라우저가 파일을 열지 않게
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => { e.preventDefault(); if (!drop.contains(e.target)) addFiles([...e.dataTransfer.files]); });
$('#build').addEventListener('click', build);
$('#reset-memory').addEventListener('click', () => {
  mem = { products: {}, priceMap: {} };
  saveMem();
  if (state.base) replan();
});

if (!XLSX || !JSZip) setMsg('라이브러리를 불러오지 못했습니다. 페이지를 새로 고쳐 주세요.', 'bad');
else setMsg('매출현황 워크북을 넣어 주세요.');
