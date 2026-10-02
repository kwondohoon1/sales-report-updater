// 워크시트 XML을 직접 다루는 최소 도구. 원본 파트를 통째로 다시 쓰지 않고
// 필요한 행·셀만 고쳐서 서식·이미지·메모·표를 그대로 보존한다.

export const esc = (s) =>
  String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export const unesc = (s) =>
  s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&amp;/g, '&');

export function colToNum(col) {
  let n = 0;
  for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function numToCol(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function splitRef(ref) {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  return { col: colToNum(m[1]), row: +m[2] };
}

// 속성 문자열 → 순서 보존 배열
export function parseAttrs(s) {
  const out = [];
  const re = /([\w:]+)="([^"]*)"/g;
  let m;
  while ((m = re.exec(s))) out.push([m[1], m[2]]);
  return out;
}

export const attrsToStr = (a) => a.map(([k, v]) => ` ${k}="${v}"`).join('');
export const getAttr = (a, k) => (a.find((x) => x[0] === k) || [])[1];
export function setAttr(a, k, v) {
  const i = a.findIndex((x) => x[0] === k);
  if (v == null) { if (i >= 0) a.splice(i, 1); return; }
  if (i >= 0) a[i][1] = v; else a.push([k, v]);
}

// ---------------------------------------------------------------- 수식 이동

// 수식의 상대 참조를 (dc, dr)만큼 옮긴다. 문자열 리터럴과 따옴표 시트명은 건드리지 않는다.
export function shiftFormula(f, dc, dr = 0) {
  let out = '';
  let i = 0;
  while (i < f.length) {
    const ch = f[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      while (j < f.length) {
        if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue; } break; }
        j++;
      }
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    let j = i;
    while (j < f.length && f[j] !== '"' && f[j] !== "'") j++;
    out += shiftRefs(f.slice(i, j), dc, dr);
    i = j;
  }
  return out;
}

function shiftRefs(s, dc, dr) {
  return s.replace(/(?<![A-Za-z0-9_.$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(A-Za-z_!])/g, (m, ac, col, ar, row) => {
    const c = ac ? colToNum(col) : colToNum(col) + dc;
    const r = ar ? +row : +row + dr;
    if (c < 1 || r < 1) return m;
    return `${ac}${numToCol(c)}${ar}${r}`;
  });
}

// ---------------------------------------------------------------- 시트 모델

// sheetData만 행/셀 단위로 풀고, 앞뒤 XML은 원문 그대로 둔다.
export function parseSheet(xml) {
  let start = xml.indexOf('<sheetData');
  const openEnd = xml.indexOf('>', start);
  let head, body, tail;
  if (xml[openEnd - 1] === '/') {
    head = xml.slice(0, start) + '<sheetData>';
    body = '';
    tail = '</sheetData>' + xml.slice(openEnd + 1);
  } else {
    const end = xml.indexOf('</sheetData>');
    head = xml.slice(0, openEnd + 1);
    body = xml.slice(openEnd + 1, end);
    tail = xml.slice(end);
  }
  const rows = new Map();
  const rowRe = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
  let m;
  while ((m = rowRe.exec(body))) {
    const attrs = parseAttrs(m[1]);
    const r = +getAttr(attrs, 'r');
    const cells = new Map();
    const cellRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let c;
    const inner = m[2] || '';
    while ((c = cellRe.exec(inner))) {
      const ca = parseAttrs(c[1]);
      const ref = getAttr(ca, 'r');
      const cell = { attrs: ca, f: null, v: null, rest: '' };
      const ci = c[2] || '';
      const fm = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(ci);
      if (fm) cell.f = { attrs: parseAttrs(fm[1]), text: fm[2] != null ? unesc(fm[2]) : null };
      const vm = /<v>([\s\S]*?)<\/v>/.exec(ci);
      if (vm) cell.v = vm[1];
      const im = /<is>[\s\S]*?<\/is>/.exec(ci);
      if (im) cell.rest = im[0];
      cells.set(splitRef(ref).col, cell);
    }
    rows.set(r, { attrs, cells });
  }
  return { head, tail, rows };
}

export function serializeSheet(sh) {
  const rowNums = [...sh.rows.keys()].sort((a, b) => a - b);
  let out = '';
  for (const r of rowNums) {
    const row = sh.rows.get(r);
    const cols = [...row.cells.keys()].sort((a, b) => a - b);
    if (!cols.length) { out += `<row${attrsToStr(row.attrs)}/>`; continue; }
    out += `<row${attrsToStr(row.attrs)}>`;
    for (const c of cols) out += cellXml(row.cells.get(c));
    out += '</row>';
  }
  return sh.head + out + sh.tail;
}

function cellXml(cell) {
  let inner = '';
  if (cell.f) {
    inner += cell.f.text == null
      ? `<f${attrsToStr(cell.f.attrs)}/>`
      : `<f${attrsToStr(cell.f.attrs)}>${esc(cell.f.text)}</f>`;
  }
  if (cell.v != null) inner += `<v>${cell.v}</v>`;
  if (cell.rest) inner += cell.rest;
  return inner ? `<c${attrsToStr(cell.attrs)}>${inner}</c>` : `<c${attrsToStr(cell.attrs)}/>`;
}

export function getCell(sh, row, col) {
  const r = sh.rows.get(row);
  return r ? r.cells.get(col) : undefined;
}

export function ensureCell(sh, row, col, styleFrom) {
  let r = sh.rows.get(row);
  if (!r) { r = { attrs: [['r', String(row)]], cells: new Map() }; sh.rows.set(row, r); }
  let cell = r.cells.get(col);
  if (!cell) {
    cell = { attrs: [['r', numToCol(col) + row]], f: null, v: null, rest: '' };
    if (styleFrom != null) cell.attrs.push(['s', String(styleFrom)]);
    r.cells.set(col, cell);
  }
  return cell;
}

export function setFormula(cell, text) {
  cell.f = { attrs: [], text };
  cell.v = null;
  cell.rest = '';
  setAttr(cell.attrs, 't', null);
}

export function setNumber(cell, n) {
  cell.f = null;
  cell.v = String(n);
  cell.rest = '';
  setAttr(cell.attrs, 't', null);
}

export function setText(cell, s) {
  cell.f = null;
  cell.v = null;
  const t = String(s);
  const sp = /^\s|\s$/.test(t) ? ' xml:space="preserve"' : '';
  cell.rest = `<is><t${sp}>${esc(t)}</t></is>`;
  setAttr(cell.attrs, 't', 'inlineStr');
}

export function clearCell(cell) {
  cell.f = null;
  cell.v = null;
  cell.rest = '';
  setAttr(cell.attrs, 't', null);
}

// 공유 수식을 개별 수식으로 푼다. 열을 새로 채우거나 지울 때 공유 범위가 깨지지 않게 하려는 것.
export function expandSharedFormulas(sh) {
  const masters = new Map();
  for (const [r, row] of sh.rows) {
    for (const [c, cell] of row.cells) {
      if (cell.f && getAttr(cell.f.attrs, 't') === 'shared' && cell.f.text != null) {
        masters.set(getAttr(cell.f.attrs, 'si'), { r, c, text: cell.f.text });
      }
    }
  }
  for (const [r, row] of sh.rows) {
    for (const [c, cell] of row.cells) {
      if (!cell.f || getAttr(cell.f.attrs, 't') !== 'shared') continue;
      const mst = masters.get(getAttr(cell.f.attrs, 'si'));
      if (!mst) continue;
      const text = cell.f.text != null ? cell.f.text : shiftFormula(mst.text, c - mst.c, r - mst.r);
      cell.f = { attrs: [], text };
    }
  }
}

// 수식 셀의 캐시 값을 지운다. 파일을 열 때 Excel이 전체 재계산한다.
export function dropFormulaCache(sh) {
  for (const row of sh.rows.values()) {
    for (const cell of row.cells.values()) {
      if (cell.f) { cell.v = null; setAttr(cell.attrs, 't', null); }
    }
  }
}

export function lastRowWithValue(sh, col) {
  let last = 0;
  for (const [r, row] of sh.rows) {
    const c = row.cells.get(col);
    if (c && (c.v != null || c.rest || c.f)) last = Math.max(last, r);
  }
  return last;
}

// sharedStrings.xml → 문자열 배열
export function parseSharedStrings(xml) {
  const out = [];
  if (!xml) return out;
  const re = /<si>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = re.exec(xml))) {
    let s = '';
    const tr = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let t;
    // 윗주(rPh) 안의 <t>는 제외
    const body = m[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, '');
    while ((t = tr.exec(body))) s += unesc(t[1]);
    out.push(s);
  }
  return out;
}

export function cellText(cell, sst) {
  if (!cell) return null;
  const t = getAttr(cell.attrs, 't');
  if (t === 's' && cell.v != null) return sst[+cell.v];
  if (t === 'inlineStr') {
    const m = /<t\b[^>]*>([\s\S]*?)<\/t>/.exec(cell.rest);
    return m ? unesc(m[1]) : '';
  }
  if (t === 'str') return cell.v != null ? unesc(cell.v) : null;
  return cell.v;
}

export function cellNumber(cell) {
  if (!cell || cell.v == null) return null;
  const t = getAttr(cell.attrs, 't');
  if (t && t !== 'n') return null;
  const n = Number(cell.v);
  return Number.isFinite(n) ? n : null;
}
