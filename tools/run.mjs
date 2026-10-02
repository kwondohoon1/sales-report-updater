// 명령줄 실행기(검증용). 웹 화면과 같은 core.js를 쓴다.
//   node tools/run.mjs --out <결과.xlsx> <파일...>
// 입력·결과 파일은 레포 밖에 둔다. 레포에는 어떤 엑셀도 커밋하지 않는다.
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import * as core from '../src/core.js';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const JSZip = require('jszip');

const args = process.argv.slice(2);
const oi = args.indexOf('--out');
const out = oi >= 0 ? args.splice(oi, 2)[1] : null;
const files = args;
if (!files.length) { console.error('usage: node tools/run.mjs --out out.xlsx <files...>'); process.exit(2); }

const repo = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1')), '..');
if (out && path.resolve(out).startsWith(repo)) { console.error('결과 파일을 레포 안에 쓰지 않습니다.'); process.exit(2); }

const inputs = { sales: [] };
for (const f of files) {
  const buf = new Uint8Array(fs.readFileSync(f));
  const kind = core.detectKind(XLSX, buf);
  console.log(`[${kind}] ${path.basename(f)}`);
  if (kind === 'base') inputs.base = await core.readBase(JSZip, buf, path.basename(f));
  else if (kind === 'sales') inputs.sales.push(...core.readSalesFile(XLSX, buf, path.basename(f)));
  else if (kind === 'stock') inputs.stock = core.readStockFile(XLSX, buf, path.basename(f));
  else if (kind === 'price') inputs.price = core.readPriceFile(XLSX, buf, path.basename(f));
  else throw new Error(`판별 실패: ${f}`);
}
const b = inputs.base;
console.log(`판매현황 마지막 행 ${b.lastRow}, 기존 주문번호 ${b.orderIds.size}, 월 시트 ${b.months.join(',')}`);
const p = core.plan(b, inputs.sales, inputs.price);
console.log(`추가 ${p.rows.length}행 / 중복 ${p.skipped.dup.length} / 업로드내중복 ${p.skipped.dupInUpload.length} / 제외 ${p.skipped.excluded.length} / 날짜없음 ${p.skipped.noDate.length}`);
for (const pr of p.products) {
  console.log(`  ${pr.code.padEnd(20)} ${pr.include ? '포함' : '제외'} ${pr.cat}/${pr.model} → ${pr.match ? `${pr.match.item.name} [${pr.match.how}]` : '가격세팅 매칭 없음'}`);
}
for (const r of p.rows) {
  console.log(`  ${r.yymm} ${r.week} ${r.market.padEnd(5)} ${r.code.padEnd(20)} x${r.qty} 결제 ${r.price} 정산 ${r.settle} (${r.settleFrom}) 원가 ${r.cost} (${r.costFrom})`);
}
if (out) {
  const t = Date.now();
  const { bytes, report } = await core.apply(JSZip, b, p, { stock: inputs.stock, price: inputs.price });
  fs.writeFileSync(out, bytes);
  console.log(`쓰기 ${Date.now() - t}ms →`, out);
  console.log(JSON.stringify(report, null, 1));
}
