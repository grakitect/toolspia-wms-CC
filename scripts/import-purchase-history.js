// 자료/2024-2026 구매내역 3.xlsx 를 읽어
//   1) 구매처 마스터 → db.partners.purchase 에 병합
//   2) 구매품목 마스터 → db.purchaseProductMaster (품목코드 기준, 품명/규격 분리)
//   3) 이카운트 창고 리스트(2번 시트) → db.ecountWarehouses
//   4) 구매이력(전체 거래 원장) → data/purchase-history.json (별도 파일, db.json에는 넣지 않음)
// 을 채운다.
//
// 기본은 미리보기(dry-run)만 하고 실제로는 아무 파일도 바꾸지 않는다.
// 실제로 반영하려면: node scripts/import-purchase-history.js --write

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const ROOT = path.join(__dirname, '..');
const SOURCE_FILE = path.join(ROOT, '자료', '2024-2026 구매내역 3.xlsx');
const DB_PATH = path.join(ROOT, 'data', 'db.json');
const PURCHASE_HISTORY_PATH = path.join(ROOT, 'data', 'purchase-history.json');
const BACKUP_DIR = path.join(ROOT, 'data', 'db-backups');

const WRITE = process.argv.includes('--write');

function splitNameSpec(raw) {
  const s = String(raw || '').trim();
  const m = s.match(/\[([^\[\]]*)\]\s*$/);
  if (!m) return { name: s, spec: '' };
  return { name: s.slice(0, m.index).trim(), spec: m[1].trim() };
}

function parseDateNo(raw) {
  // "24/01/02-1" -> { date: "2024-01-02", seq: 1 }
  const s = String(raw || '').trim();
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{2})-(\d+)$/);
  if (!m) return { date: '', seq: 0 };
  const [, yy, mm, dd, seq] = m;
  return { date: `20${yy}-${mm}-${dd}`, seq: parseInt(seq, 10) };
}

function main() {
  if (!fs.existsSync(SOURCE_FILE)) {
    console.error('원본 파일을 찾을 수 없습니다:', SOURCE_FILE);
    process.exit(1);
  }

  const wb = XLSX.readFile(SOURCE_FILE);
  const ws1 = wb.Sheets['Sheet1'];
  const rows = XLSX.utils.sheet_to_json(ws1, { defval: '', header: 1 });

  const vendorByName = new Map(); // name -> custCode
  const itemByCode = new Map(); // code -> { name, spec, barcode }
  const history = [];
  let bracketSplitCount = 0;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every((c) => c === '')) continue;
    const [dateNo, itemCode, , barcode, nameSpec, qty, unitPrice, supplyAmt, vat, total, vendorCode, vendorName, whCode, note, invoiceNo, memo, taxType] = r;

    const { date, seq } = parseDateNo(dateNo);
    const { name, spec } = splitNameSpec(nameSpec);
    if (spec) bracketSplitCount++;

    const vCode = String(vendorCode || '').trim();
    const vName = String(vendorName || '').trim();
    if (vName) {
      if (!vendorByName.has(vName)) vendorByName.set(vName, vCode);
      else if (!vendorByName.get(vName) && vCode) vendorByName.set(vName, vCode);
    }

    const iCode = String(itemCode || '').trim();
    const bcode = String(barcode || '').trim();
    if (iCode && !itemByCode.has(iCode)) {
      itemByCode.set(iCode, { name, spec, barcode: bcode });
    }

    history.push({
      id: `PH-${String(i).padStart(6, '0')}`,
      date,
      seq,
      itemCode: iCode,
      itemName: name,
      spec,
      barcode: bcode,
      qty: Number(qty) || 0,
      unitPrice: Number(unitPrice) || 0,
      supplyAmount: Number(supplyAmt) || 0,
      vat: Number(vat) || 0,
      total: Number(total) || 0,
      vendorCode: vCode,
      vendorName: vName,
      warehouseCode: String(whCode || '').trim(),
      note: String(note || '').trim(),
      invoiceNo: String(invoiceNo || '').trim(),
      memo: String(memo || '').trim(),
      taxType: String(taxType || '').trim(),
    });
  }

  const purchaseProductMaster = [...itemByCode.entries()].map(([code, v]) => ({
    code,
    name: v.name,
    spec: v.spec,
    barcode: v.barcode,
  }));

  // 창고 리스트 (2번 시트)
  const ws2 = wb.Sheets['창고리스트'];
  const whRows = XLSX.utils.sheet_to_json(ws2, { defval: '', header: 1 });
  const ecountWarehouses = whRows.slice(1)
    .filter((r) => r && r.some((c) => c !== ''))
    .map((r) => ({
      code: String(r[0] || '').trim(),
      name: String(r[1] || '').trim(),
      type: String(r[2] || '').trim(),
      processName: String(r[3] || '').trim(),
      outsourceVendor: String(r[4] || '').trim(),
      active: String(r[5] || '').trim().toLowerCase() === 'yes',
      businessUnit: String(r[6] || '').trim(),
    }));

  // ── DB 병합 ──
  const db = JSON.parse(fs.readFileSync(DB_PATH, 'utf-8'));
  if (!db.partners) db.partners = { inbound: [], outbound: [], purchase: [] };
  if (!Array.isArray(db.partners.purchase)) db.partners.purchase = [];

  let vendorsAdded = 0, vendorsUpdated = 0;
  for (const [name, custCode] of vendorByName.entries()) {
    const existing = db.partners.purchase.find((p) => p.name === name);
    if (existing) {
      if (custCode && existing.custCode !== custCode) { existing.custCode = custCode; vendorsUpdated++; }
    } else {
      db.partners.purchase.push({ name, custCode: custCode || '' });
      vendorsAdded++;
    }
  }

  db.purchaseProductMaster = purchaseProductMaster;
  db.ecountWarehouses = ecountWarehouses;

  // ── 요약 출력 ──
  console.log('=== 구매내역 임포트 요약 ===');
  console.log('원본 거래 행 수:', history.length.toLocaleString());
  console.log('구매처(이름 기준 고유):', vendorByName.size, `(신규 ${vendorsAdded} / 코드갱신 ${vendorsUpdated} / 기존유지 ${vendorByName.size - vendorsAdded - vendorsUpdated})`);
  console.log('구매품목 마스터(품목코드 기준 고유):', purchaseProductMaster.length);
  console.log('대괄호 규격 분리된 행 수:', bracketSplitCount, `(${(bracketSplitCount / history.length * 100).toFixed(1)}%)`);
  console.log('이카운트 창고 리스트:', ecountWarehouses.length);

  const historyJson = JSON.stringify(history);
  console.log('purchase-history.json 예상 크기: ' + (historyJson.length / 1024 / 1024).toFixed(2) + ' MB');

  console.log('\n샘플 구매품목(대괄호 분리됨) 5건:');
  purchaseProductMaster.filter((p) => p.spec).slice(0, 5).forEach((p) => console.log(' ', JSON.stringify(p)));

  if (!WRITE) {
    console.log('\n[미리보기 모드] 아무 파일도 변경하지 않았습니다. 실제로 반영하려면 --write 옵션을 붙여 다시 실행하세요.');
    return;
  }

  // 백업
  if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const tag = new Date().toISOString().replace(/[:.]/g, '-');
  fs.copyFileSync(DB_PATH, path.join(BACKUP_DIR, `db-before-purchase-import-${tag}.json`));

  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf-8');
  fs.writeFileSync(PURCHASE_HISTORY_PATH, historyJson, 'utf-8');

  console.log('\n[완료] db.json 및 data/purchase-history.json 에 반영했습니다.');
  console.log('백업:', path.join(BACKUP_DIR, `db-before-purchase-import-${tag}.json`));
}

main();
