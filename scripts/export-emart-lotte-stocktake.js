const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const dbPath = path.join(__dirname, '..', 'data', 'db.json');
const db = JSON.parse(fs.readFileSync(dbPath, 'utf8'));

function totalStock(productCode) {
  let stock = 0;
  for (const m of db.movements) {
    if (m.cancelled) continue;
    if (m.productCode !== productCode) continue;
    if (m.type === 'LOC_TRANSFER' || m.type === 'TRANSFER') continue; // 창고 간 이동은 총 재고에 영향 없음
    stock += Number(m.qty || 0);
  }
  return stock;
}

const vendors = ['이마트', '롯데마트'];

function buildRows() {
  return db.products
    .filter((p) => (p.deliveryVendors || []).some((v) => vendors.includes(v)))
    .map((p) => ({
      '판매처': (p.deliveryVendors || []).filter((v) => vendors.includes(v)).join(','),
      '품목코드(이카운트)': p.ecountCode || p.code || '',
      '바코드(SKU)': p.barcode || '',
      '바코드(INN)': p.middleBarcode || '',
      '바코드(CT)': p.logisticsBarcode || '',
      '상품명': p.ecountName || p.name || '',
      '규격': p.spec || '',
      '카테고리': (p.categories && p.categories[0]) || p.category || '',
      '판매상태': p.status || '',
      '전산재고': totalStock(p.code),
      '실사수량': '',
      '차이': '',
    }))
    .sort((a, b) => a['상품명'].localeCompare(b['상품명'], 'ko'));
}

const wb = XLSX.utils.book_new();
const rows = buildRows();
const ws = XLSX.utils.json_to_sheet(rows);
ws['!cols'] = [
  { wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 },
  { wch: 28 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 },
];
XLSX.utils.book_append_sheet(wb, ws, '이마트_롯데마트');
console.log(`전체: ${rows.length}개`);

const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
const outDir = path.join(__dirname, '..', '상품정보');
let outPath = path.join(outDir, `이마트_롯데마트_재고조사리스트_${today}.xlsx`);
try {
  XLSX.writeFile(wb, outPath);
} catch (e) {
  if (e.code === 'EBUSY') {
    outPath = path.join(outDir, `이마트_롯데마트_재고조사리스트_${today}_2.xlsx`);
    XLSX.writeFile(wb, outPath);
  } else {
    throw e;
  }
}
console.log('저장됨:', outPath);
