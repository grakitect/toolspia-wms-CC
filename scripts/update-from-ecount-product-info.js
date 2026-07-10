// 일회성 스크립트: "상품정보/이카운트 상품정보.xlsx" 파일(품목코드 기준)로
// 기본상품정보(products)의 품목명/구분(itemType)/규격(spec)을 갱신한다.
// BOM 컬럼은 파일에 값이 없어 대상에서 제외.
//
// 사용법:
//   node scripts/update-from-ecount-product-info.js          # dry-run
//   node scripts/update-from-ecount-product-info.js --apply  # 실제 반영

const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DB_PATH = path.join(__dirname, "..", "data", "db.json");
const XLSX_PATH = path.join(__dirname, "..", "상품정보", "이카운트 상품정보.xlsx");
const APPLY = process.argv.includes("--apply");

function normalizeLoose(s) {
  return String(s || "").replace(/ /g, " ").replace(/\s+/g, " ").trim();
}

const db = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));
const wb = XLSX.readFile(XLSX_PATH);
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

const byCode = new Map(); // normalized code -> row
for (let i = 1; i < rows.length; i++) {
  const r = rows[i];
  const code = String(r[1] ?? "").trim();
  if (!code) continue;
  byCode.set(code, r);
  byCode.set(normalizeLoose(code), r);
  byCode.set(code.replace(/\s+/g, ""), r);
}

const changes = [];
for (const p of db.products || []) {
  const candidates = [p.code, p.ecountCode, normalizeLoose(p.code), normalizeLoose(p.ecountCode)].filter(Boolean);
  let row = null;
  for (const c of candidates) {
    if (byCode.has(c)) { row = byCode.get(c); break; }
  }
  if (!row) continue;

  const newName = String(row[4] ?? "").trim();
  const newItemType = String(row[5] ?? "").trim();
  const newSpec = String(row[6] ?? "").trim();

  const before = { name: p.name, ecountName: p.ecountName, itemType: p.itemType, spec: p.spec };
  let changed = false;
  if (newName && (p.name !== newName || p.ecountName !== newName)) {
    p.name = newName;
    p.ecountName = newName;
    changed = true;
  }
  if (newItemType && p.itemType !== newItemType) {
    p.itemType = newItemType;
    changed = true;
  }
  if (newSpec && p.spec !== newSpec) {
    p.spec = newSpec;
    changed = true;
  }
  if (changed) {
    changes.push({ code: p.code, before, after: { name: p.name, ecountName: p.ecountName, itemType: p.itemType, spec: p.spec } });
  }
}

const matchedCount = (db.products || []).filter((p) => {
  const candidates = [p.code, p.ecountCode, normalizeLoose(p.code), normalizeLoose(p.ecountCode)].filter(Boolean);
  return candidates.some((c) => byCode.has(c));
}).length;

console.log(`전체 상품 ${db.products.length}건 중 이카운트 파일과 매칭된 상품 ${matchedCount}건, 실제 변경 대상 ${changes.length}건\n`);
for (const c of changes.slice(0, 25)) {
  console.log(`- ${c.code}: ${JSON.stringify(c.before)} -> ${JSON.stringify(c.after)}`);
}
if (changes.length > 25) console.log(`  ... 외 ${changes.length - 25}건`);

if (APPLY) {
  const backupPath = DB_PATH.replace(/\.json$/, `.backup-before-ecount-info-update-${Date.now()}.json`);
  fs.copyFileSync(DB_PATH, backupPath);
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), "utf8");
  console.log(`\n적용 완료. 백업: ${backupPath}`);
} else {
  console.log("\n(dry-run) 실제 반영하려면 --apply 옵션으로 다시 실행하세요.");
}
