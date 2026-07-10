// 일회성 스크립트: 그동안 쌓인 이마트 출고 업로드 데이터(productCode ↔ productName/sourceProductCode)를
// 기본상품정보의 deliveryVendorInfo(이마트 항목)에 매칭해서 채워 넣는다.
// 이미 값이 있는 항목은 덮어쓰지 않는다. 이마트 판매 이력이 있는데 deliveryVendors에 이마트가 없으면 추가한다.
//
// 사용법:
//   node scripts/backfill-emart-vendor-names.js          # dry-run (변경 내역만 출력)
//   node scripts/backfill-emart-vendor-names.js --apply   # 실제로 data/db.json에 반영

const fs = require("fs");
const path = require("path");

const DB_PATH = path.join(__dirname, "..", "data", "db.json");
const APPLY = process.argv.includes("--apply");

const db = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));

const rows = [
  ...(db.outboundOrderUpload?.uploadedRows || []),
  ...(db.outboundOrderUpload?.lines || [])
].filter((r) => r.partnerType === "emart" && r.productCode && r.productName);

function mostFrequent(map) {
  let best = "", bestCount = 0;
  for (const [k, v] of map) if (v > bestCount) { best = k; bestCount = v; }
  return best;
}

const grouped = new Map(); // productCode -> { names: Map, codes: Map }
for (const r of rows) {
  const key = String(r.productCode).trim();
  if (!grouped.has(key)) grouped.set(key, { names: new Map(), codes: new Map() });
  const g = grouped.get(key);
  const name = String(r.productName || "").trim();
  if (name) g.names.set(name, (g.names.get(name) || 0) + 1);
  const code = String(r.sourceProductCode || "").trim();
  if (code) g.codes.set(code, (g.codes.get(code) || 0) + 1);
}

const changes = [];
for (const p of db.products || []) {
  const key = grouped.has(p.code) ? p.code : (grouped.has(p.ecountCode) ? p.ecountCode : null);
  if (!key) continue;
  const g = grouped.get(key);
  const derivedName = mostFrequent(g.names);
  const derivedCode = mostFrequent(g.codes);
  if (!derivedName && !derivedCode) continue;

  if (!Array.isArray(p.deliveryVendorInfo)) p.deliveryVendorInfo = [];
  let entry = p.deliveryVendorInfo.find((v) => v.vendor === "이마트");
  const before = entry ? { ...entry } : null;
  if (!entry) {
    entry = { vendor: "이마트", code: "", itemName: "" };
    p.deliveryVendorInfo.push(entry);
  }
  let changed = false;
  if (!entry.itemName && derivedName) { entry.itemName = derivedName; changed = true; }
  if (!entry.code && derivedCode) { entry.code = derivedCode; changed = true; }
  if (!Array.isArray(p.deliveryVendors)) p.deliveryVendors = [];
  if (!p.deliveryVendors.includes("이마트")) { p.deliveryVendors.push("이마트"); changed = true; }

  if (changed) {
    changes.push({
      code: p.code,
      name: p.ecountName || p.name,
      before,
      after: { ...entry }
    });
  }
}

console.log(`총 상품 ${db.products.length}건 중 이마트 출고이력 매칭 상품 ${grouped.size}건, 실제 변경 대상 ${changes.length}건\n`);
for (const c of changes.slice(0, 20)) {
  console.log(`- ${c.code} (${c.name}): ${JSON.stringify(c.before)} -> ${JSON.stringify(c.after)}`);
}
if (changes.length > 20) console.log(`  ... 외 ${changes.length - 20}건`);

if (APPLY) {
  const backupPath = DB_PATH.replace(/\.json$/, `.backup-before-emart-backfill-${Date.now()}.json`);
  fs.copyFileSync(DB_PATH, backupPath);
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), "utf8");
  console.log(`\n적용 완료. 백업: ${backupPath}`);
} else {
  console.log("\n(dry-run) 실제 반영하려면 --apply 옵션으로 다시 실행하세요.");
}
