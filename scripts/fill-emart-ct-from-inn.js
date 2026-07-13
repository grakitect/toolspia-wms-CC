// 일회성 스크립트: 이마트 판매 상품의 바코드(CT)를 바코드(INN)과 동일하게 채운다.
//
// 사용법:
//   node scripts/fill-emart-ct-from-inn.js          # dry-run (변경 내역만 출력)
//   node scripts/fill-emart-ct-from-inn.js --apply   # 실제로 data/db.json에 반영

const fs = require("fs");
const path = require("path");

const DB_PATH = path.join(__dirname, "..", "data", "db.json");
const APPLY = process.argv.includes("--apply");

const db = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));

const changes = [];
for (const p of db.products || []) {
  const isEmart = (p.deliveryVendorInfo || []).some((v) => v.vendor === "이마트");
  if (!isEmart) continue;
  const inn = String(p.middleBarcode || "").trim();
  if (!inn) continue;
  const before = p.logisticsBarcode || "";
  if (before !== inn) {
    p.logisticsBarcode = inn;
    changes.push({ code: p.code, name: p.ecountName || p.name, before, after: inn });
  }
}

console.log(`이마트 판매 상품 중 변경 대상 ${changes.length}건\n`);
for (const c of changes.slice(0, 30)) {
  console.log(`- ${c.code} (${c.name}): "${c.before}" -> "${c.after}"`);
}
if (changes.length > 30) console.log(`  ... 외 ${changes.length - 30}건`);

if (APPLY) {
  const backupPath = DB_PATH.replace(/\.json$/, `.backup-before-emart-ct-fill-${Date.now()}.json`);
  fs.copyFileSync(DB_PATH, backupPath);
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), "utf8");
  console.log(`\n적용 완료. 백업: ${backupPath}`);
} else {
  console.log("\n(dry-run) 실제 반영하려면 --apply 옵션으로 다시 실행하세요.");
}
