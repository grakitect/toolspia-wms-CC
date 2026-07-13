// 일회성 스크립트: "이마트 상품정보 판매처상품명 정리-2026-07-13-1.xlsx" 파일 기준으로
// 기본상품정보의 deliveryVendorInfo(이마트 항목) 판매처 품목명을 업데이트한다.
// 상품코드(이카운트코드)로 매칭. 기존 값이 달라도 파일 값으로 덮어쓴다.
//
// 사용법:
//   node scripts/update-emart-vendor-names-20260713.js          # dry-run (변경 내역만 출력)
//   node scripts/update-emart-vendor-names-20260713.js --apply   # 실제로 data/db.json에 반영

const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const DB_PATH = path.join(__dirname, "..", "data", "db.json");
const XLSX_PATH = path.join(__dirname, "..", "상품정보", "이마트 상품정보 판매처상품명 정리-2026-07-13-1.xlsx");
const APPLY = process.argv.includes("--apply");

const db = JSON.parse(fs.readFileSync(DB_PATH, "utf8"));
const wb = xlsx.readFile(XLSX_PATH);
const sheet = wb.Sheets[wb.SheetNames[0]];
const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });

const updateRows = []; // { ecountCode, itemCode, name }
for (const r of rows) {
  const ecountCode = String(r["이카운트코드"] || "").trim();
  const itemCode = String(r["상품코드"] || "").trim();
  const name = String(r["업체등록 상품명"] || "").trim();
  if ((ecountCode || itemCode) && name) updateRows.push({ ecountCode, itemCode, name });
}

const changes = [];
const notFound = [];
for (const { ecountCode, itemCode, name: newName } of updateRows) {
  const code = ecountCode || itemCode;
  const p = db.products.find((p) => String(p.ecountCode || p.code) === ecountCode)
    || db.products.find((p) => String(p.ecountCode || p.code) === itemCode)
    || db.products.find((p) => String(p.barcode) === itemCode);
  if (!p) { notFound.push(code); continue; }

  if (!Array.isArray(p.deliveryVendorInfo)) p.deliveryVendorInfo = [];
  let entry = p.deliveryVendorInfo.find((v) => v.vendor === "이마트");
  const before = entry ? entry.itemName : undefined;
  if (!entry) {
    entry = { vendor: "이마트", code: "", itemName: "" };
    p.deliveryVendorInfo.push(entry);
    if (!Array.isArray(p.deliveryVendors)) p.deliveryVendors = [];
    if (!p.deliveryVendors.includes("이마트")) p.deliveryVendors.push("이마트");
  }
  if (before !== newName) {
    entry.itemName = newName;
    changes.push({ code, name: p.ecountName || p.name, before, after: newName });
  }
}

console.log(`업데이트 파일 ${updateRows.length}건 중 매칭 상품 ${updateRows.length - notFound.length}건, 실제 변경 ${changes.length}건, 미매칭 ${notFound.length}건\n`);
for (const c of changes.slice(0, 30)) {
  console.log(`- ${c.code} (${c.name}): "${c.before || ""}" -> "${c.after}"`);
}
if (changes.length > 30) console.log(`  ... 외 ${changes.length - 30}건`);
if (notFound.length) console.log(`\n미매칭 코드: ${notFound.join(", ")}`);

if (APPLY) {
  const backupPath = DB_PATH.replace(/\.json$/, `.backup-before-emart-vendorname-update-${Date.now()}.json`);
  fs.copyFileSync(DB_PATH, backupPath);
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), "utf8");
  console.log(`\n적용 완료. 백업: ${backupPath}`);
} else {
  console.log("\n(dry-run) 실제 반영하려면 --apply 옵션으로 다시 실행하세요.");
}
