// 구매처명에 특수문자 '㈜'(U+3286, 원문자)를 쓰는 곳이 있어, 같은 이카운트 거래처코드인데도
// '㈜OOO'와 '(주)OOO'로 표기가 갈려 구매처 정보 화면에서 같은 회사가 따로 나뉘어 보이는 문제를 정리한다.
//   1) data/purchase-history.json 의 모든 vendorName에서 '㈜' -> '(주)'로 치환(구매처 정보 화면은
//      이 파일을 vendorName 문자열 기준으로 그룹핑하므로, 여기를 통일해야 실제로 합쳐짐)
//   2) db.json 의 db.partners.purchase 도 같은 치환 + 이름 중복 제거(같은 이름으로 합쳐진 항목은
//      custCode가 있는 쪽을 우선 유지)
//
// 기본은 미리보기(dry-run). 실제 반영하려면: node scripts/merge-vendor-name-variants.js --write

const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const DB_PATH = path.join(ROOT, 'data', 'db.json');
const PURCHASE_HISTORY_PATH = path.join(ROOT, 'data', 'purchase-history.json');
const BACKUP_DIR = path.join(ROOT, 'data', 'db-backups');

const WRITE = process.argv.includes('--write');
const SPECIAL_CHAR = '㈜';
const REPLACEMENT = '(주)';

function normalizeName(name) {
  return String(name || '').split(SPECIAL_CHAR).join(REPLACEMENT);
}

function main() {
  const history = JSON.parse(fs.readFileSync(PURCHASE_HISTORY_PATH, 'utf-8'));
  const db = JSON.parse(fs.readFileSync(DB_PATH, 'utf-8'));

  let historyChanged = 0;
  const beforeAfter = new Map(); // before -> after (for summary)
  for (const r of history) {
    if (r.vendorName && r.vendorName.includes(SPECIAL_CHAR)) {
      const before = r.vendorName;
      const after = normalizeName(before);
      beforeAfter.set(before, after);
      r.vendorName = after;
      historyChanged++;
    }
  }

  let partnersChanged = 0;
  for (const p of db.partners.purchase) {
    if (p.name && p.name.includes(SPECIAL_CHAR)) {
      const before = p.name;
      p.name = normalizeName(before);
      beforeAfter.set(before, p.name);
      partnersChanged++;
    }
  }

  // 이름이 같아진 항목끼리 병합 (custCode 있는 쪽 우선)
  const merged = [];
  const byName = new Map();
  for (const p of db.partners.purchase) {
    const existing = byName.get(p.name);
    if (!existing) {
      byName.set(p.name, p);
      merged.push(p);
    } else if (!existing.custCode && p.custCode) {
      existing.custCode = p.custCode;
    }
  }

  console.log('=== 구매처명 특수문자(㈜) 정리 요약 ===');
  console.log('purchase-history.json 행 중 vendorName 치환:', historyChanged.toLocaleString());
  console.log('db.partners.purchase 중 이름 치환:', partnersChanged);
  console.log('치환 전후 이름 매핑:');
  for (const [before, after] of beforeAfter) console.log(`  "${before}" -> "${after}"`);
  console.log('구매처 마스터 병합 전 개수:', db.partners.purchase.length, '-> 병합 후:', merged.length);

  if (!WRITE) {
    console.log('\n[미리보기 모드] 아무 파일도 변경하지 않았습니다. 실제로 반영하려면 --write 옵션을 붙여 다시 실행하세요.');
    return;
  }

  if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const tag = new Date().toISOString().replace(/[:.]/g, '-');
  fs.copyFileSync(DB_PATH, path.join(BACKUP_DIR, `db-before-vendor-merge-${tag}.json`));
  fs.copyFileSync(PURCHASE_HISTORY_PATH, path.join(BACKUP_DIR, `purchase-history-before-vendor-merge-${tag}.json`));

  db.partners.purchase = merged;
  fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2), 'utf-8');
  fs.writeFileSync(PURCHASE_HISTORY_PATH, JSON.stringify(history), 'utf-8');

  console.log('\n[완료] db.json 및 data/purchase-history.json 에 반영했습니다.');
}

main();
