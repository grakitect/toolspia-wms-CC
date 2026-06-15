// eCvan collect 로직 — server.js 재시작 없이 이 파일만 수정하면 즉시 반영됨
// 함수 파라미터: session, browser, page, path, fs, __dirname, readDb, writeDb

const _sleep = (ms) => new Promise(r => setTimeout(r, ms));
const snapDir = path.join(__dirname, "ecvan-debug");
if (!fs.existsSync(snapDir)) fs.mkdirSync(snapDir);

// SSO 리다이렉트 완료 대기 (ssoindex.jsp → 대시보드)
session.progress = "SSO 리다이렉트 대기 중...";
{
  const allPages = await browser.pages().catch(() => []);
  for (const p of allPages) {
    try {
      await p.waitForFunction(
        () => !window.location.href.includes("ssoindex"),
        { timeout: 30000 }
      ).catch(() => {});
    } catch {}
  }
}
await _sleep(2000);

// 대시보드 탭 찾기: 모든 탭+프레임에서 body에 "납품" 포함 여부로 선택
let pg = session._newPage || page;
{
  const allPages = await browser.pages().catch(() => [pg]);
  for (const p of allPages) {
    try {
      const allFrames = [p, ...p.frames()];
      let found = false;
      for (const f of allFrames) {
        const bodyHasNapum = await f.evaluate(() =>
          (document.body?.innerText || document.body?.textContent || "").includes("납품")
        ).catch(() => false);
        if (bodyHasNapum) { pg = p; found = true; break; }
      }
      if (found) break;
    } catch {}
  }
}

const _snap = async (name) => pg.screenshot({ path: path.join(snapDir, `collect-${name}.png`), fullPage: true }).catch(() => {});
const _allFrames = () => [pg, ...pg.frames()];

// Nexacro 버튼용: 텍스트 찾아 실제 마우스 좌표로 클릭 (el.click() 대신)
const _mouseClickText = async (text, timeout = 5000) => {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    const rect = await pg.evaluate((t) => {
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
      let node;
      while ((node = walker.nextNode())) {
        if (node.textContent.trim() === t) {
          const el = node.parentElement;
          if (el) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, width: r.width, height: r.height };
          }
        }
      }
      return null;
    }, text).catch(() => null);
    if (rect) {
      await pg.mouse.move(rect.x + rect.width / 2, rect.y + rect.height / 2);
      await _sleep(200);
      await pg.mouse.click(rect.x + rect.width / 2, rect.y + rect.height / 2);
      return true;
    }
    await _sleep(400);
  }
  return false;
};

// TreeWalker 기반 텍스트 클릭: 정확히 일치하는 텍스트 노드의 부모 엘리먼트 클릭
const _findAndClick = async (text, timeout = 5000) => {
  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    for (const frame of _allFrames()) {
      try {
        const clicked = await frame.evaluate((t) => {
          if (!document.body) return false;
          const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
          let node;
          while ((node = walker.nextNode())) {
            if (node.textContent.trim() === t) {
              const el = node.parentElement;
              if (el) {
                el.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
                el.dispatchEvent(new MouseEvent("mouseenter", { bubbles: true }));
                el.click();
                return true;
              }
            }
          }
          return false;
        }, text);
        if (clicked) return true;
      } catch {}
    }
    await _sleep(400);
  }
  return false;
};

const _closePopups = async () => {
  await _sleep(800);
  // 팝업/레이어를 DOM에서 직접 제거 — 페이지 이동 없이 확실히 닫음
  // pg 탭에서만 실행 (다른 탭 건드리지 않음)
  for (let i = 0; i < 3; i++) {
    const removed = await pg.evaluate(() => {
      const sel = [
        '[class*="modal"]','[class*="popup"]','[class*="layer"]',
        '[class*="banner"]','[class*="dim"]','[class*="dialog"]',
        '[class*="alert"]','[class*="notice"]'
      ].join(",");
      const visible = [...document.querySelectorAll(sel)].filter(el => {
        const s = window.getComputedStyle(el);
        return s.display !== "none" && s.visibility !== "hidden"
          && el.offsetParent !== null && el.offsetWidth > 100 && el.offsetHeight > 100;
      });
      visible.forEach(el => el.remove());
      document.body.style.overflow = "";
      return visible.length;
    }).catch(() => 0);
    session.progress = `로그인 후 팝업 닫는 중... (${i+1}/3) DOM제거 ${removed}개`;
    if (removed === 0) break;
    await _sleep(500);
  }
};

session.status = "navigating"; session.progress = "팝업 닫는 중...";
await _closePopups();
await _sleep(800);
await _snap("01-after-login");

// 납품 버튼 위치 찾기 (ID 우선, 실패 시 텍스트 폴백, 최대 15초 대기)
session.progress = "납품 메뉴 위치 찾는 중...";
let napumRect = null;
for (let attempt = 0; attempt < 10; attempt++) {
  napumRect = await pg.evaluate(() => {
    // ID로 찾기
    const btn = document.getElementById("mainframe.VFrameSet.TopFrame.form.divTopBtn.form.btn_topMenu1");
    if (btn) { const r = btn.getBoundingClientRect(); if (r.width > 0) return { x: r.x, y: r.y, width: r.width, height: r.height }; }
    // 텍스트로 폴백: TopFrame 영역의 "납품" 텍스트
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
    let node;
    while ((node = walker.nextNode())) {
      if (node.textContent.trim() === "납품") {
        const el = node.parentElement;
        if (el) { const r = el.getBoundingClientRect(); if (r.width > 0 && r.y < 100) return { x: r.x, y: r.y, width: r.width, height: r.height }; }
      }
    }
    return null;
  }).catch(() => null);
  if (napumRect) break;
  session.progress = `납품 메뉴 대기 중... (${attempt + 1}/10)`;
  await _sleep(1500);
}

if (!napumRect) {
  await _snap("err-no-napum");
  throw new Error("납품 버튼을 찾지 못했습니다.");
}

session.progress = `납품 hover 후 click 중... (x:${Math.round(napumRect.x)}, y:${Math.round(napumRect.y)})`;
await pg.mouse.move(napumRect.x + napumRect.width / 2, napumRect.y + napumRect.height / 2);
await _sleep(300);
await pg.mouse.click(napumRect.x + napumRect.width / 2, napumRect.y + napumRect.height / 2);
await _sleep(2000);
await _snap("02-napum-click");

// 사이드바 "미납사유" — bounding rect 기반 mouse.click
session.progress = "미납사유 위치 찾는 중...";
const minapRect = await pg.evaluate(() => {
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
  let node;
  while ((node = walker.nextNode())) {
    if (node.textContent.trim() === "미납사유") {
      const el = node.parentElement;
      if (el) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
    }
  }
  return null;
}).catch(() => null);

if (!minapRect) {
  await _snap("err-no-minapRect");
  throw new Error("미납사유 위치를 찾지 못했습니다.");
}

session.progress = `미납사유 hover+click 중... (x:${Math.round(minapRect.x)}, y:${Math.round(minapRect.y)})`;
await pg.mouse.move(minapRect.x + minapRect.width / 2, minapRect.y + minapRect.height / 2);
await _sleep(400);
await pg.mouse.click(minapRect.x + minapRect.width / 2, minapRect.y + minapRect.height / 2);
await pg.waitForNavigation({ waitUntil: "networkidle2", timeout: 8000 }).catch(() => {});
await _sleep(2000);
await _snap("02b-after-minap-click");

// "미납사유등록" — bounding rect 기반 mouse.click (Nexacro 이벤트 처리)
session.progress = "미납사유등록 위치 찾는 중...";
const minapRegRect = await pg.evaluate(() => {
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
  let node;
  while ((node = walker.nextNode())) {
    if (node.textContent.trim() === "미납사유등록") {
      const el = node.parentElement;
      if (el) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, width: r.width, height: r.height };
      }
    }
  }
  return null;
}).catch(() => null);

if (!minapRegRect) {
  await _snap("err-no-minapRegRect");
  const bodyTexts = [];
  for (const frame of _allFrames()) {
    try {
      const t = await frame.evaluate(() => (document.body?.innerText || "").slice(0, 400)).catch(() => "");
      if (t) bodyTexts.push(t.replace(/\s+/g, " ").trim());
    } catch {}
  }
  throw new Error(`미납사유등록 위치를 찾지 못했습니다. body: ${bodyTexts.join(" | ")}`);
}

session.progress = `미납사유등록 click 중... (x:${Math.round(minapRegRect.x)}, y:${Math.round(minapRegRect.y)})`;
await pg.mouse.move(minapRegRect.x + minapRegRect.width / 2, minapRegRect.y + minapRegRect.height / 2);
await _sleep(300);
await pg.mouse.click(minapRegRect.x + minapRegRect.width / 2, minapRegRect.y + minapRegRect.height / 2);
await pg.waitForNavigation({ waitUntil: "networkidle2", timeout: 10000 }).catch(() => {});
// Nexacro가 div_work 안에 폼을 렌더링할 때까지 대기 (고정 sleep 대신 동적 폴링, 최대 25초)
session.progress = "미납사유등록 페이지 렌더링 대기 중...";
await pg.waitForFunction(() => {
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
  let node;
  while ((node = walker.nextNode())) {
    const t = node.textContent.trim();
    if (t === "상품별" || t === "조회") return true;
  }
  return false;
}, { timeout: 25000, polling: 500 }).catch(() => {});
await _sleep(500);
await _snap("03-minapsa-reg");

// 상품별 탭 + 조회 (Nexacro 버튼 — 실제 마우스 클릭 필요)
session.progress = "상품별 탭 클릭 중...";
const clickedSangpum = await _mouseClickText("상품별", 12000);
if (!clickedSangpum) { await _snap("err-no-sangpum"); throw new Error("상품별 탭을 찾지 못했습니다"); }
await _sleep(800);
await _snap("04-sangpum-clicked");

session.progress = "조회 중...";
const clickedJohoe = await _mouseClickText("조회", 8000);
if (!clickedJohoe) { await _snap("err-no-johoe"); throw new Error("조회 버튼을 찾지 못했습니다"); }
await _sleep(3000);
await _snap("04-result");

// 조회 결과가 나온 후 보기 행 수 500으로 설정
// Nexacro ID: ...form_NONDELIVERYRSNLIST.form.div_work.form.div_page.form.cboCnt
session.progress = "보기 행 수 500으로 설정 중...";
const cboCntRect = await pg.evaluate(() => {
  const el = document.getElementById(
    "mainframe.VFrameSet.HFrameSet.VFrameSet.WorkFrameSet.form_NONDELIVERYRSNLIST.form.div_work.form.div_page.form.cboCnt"
  );
  if (!el) return null;
  const r = el.getBoundingClientRect();
  return r.width > 0 ? { x: r.x, y: r.y, w: r.width, h: r.height } : null;
}).catch(() => null);

if (cboCntRect) {
  // cboCnt 클릭해서 드롭다운 열기
  await pg.mouse.move(cboCntRect.x + cboCntRect.w / 2, cboCntRect.y + cboCntRect.h / 2);
  await _sleep(200);
  await pg.mouse.click(cboCntRect.x + cboCntRect.w / 2, cboCntRect.y + cboCntRect.h / 2);
  await _sleep(600);
  // 드롭다운에서 500 또는 최대 페이지 크기 선택 (일반 페이지 크기: 10~1000)
  const VALID_SIZES = new Set([10, 20, 30, 50, 100, 200, 300, 500, 1000]);
  const opt500 = await pg.evaluate((validSizes) => {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
    let node;
    let maxOpt = null;
    let maxVal = 0;
    while ((node = walker.nextNode())) {
      const t = node.textContent.trim();
      const n = parseInt(t);
      // 유효한 페이지 크기만 (자릿수 4 이하, VALID_SIZES 목록)
      if (!isNaN(n) && validSizes.includes(n) && n > maxVal) {
        const el = node.parentElement;
        if (el) {
          const r = el.getBoundingClientRect();
          if (r.width > 0 && r.height > 0) { maxVal = n; maxOpt = { x: r.x, y: r.y, w: r.width, h: r.height, val: n }; }
        }
      }
    }
    return maxOpt;
  }, [...VALID_SIZES]).catch(() => null);
  if (opt500) {
    session.progress = `보기 ${opt500.val}행 선택 중...`;
    await pg.mouse.move(opt500.x + opt500.w / 2, opt500.y + opt500.h / 2);
    await _sleep(200);
    await pg.mouse.click(opt500.x + opt500.w / 2, opt500.y + opt500.h / 2);
    await _sleep(3000); // 재조회 대기
    await _snap("04b-pagesize500");
  } else {
    session.progress = "500 옵션을 찾지 못했습니다 (현재 페이지 데이터로 수집)";
  }
} else {
  session.progress = "cboCnt 콤보박스를 찾지 못했습니다 (현재 페이지 데이터로 수집)";
}

// 조회 후 HTML 덤프 (그리드 구조 파악용)
try {
  const workHtml = await pg.evaluate(() => document.documentElement.outerHTML);
  fs.writeFileSync(path.join(snapDir, "work-page.html"), workHtml, "utf8");
} catch {}

session.status = "collecting"; session.progress = "Nexacro 그리드 데이터 수집 중...";

const now = new Date().toISOString();

// 방법 1: 정확한 경로로 Nexacro Dataset API 접근
// 폼: form_NONDELIVERYRSNLIST > form > div_work > form > grd_main
const nexacroRows = await pg.evaluate(() => {
  try {
    const app = window.nexacro?.getApplication?.();
    if (!app) return null;

    // 정확한 폼 경로
    const formRoot = app.mainframe
      ?.VFrameSet?.HFrameSet?.VFrameSet?.WorkFrameSet
      ?.form_NONDELIVERYRSNLIST?.form?.div_work?.form;
    if (!formRoot) return { error: "form_NONDELIVERYRSNLIST 폼 없음" };

    // formRoot 안의 모든 Dataset(rowcount + getColumn 있는 것) 탐색
    const readDs = (ds, key) => {
      if (!ds || typeof ds.rowcount !== 'number' || typeof ds.getColumn !== 'function') return null;
      const rows = [];
      for (let i = 0; i < ds.rowcount; i++) {
        const row = {};
        for (let j = 0; j < (ds.colcount || 0); j++) {
          try { const col = ds.getColID ? ds.getColID(j) : `col${j}`; row[col] = String(ds.getColumn(i, col) ?? ""); } catch {}
        }
        rows.push(row);
      }
      return { key, rowcount: ds.rowcount, rows };
    };

    const found = [];
    for (const k of Object.keys(formRoot)) {
      try {
        const r = readDs(formRoot[k], k);
        if (r && r.rowcount > 0) found.push(r);
      } catch {}
    }

    // grd_main의 dataset 속성으로도 시도
    const grid = formRoot.grd_main;
    if (grid) {
      const ds = grid.getDataset?.() ?? grid.dataset ?? null;
      if (ds) {
        const r = readDs(ds, "grd_main_dataset");
        if (r && r.rowcount > 0) found.push(r);
      }
    }

    if (found.length === 0) return { error: "Dataset 없음", formKeys: Object.keys(formRoot).slice(0, 30) };
    // ds_main 우선 (실제 미납 상품 목록), 없으면 STR_SKU_CD 컬럼 있는 것, 없으면 rowcount 최대
    const dsMain = found.find(f => f.key === "ds_main");
    const dsWithSku = found.find(f => f.rows[0] && "STR_SKU_CD" in f.rows[0]);
    const best = dsMain || dsWithSku || found.sort((a, b) => b.rowcount - a.rowcount)[0];
    return { method: "nexacro_exact", best, allFound: found.map(f => ({ key: f.key, rowcount: f.rowcount })) };
  } catch (e) {
    return { error: e.message };
  }
}).catch(e => ({ error: e.message }));

session.progress = `Nexacro API 결과: ${JSON.stringify(nexacroRows?.allFound || nexacroRows?.error || "null")}`;
await _snap("05-nexacro-result");

let allRows = [];

if (nexacroRows?.best?.rows?.length > 0) {
  allRows = nexacroRows.best.rows.map(r => ({ ...r, uploadedAt: now }));
  session.progress = `Nexacro Dataset "${nexacroRows.best.key}" ${allRows.length}행 수집됨`;
} else {
  // 방법 2: grd_main DOM cell ID로 직접 읽기 (보기를 500으로 설정 필요)
  session.progress = "DOM 방식으로 그리드 수집 중 (보기 500 설정 후)...";

  // 보기 500 설정 (조회 결과 나온 뒤 페이지 하단 콤보박스)
  const setPageSize = async (size) => {
    // 현재 행 수 표시 콤보박스 찾기
    const comboRect = await pg.evaluate((s) => {
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
      let node;
      while ((node = walker.nextNode())) {
        const t = node.textContent.trim();
        if (/^\d+행?$/.test(t) && parseInt(t) < s) {
          const el = node.parentElement;
          if (el) { const r = el.getBoundingClientRect(); if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, w: r.width, h: r.height }; }
        }
      }
      return null;
    }, size).catch(() => null);
    if (!comboRect) return false;
    await pg.mouse.move(comboRect.x + comboRect.w / 2, comboRect.y + comboRect.h / 2);
    await _sleep(200);
    await pg.mouse.click(comboRect.x + comboRect.w / 2, comboRect.y + comboRect.h / 2);
    await _sleep(500);
    // 500 옵션 클릭
    const opt = await pg.evaluate((s) => {
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
      let node;
      while ((node = walker.nextNode())) {
        const t = node.textContent.trim();
        if (t === String(s) || t === `${s}행`) {
          const el = node.parentElement;
          if (el) { const r = el.getBoundingClientRect(); if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, w: r.width, h: r.height }; }
        }
      }
      return null;
    }, size).catch(() => null);
    if (!opt) return false;
    await pg.mouse.move(opt.x + opt.w / 2, opt.y + opt.h / 2);
    await _sleep(200);
    await pg.mouse.click(opt.x + opt.w / 2, opt.y + opt.h / 2);
    await _sleep(1500);
    return true;
  };

  const sized = await setPageSize(500);
  session.progress = `보기 500 설정: ${sized ? "성공" : "실패"}`;
  if (sized) await _snap("05b-pagesize500");

  // grd_main 안의 모든 body 셀 읽기 (cell ID 패턴 활용)
  const domRows = await pg.evaluate(() => {
    const PREFIX = "mainframe.VFrameSet.HFrameSet.VFrameSet.WorkFrameSet.form_NONDELIVERYRSNLIST.form.div_work.form.grd_main";
    // 헤더 텍스트 수집
    const headers = [];
    for (let ci = 0; ci <= 20; ci++) {
      const el = document.getElementById(`${PREFIX}.head.gridrow_-1.cell_-1_${ci}`);
      if (!el) break;
      const t = el.querySelector('.nexacontentsbox') || el;
      headers.push(t.textContent.trim());
    }
    // 데이터 행 수집
    const rows = [];
    for (let ri = 0; ri <= 1000; ri++) {
      const rowEl = document.getElementById(`${PREFIX}.body.gridrow_${ri}`);
      if (!rowEl) break;
      const row = {};
      for (let ci = 0; ci <= headers.length + 2; ci++) {
        const cellEl = document.getElementById(`${PREFIX}.body.gridrow_${ri}.cell_${ri}_${ci}`);
        if (!cellEl) break;
        const t = cellEl.querySelector('.nexacontentsbox') || cellEl;
        row[headers[ci] || `col${ci}`] = t.textContent.trim();
      }
      if (Object.values(row).some(v => v !== "")) rows.push(row);
    }
    return { headers, rows };
  }).catch(() => ({ headers: [], rows: [] }));

  if (domRows.rows.length > 0) {
    session.progress = `DOM 방식 ${domRows.rows.length}행 수집됨`;
    allRows = domRows.rows.map(r => ({ ...r, uploadedAt: now }));
  } else {
    session.progress = "데이터 추출 실패 — work-page.html 및 스냅샷 확인 필요";
    allRows = [];
  }
}

await _snap("05-collected");

if (allRows.length === 0) {
  session.status = "done"; session.progress = "미납 상품이 없습니다."; session.data = [];
} else {
  // ── Phase 2: 상품코드 클릭 → 점포별 세부 수집 ───────────────────────
  session.progress = `요약 ${allRows.length}건 확인. 세부 내역 수집 시작...`;

  const GRID_PREFIX = "mainframe.VFrameSet.HFrameSet.VFrameSet.WorkFrameSet.form_NONDELIVERYRSNLIST.form.div_work.form.grd_main";

  // Nexacro 폼에서 데이터셋 전체 읽기
  const _readNexacroDatasets = async (formFinder) => {
    return await pg.evaluate((finder) => {
      try {
        const app = window.nexacro?.getApplication?.();
        if (!app) return null;
        const vfs = app.mainframe?.VFrameSet?.HFrameSet?.VFrameSet?.WorkFrameSet;
        if (!vfs) return null;
        const formKeys = Object.keys(vfs);
        const key = formKeys.find(k => finder.some(f => k.includes(f)));
        const form = key ? vfs[key]?.form?.div_work?.form : null;
        if (!form) return { formKeys };
        const results = [];
        for (const k of Object.keys(form)) {
          try {
            const ds = form[k];
            if (!ds || typeof ds.rowcount !== "number" || typeof ds.getColumn !== "function" || ds.rowcount === 0) continue;
            const rows = [];
            for (let i = 0; i < ds.rowcount; i++) {
              const row = {};
              for (let j = 0; j < (ds.colcount || 0); j++) {
                const col = ds.getColID ? ds.getColID(j) : `col${j}`;
                row[col] = String(ds.getColumn(i, col) ?? "");
              }
              rows.push(row);
            }
            results.push({ key: k, rows });
          } catch {}
        }
        return { formKey: key, datasets: results };
      } catch (e) { return { error: e.message }; }
    }, formFinder).catch(() => null);
  };

  // 점포별 세부 데이터: barcode → [{storeName, storeCode, storeInDate, minapQty, orderQty}]
  const detailByBarcode = {};

  for (let i = 0; i < allRows.length; i++) {
    const barcode = String(allRows[i].STR_SKU_CD || allRows[i].productCode || "").trim();
    session.progress = `세부 수집 (${i + 1}/${allRows.length}) 바코드: ${barcode}`;

    // 상품코드 셀 클릭 (바코드 값이 있는 컬럼을 찾아 클릭)
    const clicked = await (async () => {
      // DOM 셀 ID로 찾기 — 상품코드 컬럼(바코드 텍스트 포함)
      const cellRect = await pg.evaluate((prefix, ri, bcode) => {
        for (let ci = 0; ci <= 5; ci++) {
          const el = document.getElementById(`${prefix}.body.gridrow_${ri}.cell_${ri}_${ci}`);
          if (!el) continue;
          const t = (el.querySelector(".nexacontentsbox") || el).textContent.trim();
          if (t === bcode || /^\d{10,15}$/.test(t)) {
            const r = el.getBoundingClientRect();
            if (r.width > 0 && r.height > 0) return { x: r.x, y: r.y, w: r.width, h: r.height };
          }
        }
        return null;
      }, GRID_PREFIX, i, barcode).catch(() => null);

      if (cellRect) {
        await pg.mouse.move(cellRect.x + cellRect.w / 2, cellRect.y + cellRect.h / 2);
        await _sleep(200);
        await pg.mouse.click(cellRect.x + cellRect.w / 2, cellRect.y + cellRect.h / 2);
        return true;
      }
      // 폴백: 바코드 텍스트로 직접 클릭
      return await _mouseClickText(barcode, 3000);
    })();

    if (!clicked) {
      session.progress = `행 ${i} 클릭 실패 (${barcode}), 건너뜀`;
      await _snap(`detail-err-${i}`);
      continue;
    }

    await _sleep(2000);
    await _snap(`detail-${String(i).padStart(3, "0")}-opened`);

    // 세부 페이지 Nexacro 데이터셋 읽기 (DTL/DETAIL 포함 폼 이름 탐색)
    const detailResult = await _readNexacroDatasets(["DTL", "DETAIL", "detail", "Dtl"]);

    let storeRows = [];

    if (detailResult?.datasets?.length) {
      // 점포명/점포코드 컬럼이 있는 데이터셋 선택
      const storeDs = detailResult.datasets.find(d =>
        d.rows[0] && ("STR_NM" in d.rows[0] || "STR_CD" in d.rows[0] || "STORE_NM" in d.rows[0])
      ) || detailResult.datasets.sort((a, b) => b.rows.length - a.rows.length)[0];

      if (storeDs) {
        storeRows = storeDs.rows.map(r => ({
          storeName: String(r.STR_NM || r.STORE_NM || r.점포명 || "").trim(),
          storeCode: String(r.STR_CD || r.STORE_CD || r.점포코드 || "").trim(),
          storeInDate: String(r.ARIV_DT || r.ARRIVE_DT || r.입점일 || "").trim().replace(/-/g, ""),
          orderDate:   String(r.ORD_DT || r.ORDER_DT || r.주문일 || "").trim().replace(/-/g, ""),
          minapQty: parseInt(String(r.RET_QTY || r.MINAP_QTY || r.미납합계 || "0").replace(/[^0-9]/g, "")) || 0,
          orderQty: parseInt(String(r.ORD_QTY || r.ORDER_QTY || r.주문합계 || "0").replace(/[^0-9]/g, "")) || 0,
        })).filter(r => r.minapQty > 0);
      }
    }

    // 폴백: DOM 테이블 스크래핑
    if (!storeRows.length) {
      const domRows = await pg.evaluate(() => {
        for (const tbl of document.querySelectorAll("table")) {
          const ths = [...tbl.querySelectorAll("thead th, thead td")].map(th => th.textContent.trim());
          if (!ths.some(h => h.includes("점포") || h.includes("STR"))) continue;
          const rows = [];
          tbl.querySelectorAll("tbody tr").forEach(tr => {
            const cells = [...tr.querySelectorAll("td")].map(td => td.textContent.trim());
            if (!cells.length) return;
            const row = {};
            ths.forEach((h, ci) => { row[h] = cells[ci] || ""; });
            rows.push(row);
          });
          if (rows.length) return rows;
        }
        return null;
      }).catch(() => null);

      if (domRows?.length) {
        storeRows = domRows.map(r => ({
          storeName: r["점포명"] || r["STR_NM"] || "",
          storeCode: r["점포코드"] || r["STR_CD"] || "",
          storeInDate: (r["입점일"] || "").replace(/-/g, ""),
          orderDate:   (r["주문일"] || "").replace(/-/g, ""),
          minapQty: parseInt(String(r["미납합계"] || r["미납수량"] || r["RET_QTY"] || "0").replace(/[^0-9]/g, "")) || 0,
          orderQty: parseInt(String(r["주문합계"] || r["주문수량"] || r["ORD_QTY"] || "0").replace(/[^0-9]/g, "")) || 0,
        })).filter(r => r.minapQty > 0);
      }
    }

    if (storeRows.length) {
      detailByBarcode[barcode] = storeRows;
      session.progress = `행 ${i} (${barcode}): 점포 ${storeRows.length}건 수집`;
    } else {
      session.progress = `행 ${i} (${barcode}): 세부 데이터 없음`;
    }

    // 목록으로 복귀
    const wentBack = await _mouseClickText("목록", 3000);
    if (!wentBack) await pg.goBack().catch(() => {});
    await _sleep(1500);
    await _snap(`detail-${String(i).padStart(3, "0")}-back`);
  }

  // ── Phase 3: 점포 → 센터 매핑 + 집계 ────────────────────────────────
  const db2 = readDb();
  const products = db2.products || [];
  const orderLines = (db2.outboundOrderUpload?.lines || []).filter(l => l.partnerType === "emart");
  const nowStr = new Date().toISOString();

  // (barcode, storeName/storeCode) → orderLine 조회 캐시
  const lineCache = new Map();
  for (const l of orderLines) {
    if (l.sourceProductCode && l.storeName)
      lineCache.set(`${l.sourceProductCode}|||${l.storeName}`, l);
    if (l.sourceProductCode && l.storeCode)
      lineCache.set(`${l.sourceProductCode}|||${l.storeCode}`, l);
  }

  const officialRows = [];

  if (Object.keys(detailByBarcode).length > 0) {
    // 세부 수집 성공: 점포별 → 센터 매핑
    for (const [barcode, storeRows] of Object.entries(detailByBarcode)) {
      const prod = products.find(p => (p.barcode && p.barcode === barcode) || p.code === barcode);

      for (const sr of storeRows) {
        const line = lineCache.get(`${barcode}|||${sr.storeName}`)
                  || lineCache.get(`${barcode}|||${sr.storeCode}`);

        officialRows.push({
          productCode: line?.productCode || prod?.code || barcode,
          productName: line?.productName || prod?.name || "",
          barcode,
          storeName:   sr.storeName,
          storeCode:   sr.storeCode,
          centerName:  line?.centerName || "",
          centerCode:  line?.centerCode || "",
          dueDate:     sr.storeInDate || "",
          orderDate:   sr.orderDate || "",
          officialQty: sr.minapQty,
          orderQty:    sr.orderQty,
          uploadedAt:  nowStr,
        });
      }
    }
    session.progress = `세부 수집 완료: 상품 ${Object.keys(detailByBarcode).length}건 / 점포 ${officialRows.length}건`;
  } else {
    // 세부 수집 실패: 요약 데이터로 폴백
    session.progress = "세부 수집 실패 — 요약 데이터로 대체 저장";
    for (const r of allRows) {
      const barcode = String(r.STR_SKU_CD || r.productCode || "").trim();
      const prod = products.find(p => (p.barcode && p.barcode === barcode) || p.code === barcode);
      const qty = parseInt(String(r.RET_QTY || "0").replace(/[^0-9]/g, "")) || 0;
      if (!barcode || !qty) continue;
      officialRows.push({
        productCode: prod?.code || barcode,
        productName: prod?.name || String(r.STR_SKU_NM || "").trim(),
        barcode,
        storeName: "", storeCode: "", centerName: "",
        dueDate: "", orderDate: "",
        officialQty: qty,
        orderQty: parseInt(String(r.ORD_QTY || "0").replace(/[^0-9]/g, "")) || 0,
        uploadedAt: nowStr,
      });
    }
  }

  session.data = officialRows;
  db2.unshipCompare = db2.unshipCompare || {};
  db2.unshipCompare.emart = db2.unshipCompare.emart || {};
  db2.unshipCompare.emart.officialRows = officialRows;
  db2.unshipCompare.emart.uploadedAt = nowStr;
  writeDb(db2);
  session.status = "done";
  session.progress = `수집 완료: ${officialRows.length}건`;
}
