const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

function loadEnvFile() {
  const envPath = path.join(__dirname, ".env");
  if (!fs.existsSync(envPath)) return;
  const raw = fs.readFileSync(envPath, "utf-8");
  for (const line of raw.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq <= 0) continue;
    const key = trimmed.slice(0, eq).trim();
    let value = trimmed.slice(eq + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    if (key && process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}

loadEnvFile();

const XLSX = require("xlsx");

const PORT = process.env.PORT || 4000;
const HOST = process.env.HOST || "0.0.0.0";

// eCvan 스크래퍼 세션 저장소
const ecvanSessions = new Map();
// 30분 초과 세션 자동 정리
setInterval(() => {
  const cutoff = Date.now() - 30 * 60 * 1000;
  for (const [sid, s] of ecvanSessions) {
    if (new Date(s.startedAt).getTime() < cutoff) {
      s.browser?.close().catch(() => {});
      ecvanSessions.delete(sid);
    }
  }
}, 5 * 60 * 1000).unref();
const PRODUCT_OPTION_KEYS = [
  "status",
  "deliveryVendors",
  "orderDept",
  "orderManagers",
  "supplyType",
  "warehouseGroup",
  "itemType",
  "categories"
];
const PUBLIC_DIR = path.join(__dirname, "public");
const DATA_DIR = path.join(__dirname, "data");
const DB_PATH = path.join(DATA_DIR, "db.json");

function ensureDb() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DB_PATH)) {
    const initial = {
      products: [],
      movements: [],
      warehouses: ["유통사업부", "다이소", "아세로직스", "가온플러스"],
      partners: {
        inbound: [],
        outbound: [],
        purchase: []
      },
      managers: ["admin"],
      seq: 1,
      inboundSlipConfirmLog: [],
      inboundPlanUpload: { lines: [], sourceFileName: "", uploadedAt: "", slipWorkflow: {} },
      outboundOrderUpload: {
        lines: [],
        uploadedRows: [],
        batches: [],
        appliedBatchByPartner: {},
        uploadedAt: "",
        slipWorkflow: {},
        confirmedLists: {}
      },
      outboundCodeMasters: []
    };
    fs.writeFileSync(DB_PATH, JSON.stringify(initial, null, 2), "utf-8");
  }
}

function sleepBusyMs(ms) {
  const end = Date.now() + Math.max(0, ms);
  while (Date.now() < end) {
    /* spin: avoid deps; short waits only */
  }
}

function isFsLockLikeError(err) {
  if (!err) return false;
  const code = String(err.code || "");
  if (["EBUSY", "EPERM", "EACCES", "ETXTBSY", "UNKNOWN"].includes(code)) return true;
  const msg = String(err.message || "");
  return /EBUSY|EPERM|EACCES|resource busy|another process|사용 중|잠금|locked/i.test(msg);
}

function readDb() {
  ensureDb();
  let lastErr;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      const raw = fs.readFileSync(DB_PATH, "utf-8");
      return JSON.parse(raw);
    } catch (e) {
      lastErr = e;
      if (e instanceof SyntaxError) throw e;
      if (attempt < 4 && isFsLockLikeError(e)) {
        sleepBusyMs(35 + attempt * 15);
        continue;
      }
      throw e;
    }
  }
  throw lastErr;
}

function writeDb(db) {
  const payload = JSON.stringify(db, null, 2);
  let lastErr;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      fs.writeFileSync(DB_PATH, payload, "utf-8");
      return;
    } catch (e) {
      lastErr = e;
      if (attempt < 4 && isFsLockLikeError(e)) {
        sleepBusyMs(35 + attempt * 15);
        continue;
      }
      throw e;
    }
  }
  throw lastErr;
}

function normalizeDb(db) {
  if (!db.partners) {
    db.partners = { inbound: [], outbound: [], purchase: [] };
  }
  if (!Array.isArray(db.partners.inbound)) db.partners.inbound = [];
  if (!Array.isArray(db.partners.outbound)) db.partners.outbound = [];
  if (!Array.isArray(db.partners.purchase)) db.partners.purchase = [];
  if (!Array.isArray(db.managers) || db.managers.length === 0) {
    db.managers = ["admin"];
  }
  if (!Array.isArray(db.warehouses) || db.warehouses.length === 0) {
    db.warehouses = ["유통사업부", "다이소", "아세로직스", "가온플러스"];
  }
  if (!db.productOptions || typeof db.productOptions !== "object") db.productOptions = {};
  for (const k of PRODUCT_OPTION_KEYS) {
    db.productOptions[k] = normalizeOptionValues(db.productOptions[k]);
  }
  for (const p of db.products || []) {
    p.deliveryVendors = toTagList(p.deliveryVendors || p.salesVendor);
    p.orderManagers = toTagList(p.orderManagers);
    p.categories = toTagList(p.categories || p.category);
    p.usedWarehouses = toTagList(p.usedWarehouses).filter((w) => db.warehouses.includes(w));
    if (!p.ecountCode) p.ecountCode = String(p.code || "");
    if (!p.ecountName) p.ecountName = String(p.name || "");
    p.middleBarcode = String(p.middleBarcode || "");
    p.status = firstToken(p.status, "판매중");
    p.supplyType = firstToken(p.supplyType, "");
    p.orderDept = firstToken(p.orderDept, "");
    p.warehouseGroup = firstToken(p.warehouseGroup, "");
    p.itemType = firstToken(p.itemType, "");
    db.productOptions.status.push(String(p.status || "").trim());
    db.productOptions.deliveryVendors.push(...toTagList(p.deliveryVendors));
    db.productOptions.orderDept.push(String(p.orderDept || "").trim());
    db.productOptions.orderManagers.push(...toTagList(p.orderManagers));
    db.productOptions.supplyType.push(String(p.supplyType || "").trim());
    db.productOptions.warehouseGroup.push(String(p.warehouseGroup || "").trim());
    db.productOptions.itemType.push(String(p.itemType || "").trim());
    db.productOptions.categories.push(...toTagList(p.categories));
  }
  for (const k of PRODUCT_OPTION_KEYS) {
    db.productOptions[k] = normalizeOptionValues(db.productOptions[k]);
  }
  for (const m of db.movements) {
    if (!m.warehouse) m.warehouse = "유통사업부";
  }
  if (!db.inboundPlanUpload || typeof db.inboundPlanUpload !== "object") {
    db.inboundPlanUpload = { lines: [], sourceFileName: "", uploadedAt: "", slipWorkflow: {} };
  }
  if (!Array.isArray(db.inboundPlanUpload.lines)) db.inboundPlanUpload.lines = [];
  if (!db.inboundPlanUpload.slipWorkflow || typeof db.inboundPlanUpload.slipWorkflow !== "object") {
    db.inboundPlanUpload.slipWorkflow = {};
  }
  if (!Array.isArray(db.inboundSlipConfirmLog)) db.inboundSlipConfirmLog = [];
  if (!db.outboundOrderUpload || typeof db.outboundOrderUpload !== "object") {
    db.outboundOrderUpload = {
      lines: [],
      uploadedRows: [],
      batches: [],
      appliedBatchByPartner: {},
      uploadedAt: "",
      slipWorkflow: {}
    };
  }
  if (!Array.isArray(db.outboundOrderUpload.lines)) db.outboundOrderUpload.lines = [];
  if (!Array.isArray(db.outboundOrderUpload.uploadedRows)) {
    // migrate legacy data: old versions stored all uploaded rows in lines
    db.outboundOrderUpload.uploadedRows = [...db.outboundOrderUpload.lines];
  }
  if (!Array.isArray(db.outboundOrderUpload.batches)) db.outboundOrderUpload.batches = [];
  if (!db.outboundOrderUpload.appliedBatchByPartner || typeof db.outboundOrderUpload.appliedBatchByPartner !== "object") {
    db.outboundOrderUpload.appliedBatchByPartner = {};
  }
  if (!db.outboundOrderUpload.slipWorkflow || typeof db.outboundOrderUpload.slipWorkflow !== "object") {
    db.outboundOrderUpload.slipWorkflow = {};
  }
  if (!db.outboundOrderUpload.confirmedLists || typeof db.outboundOrderUpload.confirmedLists !== "object") {
    db.outboundOrderUpload.confirmedLists = {};
  }
  if (!Array.isArray(db.outboundCodeMasters)) db.outboundCodeMasters = [];
  if (!db.unshipCompare || typeof db.unshipCompare !== "object") db.unshipCompare = {};
  for (const pt of ["emart", "daiso", "lotte"]) {
    if (!db.unshipCompare[pt] || typeof db.unshipCompare[pt] !== "object") {
      db.unshipCompare[pt] = { officialRows: [], adjustments: [] };
    }
    if (!Array.isArray(db.unshipCompare[pt].officialRows)) db.unshipCompare[pt].officialRows = [];
    if (!Array.isArray(db.unshipCompare[pt].adjustments)) db.unshipCompare[pt].adjustments = [];
  }
  if (!db.ecvanCredentials || typeof db.ecvanCredentials !== "object") db.ecvanCredentials = { id: "", pw: "" };
}

function sendJson(res, code, payload) {
  res.writeHead(code, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(payload));
}

function toTagList(value) {
  if (Array.isArray(value)) return value.map((x) => String(x || "").trim()).filter(Boolean);
  return String(value || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}

/** 엑셀·입력창 흔한 공백( NBSP ) 정리 */
function normalizeInputLoose(s) {
  return String(s || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** 기본정보에 등록된 창고명과 맞춤(표기만 다른 경우) */
function resolveRegisteredWarehouse(db, raw) {
  const key = normalizeInputLoose(raw);
  if (!key) return "";
  const list = db.warehouses || [];
  if (list.includes(key)) return key;
  const hit = list.find((w) => normalizeInputLoose(w) === key);
  return hit || "";
}

/** 일괄입고 등: 품목코드·이카운트코드·바코드(SKU)·중포바코드로 상품 탐색 후 대표 code 사용 */
function findProductByMovementImportCode(db, raw) {
  const pc = normalizeInputLoose(raw);
  if (!pc) return null;
  const list = db.products || [];
  return (
    list.find((p) => normalizeInputLoose(p.code) === pc) ||
    list.find((p) => normalizeInputLoose(p.ecountCode) === pc) ||
    list.find((p) => normalizeInputLoose(p.barcode) === pc) ||
    list.find((p) => normalizeInputLoose(p.middleBarcode) === pc) ||
    null
  );
}

function normalizeOptionValues(value) {
  const base = Array.isArray(value) ? value : toTagList(value);
  const flat = [];
  for (const v of base) flat.push(...toTagList(v));
  return Array.from(new Set(flat.map((x) => String(x || "").trim()).filter(Boolean)));
}

function firstToken(value, fallback = "") {
  return toTagList(value)[0] || fallback;
}

function parseBody(req) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
      if (raw.length > 5 * 1024 * 1024) {
        reject(new Error("Request too large"));
      }
    });
    req.on("end", () => {
      if (!raw) return resolve({});
      try {
        resolve(JSON.parse(raw));
      } catch (e) {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function serveStatic(req, res, pathname) {
  const safePath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.join(PUBLIC_DIR, safePath);
  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403);
    return res.end("Forbidden");
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404);
      return res.end("Not Found");
    }
    const ext = path.extname(filePath).toLowerCase();
    const type =
      ext === ".html"
        ? "text/html; charset=utf-8"
        : ext === ".css"
          ? "text/css; charset=utf-8"
          : ext === ".js"
            ? "application/javascript; charset=utf-8"
            : "application/octet-stream";
    res.writeHead(200, { "Content-Type": type });
    res.end(data);
  });
}

function upsertProduct(db, product) {
  const ecountCode = String(product.ecountCode || product.code || "").trim();
  const code = String(product.code || ecountCode).trim();
  if (!code) throw new Error("상품코드는 필수입니다.");
  const name = String(product.name || product.ecountName || "").trim();
  if (!name) throw new Error("상품명은 필수입니다.");
  const found = db.products.find((p) => p.code === code || String(p.ecountCode || "").trim() === ecountCode);
  const deliveryVendors = toTagList(product.deliveryVendors);
  const orderManagers = toTagList(product.orderManagers);
  const categories = toTagList(product.categories || product.category);
  const usedWarehouses = toTagList(product.usedWarehouses).filter((w) => db.warehouses.includes(w));
  const updated = {
    name,
    ecountCode,
    ecountName: String(product.ecountName || name),
    barcode: String(product.barcode || ""),
    middleBarcode: String(product.middleBarcode || ""),
    logisticsBarcode: String(product.logisticsBarcode || ""),
    status: firstToken(product.status, "판매중"),
    deliveryVendors,
    deliveryVendorCode: String(product.deliveryVendorCode || ""),
    deliveryItemName: String(product.deliveryItemName || ""),
    spec: String(product.spec || ""),
    purchaseVendor: String(product.purchaseVendor || ""),
    supplyType: firstToken(product.supplyType, ""),
    orderDept: firstToken(product.orderDept, ""),
    orderManagers,
    purchaseItemCode: String(product.purchaseItemCode || ""),
    purchaseItemName: String(product.purchaseItemName || ""),
    warehouseGroup: firstToken(product.warehouseGroup, ""),
    usedWarehouses,
    itemType: firstToken(product.itemType, ""),
    categories,
    // Backward compatibility for existing stock screens.
    category: categories.join(", "),
    salesVendor: deliveryVendors.join(", "),
    unit: String(product.unit || "EA"),
    safetyStock: Number(product.safetyStock || 0),
    optimalStock: Number(product.optimalStock || 0),
    note: String(product.note || "")
  };
  if (found) {
    // Keep legacy internal code to avoid breaking existing movement history references.
    updated.code = found.code;
    Object.assign(found, updated);
    found.updatedAt = new Date().toISOString();
  } else {
    db.products.push({
      code,
      ...updated,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    });
  }
}

function calculateStock(db, warehouseFilter) {
  const items = [];
  const warehouses = warehouseFilter ? [warehouseFilter] : db.warehouses;
  for (const w of warehouses) {
    const stockMap = {};
    for (const p of db.products) stockMap[p.code] = 0;
    for (const m of db.movements) {
      const qty = Number(m.qty || 0);
      if (m.type === "TRANSFER") {
        if (m.warehouse === w) stockMap[m.productCode] = (stockMap[m.productCode] || 0) - Math.abs(qty);
        if (m.toWarehouse === w) stockMap[m.productCode] = (stockMap[m.productCode] || 0) + Math.abs(qty);
        continue;
      }
      if (m.warehouse !== w) continue;
      stockMap[m.productCode] = (stockMap[m.productCode] || 0) + qty;
    }
    for (const p of db.products) {
      items.push({
        ...p,
        warehouse: w,
        stock: stockMap[p.code] || 0
      });
    }
  }
  return items;
}

function addMovement(db, row) {
  const type = row.type;
  if (!["IN", "OUT", "ADJUST", "TRANSFER"].includes(type)) {
    throw new Error("유효하지 않은 구분입니다.");
  }
  const productCodeInput = normalizeInputLoose(row.productCode);
  const qtyRaw = Number(row.qty);
  const warehouseRaw = normalizeInputLoose(row.warehouse);
  const toWarehouseRaw = normalizeInputLoose(row.toWarehouse);
  if (!productCodeInput) throw new Error("상품코드는 필수입니다.");
  if (!Number.isFinite(qtyRaw) || qtyRaw === 0) {
    throw new Error("수량은 0이 아닌 숫자여야 합니다.");
  }
  const warehouse = resolveRegisteredWarehouse(db, warehouseRaw);
  if (!warehouse) throw new Error(`등록되지 않은 창고: ${warehouseRaw || "-"}`);
  let toWarehouse = "";
  if (type === "TRANSFER") {
    toWarehouse = resolveRegisteredWarehouse(db, toWarehouseRaw);
    if (!toWarehouse) throw new Error(`등록되지 않은 창고: ${toWarehouseRaw || "-"}`);
    if (warehouse === toWarehouse) throw new Error("출발/도착 창고가 동일합니다.");
  }
  const product = findProductByMovementImportCode(db, productCodeInput);
  if (!product) {
    throw new Error(
      `등록되지 않은 상품입니다(품목코드·이카운트코드·바코드(SKU)·중포바코드로 조회): ${productCodeInput}`
    );
  }
  const productCode = normalizeInputLoose(product.code);
  if (!productCode) throw new Error("상품의 품목코드(code)가 비어 있습니다.");
  const usedWarehouses = toTagList(product.usedWarehouses);
  if (usedWarehouses.length) {
    if (!usedWarehouses.includes(warehouse)) {
      throw new Error(
        `해당 상품은 지정된 사용창고에서만 처리할 수 있습니다: ${usedWarehouses.join(", ")} (요청 창고: ${warehouse}, 품목: ${productCode} ${product.name || ""})`
      );
    }
    if (type === "TRANSFER" && !usedWarehouses.includes(toWarehouse)) {
      throw new Error(`이동 대상 창고가 사용창고에 없습니다: ${usedWarehouses.join(", ")}`);
    }
  }
  const user = String(row.user || "").trim();
  if (!user) throw new Error("담당자는 필수입니다.");
  if (!db.managers.includes(user)) {
    throw new Error(`등록되지 않은 담당자입니다: ${user}`);
  }
  const qty = type === "OUT" ? -Math.abs(qtyRaw) : type === "TRANSFER" ? Math.abs(qtyRaw) : qtyRaw;
  const stockItems = calculateStock(db, warehouse).filter((x) => x.code === productCode);
  const current = stockItems.length ? Number(stockItems[0].stock || 0) : 0;
  if (type === "OUT" && current + qty < 0) {
    throw new Error(`출고 불가(재고 부족): ${productCode}, 현재고 ${current}`);
  }
  if (type === "TRANSFER" && current - Math.abs(qty) < 0) {
    throw new Error(`이동 불가(재고 부족): ${productCode}, 현재고 ${current}`);
  }

  db.movements.push({
    id: db.seq++,
    type,
    productCode,
    qty,
    warehouse,
    toWarehouse: type === "TRANSFER" ? toWarehouse : "",
    partner: String(row.partner || ""),
    memo: String(row.memo || ""),
    slipNo: String(row.slipNo || "").trim(),
    user,
    cancelled: false,
    createdAt: row.createdAt || new Date().toISOString()
  });
}

function cancelMovement(db, id, user) {
  const target = db.movements.find((m) => Number(m.id) === Number(id));
  if (!target) throw new Error("취소할 이력이 없습니다.");
  if (target.cancelled) throw new Error("이미 취소된 이력입니다.");
  if (!db.managers.includes(user)) throw new Error("취소 담당자는 등록된 담당자여야 합니다.");
  target.cancelled = true;
  target.cancelledAt = new Date().toISOString();
  target.cancelledBy = user;

  db.movements.push({
    id: db.seq++,
    type: "CANCEL",
    productCode: target.productCode,
    qty: -Number(target.qty || 0),
    warehouse: target.warehouse || "유통사업부",
    toWarehouse: target.toWarehouse || "",
    partner: target.partner,
    memo: `원거래 ${target.id} 취소`,
    user,
    cancelled: false,
    originId: target.id,
    originType: target.type,
    createdAt: new Date().toISOString()
  });
}

function computeStockAfterEachMovement(db) {
  const stockMap = {};
  for (const w of db.warehouses) {
    stockMap[w] = {};
    for (const p of db.products) stockMap[w][p.code] = 0;
  }
  const sorted = [...db.movements].sort((a, b) => {
    const da = new Date(a.createdAt || 0).getTime();
    const dbt = new Date(b.createdAt || 0).getTime();
    if (da !== dbt) return da - dbt;
    return Number(a.id || 0) - Number(b.id || 0);
  });
  const afterMap = {};
  for (const m of sorted) {
    const from = m.warehouse || "유통사업부";
    if (!stockMap[from]) stockMap[from] = {};
    if (!(m.productCode in stockMap[from])) stockMap[from][m.productCode] = 0;
    if (m.type === "TRANSFER") {
      const to = m.toWarehouse || "";
      stockMap[from][m.productCode] -= Math.abs(Number(m.qty || 0));
      if (to) {
        if (!stockMap[to]) stockMap[to] = {};
        if (!(m.productCode in stockMap[to])) stockMap[to][m.productCode] = 0;
        stockMap[to][m.productCode] += Math.abs(Number(m.qty || 0));
      }
      afterMap[m.id] = stockMap[from][m.productCode];
    } else {
      stockMap[from][m.productCode] += Number(m.qty || 0);
      afterMap[m.id] = stockMap[from][m.productCode];
    }
  }
  return afterMap;
}

let ecountSessionCache = { value: "", expiresAt: 0 };
let ecountZoneCache = { zone: "", domain: "", expiresAt: 0 };
let ecountApiBaseCache = { value: "", expiresAt: 0 };

function normalizeDateCompact(ymd) {
  const raw = String(ymd || "").trim();
  const onlyDigits = raw.replace(/[^\d]/g, "");
  if (onlyDigits.length !== 8) return "";
  return onlyDigits;
}

function compactToUtcDate(yyyymmdd) {
  const s = normalizeDateCompact(yyyymmdd);
  if (!s) return null;
  const y = Number(s.slice(0, 4));
  const m = Number(s.slice(4, 6));
  const d = Number(s.slice(6, 8));
  if (!y || !m || !d) return null;
  return new Date(Date.UTC(y, m - 1, d));
}

function utcDateToCompact(dt) {
  const y = dt.getUTCFullYear();
  const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
  const d = String(dt.getUTCDate()).padStart(2, "0");
  return `${y}${m}${d}`;
}

function clampDateRangeToMax30Days(fromCompact, toCompact) {
  let fromDate = compactToUtcDate(fromCompact);
  let toDate = compactToUtcDate(toCompact);
  if (!fromDate || !toDate) return { fromCompact, toCompact, adjusted: false };
  // If user selects reversed dates, normalize order.
  if (fromDate.getTime() > toDate.getTime()) {
    const tmp = fromDate;
    fromDate = toDate;
    toDate = tmp;
  }
  const normalizedFrom = utcDateToCompact(fromDate);
  const normalizedTo = utcDateToCompact(toDate);
  const diffDays = Math.floor((toDate.getTime() - fromDate.getTime()) / 86400000) + 1; // inclusive
  if (diffDays <= 30) return { fromCompact: normalizedFrom, toCompact: normalizedTo, adjusted: false };
  const adjustedFrom = new Date(toDate.getTime() - 29 * 86400000); // inclusive 30 days
  return { fromCompact: utcDateToCompact(adjustedFrom), toCompact: normalizedTo, adjusted: true };
}

function extractSessionId(payload) {
  if (!payload || typeof payload !== "object") return "";
  const datasNested =
    payload.Data &&
    payload.Data.Datas &&
    typeof payload.Data.Datas === "object" &&
    (payload.Data.Datas.SESSION_ID || payload.Data.Datas.session_id || payload.Data.Datas.sessionId);
  const direct =
    payload.SESSION_ID ||
    payload.session_id ||
    payload.sessionId ||
    (payload.Data && (payload.Data.SESSION_ID || payload.Data.session_id || payload.Data.sessionId)) ||
    datasNested;
  if (direct) return String(direct);

  // Fallback: recursive search for possible session fields.
  const stack = [payload];
  while (stack.length) {
    const cur = stack.pop();
    if (!cur || typeof cur !== "object") continue;
    for (const [k, v] of Object.entries(cur)) {
      if (v && typeof v === "object") stack.push(v);
      if (/session[_-]?id/i.test(String(k)) && v) return String(v);
    }
  }
  return "";
}

async function fetchEcountSession() {
  const now = Date.now();
  if (ecountSessionCache.value && ecountSessionCache.expiresAt > now + 10_000) {
    return ecountSessionCache.value;
  }

  const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();
  const userId = String(process.env.ECOUNT_USER_ID || "").trim();
  const userPw = String(process.env.ECOUNT_USER_PW || "").trim();
  const apiKey = String(process.env.ECOUNT_API_KEY || "").trim();
  const lang = String(process.env.ECOUNT_LANG || "ko-KR").trim();
  const zoneInfo = await fetchEcountZone(comCode);

  if (!comCode || !userId || !userPw || !apiKey) {
    throw new Error(
      "ECOUNT_COM_CODE/ECOUNT_USER_ID/ECOUNT_USER_PW/ECOUNT_API_KEY 설정이 필요합니다."
    );
  }

  const loginBodies = Array.from(
    new Set([String(comCode || "").trim(), String(comCode || "").trim().padStart(6, "0")].filter(Boolean))
  ).map((cc) => ({
    COM_CODE: cc,
    USER_ID: userId,
    USER_PW: userPw,
    API_CERT_KEY: apiKey,
    LAN_TYPE: lang,
    ZONE: String(zoneInfo?.zone || "").trim(),
    DOMAIN: String(zoneInfo?.domain || "").trim()
  }));
  const loginCandidates = [];
  const forcedLoginUrl = String(process.env.ECOUNT_LOGIN_URL || "").trim();
  if (forcedLoginUrl) loginCandidates.push(forcedLoginUrl);
  for (const origin of buildEcountApiOrigins(zoneInfo)) {
    loginCandidates.push(`${origin}/OAPI/V2/OAPILogin`);
  }

  let sessionId = "";
  let lastErr = "";
  for (const loginUrl of Array.from(new Set(loginCandidates))) {
    for (const loginBody of loginBodies) {
      try {
        const resp = await fetch(loginUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(loginBody)
        });
        const txt = await resp.text();
        if (!resp.ok) {
          lastErr = `${loginUrl} [COM_CODE=${loginBody.COM_CODE}] -> (${resp.status}) ${txt.slice(0, 300)}`;
          continue;
        }
        let data = {};
        try {
          data = JSON.parse(txt || "{}");
        } catch (_) {
          data = {};
        }
        sessionId = extractSessionId(data);
        if (!sessionId) {
          lastErr = `${loginUrl} [COM_CODE=${loginBody.COM_CODE}] -> SESSION_ID not found: ${txt.slice(0, 300)}`;
          continue;
        }
        // Keep the successful API origin for subsequent calls.
        const origin = new URL(loginUrl).origin;
        ecountApiBaseCache = { value: origin, expiresAt: now + 20 * 60 * 1000 };
        break;
      } catch (err) {
        lastErr = `${loginUrl} [COM_CODE=${loginBody.COM_CODE}] -> network error: ${err?.message || String(err)}`;
        continue;
      }
    }
    if (sessionId) break;
  }
  if (!sessionId) {
    throw new Error(`이카운트 로그인 실패: ${lastErr || "응답 확인 필요"}`);
  }

  // Default cache 20 minutes (session ttl unknown); refresh early.
  ecountSessionCache = {
    value: sessionId,
    expiresAt: now + 20 * 60 * 1000
  };
  return sessionId;
}

function buildEcountApiOrigins(zoneInfo) {
  const list = [];
  const forcedBase = String(process.env.ECOUNT_API_BASE_URL || "").trim();
  if (forcedBase) {
    try {
      list.push(new URL(forcedBase).origin);
    } catch (_) {
      list.push(`https://${forcedBase.replace(/^https?:\/\//, "").replace(/\/+$/, "")}`);
    }
  }
  const domain = String(zoneInfo?.domain || "").trim();
  if (domain) {
    const cleaned = domain.replace(/^https?:\/\//, "").replace(/\/+$/, "");
    list.push(`https://${cleaned}`);
  }
  const zone = String(zoneInfo?.zone || "").trim();
  if (zone) {
    list.push(`https://oapi${zone}.ecount.com`);
    list.push(`https://sboapi${zone}.ecount.com`);
  }
  return Array.from(new Set(list));
}

async function fetchEcountApiBase(comCode) {
  const now = Date.now();
  if (ecountApiBaseCache.value && ecountApiBaseCache.expiresAt > now + 10_000) {
    return ecountApiBaseCache.value;
  }
  const zoneInfo = await fetchEcountZone(comCode);
  const origins = buildEcountApiOrigins(zoneInfo);
  if (!origins.length) throw new Error("이카운트 API 베이스 URL을 구성하지 못했습니다.");
  ecountApiBaseCache = { value: origins[0], expiresAt: now + 20 * 60 * 1000 };
  return ecountApiBaseCache.value;
}

async function fetchEcountZone(comCode) {
  const forced = String(process.env.ECOUNT_ZONE || "").trim();
  const forcedDomain = String(process.env.ECOUNT_DOMAIN || "").trim();
  if (forced) return { zone: forced, domain: forcedDomain };

  const now = Date.now();
  if (ecountZoneCache.zone && ecountZoneCache.expiresAt > now + 60_000) {
    return { zone: ecountZoneCache.zone, domain: ecountZoneCache.domain };
  }
  if (!comCode) throw new Error("ECOUNT_COM_CODE가 필요합니다.");

  const zoneUrl = "https://oapi.ecount.com/OAPI/V2/Zone";
  const comCandidates = Array.from(
    new Set([String(comCode || "").trim(), String(comCode || "").trim().padStart(6, "0")].filter(Boolean))
  );

  let zone = "";
  let domain = "";
  let lastErrText = "";
  for (const candidate of comCandidates) {
    try {
      const resp = await fetch(zoneUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json", Accept: "application/json" },
        body: JSON.stringify({ COM_CODE: candidate })
      });
      const bodyText = await resp.text();
      if (!resp.ok) {
        lastErrText = `COM_CODE=${candidate} -> (${resp.status}) ${bodyText.slice(0, 300)}`;
        continue;
      }
      let data = {};
      try {
        data = JSON.parse(bodyText || "{}");
      } catch (_) {
        data = {};
      }
      zone = String(
        pickFirst(data, ["ZONE"]) || pickFirst(data?.Data, ["ZONE"]) || pickFirst(data?.Result, ["ZONE"]) || ""
      ).trim();
      domain = String(
        pickFirst(data, ["DOMAIN"]) || pickFirst(data?.Data, ["DOMAIN"]) || pickFirst(data?.Result, ["DOMAIN"]) || ""
      ).trim();
      if (zone) break;
      lastErrText = `COM_CODE=${candidate} -> ZONE not found: ${bodyText.slice(0, 300)}`;
    } catch (err) {
      lastErrText = `COM_CODE=${candidate} -> network error: ${err?.message || String(err)}`;
      continue;
    }
  }

  if (!zone) {
    throw new Error(
      `Zone API 조회 실패. 회사코드(원본:${comCode}, 6자리:${String(comCode || "").trim().padStart(6, "0")})를 확인해주세요. 응답: ${String(
        lastErrText || ""
      ).slice(0, 300)}`
    );
  }

  // Cache for one day; Zone rarely changes.
  ecountZoneCache = { zone, domain, expiresAt: now + 24 * 60 * 60 * 1000 };
  return { zone, domain };
}

function pickFirst(obj, keys) {
  for (const k of keys) {
    if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
  }
  return "";
}

function pickFirstNonEmpty(obj, keys) {
  for (const k of keys) {
    const v = obj && obj[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return "";
}

function pickFirstByKeyPattern(obj, patternFn) {
  if (!obj || typeof obj !== "object") return "";
  for (const k of Object.keys(obj)) {
    if (!patternFn(k)) continue;
    const v = obj[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return v;
  }
  return "";
}

function pickPurchaseRequestNo(obj) {
  const ordNo = String(pickFirstNonEmpty(obj, ["ORD_NO"]) || "").trim();
  if (ordNo) return ordNo;

  const ioNo = String(pickFirstNonEmpty(obj, ["IO_NO"]) || "").trim();
  if (ioNo) return ioNo;

  // Keep purchase-request style keys as a conservative fallback only.
  return String(
    pickFirstNonEmpty(obj, [
      "PR_NO",
      "REQ_NO",
      "PUR_REQ_NO",
      "REQUEST_NO",
      "PURCHASE_REQ_NO",
      "REQNUM",
      "REQ_NUM",
      "REQNO",
      "PRNUM",
      "PR_NUM",
      "PURCHASE_REQUEST_NO",
      "구매요청번호"
    ]) || ""
  ).trim();
}

function formatOrderNoByDateAndSeq(orderNo, orderDate, fallbackDateCompact = "") {
  const no = String(orderNo || "").trim();
  const dt = String(orderDate || "").trim();
  if (!no) return no;

  // Already complete format like 26/04/20-1
  if (/^\d{2}\/\d{2}\/\d{2}-\d+$/.test(no)) return no;

  // If numeric-like suffix is present (e.g. "1", "1.0000000000"), rebuild with order date.
  const seqMatch = no.match(/(\d+)(?:\.0+)?$/);
  if (!seqMatch) return no;
  const seq = String(seqMatch[1] || "").trim();
  if (!seq) return no;
  const compact = normalizeDateCompact(dt) || String(fallbackDateCompact || "").trim();
  if (!compact || compact.length !== 8) return no;
  const yy = compact.slice(2, 4);
  const mm = compact.slice(4, 6);
  const dd = compact.slice(6, 8);
  return `${yy}/${mm}/${dd}-${seq}`;
}

function buildInboundPlanDetail(raw) {
  const x = raw && typeof raw === "object" ? raw : {};
  const entries = [
    ["전표번호", pickFirstNonEmpty(x, ["ORD_NO", "IO_NO"])],
    ["발주일자", pickFirstNonEmpty(x, ["ORD_DATE", "IO_DATE"])],
    ["납기일자", pickFirstNonEmpty(x, ["TIME_DATE"])],
    ["거래처코드", pickFirstNonEmpty(x, ["CUST"])],
    ["거래처명", pickFirstNonEmpty(x, ["CUST_DES"])],
    ["담당자", pickFirstNonEmpty(x, ["CUST_NAME", "EMP_CD", "WRITER_ID", "LOGID"])],
    ["품목명", pickFirstNonEmpty(x, ["PROD_DES", "ITEM_DES", "ITEM_NAME"])],
    ["수량", pickFirstNonEmpty(x, ["QTY", "ORDER_QTY", "PUR_QTY"])],
    ["창고코드", pickFirstNonEmpty(x, ["WH_CD"])],
    ["창고명", pickFirstNonEmpty(x, ["WH_DES"])],
    ["비고", pickFirstNonEmpty(x, ["REF_DES"])],
    ["상태코드", pickFirstNonEmpty(x, ["P_FLAG"])],
    ["전송상태", pickFirstNonEmpty(x, ["SEND_FLAG"])]
  ];
  for (let i = 1; i <= 6; i += 1) {
    entries.push([`보조항목${i}`, pickFirstNonEmpty(x, [`P_DES${i}`])]);
  }
  return entries
    .map(([label, value]) => ({ label, value: String(value || "").trim() }))
    .filter((row) => row.value);
}

/** 발주서조회API Result 한 행 → WMS 입고예정목록 행 (이카운트 필드명은 주석/문서와 동일 계열). */
function normalizeInboundPlanRows(rawRows, fallbackDateCompact = "") {
  const rows = Array.isArray(rawRows) ? rawRows : [];
  return rows
    .map((x) => {
      const poNo = String(pickPurchaseRequestNo(x));
      const poDate = String(
        pickFirstNonEmpty(x, ["ORD_DATE", "PO_DATE", "PODATE", "ORDER_DATE", "DATE", "WRITE_DATE", "IO_DATE"])
      );
      // 거래처는 CUST_DES(거래처명)를 최우선으로 사용.
      const vendor = String(
        pickFirstNonEmpty(x, ["CUST_DES", "CUST_NM", "VENDOR_NM", "VENDOR_NAME", "CLIENT_NAME", "CUST"])
      );
      const manager = String(pickFirstNonEmpty(x, ["CUST_NAME", "EMP_CD", "WRITER_ID", "LOGID", "MANAGER", "MANAGER_NM"]));
      const itemCode = String(
        pickFirstNonEmpty(x, ["PROD_CD", "ITEM_CD", "ITEM_CODE", "GOODS_CD", "PRODUCT_CODE", "ITEM_NO", "품목코드"])
      );
      const itemName = String(
        pickFirstNonEmpty(x, [
          "PROD_DES",
          "ITEM_NM",
          "ITEM_NAME",
          "GOODS_NM",
          "PRODUCT_NAME",
          "ITEM_DES",
          "품목명",
          "DES",
          "REMARK"
        ])
      );
      const spec = String(pickFirstNonEmpty(x, ["SPEC", "SIZE_DES", "ITEM_SPEC", "규격"]));
      const barcode = String(pickFirstNonEmpty(x, ["BAR_CODE", "BARCODE", "PROD_BARCODE", "ITEM_BARCODE"]));
      const boxQty = pickFirstNonEmpty(x, ["UQTY", "BOX_QTY", "QTY_BOX"]);
      const qty = pickFirstNonEmpty(x, ["QTY", "ORDER_QTY", "PUR_QTY", "PO_QTY", "STOCK_QTY", "AMT_QTY"]);
      const dueDate = String(pickFirstNonEmpty(x, ["TIME_DATE"]));
      const whName = String(pickFirstNonEmpty(x, ["WH_NM", "WAREHOUSE_NM", "WH_NAME", "WH_DES", "WAREHOUSE_NAME"]));
      const status = String(pickFirstNonEmpty(x, ["STATUS", "STAT_NM", "PROC_STATUS", "STATE", "PROGRESS"]));
      const remark = String(pickFirstNonEmpty(x, ["REF_DES", "REMARK", "NOTE", "DESC"]));
      const note = String(pickFirstNonEmpty(x, ["MEMO", "BIGO", "NOTE2"]));
      const currency = String(pickFirstNonEmpty(x, ["CODE_DES", "EXCHANGE_TYPE", "FOREIGN_FLAG"]));
      const lastModifiedAt = String(
        pickFirstNonEmpty(x, [
          "MODIFY_DATE",
          "MODIFIED_DATE",
          "MOD_DATE",
          "LAST_MOD_DATE",
          "LAST_MODIFIED_DATE",
          "LAST_UPDATE_DATE",
          "UPDATE_DATE",
          "UPD_DATE",
          "CHG_DATE",
          "EDIT_DATE",
          "WRITE_DATE"
        ]) || pickFirstByKeyPattern(x, (k) => /modi|update|upd|chg|edit|수정/i.test(String(k)))
      );
      const displayPoNo = formatOrderNoByDateAndSeq(poNo, poDate, fallbackDateCompact);
      const detail = buildInboundPlanDetail(x);
      return {
        poNo: displayPoNo,
        poDate,
        vendor,
        manager,
        itemCode,
        itemName,
        spec,
        barcode,
        boxQty,
        qty,
        dueDate,
        whName,
        status,
        remark,
        note,
        currency,
        lastModifiedAt,
        detail
      };
    });
}

function normalizeSlipNoText(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  const compact = s.replace(/\s+/g, "");
  // Treat 26/04/18-1 and 2026/04/18 -1 as the same slip key.
  const m = compact.match(/^(\d{2}|\d{4})\/?(\d{2})\/?(\d{2})-(\d+)$/);
  if (m) {
    const yy = m[1].length === 4 ? m[1].slice(2) : m[1];
    return `${yy}/${m[2]}/${m[3]}-${String(Number(m[4]))}`;
  }
  return compact;
}

function ymdOrCompactForPurchasesSave(v) {
  const c = normalizeDateCompact(v);
  if (c) return c;
  const s = String(v || "").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.replace(/-/g, "");
  return "";
}

/**
 * 발주 조회로 모은 헤더·품목행 → SavePurchases PurchasesList (1품목 = 1건).
 * opts: { ioDate, custCd, whCd } — 헤더에 코드가 없을 때 테스트용으로 덮어쓰기 가능.
 */
function buildSavePurchasesPayloadFromOrder(headerRaw, matchedLines, opts = {}) {
  const lines = Array.isArray(matchedLines) ? matchedLines : [];
  const hdr = headerRaw && typeof headerRaw === "object" ? headerRaw : {};
  const ioDate =
    ymdOrCompactForPurchasesSave(opts.ioDate) ||
    ymdOrCompactForPurchasesSave(pickFirstNonEmpty(hdr, ["ORD_DATE", "IO_DATE", "WRITE_DATE"])) ||
    ymdOrCompactForPurchasesSave(lines[0]?.poDate) ||
    normalizeDateCompact(new Date().toISOString().slice(0, 10));

  const cust = String(opts.custCd || pickFirstNonEmpty(hdr, ["CUST", "CUST_CD"]) || "").trim();
  const custDes = String(pickFirstNonEmpty(hdr, ["CUST_DES", "CUST_NM"]) || "").trim();
  const whCd = String(opts.whCd || pickFirstNonEmpty(hdr, ["WH_CD"]) || "").trim();
  const ordNo = String(pickFirstNonEmpty(hdr, ["ORD_NO"]) || "").trim();
  const ordDate = ymdOrCompactForPurchasesSave(pickFirstNonEmpty(hdr, ["ORD_DATE"]));
  const empCd = String(pickFirstNonEmpty(hdr, ["EMP_CD", "CUST_NAME", "WRITER_ID"]) || "").trim();

  const usable = lines.filter((x) => String(x.itemCode || "").trim() && String(x.qty ?? "").trim() !== "");
  if (!usable.length) {
    throw new Error(
      "구매입력용 품목 라인이 없습니다(품목코드·수량). 발주 상세에서 품목 행이 열리는지 확인하거나 .env의 ECOUNT_INBOUND_LINE_API_PATHS 등을 설정하세요."
    );
  }
  if (!cust) throw new Error("거래처코드(CUST)를 찾지 못했습니다. 발주 헤더에 없으면 요청 본문에 custCd로 넘기세요.");
  if (!whCd) throw new Error("창고코드(WH_CD)를 찾지 못했습니다. 발주 헤더에 없으면 요청 본문에 whCd로 넘기세요.");

  const PurchasesList = usable.map((line) => {
    const qty = String(line.qty ?? "")
      .replace(/,/g, "")
      .trim();
    const timeDate = ymdOrCompactForPurchasesSave(line.dueDate);
    const bulk = {
      IO_DATE: ioDate,
      CUST: cust,
      CUST_DES: custDes,
      WH_CD: whCd,
      PROD_CD: String(line.itemCode || "").trim(),
      PROD_DES: String(line.itemName || "").trim(),
      QTY: qty
    };
    if (ordNo) bulk.ORD_NO = ordNo;
    if (ordDate) bulk.ORD_DATE = ordDate;
    if (empCd) bulk.EMP_CD = empCd;
    if (timeDate) bulk.TIME_DATE = timeDate;
    const ref = String(line.remark || "").trim();
    if (ref) bulk.REF_DES = ref;
    const memo = String(line.note || "").trim();
    if (memo) bulk.MEMO = memo;
    return { BulkDatas: bulk };
  });

  return { PurchasesList };
}

/** 업로드 전표 품목 행(품목코드 순) — 수량·구매입력과 동일 순서 유지 */
function getUploadSlipLineGroupSorted(lines, slipKeyNorm) {
  const group = (Array.isArray(lines) ? lines : []).filter((x) => normalizeSlipNoText(x.slipNo) === slipKeyNorm);
  group.sort((a, b) => String(a.itemCode).localeCompare(String(b.itemCode)));
  return group;
}

function mergeLineQtyByIndexFromBody(bodyMap, group) {
  const out = {};
  for (let i = 0; i < group.length; i++) {
    const raw =
      bodyMap && bodyMap[String(i)] !== undefined && bodyMap[String(i)] !== null && String(bodyMap[String(i)]).trim() !== ""
        ? Number(String(bodyMap[String(i)]).replace(/,/g, ""))
        : Number(String(group[i].qty ?? "").replace(/,/g, "")) || 0;
    out[String(i)] = raw;
  }
  return out;
}

const OUTBOUND_PARTNER_ADAPTERS = {
  daiso: {
    parserVersion: "daiso-v1",
    alias: {
      orderNo: ["발주번호", "주문번호", "orderNo", "ORDER_NO"],
      orderDate: ["주문일자", "발주일자", "orderDate"],
      dueDate: ["납기일", "출고예정일", "dueDate"],
      productCode: ["상품코드", "품번", "SKU", "productCode"],
      productName: ["상품명", "품명", "productName"],
      orderQty: ["주문수량", "수량", "orderQty"],
      warehouse: ["창고", "출고창고", "warehouse"],
      remark: ["비고", "메모", "remark"]
    }
  },
  emart: {
    parserVersion: "emart-v1",
    alias: {
      orderNo: ["주문번호", "발주번호", "OrderNo"],
      orderDate: ["주문일", "발주일", "OrderDate"],
      storeInDate: ["점입점일자"],
      poDate: ["발주일자"],
      storeName: ["점포명"],
      dueDate: ["납품요청일", "출고일", "DueDate"],
      productCode: ["품목코드", "SKU", "상품코드", "ProductCode"],
      productName: ["품목명", "상품명", "ProductName"],
      barcode: ["바코드", "Barcode", "BARCODE"],
      orderUnit: ["발주단위"],
      lot: ["LOT"],
      orderQty: ["수량", "요청수량", "주문수량", "Qty"],
      unshipStatus: ["미출상태"],
      fixedQty: ["확정수량"],
      orderAmount: ["발주금액"],
      unitPrice: ["발주원가", "단가", "주문단가", "판매단가", "UnitPrice", "PRICE"],
      supplyAmt: ["공급가액", "공급가", "SupplyAmt", "SUPPLY_AMT"],
      vatAmt: ["부가세", "부가세액", "VatAmt", "VAT_AMT"],
      totalAmt: ["합계", "총액", "TotalAmt", "TOTAL_AMT"],
      centerInDate: ["센터입하일자"],
      centerCode: ["센터코드"],
      centerName: ["센터이름", "센터명", "센터"],
      warehouse: ["출고창고", "센터", "센터명", "warehouse"],
      remark: ["메모", "비고", "Remark"]
    }
  },
  lotte: {
    parserVersion: "lotte-v1",
    alias: {
      orderNo: ["발주No", "주문번호", "OrderNo"],
      orderDate: ["발주일자", "주문일", "OrderDate"],
      dueDate: ["납기일자", "요청출고일", "DueDate"],
      productCode: ["상품코드", "품목코드", "SKU"],
      productName: ["상품명", "품목명"],
      orderQty: ["주문수량", "지시수량", "Qty"],
      warehouse: ["창고", "출고창고"],
      remark: ["비고", "메모"]
    }
  }
};

const OUTBOUND_PARTNER_TEMPLATE_HEADERS = {
  daiso: ["발주번호", "주문일자", "납기일", "상품코드", "상품명", "주문수량", "창고", "비고"],
  emart: ["주문번호", "주문일", "납품요청일", "품목코드", "품목명", "요청수량", "출고창고", "메모"],
  lotte: ["발주No", "발주일자", "납기일자", "상품코드", "상품명", "주문수량", "창고", "비고"]
};

function normalizeOutboundPartnerType(v) {
  const s = String(v || "").trim().toLowerCase();
  if (["daiso", "emart", "lotte"].includes(s)) return s;
  return "";
}

function partnerLabelFromType(t) {
  if (t === "daiso") return "다이소";
  if (t === "emart") return "이마트";
  if (t === "lotte") return "롯데마트";
  return "";
}

function getOutboundCodeMasterMap(db, partnerType) {
  const map = new Map();
  for (const row of db.outboundCodeMasters || []) {
    if (!row || row.partnerType !== partnerType) continue;
    const src = String(row.sourceCode || "").trim();
    const dst = String(row.ecountCode || "").trim();
    if (!src || !dst) continue;
    map.set(src, dst);
    const loose = normalizeInputLoose(src);
    if (loose) map.set(loose, dst);
    const numLike = src.replace(/,/g, "").trim();
    if (/^\d+\.0+$/.test(numLike)) map.set(numLike.replace(/\.0+$/, ""), dst);
  }
  return map;
}

function parseOutboundCodeMasterRows(matrix, partnerType, sourceFileName) {
  const rows = Array.isArray(matrix) ? matrix : [];
  if (!rows.length) return { items: [], errorRows: [] };
  const headerRow = detectHeaderRowIndex(
    rows,
    [
      ["바코드", "상품코드", "원본코드", "sourceCode", "SOURCE_CODE"],
      ["이카운트코드", "이카운트품목코드", "품목코드", "ecountCode", "ECOUNT_CODE"]
    ],
    30
  );
  const headers = rows[headerRow] || [];
  const iSrc = findHeaderIndexByAliases(headers, ["바코드", "상품코드", "원본코드", "sourceCode", "SOURCE_CODE"]);
  const iDst = findHeaderIndexByAliases(headers, ["이카운트코드", "이카운트품목코드", "품목코드", "ecountCode", "ECOUNT_CODE"]);
  let iName = findHeaderIndexByAliases(headers, [
    "품명 및 규격(이카운트)",
    "품명및규격(이카운트)",
    "품목명(이카운트)",
    "상품명(이카운트)",
    "품목명",
    "상품명",
    "name"
  ]);
  if (iName < 0) {
    const h = headers.map((x) => String(x || "").trim());
    iName = h.findIndex((x) => x.includes("품명") || x.includes("상품명"));
  }
  if (iSrc < 0 || iDst < 0) throw new Error("코드 마스터 필수 열(바코드/상품코드, 이카운트코드)을 찾지 못했습니다.");
  const items = [];
  const errorRows = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const src = String(row[iSrc] ?? "").trim();
    const dst = String(row[iDst] ?? "").trim();
    const name = iName >= 0 ? String(row[iName] ?? "").trim() : "";
    if (!src && !dst) continue;
    if (!src || !dst) {
      errorRows.push({ sourceRowNo: r + 1, reason: "원본코드/이카운트코드 누락", raw: row });
      continue;
    }
    items.push({
      partnerType,
      sourceCode: src,
      ecountCode: dst,
      name,
      sourceFileName: sourceFileName || "",
      updatedAt: new Date().toISOString()
    });
  }
  return { items, errorRows };
}

function findHeaderIndexByAliases(headers, aliases) {
  const hRaw = (Array.isArray(headers) ? headers : []).map((x) => String(x || "").trim().toLowerCase());
  const hNorm = hRaw.map((x) => normalizeHeaderKey(x));
  for (const alias of aliases || []) {
    const aRaw = String(alias || "").trim().toLowerCase();
    const aNorm = normalizeHeaderKey(aRaw);
    let i = hRaw.findIndex((x) => x === aRaw);
    if (i >= 0) return i;
    i = hNorm.findIndex((x) => x === aNorm);
    if (i >= 0) return i;
  }
  return -1;
}

function detectHeaderRowIndex(rows, requiredAliasGroups, maxScanRows = 20) {
  const scanMax = Math.min(Array.isArray(rows) ? rows.length : 0, maxScanRows);
  for (let r = 0; r < scanMax; r++) {
    const headers = rows[r] || [];
    const ok = requiredAliasGroups.every((aliases) => findHeaderIndexByAliases(headers, aliases) >= 0);
    if (ok) return r;
  }
  return 0;
}

function normalizeHeaderKey(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\s\-_().[\]{}]/g, "");
}

function parseOutboundOrderRowsByPartner(matrix, partnerType, uploadBatchId, sourceFileName, db) {
  const cfg = OUTBOUND_PARTNER_ADAPTERS[partnerType];
  if (!cfg) throw new Error("지원하지 않는 판매처 타입입니다.");
  const rows = Array.isArray(matrix) ? matrix : [];
  if (!rows.length) return { okRows: [], errorRows: [] };
  const requiredAliasGroups =
    partnerType === "emart"
      ? [cfg.alias.productCode, cfg.alias.orderQty, [...(cfg.alias.orderNo || []), ...(cfg.alias.poDate || [])]]
      : [cfg.alias.orderNo, cfg.alias.productCode, cfg.alias.orderQty];
  const headerRow = detectHeaderRowIndex(rows, requiredAliasGroups, 30);
  const headers = rows[headerRow] || [];
  const idx = {
    orderNo: findHeaderIndexByAliases(headers, cfg.alias.orderNo),
    orderDate: findHeaderIndexByAliases(headers, cfg.alias.orderDate),
    storeInDate: findHeaderIndexByAliases(headers, cfg.alias.storeInDate || []),
    poDate: findHeaderIndexByAliases(headers, cfg.alias.poDate || []),
    storeName: findHeaderIndexByAliases(headers, cfg.alias.storeName || []),
    dueDate: findHeaderIndexByAliases(headers, cfg.alias.dueDate),
    productCode: findHeaderIndexByAliases(headers, cfg.alias.productCode),
    productName: findHeaderIndexByAliases(headers, cfg.alias.productName),
    barcode: findHeaderIndexByAliases(headers, cfg.alias.barcode || []),
    orderUnit: findHeaderIndexByAliases(headers, cfg.alias.orderUnit || []),
    lot: findHeaderIndexByAliases(headers, cfg.alias.lot || []),
    orderQty: findHeaderIndexByAliases(headers, cfg.alias.orderQty),
    unshipStatus: findHeaderIndexByAliases(headers, cfg.alias.unshipStatus || []),
    fixedQty: findHeaderIndexByAliases(headers, cfg.alias.fixedQty || []),
    orderAmount: findHeaderIndexByAliases(headers, cfg.alias.orderAmount || []),
    unitPrice: findHeaderIndexByAliases(headers, cfg.alias.unitPrice || []),
    supplyAmt: findHeaderIndexByAliases(headers, cfg.alias.supplyAmt || []),
    vatAmt: findHeaderIndexByAliases(headers, cfg.alias.vatAmt || []),
    totalAmt: findHeaderIndexByAliases(headers, cfg.alias.totalAmt || []),
    centerInDate: findHeaderIndexByAliases(headers, cfg.alias.centerInDate || []),
    centerCode: findHeaderIndexByAliases(headers, cfg.alias.centerCode || []),
    centerName: findHeaderIndexByAliases(headers, cfg.alias.centerName || []),
    warehouse: findHeaderIndexByAliases(headers, cfg.alias.warehouse),
    remark: findHeaderIndexByAliases(headers, cfg.alias.remark)
  };
  if ((partnerType !== "emart" && idx.orderNo < 0) || idx.productCode < 0 || idx.orderQty < 0) {
    throw new Error("필수 열(주문번호 또는 발주일자, 상품코드, 주문수량)을 찾지 못했습니다.");
  }
  const okRows = [];
  const errorRows = [];
  const partner = partnerLabelFromType(partnerType);
  const masterMap = getOutboundCodeMasterMap(db || {}, partnerType);
  const startRow = headerRow + 1;
  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r] || [];
    const orderNoRaw = idx.orderNo >= 0 ? String(row[idx.orderNo] ?? "").trim() : "";
    const poDate = idx.poDate >= 0 ? String(row[idx.poDate] ?? "").trim() : "";
    const storeInDate = idx.storeInDate >= 0 ? String(row[idx.storeInDate] ?? "").trim() : "";
    const centerCode = idx.centerCode >= 0 ? String(row[idx.centerCode] ?? "").trim() : "";
    const storeName = idx.storeName >= 0 ? String(row[idx.storeName] ?? "").trim() : "";
    const autoOrderNoBase = [poDate || storeInDate, centerCode, storeName].filter(Boolean).join("-");
    const orderNo = orderNoRaw || (partnerType === "emart" ? autoOrderNoBase || `${uploadBatchId}-${r + 1}` : "");
    const productCodeRaw = String(row[idx.productCode] ?? "").trim();
    const productCodeLoose = normalizeInputLoose(productCodeRaw);
    const productCodeNum = productCodeRaw.replace(/,/g, "").trim().replace(/\.0+$/, "");
    const productCodeMapped =
      masterMap.get(productCodeRaw) || masterMap.get(productCodeLoose) || masterMap.get(productCodeNum) || "";
    const productCode = productCodeMapped || productCodeRaw;
    const orderQty = Number(String(row[idx.orderQty] ?? "").replace(/,/g, "").trim());
    if (!orderNo && !productCodeRaw) continue;
    if (!orderNo || !productCodeRaw || !Number.isFinite(orderQty) || orderQty <= 0) {
      errorRows.push({ sourceRowNo: r + 1, reason: "필수값 누락/수량 오류", raw: row });
      continue;
    }
    const slipNo = `${partnerType.toUpperCase()}-${orderNo}`;
    okRows.push({
      partnerType,
      partner,
      partnerCode: partnerType.toUpperCase(),
      parserVersion: cfg.parserVersion,
      uploadBatchId,
      sourceFileName: sourceFileName || "",
      sourceRowNo: r + 1,
      rawSnapshot: JSON.stringify(row).slice(0, 1000),
      orderNo,
      slipNo,
      orderDate: idx.orderDate >= 0 ? String(row[idx.orderDate] ?? "").trim() : poDate || storeInDate,
      dueDate: idx.dueDate >= 0 ? String(row[idx.dueDate] ?? "").trim() : "",
      productCode,
      sourceProductCode: productCodeRaw,
      mappedProductCode: productCodeMapped,
      productName: idx.productName >= 0 ? String(row[idx.productName] ?? "").trim() : "",
      barcode: idx.barcode >= 0 ? String(row[idx.barcode] ?? "").trim() : "",
      orderQty: Math.round(orderQty),
      orderUnit: idx.orderUnit >= 0 ? String(row[idx.orderUnit] ?? "").trim() : "",
      lot: idx.lot >= 0 ? String(row[idx.lot] ?? "").trim() : "",
      unshipStatus: idx.unshipStatus >= 0 ? String(row[idx.unshipStatus] ?? "").trim() : "",
      fixedQty: idx.fixedQty >= 0 ? String(row[idx.fixedQty] ?? "").trim() : "",
      orderAmount: idx.orderAmount >= 0 ? String(row[idx.orderAmount] ?? "").trim() : "",
      unitPrice: idx.unitPrice >= 0 ? String(row[idx.unitPrice] ?? "").trim() : "",
      supplyAmt: idx.supplyAmt >= 0 ? String(row[idx.supplyAmt] ?? "").trim() : "",
      vatAmt: idx.vatAmt >= 0 ? String(row[idx.vatAmt] ?? "").trim() : "",
      totalAmt: idx.totalAmt >= 0 ? String(row[idx.totalAmt] ?? "").trim() : "",
      centerInDate: idx.centerInDate >= 0 ? String(row[idx.centerInDate] ?? "").trim() : "",
      centerCode: idx.centerCode >= 0 ? String(row[idx.centerCode] ?? "").trim() : "",
      centerName:
        idx.centerName >= 0
          ? String(row[idx.centerName] ?? "").trim()
          : idx.warehouse >= 0
            ? String(row[idx.warehouse] ?? "").trim()
            : "",
      storeInDate: storeInDate,
      poDate: poDate,
      storeName: storeName,
      remark: idx.remark >= 0 ? String(row[idx.remark] ?? "").trim() : "",
      warehouse: idx.warehouse >= 0 ? String(row[idx.warehouse] ?? "").trim() : ""
    });
  }
  return { okRows, errorRows };
}

function aggregateOutboundUploadSlips(lines, slipWorkflow, partnerTypeFilter) {
  const arr = Array.isArray(lines) ? lines : [];
  const wf = slipWorkflow && typeof slipWorkflow === "object" ? slipWorkflow : {};
  const by = new Map();
  for (const x of arr) {
    if (partnerTypeFilter && x.partnerType !== partnerTypeFilter) continue;
    const k = normalizeSlipNoText(x.slipNo);
    if (!k) continue;
    if (!by.has(k)) by.set(k, []);
    by.get(k).push(x);
  }
  const items = [];
  for (const [k, group] of by) {
    const first = group[0];
    const w = wf[k] || {};
    const qMap = w.lineQtyByIndex && typeof w.lineQtyByIndex === "object" ? w.lineQtyByIndex : {};
    const totalOrderQty = group.reduce((s, x) => s + (Number(x.orderQty) || 0), 0);
    const totalConfirmQty = group.reduce((s, x, i) => {
      const q = qMap[String(i)] != null ? Number(qMap[String(i)]) : Number(x.orderQty) || 0;
      return s + (Number.isFinite(q) ? q : 0);
    }, 0);
    items.push({
      slipNo: first.slipNo,
      orderNo: first.orderNo,
      partnerType: first.partnerType,
      partner: first.partner,
      orderDate: first.orderDate,
      dueDate: first.dueDate,
      lineCount: group.length,
      totalOrderQty,
      totalConfirmQty,
      status: w.status || "draft",
      confirmedAt: w.confirmedAt || null,
      sentAt: w.sentAt || null
    });
  }
  items.sort((a, b) => String(b.orderDate || "").localeCompare(String(a.orderDate || "")));
  return items;
}

function outboundUploadLinesToDetailRows(lines, slipKeyNorm) {
  const group = (Array.isArray(lines) ? lines : []).filter((x) => normalizeSlipNoText(x.slipNo) === slipKeyNorm);
  group.sort((a, b) => String(a.productCode || "").localeCompare(String(b.productCode || ""), "ko"));
  return group;
}

/** 출고확정 취소: 확정 시 기록된 수량만큼 ADJUST로 재고 원복 후 전표를 초안으로 되돌림 */
function performOutboundWorkflowReopen(db, slipNoInput) {
  const slipDisp = String(slipNoInput || "").trim();
  const key = normalizeSlipNoText(slipNoInput);
  const group = outboundUploadLinesToDetailRows(db.outboundOrderUpload.lines || [], key);
  if (!group.length) throw new Error(`해당 전표의 품목 행이 없습니다: ${slipDisp || key}`);
  const wfPrev = db.outboundOrderUpload.slipWorkflow[key] || { status: "draft", lineQtyByIndex: {} };
  if (wfPrev.sentAt) throw new Error("이카운트 전송 완료 전표는 확정취소할 수 없습니다.");
  if (wfPrev.status !== "confirmed") throw new Error("출고확정된 전표만 확정취소할 수 있습니다.");
  const memoSlip = String(group[0]?.slipNo || slipDisp || key).trim();
  const stockUser = String(db.managers[0] || "");
  if (!stockUser) throw new Error("등록된 담당자가 없어 재고 원복을 처리할 수 없습니다.");
  for (let i = 0; i < group.length; i++) {
    const qty = Number((wfPrev.lineQtyByIndex || {})[String(i)] ?? group[i].orderQty);
    if (!Number.isFinite(qty) || qty <= 0) continue;
    addMovement(db, {
      type: "ADJUST",
      productCode: group[i].productCode,
      qty: Math.abs(qty),
      warehouse: String(group[i].warehouse || "").trim() || String(process.env.WMS_DEFAULT_OUTBOUND_WH || "").trim() || "유통사업부",
      partner: "",
      memo: `출고확정 취소 원복 ${memoSlip}`,
      slipNo: memoSlip,
      user: stockUser
    });
  }
  db.outboundOrderUpload.slipWorkflow[key] = { ...wfPrev, status: "draft", confirmedAt: undefined };
}

function computeOutboundConfirmQMerged(group, bodyLineQtyByIndex) {
  const bodyLineQtyByIndexSafe = bodyLineQtyByIndex && typeof bodyLineQtyByIndex === "object" ? bodyLineQtyByIndex : {};
  const hasBodyLineQtyByIndex = Object.keys(bodyLineQtyByIndexSafe).length > 0;
  if (hasBodyLineQtyByIndex) {
    return mergeLineQtyByIndexFromBody(bodyLineQtyByIndexSafe, group.map((x) => ({ qty: x.orderQty })));
  }
  const out = {};
  for (let i = 0; i < group.length; i++) {
    const fixedRaw = group[i]?.fixedQty;
    const fixed =
      fixedRaw != null && String(fixedRaw).trim() !== ""
        ? Number(String(fixedRaw).replace(/,/g, ""))
        : Number(group[i]?.orderQty || 0);
    const unship = String(group[i]?.unshipStatus || "").trim() === "미출";
    const qty = unship ? 0 : Number.isFinite(fixed) ? Math.max(0, Math.trunc(fixed)) : 0;
    out[String(i)] = qty;
  }
  return out;
}

/** 단건 출고확정(전표 상세·일괄 API 공통). db 객체를 직접 수정한다. */
function applyOutboundWorkflowConfirm(db, slipNoInput, stockUserInput, bodyLineQtyByIndexOpt) {
  const slipNo = String(slipNoInput || "").trim();
  const key = normalizeSlipNoText(slipNo);
  const group = outboundUploadLinesToDetailRows(db.outboundOrderUpload.lines || [], key);
  if (!group.length) throw new Error(`해당 전표의 품목 행이 없습니다: ${slipNo}`);
  const wfPrev = db.outboundOrderUpload.slipWorkflow[key] || { status: "draft", lineQtyByIndex: {} };
  const qMerged = computeOutboundConfirmQMerged(group, bodyLineQtyByIndexOpt);

  if (wfPrev.status === "confirmed") throw new Error("이미 출고확정된 전표입니다.");
  if (wfPrev.sentAt) throw new Error("이카운트 전송 완료 전표는 확정취소 후 다시 처리하세요.");
  const stockUser =
    String(stockUserInput || "").trim() || String(group[0].manager || "").trim() || String(db.managers[0] || "");
  if (!db.managers.includes(stockUser)) throw new Error("출고 담당자를 선택하세요.");

  const now = new Date();
  const saveYmd = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}${String(now.getDate()).padStart(2, "0")}`;
  const toCenterMergedSlipNo = (row) => {
    const centerRaw = String(row?.centerName || "").trim();
    const centerBucket = resolveCenterBucket(row);
    const centerKey = (centerRaw || centerBucket || "센터미정").replace(/\s+/g, "");
    const storeInRaw = String(row?.storeInDate || row?.dueDate || row?.orderDate || row?.poDate || row?.centerInDate || "")
      .trim()
      .replace(/-/g, "");
    const storeInCompact = normalizeDateCompact(storeInRaw) || "NODATE";
    return `${centerKey}_${saveYmd}_${storeInCompact}`;
  };

  let confirmedLineCount = 0;
  for (let i = 0; i < group.length; i++) {
    const qty = Number(qMerged[String(i)]);
    if (!Number.isFinite(qty) || qty <= 0) continue;
    confirmedLineCount += 1;
    const centerMergedSlipNo = toCenterMergedSlipNo(group[i]);
    addMovement(db, {
      type: "OUT",
      productCode: group[i].productCode,
      qty,
      warehouse: String(group[i].warehouse || "").trim() || String(process.env.WMS_DEFAULT_OUTBOUND_WH || "").trim() || "유통사업부",
      partner: String(group[i].partner || "").trim(),
      memo: "",
      // 출고 이력 전표번호는 센터 통합 전표번호를 기록
      slipNo: centerMergedSlipNo,
      user: stockUser
    });
  }
  if (!confirmedLineCount) throw new Error("확정할 출고확인 수량이 0입니다. 정상으로 변경 후 다시 시도하세요.");

  db.outboundOrderUpload.slipWorkflow[key] = {
    ...wfPrev,
    status: "confirmed",
    lineQtyByIndex: qMerged,
    confirmedAt: new Date().toISOString()
  };

  const partnerTypeForRecord = String(group[0]?.partnerType || "").trim();
  if (partnerTypeForRecord) {
    db.outboundOrderUpload.confirmedLists[partnerTypeForRecord] = db.outboundOrderUpload.confirmedLists[partnerTypeForRecord] || {};
    const recMap = db.outboundOrderUpload.confirmedLists[partnerTypeForRecord];
    const recordGroups = {};
    for (let i = 0; i < group.length; i++) {
      const qty = Number(qMerged[String(i)]);
      if (!Number.isFinite(qty) || qty <= 0) continue;
      const recKey = toCenterMergedSlipNo(group[i]);
      if (!recordGroups[recKey]) recordGroups[recKey] = [];
      recordGroups[recKey].push({
        slipNo: recKey,
        sourceSlipNo: key,
        productCode: group[i].productCode,
        sourceProductCode: String(group[i].sourceProductCode || "").trim(),
        mappedProductCode: String(group[i].mappedProductCode || "").trim(),
        productName: group[i].productName || "",
        barcode: String(group[i].barcode || "").trim(),
        qty,
        warehouse: String(group[i].warehouse || "").trim(),
        lot: String(group[i].lot || "").trim(),
        orderQty: String(group[i].orderQty || "").trim(),
        fixedQty: String(group[i].fixedQty || "").trim(),
        unshipStatus: String(group[i].unshipStatus || "").trim(),
        unitPrice: String(group[i].unitPrice || "").trim(),
        supplyAmt: String(group[i].supplyAmt || "").trim(),
        vatAmt: String(group[i].vatAmt || "").trim(),
        totalAmt: String(group[i].totalAmt || "").trim()
      });
    }

    for (const [recKey, lineEntries] of Object.entries(recordGroups)) {
      const centerNameFromKey = recKey.split("_")[0] || "";
      const storeInFromKey = recKey.split("_")[2] || "";
      if (!recMap[recKey]) {
        recMap[recKey] = {
          key: recKey,
          partnerType: partnerTypeForRecord,
          centerName: centerNameFromKey,
          saveDateYmd: saveYmd,
          storeInDateYmd: storeInFromKey,
          slipNos: [],
          lines: [],
          updatedAt: now.toISOString()
        };
      }
      recMap[recKey].lines = (recMap[recKey].lines || []).filter((ln) => ln && ln.sourceSlipNo !== key);
      if (!recMap[recKey].slipNos.includes(key)) recMap[recKey].slipNos.push(key);
      recMap[recKey].lines.push(...lineEntries);
      recMap[recKey].updatedAt = now.toISOString();
    }
  }
}

function resolveCenterBucket(row) {
  const name = String(row && row.centerName ? row.centerName : "").trim();
  const code = String(row && row.centerCode ? row.centerCode : "").trim();
  if (name.includes("여주") || code === "9120") return "여주";
  if (name.includes("대구") || code === "9100") return "대구";
  if (name.includes("시화") || code === "9110") return "시화";
  return "";
}

function slipWorkflowPreservePurchaseFlags(wfPrev) {
  const o = wfPrev && typeof wfPrev === "object" ? wfPrev : {};
  return {
    ecountPurchaseSavedAt: o.ecountPurchaseSavedAt,
    wmsStockInboundAt: o.wmsStockInboundAt
  };
}

function removeInboundSlipConfirmLogByKey(db, slipKeyNorm) {
  if (!Array.isArray(db.inboundSlipConfirmLog)) return;
  db.inboundSlipConfirmLog = db.inboundSlipConfirmLog.filter(
    (e) => !e || normalizeSlipNoText(e.slipNo) !== slipKeyNorm
  );
}

function resolveInboundPlanStockUser(db, body, group) {
  const fromBody = String(body.stockUser || "").trim();
  if (fromBody) {
    if (!db.managers.includes(fromBody)) throw new Error(`WMS 입고 담당자 "${fromBody}"는 등록된 담당자가 아닙니다.`);
    return fromBody;
  }
  const fromEnv = String(process.env.INBOUND_PLAN_STOCK_USER || "").trim();
  if (fromEnv) {
    if (!db.managers.includes(fromEnv)) throw new Error(`INBOUND_PLAN_STOCK_USER("${fromEnv}")는 등록된 담당자가 아닙니다.`);
    return fromEnv;
  }
  const mgr = String(group[0]?.manager || "").trim();
  if (mgr && db.managers.includes(mgr)) return mgr;
  throw new Error(
    "WMS 입고 담당자를 지정해 주세요. 전표 상세의 「WMS 입고 담당자」에서 선택하거나 .env에 INBOUND_PLAN_STOCK_USER=등록된담당자명 을 설정하세요."
  );
}

function resolveInboundPlanRollbackUser(db, body, group, wfPrev) {
  const fromBody = String(body && body.stockUser ? body.stockUser : "").trim();
  if (fromBody && db.managers.includes(fromBody)) return fromBody;
  const fromWorkflow = String(wfPrev && wfPrev.stockUser ? wfPrev.stockUser : "").trim();
  if (fromWorkflow && db.managers.includes(fromWorkflow)) return fromWorkflow;
  const fromEnv = String(process.env.INBOUND_PLAN_STOCK_USER || "").trim();
  if (fromEnv && db.managers.includes(fromEnv)) return fromEnv;
  const mgr = String(group[0]?.manager || "").trim();
  if (mgr && db.managers.includes(mgr)) return mgr;
  if (Array.isArray(db.managers) && db.managers.length) return String(db.managers[0]).trim();
  throw new Error("등록된 담당자가 없어 재고 원복을 처리할 수 없습니다. 기본정보 > 담당자를 먼저 등록하세요.");
}

function resolveInboundPlanWarehouseName(db, group) {
  const wh = String(group[0]?.whName || "").trim();
  if (wh) {
    if (!db.warehouses.includes(wh)) {
      throw new Error(`WMS에 등록되지 않은 창고입니다: "${wh}". 마스터에서 동일한 창고명으로 등록하세요.`);
    }
    return wh;
  }
  const fallback = String(process.env.WMS_DEFAULT_INBOUND_WH || "").trim();
  if (fallback) {
    if (!db.warehouses.includes(fallback)) {
      throw new Error(`WMS_DEFAULT_INBOUND_WH("${fallback}")는 등록된 창고가 아닙니다.`);
    }
    return fallback;
  }
  throw new Error(
    "입고창고(엑셀)가 비어 있습니다. 엑셀에 창고명을 채우거나 .env에 WMS_DEFAULT_INBOUND_WH=등록된창고명 을 설정하세요."
  );
}

function assertUploadSlipWmsInboundReady(db, group, lineQtyByIndex) {
  const warehouse = resolveInboundPlanWarehouseName(db, group);
  const qtyMap = lineQtyByIndex && typeof lineQtyByIndex === "object" ? lineQtyByIndex : {};
  for (let idx = 0; idx < group.length; idx++) {
    const row = group[idx];
    const productCode = String(row.itemCode || "").trim();
    const qRaw = qtyMap[String(idx)] != null ? qtyMap[String(idx)] : row.qty;
    const n = Math.round(Number(String(qRaw ?? "").replace(/,/g, "")));
    if (!Number.isFinite(n) || n <= 0) {
      throw new Error(`행 ${idx + 1}: 입고확인 수량을 1 이상으로 입력하세요.`);
    }
    const product = db.products.find((p) => p.code === productCode);
    if (!product) throw new Error(`WMS에 없는 상품코드입니다: ${productCode} (행 ${idx + 1}). 상품 마스터를 등록하세요.`);
    const usedWarehouses = toTagList(product.usedWarehouses);
    if (usedWarehouses.length && !usedWarehouses.includes(warehouse)) {
      throw new Error(
        `상품 ${productCode}: 창고 "${warehouse}"에서 입고할 수 없습니다. 사용창고: ${usedWarehouses.join(", ")}`
      );
    }
  }
  return warehouse;
}

function applyWmsStockInboundFromUploadSlip(db, group, lineQtyByIndex, slipNoDisplay, user) {
  const warehouse = resolveInboundPlanWarehouseName(db, group);
  const qtyMap = lineQtyByIndex && typeof lineQtyByIndex === "object" ? lineQtyByIndex : {};
  for (let idx = 0; idx < group.length; idx++) {
    const row = group[idx];
    const productCode = String(row.itemCode || "").trim();
    const qRaw = qtyMap[String(idx)] != null ? qtyMap[String(idx)] : row.qty;
    const n = Math.round(Number(String(qRaw ?? "").replace(/,/g, "")));
    if (!Number.isFinite(n) || n <= 0) {
      throw new Error(`WMS 입고 수량 오류(행 ${idx + 1}, 품목 ${productCode || "-"})`);
    }
    addMovement(db, {
      type: "IN",
      productCode,
      qty: n,
      warehouse,
      partner: String(row.vendor || "").trim() || String(group[0]?.vendor || "").trim(),
      memo: "",
      slipNo: String(slipNoDisplay || "").trim(),
      user
    });
  }
}

function rollbackWmsStockInboundFromUploadSlip(db, group, lineQtyByIndex, slipNoDisplay, user) {
  const warehouse = resolveInboundPlanWarehouseName(db, group);
  const qtyMap = lineQtyByIndex && typeof lineQtyByIndex === "object" ? lineQtyByIndex : {};
  const memoBase = `입고확정 취소 원복 ${String(slipNoDisplay || "").trim()}`;
  for (let idx = 0; idx < group.length; idx++) {
    const row = group[idx];
    const productCode = String(row.itemCode || "").trim();
    const qRaw = qtyMap[String(idx)] != null ? qtyMap[String(idx)] : row.qty;
    const n = Math.round(Number(String(qRaw ?? "").replace(/,/g, "")));
    if (!Number.isFinite(n) || n <= 0) {
      throw new Error(`재고 원복 수량 오류(행 ${idx + 1}, 품목 ${productCode || "-"})`);
    }
    addMovement(db, {
      type: "ADJUST",
      productCode,
      qty: -Math.abs(n),
      warehouse,
      partner: "",
      memo: memoBase,
      slipNo: String(slipNoDisplay || "").trim(),
      user
    });
  }
}

/** 입고예정목록 2 업로드 전표 → SavePurchases (확정 상태에서만 호출) */
function buildSavePurchasesPayloadFromUploadSlip(rawLinesSorted, lineQtyByIndex) {
  const rows = Array.isArray(rawLinesSorted) ? rawLinesSorted : [];
  if (!rows.length) throw new Error("품목 행이 없습니다.");
  const first = rows[0];
  const cust = String(first.vendorCode || "").trim();
  const custDes = String(first.vendor || "").trim();
  const whCd = String(process.env.ECOUNT_DEFAULT_WH_CD || "").trim();
  const ioDate =
    ymdOrCompactForPurchasesSave(first.poDate) ||
    ymdOrCompactForPurchasesSave(slipDateFromEcountSlipCell(first.slipNo)) ||
    normalizeDateCompact(new Date().toISOString().slice(0, 10));
  if (!cust) throw new Error("거래처코드(엑셀 거래처코드)가 없습니다.");
  if (!whCd) {
    throw new Error(
      "창고코드(WH_CD)가 필요합니다. .env에 ECOUNT_DEFAULT_WH_CD=이카운트창고코드 를 설정하세요. (엑셀은 창고명만 있는 경우가 많습니다.)"
    );
  }

  const qtyMap = lineQtyByIndex && typeof lineQtyByIndex === "object" ? lineQtyByIndex : {};
  const PurchasesList = rows.map((row, idx) => {
    const qRaw = qtyMap[String(idx)] != null ? qtyMap[String(idx)] : row.qty;
    const n = Number(String(qRaw ?? "").replace(/,/g, ""));
    if (!Number.isFinite(n) || n <= 0) {
      throw new Error(`수량이 올바르지 않습니다(행 ${idx + 1}, 품목 ${row.itemCode || "-"}).`);
    }
    const bulk = {
      IO_DATE: ioDate,
      CUST: cust,
      CUST_DES: custDes,
      WH_CD: whCd,
      PROD_CD: String(row.itemCode || "").trim(),
      PROD_DES: String(row.itemName || "").trim(),
      QTY: String(Math.round(n))
    };
    const due = ymdOrCompactForPurchasesSave(row.dueDate);
    if (due) bulk.TIME_DATE = due;
    const ref = String(row.remark || "").trim();
    if (ref) bulk.REF_DES = ref;
    const memo = String(row.note || "").trim();
    if (memo) bulk.MEMO = memo;
    const emp = String(row.manager || "").trim();
    if (emp) bulk.EMP_CD = emp;
    return { BulkDatas: bulk };
  });
  return { PurchasesList };
}

/** 출고 확정 묶음 → Sale/SaveSale 본문 (SaleList[].BulkDatas). */
function resolveEcountOutboundProdForSale(db, partnerType, productCodeRaw, productNameFallback) {
  const src = String(productCodeRaw || "").trim();
  if (!src) throw new Error("상품코드가 비어 있습니다.");
  const nameFb = String(productNameFallback || "").trim();
  const codeMap = getOutboundCodeMasterMap(db, partnerType);
  const mapped = codeMap.get(src);
  if (mapped) {
    return { prodCd: mapped, prodDes: nameFb };
  }
  const p = findProductByMovementImportCode(db, src);
  if (p) {
    const prodCd = String(p.ecountCode || p.code || "").trim();
    const prodDes = String(p.name || p.ecountName || nameFb || "").trim();
    if (!prodCd) throw new Error(`품목 ${src}: 이카운트 품목코드(ecountCode)가 없습니다.`);
    return { prodCd, prodDes };
  }
  return { prodCd: src, prodDes: nameFb };
}

function outboundLineWarehouseToEcountWhCd(lineWarehouse) {
  const forced = String(process.env.ECOUNT_OUTBOUND_SALE_WH_CD || "").trim();
  if (forced) return forced;
  const w = String(lineWarehouse || "").trim();
  if (w && w.length <= 8 && !/사업부|창고|센터|물류|본사/.test(w)) return w;
  const fallback = String(process.env.ECOUNT_DEFAULT_WH_CD || "").trim();
  if (!fallback) {
    throw new Error(
      "판매입력 창고코드(WH_CD)가 필요합니다. .env에 ECOUNT_OUTBOUND_SALE_WH_CD 또는 ECOUNT_DEFAULT_WH_CD 를 설정하세요."
    );
  }
  return fallback;
}

function buildSaveSalePayloadFromConfirmRecord(db, record, partnerType) {
  const lines = Array.isArray(record.lines) ? record.lines : [];
  if (!lines.length) throw new Error("확정 품목 행이 없습니다.");
  const ioDate =
    String(record.storeInDateYmd || "").trim().replace(/-/g, "") ||
    String(record.saveDateYmd || "").trim().replace(/-/g, "") ||
    normalizeDateCompact(new Date().toISOString().slice(0, 10));
  if (!ioDate || ioDate.length !== 8) throw new Error("확정일자(saveDateYmd)가 올바르지 않습니다.");
  const custByPartner = {
    emart: String(process.env.ECOUNT_OUTBOUND_SALE_CUST_EMART || "2068650913").trim(),
    daiso: String(process.env.ECOUNT_OUTBOUND_SALE_CUST_DAISO || "").trim(),
    lotte: String(process.env.ECOUNT_OUTBOUND_SALE_CUST_LOTTE || "").trim()
  };
  const cust = String(custByPartner[partnerType] || process.env.ECOUNT_OUTBOUND_SALE_CUST || "").trim();
  const custDes = String(process.env.ECOUNT_OUTBOUND_SALE_CUST_DES || "").trim();
  const empCd = String(process.env.ECOUNT_OUTBOUND_SALE_EMP_CD || "").trim();
  const listKey = String(record.key || "").trim();
  // 기본은 자동번호(빈 문자열) 사용. 환경변수로만 고정 전표번호를 명시 허용.
  const docNo = String(process.env.ECOUNT_OUTBOUND_SALE_DOC_NO || "").trim();
  const centerRemark = String(record.centerName || "").trim();
  const remarks = (centerRemark || `WMS ${partnerType} ${listKey}`).slice(0, 200);
  const ioType = String(process.env.ECOUNT_OUTBOUND_SALE_IO_TYPE || "").trim();
  const price = String(process.env.ECOUNT_OUTBOUND_SALE_PRICE ?? "0").trim() || "0";
  const codeMap = getOutboundCodeMasterMap(db, partnerType);

  const num = (v) => {
    if (v === "" || v === null || v === undefined) return NaN;
    const n = Number(String(v).replace(/[^\d.-]/g, "").trim());
    return Number.isFinite(n) ? n : NaN;
  };

  const agg = new Map();
  for (const ln of lines) {
    const qty = Number(ln.qty);
    if (!Number.isFinite(qty) || qty <= 0) continue;
    // WMS 상품마스터 + 코드마스터를 함께 사용해 PROD_CD를 확정한다.
    const rawProductCode = String(ln.productCode || "").trim();
    const sourceProductCode = String(ln.sourceProductCode || "").trim();
    const barcodeCode = String(ln.barcode || "").trim();
    const candidates = [
      rawProductCode,
      sourceProductCode,
      barcodeCode,
      rawProductCode.replace(/-\d+$/, ""),
      sourceProductCode.replace(/-\d+$/, ""),
      barcodeCode.replace(/-\d+$/, "")
    ]
      .map((x) => String(x || "").trim())
      .filter(Boolean);

    let p = null;
    for (const c of candidates) {
      p = findProductByMovementImportCode(db, c);
      if (p) break;
    }

    let mappedCode = "";
    for (const c of candidates) {
      const hit = codeMap.get(c) || codeMap.get(normalizeInputLoose(c)) || codeMap.get(c.replace(/\.0+$/, ""));
      if (hit) {
        mappedCode = String(hit).trim();
        break;
      }
    }

    const productCodeResolved = String((p && (p.ecountCode || p.code)) || mappedCode || "").trim();
    if (!productCodeResolved) {
      throw new Error(
        `전송용 품목코드를 찾지 못했습니다. 원본:${rawProductCode || "-"} / SKU:${sourceProductCode || "-"} / 바코드:${barcodeCode || "-"}`
      );
    }
    const prodCd = String(productCodeResolved).trim();
    const prodDes = String(ln.productName || p?.name || p?.ecountName || "").trim();
    if (!prodDes) {
      throw new Error(
        `품목 ${String(ln.productCode || "").trim()}: 이카운트 판매입력에 필요한 품목명(PROD_DES)이 없습니다.`
      );
    }
    const whCd = outboundLineWarehouseToEcountWhCd(ln.warehouse);
    const unitPrice = num(ln.unitPrice);
    const supplyAmt = num(ln.supplyAmt);
    const vatAmt = num(ln.vatAmt);
    const totalAmt = num(ln.totalAmt);
    // 바코드는 PROD_CD로 전달하고, ITEM_CD/P_REMARKS1에는 넣지 않음.
    const itemCd = "";
    const k = `${prodCd}\t${whCd}\t${Number.isFinite(unitPrice) ? unitPrice : ""}\t${itemCd}`;
    const prev = agg.get(k);
    const supplyResolved = Number.isFinite(supplyAmt)
      ? supplyAmt
      : Number.isFinite(totalAmt) && Number.isFinite(vatAmt)
        ? totalAmt - vatAmt
        : Number.isFinite(totalAmt) && !Number.isFinite(vatAmt)
          ? totalAmt
        : Number.isFinite(unitPrice)
          ? unitPrice * qty
          : NaN;
    const vatResolved =
      Number.isFinite(vatAmt)
        ? vatAmt
        : Number.isFinite(totalAmt) && Number.isFinite(supplyResolved)
          ? totalAmt - supplyResolved
          : Number.isFinite(supplyResolved)
            ? Math.round(supplyResolved * 0.1)
            : 0;
    const totalResolved = Number.isFinite(totalAmt) ? totalAmt : Number.isFinite(supplyResolved) ? supplyResolved + vatResolved : NaN;
    if (prev) {
      prev.qty += qty;
      if (Number.isFinite(supplyResolved)) prev.supplyAmt += supplyResolved;
      if (Number.isFinite(vatResolved)) prev.vatAmt += vatResolved;
      if (Number.isFinite(totalResolved)) prev.totalAmt += totalResolved;
    } else {
      agg.set(k, {
        prodCd,
        prodDes,
        whCd,
        qty,
        unitPrice: Number.isFinite(unitPrice) ? unitPrice : NaN,
        supplyAmt: Number.isFinite(supplyResolved) ? supplyResolved : NaN,
        vatAmt: Number.isFinite(vatResolved) ? vatResolved : NaN,
        totalAmt: Number.isFinite(totalResolved) ? totalResolved : NaN,
        itemCd
      });
    }
  }
  if (!agg.size) throw new Error("전송할 수량이 있는 품목이 없습니다.");

  const SaleList = [...agg.values()].map((row) => ({
    BulkDatas: {
      IO_DATE: ioDate,
      UPLOAD_SER_NO: "",
      CUST: cust,
      CUST_DES: custDes,
      EMP_CD: empCd,
      WH_CD: row.whCd,
      IO_TYPE: ioType,
      EXCHANGE_TYPE: "",
      EXCHANGE_RATE: "",
      REMARKS: "",
      DOC_NO: docNo,
      PROD_CD: row.prodCd,
      PROD_DES: row.prodDes,
      QTY: String(Math.trunc(row.qty)),
      ITEM_CD: row.itemCd || "",
      PRICE: Number.isFinite(row.unitPrice) ? String(row.unitPrice) : price,
      SUPPLY_AMT: Number.isFinite(row.supplyAmt) ? String(Math.trunc(row.supplyAmt)) : "",
      VAT_AMT: Number.isFinite(row.vatAmt) ? String(Math.trunc(row.vatAmt)) : "",
      CUST_AMT: "",
      TTL_CTT: remarks,
      U_MEMO2: remarks,
      P_REMARKS1: "",
      REMARKS1: "",
      REMARKS2: "",
      REMARKS3: ""
    }
  }));
  return { SaleList };
}

function throwIfEcountSaveSaleFailed(result) {
  if (!result.ok) {
    throw new Error(`이카운트 Sale/SaveSale HTTP ${result.status}: ${String(result.text || "").slice(0, 500)}`);
  }
  const d = result.data;
  if (!d || typeof d !== "object") {
    throw new Error(`이카운트 응답 파싱 실패: ${String(result.text || "").slice(0, 400)}`);
  }
  const st = String(d.Status ?? "");
  if (st === "500" || st === "400") {
    const msg =
      (d.Error && d.Error.Message) ||
      (Array.isArray(d.Errors) && d.Errors[0] && d.Errors[0].Message) ||
      JSON.stringify(d.Errors || d.Error || {}).slice(0, 400);
    throw new Error(`이카운트 판매입력: ${msg}`);
  }
  if (st && st !== "200") {
    throw new Error(`이카운트 판매입력 Status=${st}: ${String(result.text || "").slice(0, 400)}`);
  }
  const data = d.Data;
  if (data && Number(data.FailCnt) > 0) {
    const details = Array.isArray(data.ResultDetails) ? data.ResultDetails : [];
    const msgs = details
      .filter((x) => x && x.IsSuccess === false)
      .map((x) =>
        x.TotalError ||
        (Array.isArray(x.Errors) ? x.Errors.map((e) => `${String(e.ColCd || "")}:${String(e.Message || "")}`).join("; ") : "")
      )
      .filter(Boolean);
    throw new Error(`이카운트 판매입력 라인 오류: ${msgs.join(" | ") || JSON.stringify(data).slice(0, 500)}`);
  }
}

function isOnlyItemCdRegistrationError(ecountData) {
  const details = Array.isArray(ecountData?.Data?.ResultDetails) ? ecountData.Data.ResultDetails : [];
  if (!details.length) return false;
  let hasAny = false;
  for (const d of details) {
    if (!d || d.IsSuccess === true) continue;
    const errs = Array.isArray(d.Errors) ? d.Errors : [];
    if (!errs.length) return false;
    for (const e of errs) {
      const col = String(e?.ColCd || "").trim().toUpperCase();
      const msg = String(e?.Message || "");
      if (col !== "ITEM_CD" && !/바코드|item.?cd|미등록코드/i.test(msg)) return false;
      hasAny = true;
    }
  }
  return hasAny;
}

function clearItemCdInSalePayload(payload) {
  const src = payload && typeof payload === "object" ? payload : {};
  const list = Array.isArray(src.SaleList) ? src.SaleList : [];
  const next = list.map((row) => {
    const bulk = row && row.BulkDatas && typeof row.BulkDatas === "object" ? { ...row.BulkDatas } : {};
    bulk.ITEM_CD = "";
    return { ...row, BulkDatas: bulk };
  });
  return { ...src, SaleList: next };
}

async function postEcountSaveSale(salePayload) {
  const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();
  if (!comCode) throw new Error("ECOUNT_COM_CODE가 필요합니다.");
  const sessionId = await fetchEcountSession();
  const apiBase = await fetchEcountApiBase(comCode);
  return postEcountOapiV2(apiBase, sessionId, "Sale/SaveSale", salePayload);
}

function findOutboundConfirmListKeyForSlip(db, partnerType, slipNoInput) {
  const key = normalizeSlipNoText(slipNoInput);
  if (!key) return "";
  const recMap = db.outboundOrderUpload.confirmedLists && db.outboundOrderUpload.confirmedLists[partnerType];
  if (!recMap || typeof recMap !== "object") return "";
  for (const [listKey, rec] of Object.entries(recMap)) {
    const nums = Array.isArray(rec && rec.slipNos) ? rec.slipNos : [];
    for (const s of nums) {
      if (normalizeSlipNoText(String(s || "")) === key) return String(listKey || "").trim();
    }
  }
  return "";
}

/**
 * 확정리스트 묶음 단위로 이카운트 Sale/SaveSale 호출 후 slipWorkflow 전송 처리.
 * @returns {{ ok: boolean, sentNow: number, total: number, ecount?: object, alreadyComplete?: boolean, salePayload?: object }}
 */
async function runOutboundConfirmListSalesToEcount(db, partnerType, listKey) {
  db.outboundOrderUpload.confirmedLists = db.outboundOrderUpload.confirmedLists || {};
  db.outboundOrderUpload.confirmedLists[partnerType] = db.outboundOrderUpload.confirmedLists[partnerType] || {};
  const recMap = db.outboundOrderUpload.confirmedLists[partnerType];
  const record = recMap[listKey];
  if (!record) throw new Error("해당 확정 저장 내역을 찾을 수 없습니다.");

  const slipKeys = Array.isArray(record.slipNos)
    ? [...new Set(record.slipNos.map((s) => normalizeSlipNoText(String(s || ""))).filter(Boolean))]
    : [];
  if (!slipKeys.length) throw new Error("해당 확정 묶음에 전표가 없습니다.");

  for (const sk of slipKeys) {
    const wfPrev = db.outboundOrderUpload.slipWorkflow[sk] || { status: "draft", lineQtyByIndex: {} };
    if (!["confirmed", "sent"].includes(String(wfPrev.status || ""))) {
      throw new Error(`출고확정 후에 판매입력을 전송할 수 있습니다: ${sk}`);
    }
  }

  let salePayload = buildSaveSalePayloadFromConfirmRecord(db, record, partnerType);
  let result = await postEcountSaveSale(salePayload);
  let retriedWithoutItemCd = false;
  const allowEmptyItemCdRetry = ["1", "true", "yes", "on"].includes(
    String(process.env.ECOUNT_OUTBOUND_SALE_ITEMCD_RETRY_EMPTY || "").trim().toLowerCase()
  );
  if (allowEmptyItemCdRetry && isOnlyItemCdRegistrationError(result.data)) {
    salePayload = clearItemCdInSalePayload(salePayload);
    result = await postEcountSaveSale(salePayload);
    retriedWithoutItemCd = true;
  }
  throwIfEcountSaveSaleFailed(result);

  let sentNow = 0;
  for (const sk of slipKeys) {
    const wfPrev = db.outboundOrderUpload.slipWorkflow[sk] || { status: "draft", lineQtyByIndex: {} };
    db.outboundOrderUpload.slipWorkflow[sk] = { ...wfPrev, status: "sent", sentAt: new Date().toISOString() };
    sentNow += 1;
  }

  db.outboundOrderUpload.uploadedAt = new Date().toISOString();
  writeDb(db);
  return {
    ok: true,
    sentNow,
    total: slipKeys.length,
    ecount: result.data,
    salePayload,
    retriedWithoutItemCd
  };
}

/** 이카운트 발주서현황 엑셀의 「일자-No.」에서 날짜 부분 → YYYY-MM-DD */
function slipDateFromEcountSlipCell(slipRaw) {
  const s = String(slipRaw || "").trim();
  const m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (!m) return "";
  return `${m[1]}-${m[2]}-${m[3]}`;
}

function splitItemNameSpecFromEcountCell(raw) {
  const s = String(raw || "").trim();
  const m = s.match(/^(.+?)\s*\[([^\]]*)\]\s*$/);
  if (m) return { itemName: m[1].trim(), spec: m[2].trim() };
  return { itemName: s, spec: "" };
}

function numOrEmpty(v) {
  if (v === "" || v === null || v === undefined) return "";
  const n = Number(String(v).replace(/,/g, ""));
  return Number.isFinite(n) ? n : "";
}

function normalizeDateTimeDigits(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  return s.replace(/[^\d]/g, "");
}

function pickLatestDateTimeText(values) {
  const list = Array.isArray(values) ? values : [];
  let bestText = "";
  let bestKey = "";
  for (const raw of list) {
    const txt = String(raw || "").trim();
    if (!txt) continue;
    const key = normalizeDateTimeDigits(txt);
    if (!key) continue;
    if (!bestKey || key > bestKey) {
      bestKey = key;
      bestText = txt;
    }
  }
  return bestText;
}

/**
 * 이카운트 「발주서현황」 내보내기 시트(1행 제목·2행 헤더) → 라인 객체 배열
 * @param {unknown[][]} matrix sheet_to_json(..., { header:1 }) 결과
 */
function parseInboundPlanUploadMatrix(matrix) {
  const rows = Array.isArray(matrix) ? matrix : [];
  let headerRow = -1;
  for (let i = 0; i < Math.min(rows.length, 80); i++) {
    const r = rows[i] || [];
    const c0 = String(r[0] ?? "")
      .replace(/\s+/g, " ")
      .trim();
    const c1 = String(r[1] ?? "").trim();
    const headLike =
      (c0.includes("일자") && c1 === "품목코드") ||
      c0 === "일자-No." ||
      /^일자[-–]\s*No\.?$/i.test(c0);
    if (headLike && (c1 === "품목코드" || c1.includes("품목코드"))) {
      headerRow = i;
      break;
    }
  }
  if (headerRow < 0) {
    throw new Error("발주서현황 양식을 찾지 못했습니다. 1행 제목·2행에 「일자-No.」「품목코드」 헤더가 있는지 확인하세요.");
  }
  const H = (rows[headerRow] || []).map((c) => String(c ?? "").trim());
  const col = (...labels) => {
    for (const lb of labels) {
      const j = H.indexOf(lb);
      if (j >= 0) return j;
    }
    return -1;
  };
  const iSlip = col("일자-No.");
  const iCode = col("품목코드");
  const iBc = col("바코드");
  const iWh = col("창고명");
  const iName = col("품목명[규격]");
  const iQty = col("수량");
  const iMgr = col("담당자");
  const iVc = col("거래처코드");
  const iVn = col("거래처");
  const iDue = col("납기일자");
  const iPrice = col("단가");
  const iAmt = col("공급가액");
  const iRef = col("적요");
  const iNote = col("비고");
  const iLastModified = col("최종수정일자", "최종수정일시", "수정일자", "수정일시", "수정일");
  if (iSlip < 0 || iCode < 0) {
    throw new Error("필수 열(일자-No., 품목코드)을 찾지 못했습니다.");
  }
  const lines = [];
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const slipRaw = String(row[iSlip] ?? "").trim();
    const itemCode = String(row[iCode] ?? "").trim();
    if (!slipRaw && !itemCode) continue;
    if (!slipRaw || !itemCode) continue;
    const { itemName, spec } = splitItemNameSpecFromEcountCell(iName >= 0 ? row[iName] : "");
    const poDate = slipDateFromEcountSlipCell(slipRaw);
    lines.push({
      slipNo: slipRaw.replace(/\s+/g, " ").trim(),
      poDate,
      itemCode,
      barcode: iBc >= 0 ? String(row[iBc] ?? "").trim() : "",
      whName: iWh >= 0 ? String(row[iWh] ?? "").trim() : "",
      itemName,
      spec,
      qty: numOrEmpty(iQty >= 0 ? row[iQty] : ""),
      manager: iMgr >= 0 ? String(row[iMgr] ?? "").trim() : "",
      vendorCode: iVc >= 0 ? String(row[iVc] ?? "").trim() : "",
      vendor: iVn >= 0 ? String(row[iVn] ?? "").trim() : "",
      dueDate: iDue >= 0 ? String(row[iDue] ?? "").trim().replace(/\s+$/, "") : "",
      unitPrice: numOrEmpty(iPrice >= 0 ? row[iPrice] : ""),
      supplyAmount: numOrEmpty(iAmt >= 0 ? row[iAmt] : ""),
      remark: iRef >= 0 ? String(row[iRef] ?? "").trim() : "",
      note: iNote >= 0 ? String(row[iNote] ?? "").trim() : "",
      lastModifiedAt: iLastModified >= 0 ? String(row[iLastModified] ?? "").trim() : ""
    });
  }
  if (!lines.length) throw new Error("데이터 행이 없습니다. 헤더 아래에 품목 행이 있는지 확인하세요.");
  return lines;
}

/**
 * 저장소에 이미 있는 발주번호(일자-No. 정규화 키)와 동일한 전표는 파일 쪽 행을 버림.
 * 없는 발주번호의 품목 행만 기존 lines 뒤에 붙임 — 기존 전표 내용은 덮어쓰지 않음.
 */
function mergeInboundPlanUploadByNewSlipsOnly(existingLines, incomingLines) {
  const existing = Array.isArray(existingLines) ? existingLines : [];
  const incoming = Array.isArray(incomingLines) ? incomingLines : [];
  const existingSlipKeys = new Set();
  for (const row of existing) {
    const k = normalizeSlipNoText(row.slipNo);
    if (k) existingSlipKeys.add(k);
  }
  const added = [];
  let skippedLineCount = 0;
  const skippedSlipKeys = new Set();
  for (const row of incoming) {
    const k = normalizeSlipNoText(row.slipNo);
    if (!k) continue;
    if (existingSlipKeys.has(k)) {
      skippedLineCount += 1;
      skippedSlipKeys.add(k);
      continue;
    }
    added.push(row);
  }
  const newSlipKeys = new Set();
  for (const row of added) {
    const k = normalizeSlipNoText(row.slipNo);
    if (k) newSlipKeys.add(k);
  }
  const merged = [...existing, ...added];
  return {
    merged,
    addedLineCount: added.length,
    newSlipCount: newSlipKeys.size,
    skippedLineCount,
    skippedExistingSlipCount: skippedSlipKeys.size,
    skippedSlipKeysSample: Array.from(skippedSlipKeys).slice(0, 40)
  };
}

function aggregateInboundPlanUploadLines(lines, slipWorkflow = {}) {
  const arr = Array.isArray(lines) ? lines : [];
  const wf = slipWorkflow && typeof slipWorkflow === "object" ? slipWorkflow : {};
  const by = new Map();
  for (const row of arr) {
    const k = normalizeSlipNoText(row.slipNo);
    if (!k) continue;
    if (!by.has(k)) by.set(k, []);
    by.get(k).push(row);
  }
  const items = [];
  for (const [slipKey, group] of by) {
    group.sort((a, b) => String(a.itemCode).localeCompare(String(b.itemCode)));
    const first = group[0];
    const w = wf[slipKey];
    const qMap = w && w.lineQtyByIndex && typeof w.lineQtyByIndex === "object" ? w.lineQtyByIndex : {};
    const totalQty = group.reduce((s, x, idx) => {
      const q = qMap[String(idx)] != null ? Number(qMap[String(idx)]) : Number(x.qty) || 0;
      return s + (Number.isFinite(q) ? q : 0);
    }, 0);
    const names = group.map((x) => x.itemName).filter(Boolean);
    const itemSummary = names.length <= 1 ? (names[0] || "") : `${names[0]} 외 ${names.length - 1}건`;
    items.push({
      poNo: first.slipNo,
      poDate: first.poDate || slipDateFromEcountSlipCell(first.slipNo),
      vendor: first.vendor || "",
      manager: first.manager || "",
      itemName: itemSummary,
      qty: totalQty,
      dueDate: first.dueDate || "",
      whName: first.whName || "",
      status: "업로드",
      lineCount: group.length,
      lastModifiedAt: pickLatestDateTimeText(group.map((x) => x.lastModifiedAt)),
      workflowStatus: w && w.status === "confirmed" ? "confirmed" : "draft"
    });
  }
  items.sort((a, b) => String(a.poNo).localeCompare(String(b.poNo), "ko"));
  return items;
}

function inboundPlanUploadLinesToDetailRows(lines, slipKeyNorm) {
  const group = (Array.isArray(lines) ? lines : []).filter((x) => normalizeSlipNoText(x.slipNo) === slipKeyNorm);
  group.sort((a, b) => String(a.itemCode).localeCompare(String(b.itemCode)));
  return group.map((x) => ({
    poNo: x.slipNo,
    poDate: x.poDate,
    vendor: x.vendor,
    manager: x.manager,
    itemCode: x.itemCode,
    itemName: x.itemName,
    spec: x.spec,
    barcode: x.barcode,
    boxQty: "",
    orderQty: x.qty,
    qty: x.qty,
    dueDate: x.dueDate,
    whName: x.whName,
    status: "",
    remark: x.remark,
    note: x.note,
    currency: "",
    unitPrice: x.unitPrice,
    supplyAmount: x.supplyAmount,
    vendorCode: x.vendorCode,
    lastModifiedAt: x.lastModifiedAt || "",
    detail: []
  }));
}

/** ECOUNT list row may embed line items under various keys; expand for detail view. */
const NESTED_LINE_ARRAY_KEYS = [
  "Details",
  "Detail",
  "DetailList",
  "DETAIL",
  "DETAILS",
  "ItemList",
  "Items",
  "Lines",
  "LineList",
  "DTL",
  "DTL_LIST",
  "ORDER_DETAIL",
  "ORDER_DETAILS",
  "PurchasesOrderDetail",
  "ResultDetail",
  "Datas"
];

function rowLooksLikeOrderLine(obj) {
  if (!obj || typeof obj !== "object") return false;
  if (
    pickFirstNonEmpty(obj, ["PROD_CD", "ITEM_CD", "LINE_NO", "LINE_NUM", "SER_NO", "SEQ", "LINE_SEQ"])
  ) {
    return true;
  }
  const hasName = Boolean(pickFirstNonEmpty(obj, ["PROD_DES", "ITEM_DES", "ITEM_NAME"]));
  const hasQty = Boolean(pickFirstNonEmpty(obj, ["QTY", "ORDER_QTY", "PUR_QTY", "PO_QTY"]));
  const hasOrd = Boolean(pickFirstNonEmpty(obj, ["ORD_NO", "IO_NO"]));
  return hasName && hasQty && !hasOrd;
}

function rowLooksLikeOrderHeader(obj) {
  if (!obj || typeof obj !== "object") return false;
  const hasOrd = Boolean(pickFirstNonEmpty(obj, ["ORD_NO", "IO_NO"]));
  const hasLineHint = rowLooksLikeOrderLine(obj);
  return hasOrd && !hasLineHint;
}

function extractNestedLineRowsFromOrderRow(orderRow, depth = 0) {
  if (!orderRow || typeof orderRow !== "object" || depth > 8) return [];

  for (const k of NESTED_LINE_ARRAY_KEYS) {
    const v = orderRow[k];
    if (Array.isArray(v) && v.length && v.every((x) => x && typeof x === "object")) {
      const lineish = v.filter((x) => rowLooksLikeOrderLine(x));
      if (lineish.length >= 1) return lineish;
    }
  }

  for (const k of Object.keys(orderRow)) {
    const v = orderRow[k];
    if (!Array.isArray(v) || v.length < 1) continue;
    if (!v.every((x) => x && typeof x === "object")) continue;
    const lineish = v.filter((x) => rowLooksLikeOrderLine(x));
    if (lineish.length >= 2 || (lineish.length === 1 && !rowLooksLikeOrderHeader(lineish[0]))) {
      return lineish;
    }
  }

  for (const k of Object.keys(orderRow)) {
    const v = orderRow[k];
    if (v && typeof v === "object" && !Array.isArray(v)) {
      const found = extractNestedLineRowsFromOrderRow(v, depth + 1);
      if (found.length) return found;
    }
  }

  return [];
}

function mergeOrderHeaderWithLines(headerRaw, lineRows) {
  const lines = Array.isArray(lineRows) ? lineRows : [];
  return lines.map((line) => ({ ...headerRaw, ...line }));
}

/** For Network debug: show which keys exist on the slip header (arrays = possible line containers). */
function buildHeaderDebugSummary(headerRaw, maxKeys = 45) {
  if (!headerRaw || typeof headerRaw !== "object") return null;
  const keys = Object.keys(headerRaw).slice(0, maxKeys);
  const summary = {};
  for (const k of keys) {
    const v = headerRaw[k];
    if (v === null || v === undefined) summary[k] = "null";
    else if (Array.isArray(v)) summary[k] = `array[${v.length}]`;
    else if (typeof v === "object") summary[k] = `object{${Object.keys(v).slice(0, 8).join(",")}}`;
    else summary[k] = typeof v;
  }
  return summary;
}

function findMatchingRawOrderRow(rawRows, targetSlipNorm, seq, fallbackDateCompact) {
  const rows = Array.isArray(rawRows) ? rawRows : [];
  for (const r of rows) {
    const poNoRaw = String(pickPurchaseRequestNo(r) || "").trim();
    const poDate = String(
      pickFirstNonEmpty(r, ["ORD_DATE", "PO_DATE", "PODATE", "ORDER_DATE", "DATE", "WRITE_DATE", "IO_DATE"])
    );
    const displayPo = formatOrderNoByDateAndSeq(poNoRaw, poDate, fallbackDateCompact);
    const norm = normalizeSlipNoText(displayPo);
    if (norm && norm === targetSlipNorm) return r;
    if (seq && norm.endsWith(`-${seq}`)) return r;
  }
  return null;
}

function extractArrayFromUnknown(payload) {
  if (Array.isArray(payload)) return payload;
  if (!payload || typeof payload !== "object") return [];

  const directCandidates = [
    payload.Datas,
    payload.Data,
    payload.List,
    payload.Result,
    payload.items,
    payload.Rows,
    payload.records
  ];
  for (const c of directCandidates) {
    if (Array.isArray(c)) return c;
  }
  for (const c of directCandidates) {
    if (c && typeof c === "object") {
      const nested = extractArrayFromUnknown(c);
      if (nested.length) return nested;
      // Some APIs return map-like objects instead of arrays.
      const vals = Object.values(c);
      if (vals.length && vals.every((v) => v && typeof v === "object")) return vals;
    }
  }
  return [];
}

/** 발주/구매전표 키로 상세·라인 조회 시도용 요청 바디 후보 (루트 vs ListParam). */
function buildLineLookupBodies(basePayload, slipNo, ordNoRaw, ioNoRaw) {
  const lp = basePayload.ListParam ? { ...basePayload.ListParam } : {};
  const out = [];
  const ord = String(ordNoRaw || "").trim();
  const io = String(ioNoRaw || "").trim();
  const slip = String(slipNo || "").trim();
  const add = (body) => out.push(body);

  if (ord) {
    add({ ...basePayload, ORD_NO: ord });
    add({ ...basePayload, ListParam: { ...lp, ORD_NO: ord } });
  }
  if (io) {
    add({ ...basePayload, IO_NO: io });
    add({ ...basePayload, ListParam: { ...lp, IO_NO: io } });
  }
  if (slip) {
    if (!ord) add({ ...basePayload, ORD_NO: slip });
    if (!ord) add({ ...basePayload, ListParam: { ...lp, ORD_NO: slip } });
    add({ ...basePayload, ListParam: { ...lp, DOC_NO: slip } });
    add({ ...basePayload, DOC_NO: slip });
    add({ ...basePayload, SLIP_NO: slip });
    add({ ...basePayload, ListParam: { ...lp, SLIP_NO: slip } });
  }

  const seen = new Set();
  return out.filter((b) => {
    const s = JSON.stringify(b);
    if (seen.has(s)) return false;
    seen.add(s);
    return true;
  });
}

function parseInboundLineApiPathSpecs(raw) {
  return String(raw || "")
    .split(/[,;]+/)
    .map((s) => s.trim())
    .filter(Boolean)
    .map((spec) => {
      const t = spec.replace(/^\/+/, "");
      if (!t.includes("/")) return `Purchases/${t}`;
      return t;
    });
}

/** ECOUNT 상세/라인 JSON → normalizeInboundPlanRows와 동일 규칙으로 병합. */
function matchDetailRowsToInboundPlans(data, rangeCompact) {
  const rowsCandidate = extractArrayFromUnknown(data);
  if (!Array.isArray(rowsCandidate) || !rowsCandidate.length) return { matched: [], kind: "empty" };
  if (rowsCandidate.length > 1 && rowsCandidate.every((row) => rowLooksLikeOrderLine(row))) {
    return {
      matched: normalizeInboundPlanRows(rowsCandidate, rangeCompact),
      kind: "multi-line"
    };
  }
  const nested = extractNestedLineRowsFromOrderRow(rowsCandidate[0]);
  if (nested.length) {
    const merged = mergeOrderHeaderWithLines(rowsCandidate[0], nested);
    return {
      matched: normalizeInboundPlanRows(merged, rangeCompact),
      kind: "nested"
    };
  }
  return {
    matched: normalizeInboundPlanRows(rowsCandidate, rangeCompact),
    kind: "flat"
  };
}

function filterLikelyItemLines(matched) {
  const rows = Array.isArray(matched) ? matched : [];
  return rows.filter((x) => {
    const hasCode = String(x.itemCode || "").trim() !== "";
    const hasQty = String(x.qty || "").trim() !== "";
    const isSummaryTitle = /외\s*\d+건/.test(String(x.itemName || ""));
    return (hasCode || hasQty) && !isSummaryTitle;
  });
}

/**
 * GetPurchases 등 내장 후보 API 일괄 호출을 할지 여부.
 * 기본은 호출하지 않음(대부분 존에서 404·412·긴 지연만 발생). 다시 켜려면 ECOUNT_SKIP_BUILTIN_LINE_APIS=0
 */
function shouldSkipBuiltinLineApis() {
  const v = String(process.env.ECOUNT_SKIP_BUILTIN_LINE_APIS ?? "").trim().toLowerCase();
  if (v === "0" || v === "false" || v === "no" || v === "off") return false;
  return true;
}

/** 디버그용: 같은 API 경로·HTTP상태(또는 no-item-lines)로 여러 바디를 시도한 로그를 한 줄로 합침. */
function compressEndpointErrorsForDebug(lines, maxOut = 14) {
  const arr = Array.isArray(lines) ? lines : [];
  const groups = new Map();
  for (const line of arr) {
    let sig = line;
    // Module/method paths contain slashes (e.g. Purchases/GetPurchases); use .+ not [^:]+
    const mHttp = /^(.+):\((\d+)\)/.exec(line);
    if (mHttp) sig = `${mHttp[1]}:${mHttp[2]}`;
    else {
      const idx = line.indexOf(":no-item-lines");
      if (idx > 0) sig = `${line.slice(0, idx)}:no-item-lines`;
    }
    const cur = groups.get(sig);
    if (cur) cur.count += 1;
    else groups.set(sig, { count: 1, sample: line });
  }
  const out = [];
  for (const { count, sample } of groups.values()) {
    const s = String(sample);
    const short = s.length > 200 ? `${s.slice(0, 200)}…` : s;
    out.push(count > 1 ? `${short} (같은 유형 ${count}회 시도)` : short);
  }
  return out.slice(0, maxOut);
}

async function postEcountOapiV2(apiBase, sessionId, v2Path, body) {
  const path = String(v2Path || "").replace(/^\/+/, "");
  const url = `${apiBase}/OAPI/V2/${path}?SESSION_ID=${encodeURIComponent(sessionId)}`;
  const doOnce = async () => {
    const r = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json", Accept: "application/json" },
      body: JSON.stringify(body)
    });
    const text = await r.text();
    let data = {};
    try {
      data = JSON.parse(text || "{}");
    } catch (_) {
      data = {};
    }
    return { ok: r.ok, status: r.status, data, text };
  };
  let result = await doOnce();
  if (result.status === 412) {
    await new Promise((resolve) => setTimeout(resolve, 1600));
    result = await doOnce();
  }
  return result;
}

/** GET /api/inbound-plans/detail 과 동일하게 발주 slipNo 의 품목 행까지 조회 */
async function fetchEcountPurchaseOrderMatchedLines(slipNo, from, to) {
  const slipNoTrim = String(slipNo || "").trim();
  if (!slipNoTrim) throw new Error("전표번호(slipNo)가 필요합니다.");

  const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();
  const fromCompact = normalizeDateCompact(from);
  const toCompact = normalizeDateCompact(to);
  const baseDate = toCompact || fromCompact || normalizeDateCompact(new Date().toISOString().slice(0, 10));
  const range = clampDateRangeToMax30Days(fromCompact || baseDate, toCompact || baseDate);

  const sessionId = await fetchEcountSession();
  const apiBase = await fetchEcountApiBase(comCode);
  const basePayload = {
    PROD_CD: "",
    CUST_CD: "",
    ListParam: {
      BASE_DATE_FROM: range.fromCompact,
      BASE_DATE_TO: range.toCompact,
      PAGE_CURRENT: 1,
      PAGE_SIZE: 100
    }
  };
  const target = normalizeSlipNoText(slipNoTrim);
  const seq = target.split("-").pop() || "";
  const endpointErrors = [];
  const skipBuiltinLineApis = shouldSkipBuiltinLineApis();
  let usedEndpoint = "GetPurchasesOrderList";
  let rawRows = [];

  const listUrl = `${apiBase}/OAPI/V2/Purchases/GetPurchasesOrderList?SESSION_ID=${encodeURIComponent(sessionId)}`;
  const listResp = await fetch(listUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify(basePayload)
  });
  if (!listResp.ok) {
    const txt = await listResp.text();
    throw new Error(`발주 목록 조회 실패(${listResp.status}): ${txt.slice(0, 300)}`);
  }
  const listData = await listResp.json();
  rawRows = extractArrayFromUnknown(listData);

  const headerRaw = findMatchingRawOrderRow(rawRows, target, seq, range.toCompact);
  const ordNoRawEarly = headerRaw ? String(pickFirstNonEmpty(headerRaw, ["ORD_NO"]) || "").trim() : "";
  const ioNoRawEarly = headerRaw ? String(pickFirstNonEmpty(headerRaw, ["IO_NO"]) || "").trim() : "";
  let matched = [];
  let lineSource = "header-flat";
  let nestedLineCount = 0;
  let narrowListTried = false;

  if (headerRaw) {
    const nestedLines = extractNestedLineRowsFromOrderRow(headerRaw);
    nestedLineCount = nestedLines.length;
    if (nestedLines.length) {
      const merged = mergeOrderHeaderWithLines(headerRaw, nestedLines);
      matched = normalizeInboundPlanRows(merged, range.toCompact);
      lineSource = "nested-lines";
    }
  }

  if (!matched.length && headerRaw) {
    const narrowAttempts = [];
    if (ordNoRawEarly) {
      narrowAttempts.push({
        label: "ListParam.ORD_NO",
        body: { ...basePayload, ListParam: { ...basePayload.ListParam, ORD_NO: ordNoRawEarly } }
      });
      narrowAttempts.push({ label: "root.ORD_NO", body: { ...basePayload, ORD_NO: ordNoRawEarly } });
    }
    if (ioNoRawEarly) {
      narrowAttempts.push({
        label: "ListParam.IO_NO",
        body: { ...basePayload, ListParam: { ...basePayload.ListParam, IO_NO: ioNoRawEarly } }
      });
    }
    for (const na of narrowAttempts) {
      narrowListTried = true;
      try {
        const nr = await fetch(listUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(na.body)
        });
        if (!nr.ok) continue;
        const nd = await nr.json();
        const nrows = extractArrayFromUnknown(nd);
        const h2 = findMatchingRawOrderRow(nrows, target, seq, range.toCompact) || nrows[0];
        if (!h2) continue;
        const nl = extractNestedLineRowsFromOrderRow(h2);
        if (nl.length) {
          nestedLineCount = nl.length;
          matched = normalizeInboundPlanRows(mergeOrderHeaderWithLines(h2, nl), range.toCompact);
          lineSource = `narrow-list-${na.label}`;
          break;
        }
      } catch (_) {
        // ignore narrow failures
      }
    }
  }

  const envLinePaths = parseInboundLineApiPathSpecs(process.env.ECOUNT_INBOUND_LINE_API_PATHS || "");
  if (!matched.length && envLinePaths.length) {
    const ordNoL = ordNoRawEarly || seq;
    const ioNoL = ioNoRawEarly;
    const bodiesL = buildLineLookupBodies(basePayload, slipNoTrim, ordNoL, ioNoL);
    outerEnvLines: for (const v2Path of envLinePaths) {
      for (const body of bodiesL.length ? bodiesL : [basePayload]) {
        try {
          const { ok, status, data, text } = await postEcountOapiV2(apiBase, sessionId, v2Path, body);
          if (!ok) {
            endpointErrors.push(`${v2Path}:(${status}) ${String(text || "").slice(0, 120)}`);
            continue;
          }
          const { matched: cand, kind } = matchDetailRowsToInboundPlans(data, range.toCompact);
          if (kind === "empty" || !filterLikelyItemLines(cand).length) {
            endpointErrors.push(`${v2Path}:no-item-lines`);
            continue;
          }
          matched = cand;
          usedEndpoint = v2Path;
          lineSource = `env-line-api-${kind}`;
          break outerEnvLines;
        } catch (err) {
          endpointErrors.push(`${v2Path}:network ${err?.message || String(err)}`);
        }
      }
    }
  }

  const customDetailApi = String(process.env.ECOUNT_INBOUND_DETAIL_API || "").trim();
  if (!matched.length && customDetailApi) {
    const ordNoRaw = ordNoRawEarly || seq;
    const ioNoRaw = ioNoRawEarly;
    const customV2Path = customDetailApi.includes("/") ? customDetailApi.replace(/^\/+/, "") : `Purchases/${customDetailApi}`;
    const customBodies = buildLineLookupBodies(basePayload, slipNoTrim, ordNoRaw, ioNoRaw);
    outerCustom: for (const body of customBodies.length ? customBodies : [basePayload]) {
      try {
        const { ok, status, data, text } = await postEcountOapiV2(apiBase, sessionId, customV2Path, body);
        if (!ok) {
          endpointErrors.push(`${customV2Path}:(${status}) ${String(text || "").slice(0, 120)}`);
          continue;
        }
        const { matched: cand, kind } = matchDetailRowsToInboundPlans(data, range.toCompact);
        if (kind === "empty" || !filterLikelyItemLines(cand).length) {
          endpointErrors.push(`${customV2Path}:no-item-lines`);
          continue;
        }
        matched = cand;
        usedEndpoint = customV2Path;
        lineSource = `custom-detail-api-${kind}`;
        break outerCustom;
      } catch (err) {
        endpointErrors.push(`${customV2Path}:network ${err?.message || String(err)}`);
      }
    }
  }

  if (!matched.length && !skipBuiltinLineApis) {
    const ordNoRaw = headerRaw ? String(pickFirstNonEmpty(headerRaw, ["ORD_NO"]) || "").trim() : seq;
    const ioNoRaw = headerRaw ? String(pickFirstNonEmpty(headerRaw, ["IO_NO"]) || "").trim() : "";
    const lineBodies = buildLineLookupBodies(basePayload, slipNoTrim, ordNoRaw, ioNoRaw);
    const purchaseMethods = [
      "GetPurchases",
      "GetPurchasesOrderDetail",
      "GetPurchaseOrderDetail",
      "GetPurchasesOrderInfo",
      "GetPurchasesDetail"
    ];
    const extraModulePaths = [
      "Purchases/GetReadPurchases",
      "Purchases/ReadPurchases",
      "Purchases/GetPurchasesList",
      "Stock/GetPurchases",
      "Inventory/GetPurchases",
      "Inventory/GetPurchaseList"
    ];
    const detailAttempts = [];
    for (const method of purchaseMethods) {
      for (const body of lineBodies) {
        detailAttempts.push({ path: `Purchases/${method}`, body });
      }
    }
    for (const path of extraModulePaths) {
      for (const body of lineBodies) {
        detailAttempts.push({ path, body });
      }
    }

    outerBuiltin: for (const c of detailAttempts) {
      try {
        const { ok, status, data, text } = await postEcountOapiV2(apiBase, sessionId, c.path, c.body);
        if (!ok) {
          endpointErrors.push(`${c.path}:(${status}) ${String(text || "").slice(0, 120)}`);
          continue;
        }
        const { matched: cand, kind } = matchDetailRowsToInboundPlans(data, range.toCompact);
        if (kind === "empty" || !filterLikelyItemLines(cand).length) {
          endpointErrors.push(`${c.path}:no-item-lines`);
          continue;
        }
        matched = cand;
        usedEndpoint = c.path;
        lineSource = `detail-api-${kind}`;
        break outerBuiltin;
      } catch (err) {
        endpointErrors.push(`${c.path}:network ${err?.message || String(err)}`);
      }
    }
  }

  if (!matched.length) {
    const items = normalizeInboundPlanRows(rawRows, range.toCompact);
    matched = items.filter((x) => normalizeSlipNoText(x.poNo) === target);
    if (!matched.length) {
      matched = items.filter((x) => normalizeSlipNoText(x.poNo).endsWith(`-${seq}`));
    }
    lineSource = "list-summary-fallback";
  }

  const lineLike = matched.filter((x) => {
    const hasCode = String(x.itemCode || "").trim() !== "";
    const hasQty = String(x.qty || "").trim() !== "";
    const isSummaryTitle = /외\s*\d+건/.test(String(x.itemName || ""));
    return (hasCode || hasQty) && !isSummaryTitle;
  });
  if (lineLike.length) matched = lineLike;

  return {
    matched,
    headerRaw,
    rawRows,
    range,
    sessionId,
    apiBase,
    usedEndpoint,
    lineSourceFinal: lineSource,
    endpointErrors,
    nestedLineCount,
    narrowListTried,
    skipBuiltinLineApis,
    target
  };
}

async function handleApi(req, res, urlObj) {
  const pathname = urlObj.pathname;
  const db = readDb();
  normalizeDb(db);

  if (req.method === "GET" && pathname === "/api/products") {
    return sendJson(res, 200, { items: db.products });
  }

  if (req.method === "GET" && pathname === "/api/product-options") {
    return sendJson(res, 200, { items: db.productOptions });
  }

  if (req.method === "GET" && pathname === "/api/warehouses") {
    return sendJson(res, 200, { items: db.warehouses });
  }

  if (req.method === "POST" && pathname === "/api/warehouses") {
    try {
      const body = await parseBody(req);
      const name = String(body.name || "").trim();
      if (!name) throw new Error("창고명은 필수입니다.");
      if (!db.warehouses.includes(name)) db.warehouses.push(name);
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.warehouses });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "DELETE" && pathname === "/api/warehouses") {
    try {
      const body = await parseBody(req);
      const name = String(body.name || "").trim();
      const approver = String(body.approver || "").trim();
      if (!name) throw new Error("창고명은 필수입니다.");
      if (!approver) throw new Error("담당자명 인증이 필요합니다.");
      if (approver !== "박유정") throw new Error("창고 삭제 권한은 박유정에게만 있습니다.");
      if (!db.managers.includes(approver)) throw new Error("등록된 담당자만 삭제할 수 있습니다.");
      if (name === "유통사업부") throw new Error("기본 창고(유통사업부)는 삭제할 수 없습니다.");
      const stocks = calculateStock(db, name);
      const hasRemaining = stocks.some((s) => Number(s.stock || 0) !== 0);
      if (hasRemaining) {
        throw new Error("해당 창고에 잔여 재고가 있습니다. 모두 이동/조정 후 삭제하세요.");
      }
      db.warehouses = db.warehouses.filter((w) => w !== name);
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.warehouses });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "PUT" && pathname === "/api/warehouses") {
    try {
      const body = await parseBody(req);
      const oldName = String(body.oldName || "").trim();
      const newName = String(body.newName || "").trim();
      if (!oldName || !newName) throw new Error("기존명/변경명은 필수입니다.");
      if (!db.warehouses.includes(oldName)) throw new Error("기존 창고를 찾을 수 없습니다.");
      if (db.warehouses.includes(newName)) throw new Error("이미 존재하는 창고명입니다.");

      db.warehouses = db.warehouses.map((w) => (w === oldName ? newName : w));
      for (const m of db.movements) {
        if (m.warehouse === oldName) m.warehouse = newName;
        if (m.toWarehouse === oldName) m.toWarehouse = newName;
      }
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.warehouses });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/products") {
    try {
      const body = await parseBody(req);
      upsertProduct(db, body);
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/products/bulk") {
    try {
      const body = await parseBody(req);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      for (const r of rows) upsertProduct(db, r);
      writeDb(db);
      return sendJson(res, 200, { ok: true, count: rows.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "DELETE" && pathname === "/api/products") {
    try {
      const body = await parseBody(req);
      const codes = Array.isArray(body.codes) ? body.codes.map((x) => String(x || "").trim()).filter(Boolean) : [];
      if (!codes.length) throw new Error("삭제할 상품코드를 선택하세요.");
      db.products = db.products.filter((p) => !codes.includes(String(p.code)));
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/product-options") {
    try {
      const body = await parseBody(req);
      const field = String(body.field || "").trim();
      if (!PRODUCT_OPTION_KEYS.includes(field)) throw new Error("유효하지 않은 옵션 항목입니다.");
      const values = normalizeOptionValues(body.values || body.value);
      const replace = Boolean(body.replace);
      db.productOptions[field] = replace
        ? values
        : normalizeOptionValues([...(db.productOptions[field] || []), ...values]);
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.productOptions });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/partners") {
    return sendJson(res, 200, { items: db.partners });
  }

  if (req.method === "POST" && pathname === "/api/partners") {
    try {
      const body = await parseBody(req);
      const type = String(body.type || "").trim();
      const name = String(body.name || "").trim();
      if (!["inbound", "outbound", "purchase"].includes(type)) {
        throw new Error("유효하지 않은 거래처 구분입니다.");
      }
      if (!name) throw new Error("거래처명은 필수입니다.");
      if (!db.partners[type].includes(name)) db.partners[type].push(name);
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.partners[type] });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "DELETE" && pathname === "/api/partners") {
    try {
      const body = await parseBody(req);
      const type = String(body.type || "").trim();
      const name = String(body.name || "").trim();
      if (!["inbound", "outbound", "purchase"].includes(type)) {
        throw new Error("유효하지 않은 거래처 구분입니다.");
      }
      if (!name) throw new Error("거래처명은 필수입니다.");
      db.partners[type] = db.partners[type].filter((x) => x !== name);
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/managers") {
    return sendJson(res, 200, { items: db.managers });
  }

  if (req.method === "POST" && pathname === "/api/managers") {
    try {
      const body = await parseBody(req);
      const name = String(body.name || "").trim();
      if (!name) throw new Error("담당자명은 필수입니다.");
      if (!db.managers.includes(name)) db.managers.push(name);
      writeDb(db);
      return sendJson(res, 200, { ok: true, items: db.managers });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "DELETE" && pathname === "/api/managers") {
    try {
      const body = await parseBody(req);
      const name = String(body.name || "").trim();
      if (!name) throw new Error("담당자명은 필수입니다.");
      db.managers = db.managers.filter((x) => x !== name);
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/stock") {
    const warehouse = String(urlObj.searchParams.get("warehouse") || "").trim();
    return sendJson(res, 200, {
      items: calculateStock(db, warehouse || undefined),
      warehouse: warehouse || ""
    });
  }

  if (req.method === "POST" && pathname === "/api/movements") {
    try {
      const body = await parseBody(req);
      addMovement(db, body);
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/movements/bulk") {
    try {
      const body = await parseBody(req);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      for (const r of rows) addMovement(db, r);
      writeDb(db);
      return sendJson(res, 200, { ok: true, count: rows.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/movements/cancel") {
    try {
      const body = await parseBody(req);
      const id = Number(body.id);
      const user = String(body.user || "").trim();
      if (!Number.isFinite(id)) throw new Error("유효하지 않은 이력 ID입니다.");
      cancelMovement(db, id, user);
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/history") {
    const q = String(urlObj.searchParams.get("q") || "").toLowerCase();
    const stockAfter = computeStockAfterEachMovement(db);
    const sourceToCenterMergedSlipMap = new Map();
    const confirmedLists = db.outboundOrderUpload && db.outboundOrderUpload.confirmedLists;
    if (confirmedLists && typeof confirmedLists === "object") {
      for (const recMap of Object.values(confirmedLists)) {
        if (!recMap || typeof recMap !== "object") continue;
        for (const rec of Object.values(recMap)) {
          if (!rec || typeof rec !== "object") continue;
          const centerMergedSlipNo = String(rec.key || "").trim();
          if (!centerMergedSlipNo) continue;
          const lines = Array.isArray(rec.lines) ? rec.lines : [];
          for (const ln of lines) {
            const source = normalizeSlipNoText(String((ln && (ln.sourceSlipNo || ln.slipNo)) || "").trim());
            if (!source) continue;
            if (!sourceToCenterMergedSlipMap.has(source)) sourceToCenterMergedSlipMap.set(source, centerMergedSlipNo);
          }
        }
      }
    }
    const joined = db.movements
      .map((m) => {
        const p = db.products.find((x) => x.code === m.productCode);
        const rawSlipNo = String(m.slipNo || "").trim();
        const mappedCenterSlip = sourceToCenterMergedSlipMap.get(normalizeSlipNoText(rawSlipNo)) || "";
        const slipNoDisplay = mappedCenterSlip || rawSlipNo;
        return {
          ...m,
          slipNoDisplay,
          productName: p ? p.name : "",
          ecountCode: p ? p.ecountCode || "" : "",
          stockAfter: stockAfter[m.id] ?? 0
        };
      })
      .filter((x) => {
        if (!q) return true;
        return (
          String(x.productCode).toLowerCase().includes(q) ||
          String(x.productName).toLowerCase().includes(q) ||
          String(x.ecountCode || "").toLowerCase().includes(q)
        );
      })
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
    const outSlipTotals = new Map();
    for (const x of joined) {
      if (x.type !== "OUT") continue;
      const slip = String(x.slipNoDisplay || "").trim();
      if (!slip) continue;
      outSlipTotals.set(slip, (outSlipTotals.get(slip) || 0) + Math.abs(Number(x.qty || 0)));
    }
    const items = joined.map((x) => {
      if (x.type !== "OUT") return x;
      const slip = String(x.slipNoDisplay || "").trim();
      if (!slip) return x;
      return {
        ...x,
        qtyDisplay: outSlipTotals.get(slip) ?? Math.abs(Number(x.qty || 0))
      };
    });
    return sendJson(res, 200, { items });
  }

  if (req.method === "GET" && pathname === "/api/recent") {
    const type = String(urlObj.searchParams.get("type") || "").trim();
    const limit = Math.max(1, Number(urlObj.searchParams.get("limit") || 5));
    if (!["IN", "OUT"].includes(type)) {
      return sendJson(res, 400, { error: "유효하지 않은 조회 타입입니다." });
    }
    const stockAfter = computeStockAfterEachMovement(db);
    if (type === "IN") {
      const cancelledSlipSet = new Set(
        db.movements
          .filter(
            (m) =>
              m &&
              m.type === "ADJUST" &&
              !m.cancelled &&
              String(m.memo || "").includes("입고확정 취소 원복") &&
              String(m.slipNo || "").trim()
          )
          .map((m) => normalizeSlipNoText(m.slipNo))
          .filter(Boolean)
      );
      const items = db.movements
        .filter((m) => m.type === "IN" && !m.cancelled)
        .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
        .slice(0, limit)
        .map((m) => {
          const p = db.products.find((x) => x.code === m.productCode);
          const slipKey = normalizeSlipNoText(m.slipNo);
          return {
            ...m,
            slipConfirm: false,
            slipCancelled: Boolean(slipKey && cancelledSlipSet.has(slipKey)),
            productName: p ? p.name : "",
            ecountCode: p ? p.ecountCode || "" : "",
            stockAfter: stockAfter[m.id] ?? 0
          };
        });
      return sendJson(res, 200, { items });
    }
    const sourceToCenterMergedSlipMap = new Map();
    const confirmedLists = db.outboundOrderUpload && db.outboundOrderUpload.confirmedLists;
    if (confirmedLists && typeof confirmedLists === "object") {
      for (const recMap of Object.values(confirmedLists)) {
        if (!recMap || typeof recMap !== "object") continue;
        for (const rec of Object.values(recMap)) {
          if (!rec || typeof rec !== "object") continue;
          const centerMergedSlipNo = String(rec.key || "").trim();
          if (!centerMergedSlipNo) continue;
          const lines = Array.isArray(rec.lines) ? rec.lines : [];
          for (const ln of lines) {
            const source = normalizeSlipNoText(String((ln && (ln.sourceSlipNo || ln.slipNo)) || "").trim());
            if (!source) continue;
            if (!sourceToCenterMergedSlipMap.has(source)) sourceToCenterMergedSlipMap.set(source, centerMergedSlipNo);
          }
        }
      }
    }
    const itemsRaw = db.movements
      .filter((m) => m.type === type && !m.cancelled)
      .sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt))
      .slice(0, limit)
      .map((m) => {
        const p = db.products.find((x) => x.code === m.productCode);
        const rawSlipNo = String(m.slipNo || "").trim();
        const mappedCenterSlip = sourceToCenterMergedSlipMap.get(normalizeSlipNoText(rawSlipNo)) || "";
        const slipNoDisplay = mappedCenterSlip || rawSlipNo;
        return {
          ...m,
          slipNoDisplay,
          slipConfirm: false,
          productName: p ? p.name : "",
          ecountCode: p ? p.ecountCode || "" : "",
          stockAfter: stockAfter[m.id] ?? 0
        };
      });
    const outSlipTotals = new Map();
    for (const x of itemsRaw) {
      const slip = String(x.slipNoDisplay || "").trim();
      if (!slip) continue;
      outSlipTotals.set(slip, (outSlipTotals.get(slip) || 0) + Math.abs(Number(x.qty || 0)));
    }
    const items = itemsRaw.map((x) => {
      const slip = String(x.slipNoDisplay || "").trim();
      if (!slip) return x;
      return {
        ...x,
        qtyDisplay: outSlipTotals.get(slip) ?? Math.abs(Number(x.qty || 0))
      };
    });
    return sendJson(res, 200, { items });
  }

  /* 판매처별 출고 발주 업로드 */
  if (req.method === "POST" && pathname === "/api/outbound-order-upload") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      const sourceFileName = String(body.sourceFileName || "").trim().slice(0, 240);
      let matrix = [];
      if (Array.isArray(body.matrix) && body.matrix.length) {
        matrix = body.matrix;
      } else if (body.sheetBase64) {
        const buf = Buffer.from(String(body.sheetBase64), "base64");
        const wb = XLSX.read(buf, { type: "buffer" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      } else {
        throw new Error("matrix 또는 sheetBase64가 필요합니다.");
      }
      const uploadBatchId = `OUT-${Date.now()}`;
      const { okRows, errorRows } = parseOutboundOrderRowsByPartner(matrix, partnerType, uploadBatchId, sourceFileName, db);
      db.outboundOrderUpload.uploadedRows = [...(db.outboundOrderUpload.uploadedRows || []), ...okRows];
      db.outboundOrderUpload.batches = [
        {
          uploadBatchId,
          partnerType,
          partner: partnerLabelFromType(partnerType),
          sourceFileName,
          parserVersion: OUTBOUND_PARTNER_ADAPTERS[partnerType].parserVersion,
          okCount: okRows.length,
          errorCount: errorRows.length,
          uploadedAt: new Date().toISOString()
        },
        ...(db.outboundOrderUpload.batches || [])
      ].slice(0, 120);
      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      const items = aggregateOutboundUploadSlips(db.outboundOrderUpload.lines || [], db.outboundOrderUpload.slipWorkflow || {}, partnerType);
      return sendJson(res, 200, {
        ok: true,
        uploadBatchId,
        items,
        okCount: okRows.length,
        errorRows,
        needsApply: true
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    const items = aggregateOutboundUploadSlips(
      db.outboundOrderUpload.lines || [],
      db.outboundOrderUpload.slipWorkflow || {},
      partnerType
    );
    const batches = (db.outboundOrderUpload.batches || []).filter((x) => x.partnerType === partnerType).slice(0, 30);
    const appliedBatchId = String((db.outboundOrderUpload.appliedBatchByPartner || {})[partnerType] || "").trim();
    return sendJson(res, 200, { items, batches, appliedBatchId, uploadedAt: db.outboundOrderUpload.uploadedAt || "" });
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/apply-batch") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      const uploadBatchId = String(body.uploadBatchId || "").trim();
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      if (!uploadBatchId) throw new Error("적용할 uploadBatchId가 필요합니다.");

      const targetRows = (db.outboundOrderUpload.uploadedRows || []).filter(
        (row) => row.partnerType === partnerType && String(row.uploadBatchId || "").trim() === uploadBatchId
      );
      if (!targetRows.length) throw new Error("선택한 업로드 파일의 데이터가 없습니다.");

      db.outboundOrderUpload.lines = (db.outboundOrderUpload.lines || []).filter((row) => row.partnerType !== partnerType);
      db.outboundOrderUpload.lines.push(...targetRows);

      // 적용 파일이 바뀌면 해당 판매처 워크플로우/확정리스트를 초기화한다.
      const aliveKeys = new Set(targetRows.map((x) => normalizeSlipNoText(x.slipNo)).filter(Boolean));
      const wf = db.outboundOrderUpload.slipWorkflow || {};
      for (const k of Object.keys(wf)) {
        if (!aliveKeys.has(k)) delete wf[k];
      }
      db.outboundOrderUpload.slipWorkflow = wf;
      db.outboundOrderUpload.confirmedLists = db.outboundOrderUpload.confirmedLists || {};
      db.outboundOrderUpload.confirmedLists[partnerType] = {};

      db.outboundOrderUpload.appliedBatchByPartner = db.outboundOrderUpload.appliedBatchByPartner || {};
      db.outboundOrderUpload.appliedBatchByPartner[partnerType] = uploadBatchId;
      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);

      const items = aggregateOutboundUploadSlips(db.outboundOrderUpload.lines || [], db.outboundOrderUpload.slipWorkflow || {}, partnerType);
      return sendJson(res, 200, { ok: true, appliedBatchId: uploadBatchId, items, appliedCount: targetRows.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/delete") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      const all = Boolean(body.all);
      const keys = new Set(
        (Array.isArray(body.slipNos) ? body.slipNos : [])
          .map((x) => normalizeSlipNoText(x))
          .filter(Boolean)
      );
      if (!all && !keys.size) throw new Error("삭제할 전표번호가 없습니다.");

      const before = db.outboundOrderUpload.lines || [];
      const keep = before.filter((row) => {
        if (row.partnerType !== partnerType) return true;
        if (all) return false;
        return !keys.has(normalizeSlipNoText(row.slipNo));
      });
      const deletedLineCount = before.length - keep.length;
      db.outboundOrderUpload.lines = keep;

      const wf = db.outboundOrderUpload.slipWorkflow || {};
      if (all) {
        for (const k of Object.keys(wf)) {
          const row = before.find((x) => normalizeSlipNoText(x.slipNo) === k);
          if (row && row.partnerType === partnerType) delete wf[k];
        }
      } else {
        for (const k of keys) delete wf[k];
      }
      db.outboundOrderUpload.slipWorkflow = wf;

      if (all) {
        db.outboundOrderUpload.uploadedRows = (db.outboundOrderUpload.uploadedRows || []).filter((x) => x.partnerType !== partnerType);
        db.outboundOrderUpload.batches = (db.outboundOrderUpload.batches || []).filter((x) => x.partnerType !== partnerType);
        if (db.outboundOrderUpload.appliedBatchByPartner) delete db.outboundOrderUpload.appliedBatchByPartner[partnerType];
        db.outboundOrderUpload.confirmedLists = db.outboundOrderUpload.confirmedLists || {};
        db.outboundOrderUpload.confirmedLists[partnerType] = {};
      }
      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);

      const items = aggregateOutboundUploadSlips(
        db.outboundOrderUpload.lines || [],
        db.outboundOrderUpload.slipWorkflow || {},
        partnerType
      );
      return sendJson(res, 200, { ok: true, deletedLineCount, items });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/delete-batches") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      const batchIds = new Set((Array.isArray(body.uploadBatchIds) ? body.uploadBatchIds : []).map((x) => String(x || "").trim()).filter(Boolean));
      if (!batchIds.size) throw new Error("삭제할 업로드 배치가 없습니다.");

      const beforeUploaded = db.outboundOrderUpload.uploadedRows || [];
      db.outboundOrderUpload.uploadedRows = beforeUploaded.filter(
        (row) => !(row.partnerType === partnerType && batchIds.has(String(row.uploadBatchId || "").trim()))
      );
      const deletedLineCount = beforeUploaded.length - db.outboundOrderUpload.uploadedRows.length;

      const beforeActive = db.outboundOrderUpload.lines || [];
      db.outboundOrderUpload.lines = beforeActive.filter(
        (row) => !(row.partnerType === partnerType && batchIds.has(String(row.uploadBatchId || "").trim()))
      );

      db.outboundOrderUpload.batches = (db.outboundOrderUpload.batches || []).filter(
        (b) => !(b.partnerType === partnerType && batchIds.has(String(b.uploadBatchId || "").trim()))
      );

      const aliveSlipSet = new Set(
        (db.outboundOrderUpload.lines || [])
          .filter((x) => x.partnerType === partnerType)
          .map((x) => normalizeSlipNoText(x.slipNo))
          .filter(Boolean)
      );
      for (const k of Object.keys(db.outboundOrderUpload.slipWorkflow || {})) {
        if (!aliveSlipSet.has(k)) delete db.outboundOrderUpload.slipWorkflow[k];
      }

      const appliedMap = db.outboundOrderUpload.appliedBatchByPartner || {};
      const appliedId = String(appliedMap[partnerType] || "").trim();
      if (appliedId && batchIds.has(appliedId)) {
        delete appliedMap[partnerType];
        db.outboundOrderUpload.confirmedLists = db.outboundOrderUpload.confirmedLists || {};
        db.outboundOrderUpload.confirmedLists[partnerType] = {};
      }

      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      return sendJson(res, 200, { ok: true, deletedLineCount });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/lines") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    const centerTab = String(urlObj.searchParams.get("centerTab") || "all").trim();
    const baseRows = (db.outboundOrderUpload.lines || []).filter((x) => x.partnerType === partnerType);
    const counts =
      partnerType === "emart"
        ? {
            all: baseRows.length,
            여주: baseRows.filter((x) => resolveCenterBucket(x) === "여주").length,
            대구: baseRows.filter((x) => resolveCenterBucket(x) === "대구").length,
            시화: baseRows.filter((x) => resolveCenterBucket(x) === "시화").length
          }
        : { all: baseRows.length };
    let rows = baseRows;
    if (partnerType === "emart" && centerTab !== "all") {
      rows = rows.filter((x) => resolveCenterBucket(x) === centerTab);
    }
    return sendJson(res, 200, { items: rows, counts });
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/unship-lines") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    const centerTab = String(urlObj.searchParams.get("centerTab") || "all").trim();
    const baseRows = (db.outboundOrderUpload.lines || []).filter(
      (x) => x.partnerType === partnerType && String(x.unshipStatus || "").trim() === "미출"
    );
    const counts =
      partnerType === "emart"
        ? {
            all: baseRows.length,
            여주: baseRows.filter((x) => resolveCenterBucket(x) === "여주").length,
            대구: baseRows.filter((x) => resolveCenterBucket(x) === "대구").length,
            시화: baseRows.filter((x) => resolveCenterBucket(x) === "시화").length
          }
        : { all: baseRows.length };
    let rows = baseRows;
    if (partnerType === "emart" && centerTab !== "all") {
      rows = rows.filter((x) => resolveCenterBucket(x) === centerTab);
    }
    const items = rows.map((x) => ({
      slipNo: x.slipNo || "",
      dueDate: x.dueDate || "",
      storeInDate: x.storeInDate || "",
      poDate: x.poDate || "",
      orderDate: x.orderDate || "",
      centerName: x.centerName || "",
      centerCode: x.centerCode || "",
      storeName: x.storeName || "",
      productCode: x.productCode || "",
      sourceProductCode: x.sourceProductCode || "",
      productName: x.productName || "",
      orderQty: x.orderQty,
      fixedQty: x.fixedQty != null && String(x.fixedQty).trim() !== "" ? x.fixedQty : 0,
      lot: x.lot || "",
      warehouse: x.warehouse || "",
      uploadBatchId: x.uploadBatchId || "",
      sourceRowNo: x.sourceRowNo
    }));
    return sendJson(res, 200, { items, counts });
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/lines/update") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      const updatesRaw = Array.isArray(body.updates) ? body.updates : [];
      const updates = updatesRaw
        .map((x) => {
          const uploadBatchId = String(x?.uploadBatchId || "").trim();
          const sourceRowNoRaw = Number(x?.sourceRowNo);
          const sourceRowNo = Number.isFinite(sourceRowNoRaw) ? Math.trunc(sourceRowNoRaw) : NaN;
          const status = String(x?.unshipStatus || "").trim() === "미출" ? "미출" : "정상";
          let fixedQtyRaw = Number(String(x?.fixedQty ?? "").replace(/,/g, "").trim());
          if (!Number.isFinite(fixedQtyRaw)) fixedQtyRaw = 0;
          const fixedQty = Math.max(0, Math.trunc(fixedQtyRaw));
          if (!uploadBatchId || !Number.isFinite(sourceRowNo) || sourceRowNo <= 0) return null;
          return { uploadBatchId, sourceRowNo, unshipStatus: status, fixedQty };
        })
        .filter(Boolean);
      if (!updates.length) throw new Error("저장할 라인 변경값이 없습니다.");

      const updateMap = new Map(updates.map((u) => [`${u.uploadBatchId}|${u.sourceRowNo}`, u]));
      let changedCount = 0;
      db.outboundOrderUpload.lines = (db.outboundOrderUpload.lines || []).map((row) => {
        if (!row || row.partnerType !== partnerType) return row;
        const key = `${String(row.uploadBatchId || "").trim()}|${Number(row.sourceRowNo)}`;
        const next = updateMap.get(key);
        if (!next) return row;
        changedCount += 1;
        return {
          ...row,
          unshipStatus: next.unshipStatus,
          fixedQty: String(next.fixedQty)
        };
      });

      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      return sendJson(res, 200, { ok: true, changedCount });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-code-master") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    const items = (db.outboundCodeMasters || []).filter((x) => x.partnerType === partnerType).slice(0, 2000);
    return sendJson(res, 200, { items });
  }

  if (req.method === "POST" && pathname === "/api/outbound-code-master/upload") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType);
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      let matrix = [];
      if (Array.isArray(body.matrix) && body.matrix.length) matrix = body.matrix;
      else if (body.sheetBase64) {
        const buf = Buffer.from(String(body.sheetBase64), "base64");
        const wb = XLSX.read(buf, { type: "buffer" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      } else throw new Error("matrix 또는 sheetBase64가 필요합니다.");

      const sourceFileName = String(body.sourceFileName || "").trim().slice(0, 240);
      const { items, errorRows } = parseOutboundCodeMasterRows(matrix, partnerType, sourceFileName);
      const keep = (db.outboundCodeMasters || []).filter((x) => x.partnerType !== partnerType);
      const merged = new Map();
      for (const x of items) merged.set(String(x.sourceCode), x);
      db.outboundCodeMasters = [...keep, ...Array.from(merged.values())];
      writeDb(db);
      return sendJson(res, 200, { ok: true, count: items.length, errorRows });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/template") {
    try {
      const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      const headers = OUTBOUND_PARTNER_TEMPLATE_HEADERS[partnerType];
      if (!Array.isArray(headers) || !headers.length) throw new Error("템플릿 헤더를 찾지 못했습니다.");
      const sampleOrderNo =
        partnerType === "daiso" ? "DS-ORDER-001" : partnerType === "emart" ? "EM-ORDER-001" : "LT-ORDER-001";
      const sampleRow = [
        sampleOrderNo,
        "2026-04-27",
        "2026-04-30",
        "SAMPLE-001",
        "샘플상품",
        10,
        "유통사업부",
        "샘플 비고"
      ];
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet([headers, sampleRow]);
      XLSX.utils.book_append_sheet(wb, ws, "template");
      const out = XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
      const filename = `outbound-upload-template-${partnerType}.xlsx`;
      res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Content-Length": Buffer.byteLength(out)
      });
      return res.end(out);
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/detail") {
    const slipNo = String(urlObj.searchParams.get("slipNo") || "").trim();
    if (!slipNo) return sendJson(res, 400, { error: "slipNo가 필요합니다." });
    const key = normalizeSlipNoText(slipNo);
    const detailLines = outboundUploadLinesToDetailRows(db.outboundOrderUpload.lines || [], key);
    if (!detailLines.length) return sendJson(res, 200, { item: null, items: [], workflow: { status: "draft", lineQtyByIndex: {} } });
    const wf =
      (db.outboundOrderUpload.slipWorkflow && db.outboundOrderUpload.slipWorkflow[key]) || { status: "draft", lineQtyByIndex: {} };
    const qMap = wf.lineQtyByIndex && typeof wf.lineQtyByIndex === "object" ? wf.lineQtyByIndex : {};
    const items = detailLines.map((x, idx) => ({
      ...x,
      confirmQty: qMap[String(idx)] != null ? qMap[String(idx)] : x.orderQty
    }));
    return sendJson(res, 200, {
      item: items[0],
      items,
      workflow: {
        status: wf.status || "draft",
        lineQtyByIndex: qMap,
        confirmedAt: wf.confirmedAt || null,
        sentAt: wf.sentAt || null
      }
    });
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/workflow") {
    try {
      const body = await parseBody(req);
      const slipNo = String(body.slipNo || "").trim();
      const action = String(body.action || "").trim();
      if (!slipNo) throw new Error("slipNo가 필요합니다.");
      const key = normalizeSlipNoText(slipNo);
      const group = outboundUploadLinesToDetailRows(db.outboundOrderUpload.lines || [], key);
      if (!group.length) throw new Error("해당 전표의 품목 행이 없습니다.");
      const wfPrev = db.outboundOrderUpload.slipWorkflow[key] || { status: "draft", lineQtyByIndex: {} };
      const bodyLineQtyByIndex = body.lineQtyByIndex && typeof body.lineQtyByIndex === "object" ? body.lineQtyByIndex : {};

      if (action === "save-check") {
        const qMerged = mergeLineQtyByIndexFromBody(bodyLineQtyByIndex, group.map((x) => ({ qty: x.orderQty })));
        db.outboundOrderUpload.slipWorkflow[key] = { ...wfPrev, status: "draft", lineQtyByIndex: qMerged, confirmedAt: undefined };
      } else if (action === "confirm") {
        applyOutboundWorkflowConfirm(db, slipNo, String(body.stockUser || "").trim(), body.lineQtyByIndex);
      } else if (action === "reopen") {
        performOutboundWorkflowReopen(db, slipNo);
      } else if (action === "send-sales") {
        if (wfPrev.status !== "confirmed") throw new Error("출고확정 후에 판매입력을 전송할 수 있습니다.");
        const enabled = String(process.env.ECOUNT_TEST_SAVE_SALES ?? "").trim().toLowerCase();
        if (!["1", "true", "yes", "on"].includes(enabled)) {
          throw new Error("비활성화됨. 테스트 시 .env에 ECOUNT_TEST_SAVE_SALES=1 후 서버 재시작하세요.");
        }
        const ptype = String(group[0]?.partnerType || "").trim();
        if (!ptype) throw new Error("판매처(partnerType)를 찾을 수 없습니다.");
        const listKey = findOutboundConfirmListKeyForSlip(db, ptype, slipNo);
        if (!listKey) {
          throw new Error("확정리스트에 해당 전표 묶음이 없습니다. 확정리스트 탭에서 판매입력을 실행하세요.");
        }
        const ecountOut = await runOutboundConfirmListSalesToEcount(db, ptype, listKey);
        return sendJson(res, 200, {
          ok: true,
          workflow: db.outboundOrderUpload.slipWorkflow[key],
          ecount: ecountOut
        });
      } else {
        throw new Error("action은 save-check, confirm, reopen, send-sales 중 하나여야 합니다.");
      }
      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      return sendJson(res, 200, { ok: true, workflow: db.outboundOrderUpload.slipWorkflow[key] });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/workflow/bulk-confirm") {
    try {
      const body = await parseBody(req);
      const stockUser = String(body.stockUser || "").trim();
      const slipNosRaw = Array.isArray(body.slipNos) ? body.slipNos : [];
      if (!stockUser) throw new Error("출고 담당자(stockUser)가 필요합니다.");
      const slipByNorm = new Map();
      for (const s of slipNosRaw) {
        const slipOne = String(s || "").trim();
        if (!slipOne) continue;
        const n = normalizeSlipNoText(slipOne);
        if (!n) continue;
        if (!slipByNorm.has(n)) slipByNorm.set(n, slipOne);
      }
      const slipNos = [...slipByNorm.values()];
      if (!slipNos.length) throw new Error("확정할 전표(slipNos)가 없습니다.");
      if (!db.managers.includes(stockUser)) throw new Error("출고 담당자를 선택하세요.");

      const dbWork = JSON.parse(JSON.stringify(db));
      for (const slipNo of slipNos) {
        try {
          applyOutboundWorkflowConfirm(dbWork, slipNo, stockUser, undefined);
        } catch (e) {
          throw new Error(`${slipNo}: ${e.message || e}`);
        }
      }
      dbWork.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(dbWork);
      return sendJson(res, 200, { ok: true, count: slipNos.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/confirm-list/detail") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    const listKey = String(urlObj.searchParams.get("listKey") || "").trim();
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    if (!listKey) return sendJson(res, 400, { error: "listKey가 필요합니다." });
    const recMap = db.outboundOrderUpload.confirmedLists && db.outboundOrderUpload.confirmedLists[partnerType];
    if (!recMap || typeof recMap !== "object") {
      return sendJson(res, 404, { error: "확정 저장 내역이 없습니다." });
    }
    const record = recMap[listKey];
    if (!record) return sendJson(res, 404, { error: "해당 키의 확정 내역을 찾을 수 없습니다." });
    const slipNos = Array.isArray(record.slipNos) ? record.slipNos.map((s) => String(s || "").trim()).filter(Boolean) : [];
    const lines = Array.isArray(record.lines)
      ? record.lines.map((ln) => ({
          slipNo: String(ln.slipNo || "").trim(),
          sourceSlipNo: String(ln.sourceSlipNo || "").trim(),
          productCode: String(ln.productCode || "").trim(),
          productName: String(ln.productName || "").trim(),
          qty: ln.qty != null ? ln.qty : "",
          warehouse: String(ln.warehouse || "").trim(),
          lot: String(ln.lot || "").trim(),
          orderQty: String(ln.orderQty || "").trim(),
          fixedQty: String(ln.fixedQty || "").trim(),
          unshipStatus: String(ln.unshipStatus || "").trim()
        }))
      : [];
    return sendJson(res, 200, {
      item: {
        key: record.key || listKey,
        centerMergedSlipNo: record.key || listKey,
        partnerType: record.partnerType || partnerType,
        centerName: record.centerName || "",
        saveDateYmd: record.saveDateYmd || "",
        storeInDateYmd: record.storeInDateYmd || "",
        updatedAt: record.updatedAt || "",
        slipNos,
        lines
      }
    });
  }

  if (req.method === "GET" && pathname === "/api/outbound-order-upload/confirm-list") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType(daiso/emart/lotte)가 필요합니다." });
    const recMap = db.outboundOrderUpload.confirmedLists && db.outboundOrderUpload.confirmedLists[partnerType];
    if (!recMap || typeof recMap !== "object") {
      return sendJson(res, 200, { items: [] });
    }
    const items = Object.values(recMap)
      .filter((x) => x && x.key)
      .map((x) => {
        const slipKeys = Array.isArray(x.slipNos)
          ? [...new Set(x.slipNos.map((s) => normalizeSlipNoText(String(s || ""))).filter(Boolean))]
          : [];
        const sentCount = slipKeys.filter((sk) => {
          const wf = db.outboundOrderUpload.slipWorkflow && db.outboundOrderUpload.slipWorkflow[sk];
          return Boolean(wf && wf.sentAt);
        }).length;
        return {
          key: x.key,
          centerMergedSlipNo: x.key,
          centerName: x.centerName || "",
          saveDateYmd: x.saveDateYmd || "",
          storeInDateYmd: x.storeInDateYmd || "",
          slipCount: Array.isArray(x.slipNos) ? x.slipNos.length : 0,
          lineCount: Array.isArray(x.lines) ? x.lines.length : 0,
          productCount: (() => {
            const set = new Set(
              (Array.isArray(x.lines) ? x.lines : [])
                .map((ln) => String((ln && ln.productCode) || "").trim())
                .filter(Boolean)
            );
            return set.size;
          })(),
          totalQty: (Array.isArray(x.lines) ? x.lines : []).reduce((sum, ln) => {
            const n = Number((ln && ln.qty) ?? 0);
            return sum + (Number.isFinite(n) ? n : 0);
          }, 0),
          sentCount,
          allSent: slipKeys.length > 0 && sentCount === slipKeys.length,
          updatedAt: x.updatedAt || ""
        };
      })
      .sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")));
    return sendJson(res, 200, { items });
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/confirm-list/send-sales") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType || "");
      const listKey = String(body.listKey || "").trim();
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      if (!listKey) throw new Error("확정리스트 키(listKey)가 필요합니다.");

      const enabled = String(process.env.ECOUNT_TEST_SAVE_SALES ?? "").trim().toLowerCase();
      if (!["1", "true", "yes", "on"].includes(enabled)) {
        throw new Error("비활성화됨. 테스트 시 .env에 ECOUNT_TEST_SAVE_SALES=1 후 서버 재시작하세요.");
      }

      const out = await runOutboundConfirmListSalesToEcount(db, partnerType, listKey);
      return sendJson(res, 200, out);
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/outbound-order-upload/confirm-list/cancel") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType || "");
      const listKey = String(body.listKey || "").trim();
      if (!partnerType) throw new Error("partnerType(daiso/emart/lotte)가 필요합니다.");
      if (!listKey) throw new Error("확정리스트 키(listKey)가 필요합니다.");
      db.outboundOrderUpload.confirmedLists = db.outboundOrderUpload.confirmedLists || {};
      db.outboundOrderUpload.confirmedLists[partnerType] = db.outboundOrderUpload.confirmedLists[partnerType] || {};
      const recMap = db.outboundOrderUpload.confirmedLists[partnerType];
      const record = recMap[listKey];
      if (!record) throw new Error("해당 확정 저장 내역을 찾을 수 없습니다.");

      const slipKeys = Array.isArray(record.slipNos)
        ? [...new Set(record.slipNos.map((s) => normalizeSlipNoText(String(s || ""))).filter(Boolean))]
        : [];

      for (const sk of slipKeys) {
        const wf = db.outboundOrderUpload.slipWorkflow[sk] || {};
        if (wf.sentAt) throw new Error(`판매입력 전송이 완료된 전표가 포함되어 있어 취소할 수 없습니다: ${sk}`);
      }

      for (const sk of slipKeys) {
        const wf = db.outboundOrderUpload.slipWorkflow[sk] || {};
        if (wf.status === "confirmed") performOutboundWorkflowReopen(db, sk);
      }

      delete recMap[listKey];
      db.outboundOrderUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      return sendJson(res, 200, { ok: true, reopenedSlips: slipKeys.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  /*
   * 입고예정목록 2 — 이카운트 「발주서현황」 엑셀 업로드 → data/db.json inboundPlanUpload
   * 병합: 이미 저장된 발주번호와 동일한 전표는 건너뛰고, 신규 발주번호 품목 행만 추가(기존 행 덮어쓰기 없음).
   */
  if (req.method === "POST" && pathname === "/api/inbound-plan-upload") {
    try {
      const body = await parseBody(req);
      const sourceFileName = String(body.sourceFileName || "").trim().slice(0, 240);
      let incoming = [];
      if (Array.isArray(body.lines) && body.lines.length) {
        incoming = body.lines;
      } else if (Array.isArray(body.matrix) && body.matrix.length) {
        incoming = parseInboundPlanUploadMatrix(body.matrix);
      } else if (body.sheetBase64 && typeof body.sheetBase64 === "string") {
        const buf = Buffer.from(String(body.sheetBase64).trim(), "base64");
        const wb = XLSX.read(buf, { type: "buffer" });
        const name = wb.SheetNames[0];
        if (!name) throw new Error("시트가 없습니다.");
        const ws = wb.Sheets[name];
        const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        incoming = parseInboundPlanUploadMatrix(matrix);
      } else {
        throw new Error("요청에 matrix(2차원 배열), lines, 또는 sheetBase64가 필요합니다.");
      }
      if (incoming.length > 20000) throw new Error("파일 품목 행이 너무 많습니다(한 번에 최대 20000).");
      const existing = db.inboundPlanUpload.lines || [];
      const {
        merged,
        addedLineCount,
        newSlipCount,
        skippedLineCount,
        skippedExistingSlipCount,
        skippedSlipKeysSample
      } = mergeInboundPlanUploadByNewSlipsOnly(existing, incoming);
      if (merged.length > 50000) throw new Error("저장된 총 품목 행이 상한(50000)을 넘습니다. 관리자에게 문의하세요.");
      db.inboundPlanUpload.lines = merged;
      db.inboundPlanUpload.sourceFileName = sourceFileName || db.inboundPlanUpload.sourceFileName || "upload.xlsx";
      db.inboundPlanUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      const items = aggregateInboundPlanUploadLines(merged, db.inboundPlanUpload.slipWorkflow || {});
      return sendJson(res, 200, {
        ok: true,
        lineCount: merged.length,
        slipCount: items.length,
        addedLineCount,
        newSlipCount,
        skippedLineCount,
        skippedExistingSlipCount,
        skippedSlipKeysSample,
        items,
        sourceFileName: db.inboundPlanUpload.sourceFileName,
        uploadedAt: db.inboundPlanUpload.uploadedAt
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/inbound-plan-upload") {
    const lines = db.inboundPlanUpload.lines || [];
    const items = aggregateInboundPlanUploadLines(lines, db.inboundPlanUpload.slipWorkflow || {});
    return sendJson(res, 200, {
      items,
      lineCount: lines.length,
      sourceFileName: db.inboundPlanUpload.sourceFileName || "",
      uploadedAt: db.inboundPlanUpload.uploadedAt || ""
    });
  }

  const handleInboundPlanUploadDelete = async () => {
    try {
      const body = await parseBody(req);
      const slipNos = Array.isArray(body.slipNos) ? body.slipNos : [];
      const keys = new Set(
        slipNos
          .map((x) => normalizeSlipNoText(x))
          .filter(Boolean)
      );
      if (!keys.size) throw new Error("삭제할 발주번호가 없습니다.");
      const before = db.inboundPlanUpload.lines || [];
      const beforeSlipSet = new Set(before.map((row) => normalizeSlipNoText(row.slipNo)).filter(Boolean));
      const remain = before.filter((row) => !keys.has(normalizeSlipNoText(row.slipNo)));
      const deletedLineCount = before.length - remain.length;
      const remainSlipSet = new Set(remain.map((row) => normalizeSlipNoText(row.slipNo)).filter(Boolean));
      let deletedSlipCount = 0;
      for (const k of beforeSlipSet) {
        if (!remainSlipSet.has(k)) deletedSlipCount += 1;
      }
      db.inboundPlanUpload.lines = remain;
      db.inboundPlanUpload.uploadedAt = new Date().toISOString();
      const wf = db.inboundPlanUpload.slipWorkflow;
      if (wf && typeof wf === "object") {
        for (const k of keys) delete wf[k];
      }
      writeDb(db);
      const items = aggregateInboundPlanUploadLines(remain, db.inboundPlanUpload.slipWorkflow || {});
      return sendJson(res, 200, {
        ok: true,
        deletedSlipCount,
        deletedLineCount,
        lineCount: remain.length,
        slipCount: items.length,
        items,
        sourceFileName: db.inboundPlanUpload.sourceFileName || "",
        uploadedAt: db.inboundPlanUpload.uploadedAt
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  };

  if (req.method === "DELETE" && pathname === "/api/inbound-plan-upload") {
    return handleInboundPlanUploadDelete();
  }

  // 일부 환경에서 DELETE body가 누락될 수 있어 POST fallback도 허용
  if (req.method === "POST" && pathname === "/api/inbound-plan-upload/delete") {
    return handleInboundPlanUploadDelete();
  }

  if (req.method === "GET" && pathname === "/api/inbound-plan-upload/detail") {
    const slipNo = String(urlObj.searchParams.get("slipNo") || "").trim();
    if (!slipNo) return sendJson(res, 400, { error: "slipNo가 필요합니다." });
    const key = normalizeSlipNoText(slipNo);
    const lines = db.inboundPlanUpload.lines || [];
    const detailLines = inboundPlanUploadLinesToDetailRows(lines, key);
    if (!detailLines.length) {
      return sendJson(res, 200, { item: null, items: [], source: "upload", workflow: { status: "draft", lineQtyByIndex: {} } });
    }
    const wf =
      (db.inboundPlanUpload.slipWorkflow && db.inboundPlanUpload.slipWorkflow[key]) || {
        status: "draft",
        lineQtyByIndex: {}
      };
    const qMap = wf.lineQtyByIndex && typeof wf.lineQtyByIndex === "object" ? wf.lineQtyByIndex : {};
    const itemsWithQty = detailLines.map((row, idx) => {
      const orderQty = row.orderQty != null ? row.orderQty : row.qty;
      const inboundCheckQty = qMap[String(idx)] != null ? qMap[String(idx)] : orderQty;
      return { ...row, orderQty, inboundCheckQty };
    });
    return sendJson(res, 200, {
      item: itemsWithQty[0],
      items: itemsWithQty,
      source: "upload",
      workflow: {
        status: wf.status || "draft",
        lineQtyByIndex: qMap,
        confirmedAt: wf.confirmedAt || null,
        ecountPurchaseSavedAt: wf.ecountPurchaseSavedAt || null,
        wmsStockInboundAt: wf.wmsStockInboundAt || null
      }
    });
  }

  if (req.method === "POST" && pathname === "/api/inbound-plan-upload/slip-workflow") {
    try {
      const body = await parseBody(req);
      const slipNo = String(body.slipNo || "").trim();
      const action = String(body.action || "").trim();
      if (!slipNo) throw new Error("slipNo가 필요합니다.");
      const key = normalizeSlipNoText(slipNo);
      const group = getUploadSlipLineGroupSorted(db.inboundPlanUpload.lines || [], key);
      if (!group.length) throw new Error("해당 전표의 품목 행이 없습니다.");

      if (!db.inboundPlanUpload.slipWorkflow || typeof db.inboundPlanUpload.slipWorkflow !== "object") {
        db.inboundPlanUpload.slipWorkflow = {};
      }
      const wfPrev = db.inboundPlanUpload.slipWorkflow[key] || { status: "draft", lineQtyByIndex: {} };
      const prevQty =
        wfPrev.lineQtyByIndex && typeof wfPrev.lineQtyByIndex === "object" ? wfPrev.lineQtyByIndex : {};

      if (action === "save-qty") {
        const merged = mergeLineQtyByIndexFromBody(body.lineQtyByIndex || {}, group);
        db.inboundPlanUpload.slipWorkflow[key] = {
          status: "draft",
          lineQtyByIndex: merged,
          confirmedAt: undefined,
          ...slipWorkflowPreservePurchaseFlags(wfPrev)
        };
      } else if (action === "confirm") {
        const merged = mergeLineQtyByIndexFromBody(body.lineQtyByIndex || {}, group);
        for (let i = 0; i < group.length; i++) {
          const n = Number(merged[String(i)]);
          if (!Number.isFinite(n) || n <= 0) {
            throw new Error(`입고 확정 불가: 행 ${i + 1} 입고확인 수량을 1 이상으로 입력하세요.`);
          }
        }
        const stockUser = resolveInboundPlanStockUser(db, body, group);
        assertUploadSlipWmsInboundReady(db, group, merged);
        applyWmsStockInboundFromUploadSlip(db, group, merged, slipNo, stockUser);
        db.inboundPlanUpload.slipWorkflow[key] = {
          status: "confirmed",
          lineQtyByIndex: merged,
          confirmedAt: new Date().toISOString(),
          stockUser,
          ecountPurchaseSavedAt: wfPrev.ecountPurchaseSavedAt,
          wmsStockInboundAt: new Date().toISOString()
        };
      } else if (action === "reopen") {
        const allowReopenAfterEcountCancel = Boolean(body.allowReopenAfterEcountCancel);
        if (wfPrev.ecountPurchaseSavedAt && !allowReopenAfterEcountCancel) {
          throw new Error(
            "이카운트 구매입력 이력이 있습니다. 이카운트 전표를 삭제했다면 확정 취소 시 '구매입력 전표 삭제 완료' 확인 후 다시 시도하세요."
          );
        }
        if (wfPrev.wmsStockInboundAt) {
          const stockUser = resolveInboundPlanRollbackUser(db, body, group, wfPrev);
          rollbackWmsStockInboundFromUploadSlip(db, group, prevQty, slipNo, stockUser);
        }
        db.inboundPlanUpload.slipWorkflow[key] = {
          ...wfPrev,
          status: "draft",
          lineQtyByIndex: { ...prevQty },
          confirmedAt: undefined,
          ecountPurchaseSavedAt: allowReopenAfterEcountCancel ? undefined : wfPrev.ecountPurchaseSavedAt,
          wmsStockInboundAt: undefined
        };
      } else {
        throw new Error("action은 save-qty, confirm, reopen 중 하나여야 합니다.");
      }

      db.inboundPlanUpload.uploadedAt = new Date().toISOString();
      writeDb(db);
      const next = db.inboundPlanUpload.slipWorkflow[key];
      return sendJson(res, 200, { ok: true, workflow: next });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  /* 이카운트 구매입력(SavePurchases) 연결 테스트 — 실제 전표가 생성될 수 있음. ECOUNT_TEST_SAVE_PURCHASES=1 일 때만 허용. */
  if (req.method === "POST" && pathname === "/api/ecount/test-save-purchases") {
    try {
      const enabled = String(process.env.ECOUNT_TEST_SAVE_PURCHASES ?? "").trim().toLowerCase();
      if (!["1", "true", "yes", "on"].includes(enabled)) {
        return sendJson(res, 403, {
          error: "비활성화됨. 테스트 시 .env에 ECOUNT_TEST_SAVE_PURCHASES=1 후 서버 재시작하세요."
        });
      }
      const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();
      if (!comCode) throw new Error("ECOUNT_COM_CODE가 필요합니다.");
      const payload = await parseBody(req);
      if (!payload || typeof payload !== "object") throw new Error("JSON 본문이 필요합니다.");
      if (!Array.isArray(payload.PurchasesList)) throw new Error("본문에 PurchasesList 배열이 필요합니다. (이카운트 SavePurchases와 동일 구조)");
      const sessionId = await fetchEcountSession();
      const apiBase = await fetchEcountApiBase(comCode);
      const result = await postEcountOapiV2(apiBase, sessionId, "Purchases/SavePurchases", payload);
      return sendJson(res, 200, {
        ok: result.ok,
        httpStatus: result.status,
        ecount: result.data,
        rawPreview: String(result.text || "").slice(0, 2000)
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  /*
   * 발주서(slipNo) 조회 → PurchasesList 조립 → SavePurchases (ECOUNT_TEST_SAVE_PURCHASES=1 일 때만)
   * 본문: { slipNo, from, to, dryRun?, ioDate?, custCd?, whCd? }
   */
  if (req.method === "POST" && pathname === "/api/ecount/save-purchases-from-order") {
    try {
      const enabled = String(process.env.ECOUNT_TEST_SAVE_PURCHASES ?? "").trim().toLowerCase();
      if (!["1", "true", "yes", "on"].includes(enabled)) {
        return sendJson(res, 403, {
          error: "비활성화됨. 테스트 시 .env에 ECOUNT_TEST_SAVE_PURCHASES=1 후 서버 재시작하세요."
        });
      }
      const body = await parseBody(req);
      const slipNo = String(body.slipNo || "").trim();
      const from = String(body.from || "").trim();
      const to = String(body.to || "").trim();
      const dryRun = Boolean(body.dryRun);
      const ioDate = body.ioDate != null ? String(body.ioDate) : "";
      const custCd = body.custCd != null ? String(body.custCd) : "";
      const whCd = body.whCd != null ? String(body.whCd) : "";
      if (!slipNo) throw new Error("slipNo(발주 전표번호)가 필요합니다.");
      if (!from || !to) throw new Error("from, to(YYYY-MM-DD) 조회 기간이 필요합니다. 입고예정목록 1과 동일하게 넓게 잡으세요.");

      const pack = await fetchEcountPurchaseOrderMatchedLines(slipNo, from, to);
      if (!pack.matched.length) throw new Error("해당 기간에서 발주를 찾지 못했습니다. 기간·전표번호를 확인하세요.");

      const purchasesPayload = buildSavePurchasesPayloadFromOrder(pack.headerRaw, pack.matched, {
        ioDate,
        custCd,
        whCd
      });

      if (dryRun) {
        return sendJson(res, 200, {
          dryRun: true,
          slipNo,
          lineSource: pack.lineSourceFinal,
          matchedLineCount: pack.matched.length,
          purchasesPayload
        });
      }

      const result = await postEcountOapiV2(
        pack.apiBase,
        pack.sessionId,
        "Purchases/SavePurchases",
        purchasesPayload
      );
      return sendJson(res, 200, {
        ok: result.ok,
        httpStatus: result.status,
        ecount: result.data,
        rawPreview: String(result.text || "").slice(0, 2000),
        lineSource: pack.lineSourceFinal,
        purchasesPayload
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  /** UI용: 구매입력(SavePurchases) 전송에 필요한 .env 준비 여부 (비밀값 노출 없음) */
  if (req.method === "GET" && pathname === "/api/ecount/save-purchases-env") {
    const enabled = String(process.env.ECOUNT_TEST_SAVE_PURCHASES ?? "").trim().toLowerCase();
    const testSavePurchases = ["1", "true", "yes", "on"].includes(enabled);
    const defaultWhCdConfigured = Boolean(String(process.env.ECOUNT_DEFAULT_WH_CD || "").trim());
    const comCodeConfigured = Boolean(String(process.env.ECOUNT_COM_CODE || "").trim());
    return sendJson(res, 200, { testSavePurchases, defaultWhCdConfigured, comCodeConfigured });
  }

  /*
   * 입고예정목록 2 업로드 전표 — 입고 확정(slipWorkflow) 후 SavePurchases (ECOUNT_TEST_SAVE_PURCHASES=1)
   * 본문: { slipNo }
   */
  if (req.method === "POST" && pathname === "/api/ecount/save-purchases-from-upload-slip") {
    try {
      const enabled = String(process.env.ECOUNT_TEST_SAVE_PURCHASES ?? "").trim().toLowerCase();
      if (!["1", "true", "yes", "on"].includes(enabled)) {
        return sendJson(res, 403, {
          error: "비활성화됨. 테스트 시 .env에 ECOUNT_TEST_SAVE_PURCHASES=1 후 서버 재시작하세요."
        });
      }
      const body = await parseBody(req);
      const slipNo = String(body.slipNo || "").trim();
      if (!slipNo) throw new Error("slipNo가 필요합니다.");
      const key = normalizeSlipNoText(slipNo);
      const wf = db.inboundPlanUpload.slipWorkflow && db.inboundPlanUpload.slipWorkflow[key];
      if (!wf || wf.status !== "confirmed") {
        throw new Error("입고 확정된 전표만 구매입력할 수 있습니다. 전표 상세에서 입고 확정을 먼저 하세요.");
      }
      const group = getUploadSlipLineGroupSorted(db.inboundPlanUpload.lines || [], key);
      if (!group.length) throw new Error("품목 행이 없습니다.");
      const lineQtyByIndex = wf.lineQtyByIndex && typeof wf.lineQtyByIndex === "object" ? wf.lineQtyByIndex : {};
      const purchasesPayload = buildSavePurchasesPayloadFromUploadSlip(group, lineQtyByIndex);

      if (wf.ecountPurchaseSavedAt) {
        return sendJson(res, 200, {
          ok: true,
          alreadyComplete: true,
          ecount: null,
          purchasesPayload,
          message: "이미 이카운트 구매입력이 완료된 전표입니다."
        });
      }

      const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();
      if (!comCode) throw new Error("ECOUNT_COM_CODE가 필요합니다.");
      const sessionId = await fetchEcountSession();
      const apiBase = await fetchEcountApiBase(comCode);
      const result = await postEcountOapiV2(apiBase, sessionId, "Purchases/SavePurchases", purchasesPayload);
      if (!result.ok) {
        throw new Error(
          `이카운트 SavePurchases 실패(HTTP ${result.status}). ${String(result.text || "").slice(0, 500)}`
        );
      }
      wf.ecountPurchaseSavedAt = new Date().toISOString();
      writeDb(db);

      return sendJson(res, 200, {
        ok: true,
        httpStatus: result.status,
        ecount: result.data,
        rawPreview: String(result.text || "").slice(0, 2000),
        purchasesPayload
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  /*
   * 입고예정목록 = 이카운트 발주서 조회(발주서조회API) 연동 요약
   * - 목록: POST {apiBase}/OAPI/V2/Purchases/GetPurchasesOrderList?SESSION_ID=...
   * - 요청(문서 기준): SESSION_ID, PROD_CD(선택), CUST_CD(선택), ListParam.BASE_DATE_FROM/TO(YYYYMMDD), PAGE_*
   * - 응답: Data.Result[] 등 (extractArrayFromUnknown) → normalizeInboundPlanRows()에서 필드 매핑
   * - 주요 Result 필드(문서/화면 기준): ORD_NO 발주번호, ORD_DATE 일자, CUST/CUST_DES 거래처,
   *   CUST_NAME 담당자, WH_CD/WH_DES 창고, TIME_DATE 납기일, PROD_DES 품목요약(전표당 1줄일 때 제목성),
   *   QTY 합계수량, TTL_CTT 제목, P_FLAG 상태 등
   * - 프론트: public/app.js renderInboundPlan → GET /api/inbound-plans, 클릭 시 /api/inbound-plans/detail
   * - 품목 라인별 상세: 목록에 라인이 없으면 구매/입고 계열 다른 OAPI 시도
   *   · ECOUNT_INBOUND_LINE_API_PATHS=Purchases/GetX,Inventory/GetY (쉼표·세미콜론, 모듈/메서드 또는 메서드만)
   *   · ECOUNT_INBOUND_DETAIL_API=API명 또는 Module/Method (기존 단일 상세 후보)
   *   · ECOUNT_SKIP_BUILTIN_LINE_APIS — 기본: 내장 후보 호출 안 함. 예전처럼 여러 URL을 자동 시도하려면 =0 (false/no/off)
   */
  if (req.method === "GET" && pathname === "/api/inbound-plans") {
    try {
      const from = String(urlObj.searchParams.get("from") || "").trim();
      const to = String(urlObj.searchParams.get("to") || "").trim();

      // 1) explicit mock mode for UI/testing without ECOUNT credentials
      if (String(process.env.ECOUNT_MOCK_MODE || "").trim() === "1") {
        const items = [
          {
            poNo: "PO-2026-0001",
            poDate: from || "2026-04-01",
            vendor: "샘플거래처A",
            itemCode: "EC-1001",
            itemName: "샘플상품",
            qty: 120,
            dueDate: to || "2026-04-30",
            whName: "유통사업부",
            status: "발주완료"
          }
        ];
        return sendJson(res, 200, { items, source: "mock" });
      }

      // 2) Real ECOUNT mode
      const comCode = String(process.env.ECOUNT_COM_CODE || "").trim();

      const fromCompact = normalizeDateCompact(from);
      const toCompact = normalizeDateCompact(to);
      if (!fromCompact || !toCompact) {
        throw new Error("조회 기간 형식이 올바르지 않습니다. (YYYY-MM-DD)");
      }
      const range = clampDateRangeToMax30Days(fromCompact, toCompact);

      const sessionId = await fetchEcountSession();
      const apiBase = await fetchEcountApiBase(comCode);
      const url = `${apiBase}/OAPI/V2/Purchases/GetPurchasesOrderList?SESSION_ID=${encodeURIComponent(sessionId)}`;
      const payload = {
        PROD_CD: "",
        CUST_CD: "",
        ListParam: {
          BASE_DATE_FROM: range.fromCompact,
          BASE_DATE_TO: range.toCompact,
          PAGE_CURRENT: 1,
          PAGE_SIZE: 100
        }
      };

      const callEcountList = async () =>
        fetch(url, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(payload)
        });

      let r = await callEcountList();

      if (!r.ok) {
        const txt = await r.text();
        // session can expire/revoke; clear cache for retry on next call
        ecountSessionCache = { value: "", expiresAt: 0 };
        throw new Error(`발주서 조회 실패(${r.status}): ${txt.slice(0, 300)}`);
      }
      let data = await r.json();
      const bodyMsg = String(data?.Data?.Message || data?.Message || data?.Error?.Message || "");
      const needsRetry =
        bodyMsg.includes("세션") ||
        bodyMsg.toLowerCase().includes("session") ||
        String(data?.Status || "") === "401";
      if (needsRetry) {
        ecountSessionCache = { value: "", expiresAt: 0 };
        const freshSession = await fetchEcountSession();
        const retryUrl = `${apiBase}/OAPI/V2/Purchases/GetPurchasesOrderList?SESSION_ID=${encodeURIComponent(freshSession)}`;
        r = await fetch(retryUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "application/json" },
          body: JSON.stringify(payload)
        });
        if (!r.ok) {
          const txt = await r.text();
          throw new Error(`발주서 재조회 실패(${r.status}): ${txt.slice(0, 300)}`);
        }
        data = await r.json();
      }
      const rawRows = extractArrayFromUnknown(data);
      const items = normalizeInboundPlanRows(rawRows, range.toCompact);
      const topKeys = data && typeof data === "object" ? Object.keys(data).slice(0, 20) : [];
      const dataKeys =
        data && data.Data && typeof data.Data === "object" ? Object.keys(data.Data).slice(0, 20) : [];
      const dataType = data?.Data === null ? "null" : Array.isArray(data?.Data) ? "array" : typeof data?.Data;
      const dataPreview =
        data?.Data && typeof data.Data === "object" ? JSON.stringify(data.Data).slice(0, 500) : String(data?.Data || "");
      const errorPreview = data?.Error ? JSON.stringify(data.Error).slice(0, 500) : "";
      const errorsPreview = data?.Errors ? JSON.stringify(data.Errors).slice(0, 500) : "";
      return sendJson(res, 200, {
        items,
        source: "ecount-live",
        debug: {
          countRaw: Array.isArray(rawRows) ? rawRows.length : 0,
          status: data?.Status,
          code: data?.Data?.Code || data?.Code || "",
          message: data?.Data?.Message || data?.Message || "",
          topKeys,
          dataKeys,
          dataType,
          dataPreview,
          errorPreview,
          errorsPreview,
          adjustedDateRange: range.adjusted ? { from: range.fromCompact, to: range.toCompact } : null
        }
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/inbound-plans/detail") {
    try {
      const slipNo = String(urlObj.searchParams.get("slipNo") || "").trim();
      const from = String(urlObj.searchParams.get("from") || "").trim();
      const to = String(urlObj.searchParams.get("to") || "").trim();
      if (!slipNo) throw new Error("전표번호(slipNo)가 필요합니다.");

      const pack = await fetchEcountPurchaseOrderMatchedLines(slipNo, from, to);
      const {
        matched,
        headerRaw,
        rawRows,
        usedEndpoint,
        lineSourceFinal,
        nestedLineCount,
        narrowListTried,
        endpointErrors,
        skipBuiltinLineApis
      } = pack;

      let hint = "";
      if (lineSourceFinal === "list-summary-fallback") {
        const parts = [
          "발주 목록 API만으로는 품목 라인(행별 PROD_CD 등)이 없습니다. 이카운트 개발자센터 OAPI에서 「한 전표·다품목」 조회 API의 정확한 Module/Method 를 확인한 뒤 .env에 ECOUNT_INBOUND_LINE_API_PATHS 또는 ECOUNT_INBOUND_DETAIL_API 로 지정하세요. debug.headerSummary에 라인 배열 후보 키가 보이면 알려주시면 매핑을 이어갈 수 있습니다."
        ];
        if (endpointErrors.length) {
          parts.push(
            "endpointErrors: (404)·EXP00001·No HTTP resource 는 해당 존에 없는 URL, (412)·HTML 은 호출 제한·차단 가능성입니다."
          );
          if (!skipBuiltinLineApis) {
            parts.push(
              "시도가 많다면 .env의 ECOUNT_SKIP_BUILTIN_LINE_APIS=0 을 제거하고 서버를 재시작하세요(기본은 내장 후보 비활성)."
            );
          }
        } else if (skipBuiltinLineApis) {
          parts.push(
            "내장 후보 API는 호출하지 않은 상태입니다(endpointErrors 비어 있음). 품목별 행은 문서에 맞는 경로를 ECOUNT_INBOUND_LINE_API_PATHS 등에 넣어야 합니다."
          );
        }
        hint = parts.join(" ");
      }

      return sendJson(res, 200, {
        item: matched[0] || null,
        items: matched,
        source: "ecount-live-detail",
        debug: {
          requestedSlipNo: slipNo,
          countRaw: Array.isArray(rawRows) ? rawRows.length : 0,
          countMatched: matched.length,
          usedEndpoint,
          lineSource: lineSourceFinal,
          hasHeaderRaw: Boolean(headerRaw),
          nestedLineCount,
          narrowListTried,
          headerSummary: buildHeaderDebugSummary(headerRaw),
          customDetailApi: String(process.env.ECOUNT_INBOUND_DETAIL_API || "").trim() || null,
          inboundLineApiPaths: parseInboundLineApiPathSpecs(process.env.ECOUNT_INBOUND_LINE_API_PATHS || ""),
          skipBuiltinLineApis,
          hint,
          endpointErrorAttempts: endpointErrors.length,
          endpointErrors: compressEndpointErrorsForDebug(endpointErrors, 14)
        }
      });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "GET" && pathname === "/api/dashboard") {
    const stockItems = calculateStock(db);
    const lowStock = stockItems.filter((x) => x.stock <= x.safetyStock).length;
    const totalProducts = db.products.length;
    const today = new Date().toISOString().slice(0, 10);
    const todayMoves = db.movements.filter((m) => (m.createdAt || "").startsWith(today)).length;
    return sendJson(res, 200, {
      totalProducts,
      lowStock,
      todayMoves
    });
  }

  // ── 미출 대조 ────────────────────────────────────────────────────
  if (req.method === "GET" && pathname === "/api/unship-compare") {
    const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
    if (!partnerType) return sendJson(res, 400, { error: "partnerType 필요" });
    const startDate = String(urlObj.searchParams.get("startDate") || "").trim();
    const endDate = String(urlObj.searchParams.get("endDate") || "").trim();

    // WMS 미출 집계 (dueDate + productCode + centerName)
    const wmsLines = (db.outboundOrderUpload.lines || []).filter(
      (x) => x.partnerType === partnerType && String(x.unshipStatus || "").trim() === "미출"
    );
    const wmsMap = new Map();
    for (const x of wmsLines) {
      const dd = String(x.dueDate || "").trim();
      if (startDate && dd < startDate) continue;
      if (endDate && dd > endDate) continue;
      const k = `${dd}|||${String(x.productCode || "").trim()}|||${String(x.centerName || "").trim()}`;
      if (!wmsMap.has(k)) wmsMap.set(k, { dueDate: dd, productCode: String(x.productCode || "").trim(), productName: String(x.productName || "").trim(), centerName: String(x.centerName || "").trim(), wmsQty: 0 });
      wmsMap.get(k).wmsQty += Number(x.orderQty) || 0;
    }

    // 이마트 공식 미출 집계
    const officialRows = (db.unshipCompare[partnerType]?.officialRows || []);
    const officialMap = new Map();
    for (const r of officialRows) {
      const dd = String(r.dueDate || "").trim();
      if (startDate && dd < startDate) continue;
      if (endDate && dd > endDate) continue;
      const k = `${dd}|||${String(r.productCode || "").trim()}|||${String(r.centerName || "").trim()}`;
      if (!officialMap.has(k)) officialMap.set(k, { dueDate: dd, productCode: String(r.productCode || "").trim(), productName: String(r.productName || "").trim(), centerName: String(r.centerName || "").trim(), officialQty: 0 });
      officialMap.get(k).officialQty += Number(r.officialQty) || 0;
    }

    // 조정 기록 (최신 1건)
    const adjMap = new Map();
    for (const a of (db.unshipCompare[partnerType]?.adjustments || [])) {
      const k = `${String(a.dueDate||"").trim()}|||${String(a.productCode||"").trim()}|||${String(a.centerName||"").trim()}`;
      if (!adjMap.has(k) || a.appliedAt > adjMap.get(k).appliedAt) adjMap.set(k, a);
    }

    // 전체 키 합집합
    const allKeys = new Set([...wmsMap.keys(), ...officialMap.keys()]);
    const rows = [];
    for (const k of allKeys) {
      const wms = wmsMap.get(k) || {};
      const off = officialMap.get(k) || {};
      const adj = adjMap.get(k) || null;
      const dueDate = wms.dueDate || off.dueDate || "";
      const productCode = wms.productCode || off.productCode || "";
      const productName = wms.productName || off.productName || "";
      const centerName = wms.centerName || off.centerName || "";
      const wmsQty = wms.wmsQty || 0;
      const officialQty = off.officialQty || 0;
      const diff = officialQty - wmsQty;
      rows.push({ dueDate, productCode, productName, centerName, wmsQty, officialQty, diff, adjustment: adj });
    }
    rows.sort((a, b) => {
      const dd = String(a.dueDate || "").localeCompare(String(b.dueDate || ""));
      if (dd !== 0) return dd;
      return String(a.productCode || "").localeCompare(String(b.productCode || ""), "ko");
    });

    const hasOfficial = officialRows.length > 0;
    return sendJson(res, 200, { rows, hasOfficial, officialRowCount: officialRows.length });
  }

  if (req.method === "POST" && pathname === "/api/unship-compare/upload") {
    try {
      const partnerType = normalizeOutboundPartnerType(urlObj.searchParams.get("partnerType") || "");
      if (!partnerType) throw new Error("partnerType 필요");
      const body = await parseBody(req);
      const rows = Array.isArray(body.rows) ? body.rows : [];
      if (!rows.length) throw new Error("업로드할 데이터가 없습니다.");
      const now = new Date().toISOString();
      const parsed = rows.map((r) => ({
        dueDate: String(r.dueDate || "").trim(),
        productCode: String(r.productCode || "").trim(),
        productName: String(r.productName || "").trim(),
        centerName: String(r.centerName || "").trim(),
        officialQty: Math.max(0, Math.trunc(Number(String(r.officialQty || "0").replace(/,/g, "")) || 0)),
        uploadedAt: now
      })).filter((r) => r.productCode && r.officialQty > 0);
      if (!parsed.length) throw new Error("유효한 행이 없습니다. 상품코드·미출수량을 확인해주세요.");
      db.unshipCompare[partnerType].officialRows = parsed;
      writeDb(db);
      return sendJson(res, 200, { ok: true, count: parsed.length });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/unship-compare/adjust") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType || "");
      if (!partnerType) throw new Error("partnerType 필요");
      const { dueDate, productCode, productName, centerName, action, newQty, note } = body;
      if (!dueDate || !productCode || !action) throw new Error("dueDate, productCode, action 필요");
      const validActions = ["modify_qty", "add_product", "confirmed", "none"];
      if (!validActions.includes(action)) throw new Error("action 값 오류");
      const now = new Date().toISOString();
      const adj = {
        id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        dueDate: String(dueDate).trim(),
        productCode: String(productCode).trim(),
        productName: String(productName || "").trim(),
        centerName: String(centerName || "").trim(),
        action,
        newQty: action !== "confirmed" ? Math.max(0, Math.trunc(Number(newQty) || 0)) : null,
        note: String(note || "").trim(),
        appliedAt: now
      };
      db.unshipCompare[partnerType].adjustments = db.unshipCompare[partnerType].adjustments || [];
      // 이전 동일 키 기록 유지 + 신규 추가
      db.unshipCompare[partnerType].adjustments.push(adj);
      writeDb(db);
      return sendJson(res, 200, { ok: true, adj });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  if (req.method === "POST" && pathname === "/api/unship-compare/clear") {
    try {
      const body = await parseBody(req);
      const partnerType = normalizeOutboundPartnerType(body.partnerType || "");
      if (!partnerType) throw new Error("partnerType 필요");
      db.unshipCompare[partnerType] = { officialRows: [], adjustments: [] };
      writeDb(db);
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 400, { error: e.message });
    }
  }

  // ── eCvan 자격증명 저장/조회 ────────────────────────────────────
  if (req.method === "GET" && pathname === "/api/ecvan/credentials") {
    const db2 = readDb();
    return sendJson(res, 200, { id: db2.ecvanCredentials?.id || "", hasPw: !!(db2.ecvanCredentials?.pw) });
  }
  if (req.method === "POST" && pathname === "/api/ecvan/credentials") {
    try {
      const body = await parseBody(req);
      const db2 = readDb();
      if (!db2.ecvanCredentials) db2.ecvanCredentials = {};
      if (body.id !== undefined) db2.ecvanCredentials.id = String(body.id).trim();
      if (body.pw !== undefined) db2.ecvanCredentials.pw = String(body.pw);
      writeDb(db2);
      return sendJson(res, 200, { ok: true });
    } catch (e) { return sendJson(res, 400, { error: e.message }); }
  }

  // ── eCvan 자동수집 ───────────────────────────────────────────────
  if (req.method === "POST" && pathname === "/api/ecvan/start") {
    try {
      const body = await parseBody(req);
      const db2 = readDb();
      const ecvanId = String(body.id || db2.ecvanCredentials?.id || "").trim();
      const ecvanPw = String(body.pw || db2.ecvanCredentials?.pw || "").trim();
      if (!ecvanId || !ecvanPw) throw new Error("아이디와 비밀번호를 입력해주세요.");

      // 자격증명 저장
      db2.ecvanCredentials = { id: ecvanId, pw: ecvanPw };
      writeDb(db2);

      let pup;
      try { pup = require("puppeteer-core"); } catch { throw new Error("puppeteer-core가 설치되어 있지 않습니다. npm install puppeteer-core 실행 후 재시작해주세요."); }

      // Chrome 경로 탐색
      const chromePaths = [
        process.env.CHROME_PATH,
        "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
        String(process.env.LOCALAPPDATA || "") + "\\Google\\Chrome\\Application\\chrome.exe",
        "/usr/bin/google-chrome", "/usr/bin/chromium-browser", "/usr/bin/chromium"
      ].filter(Boolean);
      const chromePath = chromePaths.find(p => { try { return fs.existsSync(p); } catch { return false; } });
      if (!chromePath) throw new Error("Chrome을 찾을 수 없습니다. .env에 CHROME_PATH=경로 를 추가해주세요.");

      const sessionId = `ecvan-${Date.now()}`;

      // 매번 새 세션으로 시작 (OTP 항상 요구)
      const userDataDir = path.join(__dirname, "ecvan-chrome-profile");
      try { fs.rmSync(userDataDir, { recursive: true, force: true }); } catch {}
      fs.mkdirSync(userDataDir, { recursive: true });

      const browser = await pup.launch({
        executablePath: chromePath,
        headless: false,
        userDataDir,
        defaultViewport: null,
        args: [
          "--no-sandbox", "--disable-setuid-sandbox",
          "--lang=ko-KR,ko", "--disable-blink-features=AutomationControlled",
          "--start-maximized",
        ],
      });
      const page = await browser.newPage();
      await page.setExtraHTTPHeaders({ "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8" });

      const session = { browser, page, status: "logging_in", progress: "로그인 페이지 접속 중...", data: null, error: null, startedAt: new Date().toISOString() };
      ecvanSessions.set(sessionId, session);

      // ── 수집 로직 (OTP 후 또는 직접 로그인 후 공통) ──
      session.collect = async () => {
        const _sleep = (ms) => new Promise(r => setTimeout(r, ms));
        const _click = async (text) => page.evaluate((t) => {
          const el = [...document.querySelectorAll("button,a,input[type=button],input[type=submit]")]
            .find(e => (e.textContent || e.value || "").trim() === t);
          if (el) { el.click(); return true; } return false;
        }, text);
        const _closePopups = async () => {
          for (let i = 0; i < 3; i++) {
            await _sleep(500);
            const closed = await page.evaluate(() => {
              const keywords = ["닫기", "확인", "오늘 하루 보지 않기", "7일동안 보지 않기"];
              const btns = [...document.querySelectorAll("button,a,input[type=button]")];
              let any = false;
              btns.forEach(b => {
                const t = (b.textContent || b.value || "").trim();
                if (keywords.some(k => t === k || t.includes(k))) { b.click(); any = true; }
              });
              return any;
            }).catch(() => false);
            if (!closed) break;
          }
        };

        session.status = "navigating"; session.progress = "팝업 닫는 중...";
        await _closePopups();
        await _sleep(500);
        session.progress = "미납사유 메뉴로 이동 중...";

        // 납품 메뉴 클릭
        await page.evaluate(() => {
          const el = [...document.querySelectorAll("a")].find(a => a.textContent.trim() === "납품");
          if (el) el.click();
        });
        await _sleep(1200);

        // 미납사유등록 클릭
        await page.evaluate(() => {
          const el = [...document.querySelectorAll("a")].find(a => a.textContent.trim() === "미납사유등록");
          if (el) el.click();
        });
        await page.waitForNavigation({ waitUntil: "networkidle2", timeout: 20000 }).catch(() => {});
        await _sleep(1500);

        await _click("상품별"); await _sleep(500);
        await _click("조회"); await _sleep(2500);

        session.status = "collecting"; session.progress = "상품 목록 수집 중...";

        const products = await page.evaluate(() => {
          const rows = [...document.querySelectorAll("table tbody tr")];
          return rows.map(tr => {
            const tds = tr.querySelectorAll("td");
            const link = tr.querySelector("a");
            if (!link || tds.length < 3) return null;
            return { code: tds[1]?.textContent.trim() || "", name: tds[2]?.textContent.trim() || "", href: link.href };
          }).filter(r => r && r.code && r.href);
        });

        if (!products.length) {
          session.status = "done"; session.progress = "미납 상품이 없습니다."; session.data = [];
          browser.close().catch(() => {}); return;
        }

        session.progress = `${products.length}개 상품 발견, 상세 수집 중...`;
        const now = new Date().toISOString();
        const allRows = [];

        for (let i = 0; i < products.length; i++) {
          const prod = products[i];
          session.progress = `(${i + 1}/${products.length}) ${prod.name || prod.code} 수집 중...`;
          try {
            await page.goto(prod.href, { waitUntil: "networkidle2", timeout: 20000 });
            await _sleep(1000);
            let pageNo = 1;
            while (true) {
              const rows = await page.evaluate((prod, now) => {
                const tables = [...document.querySelectorAll("table")];
                const tbl = tables.find(t => t.textContent.includes("점포명") && t.textContent.includes("미납"));
                if (!tbl) return [];
                const n = (td) => parseInt((td?.textContent || "0").replace(/[^\d]/g, "")) || 0;
                return [...tbl.querySelectorAll("tbody tr")].map(tr => {
                  const tds = tr.querySelectorAll("td");
                  if (tds.length < 8) return null;
                  return {
                    productCode: prod.code, productName: prod.name,
                    centerName: tds[1]?.textContent.trim() || "",
                    storeCode: tds[2]?.textContent.trim() || "",
                    orderDate: tds[3]?.textContent.trim() || "",
                    dueDate: tds[4]?.textContent.trim() || "",
                    orderQty: n(tds[5]), storeInQty: n(tds[7]),
                    officialQty: n(tds[9]) || (n(tds[7]) === 0 ? n(tds[9]) : n(tds[5]) - n(tds[7])),
                    uploadedAt: now
                  };
                }).filter(r => r && r.productCode);
              }, prod, now);
              allRows.push(...rows);
              const hasNext = await page.evaluate(() => {
                const pager = document.querySelector(".paging, .pagination, [class*='pager']");
                if (!pager) return false;
                const next = [...pager.querySelectorAll("a, button")].find(b =>
                  b.textContent.trim() === ">" || b.textContent.trim() === "다음" || b.classList.contains("next")
                );
                return next && !next.disabled && !next.classList.contains("disabled");
              });
              if (!hasNext) break;
              await page.evaluate(() => {
                const pager = document.querySelector(".paging, .pagination, [class*='pager']");
                const next = pager && [...pager.querySelectorAll("a, button")].find(b =>
                  b.textContent.trim() === ">" || b.textContent.trim() === "다음" || b.classList.contains("next")
                );
                if (next) next.click();
              });
              await _sleep(1200); pageNo++;
            }
            await page.goBack({ waitUntil: "networkidle2", timeout: 15000 }).catch(() => {});
            await _sleep(700);
          } catch (e2) { console.error(`[eCvan] ${prod.code} 수집 오류:`, e2.message); }
        }

        session.data = allRows;
        const db2 = readDb();
        db2.unshipCompare = db2.unshipCompare || {};
        db2.unshipCompare.emart = db2.unshipCompare.emart || {};
        db2.unshipCompare.emart.officialRows = allRows;
        db2.unshipCompare.emart.uploadedAt = now;
        writeDb(db2);
        session.status = "done";
        session.progress = `수집 완료: ${allRows.length}건`;
        browser.close().catch(() => {});
      };

      const sleep = (ms) => new Promise(r => setTimeout(r, ms));

      // input에 값을 채우는 함수
      const fillInput = async (el, value) => {
        await el.click({ clickCount: 3 });  // 기존 값 전체 선택
        await sleep(50);
        await el.type(value, { delay: 40 }); // 요소에 직접 타이핑
        await sleep(80);
      };

      // 여러 selector 중 첫 번째로 찾은 element 반환 (iframe 포함)
      let activeFrame = page;
      const findEl = async (selectors, timeout = 8000) => {
        const deadline = Date.now() + timeout;
        while (Date.now() < deadline) {
          // 모든 frame (메인 + iframe) 탐색
          const frames = [page, ...page.frames().filter(f => f !== page)];
          for (const frame of frames) {
            for (const sel of selectors) {
              try {
                const el = await frame.$(sel);
                if (el) { activeFrame = frame; return el; }
              } catch {}
            }
          }
          await sleep(300);
        }
        return null;
      };

      // 비동기 로그인
      (async () => {
        try {
          page.on("dialog", async dlg => {
            try { await sleep(150); await dlg.accept(); } catch {}
          });

          const screenshotDir = path.join(__dirname, "ecvan-debug");
          if (!fs.existsSync(screenshotDir)) fs.mkdirSync(screenshotDir);
          const snap = async (name) => page.screenshot({ path: path.join(screenshotDir, `${name}.png`), fullPage: true }).catch(() => {});

          session.progress = "eCvan 접속 중... (크롬 창이 열립니다)";
          await page.goto("https://emart.ecvan.co.kr/pre_ecvan_ui/index.jsp", {
            waitUntil: "domcontentloaded", timeout: 30000
          });
          await sleep(2000);

          // 페이지 로드 후 팝업 닫기 - puppeteer 실제 클릭 사용
          session.progress = "팝업 닫는 중...";
          for (let i = 0; i < 5; i++) {
            await sleep(600);
            // 팝업 닫기 버튼 후보 (텍스트 또는 X 아이콘)
            const closeEl = await page.evaluateHandle(() => {
              const btns = [...document.querySelectorAll("button,a,input[type=button],span[role='button'],.close,svg")];
              const byText = btns.find(b => {
                const t = (b.textContent || b.value || "").trim();
                return t === "닫기" || t === "×" || t === "✕" || t === "X" || t === "close";
              });
              if (byText) return byText;
              // X 모양 버튼 찾기 (클래스에 close, btn-close 포함)
              return document.querySelector(".close, .btn-close, [class*='close'], [class*='Close'], [aria-label*='닫'], [title*='닫']");
            });
            const el = closeEl.asElement();
            if (el) {
              await el.click().catch(() => {});
            } else break;
          }
          await sleep(300);
          await snap("01-loaded");

          session.progress = "로그인 폼 탐색 중...";
          // 로그인 폼이 바로 없으면 로그인 링크/버튼 클릭 후 다시 탐색
          let idEl = await findEl(
            ["#id", "input[name='id']", "input[name='userId']", "input[placeholder*='아이디']", "input[placeholder*='ID']"],
            3000
          );
          if (!idEl) {
            // 로그인 링크 클릭 시도
            await page.evaluate(() => {
              const el = [...document.querySelectorAll("a,button")].find(e => {
                const t = (e.textContent || "").trim();
                return t === "로그인" || t === "LOGIN" || t === "Log In";
              });
              if (el) el.click();
            });
            await sleep(1500);
            idEl = await findEl(
              ["#id", "input[name='id']", "input[name='userId']", "input[placeholder*='아이디']", "input[placeholder*='ID']", "input[type='text']"],
              8000
            );
          }
          if (!idEl) {
            await snap("err-no-id");
            const bodyText = await page.evaluate(() => document.body.innerText.slice(0, 300)).catch(() => "");
            throw new Error(`아이디 입력창을 찾지 못했습니다. 페이지: ${bodyText}`);
          }

          session.progress = "아이디 입력 중...";

          // ID 입력
          await idEl.click({ clickCount: 3 });
          await sleep(50);
          await page.keyboard.type(ecvanId, { delay: 20 });
          await sleep(200);
          await snap("02-id-filled");

          // PW 필드 - 모든 frame 재탐색 (activeFrame 포함)
          session.progress = "비밀번호 입력 중...";
          const pwSels = ["#pw", "input[name='pw']", "input[name='password']", "input[type='password']"];
          let pwEl = null;
          const allFramesForPw = [page, ...page.frames().filter(f => f !== page)];
          for (const frame of allFramesForPw) {
            for (const sel of pwSels) {
              try {
                const el = await frame.$(sel);
                if (el) { pwEl = el; activeFrame = frame; break; }
              } catch {}
            }
            if (pwEl) break;
          }
          if (!pwEl) {
            const bodyText = await page.evaluate(() => document.body?.innerText?.slice(0, 300) || "").catch(() => "");
            throw new Error(`비밀번호 입력창을 찾지 못했습니다. 현재 화면: ${bodyText}`);
          }

          // PW 입력 - 방법1: click + keyboard.type (에러 무시)
          let pwVal = "";
          try {
            await pwEl.click({ clickCount: 3 });
            await sleep(150);
            await page.keyboard.type(ecvanPw, { delay: 30 });
            await sleep(200);
            pwVal = await pwEl.evaluate(el => el.value).catch(() => "");
          } catch (_) {}
          session.progress = `비밀번호 시도1: ${pwVal ? `성공(${pwVal.length}자)` : "실패→시도2"}`;

          if (!pwVal) {
            // 방법2: elementHandle.type() - 내부적으로 focus 후 키 입력
            try {
              await pwEl.type(ecvanPw, { delay: 30 });
              await sleep(200);
              pwVal = await pwEl.evaluate(el => el.value).catch(() => "");
            } catch (_) {}
            session.progress = `비밀번호 시도2: ${pwVal ? `성공(${pwVal.length}자)` : "실패→시도3"}`;
          }

          if (!pwVal) {
            // 방법3: JS evaluate로 직접 value 설정 (React/Vue 호환)
            try {
              await pwEl.evaluate((el, pw) => {
                const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
                if (setter) setter.call(el, pw); else el.value = pw;
                el.dispatchEvent(new Event("input", { bubbles: true }));
                el.dispatchEvent(new Event("change", { bubbles: true }));
              }, ecvanPw);
              await sleep(100);
              pwVal = await pwEl.evaluate(el => el.value).catch(() => "");
            } catch (_) {}
            session.progress = `비밀번호 시도3: ${pwVal ? `성공(${pwVal.length}자)` : "실패"}`;
          }

          await snap("03-pw-filled");
          if (!pwVal) throw new Error("비밀번호 입력 실패 (3가지 방법 모두 값이 설정되지 않았습니다)");

          // 모든 frame 텍스트 수집 (여러 곳에서 사용)
          const getAllText = async () => {
            const texts = await Promise.all(
              [page, ...page.frames()].map(f => f.evaluate(() => document.body?.innerText || "").catch(() => ""))
            );
            return texts.join("\n");
          };

          session.progress = "로그인 버튼 클릭 중...";
          const loginUrl = page.url();

          // 로그인 버튼 찾기 - 텍스트 공백 무시, 여러 방법 시도
          const loginBtnEl = await activeFrame.evaluateHandle(() => {
            const all = [...document.querySelectorAll("button,input[type='submit'],input[type='button'],a")];
            return all.find(e => {
              const t = (e.textContent || e.value || "").replace(/\s+/g, "").trim();
              return t === "로그인" || t === "LOGIN" || t === "login";
            }) || null;
          });
          const loginEl = loginBtnEl.asElement();
          if (loginEl) {
            await loginEl.click();
            session.progress = "로그인 버튼 클릭 완료, 페이지 이동 대기 중...";
          } else {
            // 버튼을 못 찾으면 PW 필드에서 Enter
            await page.keyboard.press("Enter");
            session.progress = "Enter 키로 로그인 시도 중...";
          }

          // 페이지가 실제로 이동했는지 확인 (최대 8초)
          let navigated = false;
          for (let i = 0; i < 8; i++) {
            await sleep(1000);
            const currentUrl = page.url();
            if (currentUrl !== loginUrl) { navigated = true; break; }
          }
          await snap("04-after-login");

          if (!navigated) {
            // URL이 안 바뀌었으면 로그인 버튼 클릭이 실패한 것 → Enter로 재시도
            session.progress = "페이지 이동 없음, Enter로 재시도...";
            await page.keyboard.press("Enter");
            for (let i = 0; i < 8; i++) {
              await sleep(1000);
              if (page.url() !== loginUrl) { navigated = true; break; }
            }
          }

          if (!navigated) {
            const bodyText = await getAllText();
            throw new Error(`로그인 후 페이지 이동이 없습니다. 로그인 버튼을 찾지 못했거나 자격증명 오류. 화면: ${bodyText.slice(0, 200)}`);
          }

          session.progress = "페이지 이동 확인, SMS 인증 화면 대기 중...";

          // 팝업 닫기 (닫기/확인/오늘하루보지않기 등 버튼 모두 클릭)
          const closePopups = async () => {
            await page.evaluate(() => {
              const keywords = ["닫기", "닫 기", "확인", "close", "Close", "오늘 하루 보지 않기", "7일동안 보지 않기"];
              const btns = [...document.querySelectorAll("button,a,input[type=button]")];
              btns.forEach(b => {
                const t = (b.textContent || b.value || "").trim();
                if (keywords.some(k => t === k || t.includes(k))) b.click();
              });
            }).catch(() => {});
          };

          // OTP 화면 대기: URL이 이미 바뀐 상태이므로 텍스트와 OTP 입력란 존재 여부 확인
          let otpReached = false;
          const deadline = Date.now() + 45000;
          while (Date.now() < deadline) {
            await closePopups();
            const t = await getAllText();
            // OTP 입력 필드가 실제로 있는지 확인 (6자리 숫자 입력란)
            const hasOtpInput = await page.$("input[maxlength='6'],input[maxlength='8'],input[name*='otp'],input[name*='auth'],input[id*='otp'],input[id*='auth'],input[id*='certNo'],input[name*='certNo']").catch(() => null);
            if (hasOtpInput || t.includes("인증번호를 입력") || t.includes("인증번호 입력") || t.includes("휴대폰 인증") || t.includes("인증코드를 입력")) {
              otpReached = true; break;
            }
            await sleep(1000);
          }

          await snap("05-otp-check");

          if (!otpReached) {
            const bodyText = await getAllText();
            const isLoginFail = bodyText.includes("비밀번호") && (bodyText.includes("오류") || bodyText.includes("실패") || bodyText.includes("틀") || bodyText.includes("잘못"));
            throw new Error(isLoginFail
              ? `아이디/비밀번호 오류. 화면: ${bodyText.slice(0, 150)}`
              : `SMS 인증 화면이 나타나지 않았습니다. 현재 화면: ${bodyText.slice(0, 200)}`
            );
          }

          await sleep(500);
          session.status = "waiting_otp";
          session.progress = "휴대폰으로 발송된 SMS 인증번호를 WMS에 입력해주세요";
        } catch (e) {
          const url = await page.url().catch(() => "");
          session.status = "error";
          session.error = `${e.message}${url ? ` [URL: ${url}]` : ""}`;
          browser.close().catch(() => {});
          // 세션은 클라이언트가 에러를 읽을 수 있도록 유지 (30분 후 자동 정리)
        }
      })();

      return sendJson(res, 200, { sessionId });
    } catch (e) { return sendJson(res, 400, { error: e.message }); }
  }

  if (req.method === "POST" && pathname === "/api/ecvan/otp") {
    try {
      const body = await parseBody(req);
      const sessionId = String(body.sessionId || "");
      const otp = String(body.otp || "").trim();
      const session = ecvanSessions.get(sessionId);
      if (!session) throw new Error("세션이 없거나 만료되었습니다. 다시 시작해주세요.");
      if (session.status !== "waiting_otp") throw new Error(`현재 상태: ${session.status}`);
      if (!otp) throw new Error("인증번호를 입력해주세요.");

      const { page, browser } = session;
      const sleep = (ms) => new Promise(r => setTimeout(r, ms));
      const clickByText = async (text, sel = "button,a,input[type=button],input[type=submit]") => {
        return page.evaluate((t, s) => {
          const el = [...document.querySelectorAll(s)].find(e => (e.textContent || e.value || "").trim() === t);
          if (el) { el.click(); return true; } return false;
        }, text, sel);
      };

      session.status = "submitting_otp"; session.progress = "인증번호 제출 중...";

      (async () => {
        try {
          // OTP 입력
          const otpSel = "input[placeholder*='인증번호'], input[name*='cert'], input[name*='otp'], input[id*='cert']";
          try { await page.waitForSelector(otpSel, { timeout: 5000 }); } catch {}
          const filled = await page.evaluate((otp, sel) => {
            const el = document.querySelector(sel);
            if (!el) { // 여러 text input 중 첫 번째 빈 것 시도
              const inputs = [...document.querySelectorAll("input[type=text], input[type=number]")];
              const target = inputs.find(i => !i.value);
              if (target) { target.focus(); target.value = otp; target.dispatchEvent(new Event("input", {bubbles:true})); return true; }
              return false;
            }
            el.focus(); el.value = otp; el.dispatchEvent(new Event("input", {bubbles:true})); return true;
          }, otp, otpSel);
          if (!filled) throw new Error("인증번호 입력창을 찾을 수 없습니다.");

          await sleep(300);
          await clickByText("확인");

          // 확인 후 팝업 닫기 (닫힐 때까지만, 최대 3회)
          session.progress = "팝업 닫는 중...";
          for (let i = 0; i < 3; i++) {
            await sleep(600);
            const closed = await page.evaluate(() => {
              const keywords = ["닫기", "확인", "오늘 하루 보지 않기", "7일동안 보지 않기"];
              const btns = [...document.querySelectorAll("button,a,input[type=button]")];
              let any = false;
              btns.forEach(b => {
                const t = (b.textContent || b.value || "").trim();
                if (keywords.some(k => t === k || t.includes(k))) { b.click(); any = true; }
              });
              return any;
            }).catch(() => false);
            if (!closed) break;
          }

          await session.collect();
        } catch (e) {
          const url = await page.url().catch(() => "");
          session.status = "error";
          session.error = `${e.message}${url ? ` [URL: ${url}]` : ""}`;
          browser.close().catch(() => {});
        }
      })();

      return sendJson(res, 200, { ok: true });
    } catch (e) { return sendJson(res, 400, { error: e.message }); }
  }

  if (req.method === "GET" && pathname === "/api/ecvan/status") {
    const sessionId = urlObj.searchParams.get("sessionId") || "";
    const session = ecvanSessions.get(sessionId);
    if (!session) return sendJson(res, 200, { status: "not_found" });
    return sendJson(res, 200, { status: session.status, progress: session.progress || "", error: session.error || null, rowCount: session.data ? session.data.length : 0 });
  }

  return sendJson(res, 404, { error: "Not found" });
}

ensureDb();

/** /api 요청을 한 번에 하나만 처리해 db.json 동시 접근(Windows 잠금·덮어쓰기 유실)을 막음 */
let dbApiQueue = Promise.resolve();

function runSerializedApi(handler) {
  const next = dbApiQueue.then(() => handler());
  dbApiQueue = next.then(
    () => undefined,
    () => undefined
  );
  return next;
}

const server = http.createServer(async (req, res) => {
  const urlObj = new URL(req.url, `http://${req.headers.host}`);
  if (urlObj.pathname.startsWith("/api/")) {
    return runSerializedApi(() => handleApi(req, res, urlObj));
  }
  return serveStatic(req, res, urlObj.pathname);
});

server.listen(PORT, HOST, () => {
  console.log(`WMS: http://localhost:${PORT}`);
  if (HOST === "0.0.0.0") console.log("Internal share enabled: http://<this-pc-ip>:" + PORT);
});
