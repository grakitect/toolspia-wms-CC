const views = [
  "dashboard",
  "master",
  "products",
  "stock",
  "inbound",
  "inbound-plan",
  "inbound-plan-2",
  "outbound",
  "outbound-plan",
  "outbound-upload-daiso",
  "outbound-upload-emart",
  "outbound-upload-lotte",
  "adjust",
  "history",
  "alert"
];
const LAST_VIEW_KEY = "wms:lastView";
const PRODUCT_HIDDEN_COLS_KEY = "wms:productHiddenCols";
const TABLE_FILTER_MEMORY = {};
let productListPageIndex = 0;
const state = {
  products: [],
  stock: [],
  warehouses: [],
  productOptions: {
    status: [],
    deliveryVendors: [],
    orderDept: [],
    orderManagers: [],
    supplyType: [],
    warehouseGroup: [],
    itemType: [],
    categories: []
  },
  partners: { inbound: [], outbound: [], purchase: [] },
  managers: []
};

function qs(sel) {
  return document.querySelector(sel);
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

function formatYmd(v) {
  const raw = String(v || "").trim();
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "");
  if (digits.length !== 8) return raw;
  return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
}

/** YYYY/MM/DD 또는 YYYYMMDD 등 날짜 표시용 */
function formatYmdLoose(v) {
  const s = String(v || "").trim();
  const m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  return formatYmd(v);
}

function formatDateTimeDisplay(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  const d = new Date(s);
  if (!Number.isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    const ss = String(d.getSeconds()).padStart(2, "0");
    return `${y}-${m}-${day} ${hh}:${mm}:${ss}`;
  }
  return s.replace("T", " ").replace("Z", "");
}

function normalizeDateTimeDigits(v) {
  return String(v || "")
    .trim()
    .replace(/[^\d]/g, "");
}

/** 브라우저 로컬 날짜 기준 YYYY-MM-DD (date input·API 신규 확인 기간용) */
function localDateYmd(d) {
  const x = d instanceof Date ? d : new Date(d);
  const y = x.getFullYear();
  const m = String(x.getMonth() + 1).padStart(2, "0");
  const day = String(x.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function normalizeSlipNoText(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  const compact = s.replace(/\s+/g, "");
  const m = compact.match(/^(\d{2}|\d{4})\/?(\d{2})\/?(\d{2})-(\d+)$/);
  if (m) {
    const yy = m[1].length === 4 ? m[1].slice(2) : m[1];
    return `${yy}/${m[2]}/${m[3]}-${String(Number(m[4]))}`;
  }
  return compact;
}

function toTagList(v) {
  if (Array.isArray(v)) return v.map((x) => String(x || "").trim()).filter(Boolean);
  return String(v || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);
}

function renderTagChips(v, cls = "tag tag-orange") {
  const list = toTagList(v);
  if (!list.length) return "<span class='muted'>-</span>";
  return list.map((x) => `<span class="${cls}">${esc(x)}</span>`).join("");
}

function preventEnterSubmit(formEl) {
  if (!formEl) return;
  formEl.addEventListener("keydown", (e) => {
    if (e.key === "Enter") e.preventDefault();
  });
}

function applyExcelLikeFilter(tableSelector, afterApply) {
  const table = qs(tableSelector);
  if (!table || !table.tHead || !table.tBodies || !table.tBodies[0]) return;
  const headRow = table.tHead.rows[0];
  if (!headRow) return;
  const tbody = table.tBodies[0];
  const rows = Array.from(tbody.rows || []);
  const originalRows = [...rows];
  const colCount = headRow.cells.length;
  if (!rows.length || !colCount) return;

  function readCellFilterValue(cell) {
    if (!cell) return "";
    const sel = cell.querySelector("select");
    if (sel) {
      const opt = sel.options && sel.selectedIndex >= 0 ? sel.options[sel.selectedIndex] : null;
      return String(opt ? opt.textContent : sel.value || "").trim();
    }
    const input = cell.querySelector("input");
    if (input) return String(input.value || "").trim();
    return String(cell.textContent || "").trim();
  }

  const colValues = Array.from({ length: colCount }, () => new Set());
  rows.forEach((row) => {
    for (let i = 0; i < colCount; i += 1) {
      colValues[i].add(readCellFilterValue(row.cells[i]));
    }
  });

  const memory = TABLE_FILTER_MEMORY[tableSelector];
  const filterState = colValues.map((set, idx) => {
    const allValues = Array.from(set);
    let selected = new Set(allValues);
    const saved = memory && Array.isArray(memory.filters) ? memory.filters[idx] : null;
    if (saved && Array.isArray(saved.selected)) {
      const kept = saved.selected.filter((v) => allValues.includes(v));
      selected = new Set(kept.length ? kept : allValues);
    }
    return { allValues, selected };
  });
  const sortState = memory?.sort ? { col: Number(memory.sort.col), dir: String(memory.sort.dir || "") } : { col: -1, dir: "" };
  if (!Number.isInteger(sortState.col) || sortState.col < 0 || sortState.col >= colCount) sortState.col = -1;
  if (!["asc", "desc", ""].includes(sortState.dir)) sortState.dir = "";
  let openedMenu = null;

  function cellValue(row, col) {
    return readCellFilterValue(row.cells[col]);
  }

  function compareValue(a, b, dir) {
    const na = Number(String(a).replace(/,/g, ""));
    const nb = Number(String(b).replace(/,/g, ""));
    let cmp = 0;
    if (Number.isFinite(na) && Number.isFinite(nb) && String(a) !== "" && String(b) !== "") {
      cmp = na - nb;
    } else {
      cmp = String(a).localeCompare(String(b), "ko", { numeric: true });
    }
    return dir === "asc" ? cmp : -cmp;
  }

  function apply() {
    rows.forEach((row) => {
      let visible = true;
      for (let i = 0; i < colCount; i += 1) {
        const val = cellValue(row, i);
        const st = filterState[i];
        if (st.selected.size !== st.allValues.length && !st.selected.has(val)) {
          visible = false;
          break;
        }
      }
      row.dataset.wmsExcelVisible = visible ? "1" : "0";
      row.style.display = visible ? "" : "none";
    });

    if (sortState.col >= 0 && (sortState.dir === "asc" || sortState.dir === "desc")) {
      const sorted = [...rows].sort((r1, r2) =>
        compareValue(cellValue(r1, sortState.col), cellValue(r2, sortState.col), sortState.dir)
      );
      sorted.forEach((r) => tbody.appendChild(r));
    } else {
      originalRows.forEach((r) => tbody.appendChild(r));
    }
    TABLE_FILTER_MEMORY[tableSelector] = {
      filters: filterState.map((st) => ({ selected: Array.from(st.selected) })),
      sort: { col: sortState.col, dir: sortState.dir }
    };
    if (typeof afterApply === "function") afterApply();
  }

  function closeMenu() {
    if (openedMenu) {
      openedMenu.remove();
      openedMenu = null;
    }
  }

  Array.from(headRow.cells).forEach((th, colIdx) => {
    // Keep checkbox header cell intact (e.g., product select-all column).
    if (th.querySelector('input[type="checkbox"]')) return;

    const baseLabel = th.getAttribute("data-base-label") || String(th.textContent || "").trim();
    th.setAttribute("data-base-label", baseLabel);
    th.innerHTML = "";
    th.classList.add("excel-th");

    const label = document.createElement("span");
    label.className = "excel-th-label";
    label.textContent = baseLabel;

    const trigger = document.createElement("button");
    trigger.type = "button";
    trigger.className = "excel-th-trigger";
    trigger.textContent = "▼";
    trigger.title = "필터";

    trigger.onclick = (e) => {
      e.stopPropagation();
      closeMenu();

      const menu = document.createElement("div");
      menu.className = "excel-menu";
      menu.onclick = (evt) => evt.stopPropagation();

      const sortAsc = document.createElement("button");
      sortAsc.type = "button";
      sortAsc.className = "excel-menu-btn";
      sortAsc.textContent = "오름차순";
      sortAsc.onclick = (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        sortState.col = colIdx;
        sortState.dir = "asc";
        apply();
        closeMenu();
      };

      const sortDesc = document.createElement("button");
      sortDesc.type = "button";
      sortDesc.className = "excel-menu-btn";
      sortDesc.textContent = "내림차순";
      sortDesc.onclick = (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        sortState.col = colIdx;
        sortState.dir = "desc";
        apply();
        closeMenu();
      };

      const resetSort = document.createElement("button");
      resetSort.type = "button";
      resetSort.className = "excel-menu-btn";
      resetSort.textContent = "정렬해제";
      resetSort.onclick = (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        sortState.col = -1;
        sortState.dir = "";
        apply();
        closeMenu();
      };

      const valueWrap = document.createElement("div");
      valueWrap.className = "excel-menu-values";
      const st = filterState[colIdx];
      const values = [...st.allValues].sort((a, b) => a.localeCompare(b, "ko", { numeric: true }));
      const valueCheckboxes = [];

      const allItem = document.createElement("label");
      allItem.className = "excel-menu-item excel-menu-item-all";
      const allCb = document.createElement("input");
      allCb.type = "checkbox";
      allCb.checked = st.selected.size === st.allValues.length;
      allCb.indeterminate = st.selected.size > 0 && st.selected.size < st.allValues.length;
      const allTxt = document.createElement("span");
      allTxt.textContent = "전체 선택";
      allItem.appendChild(allCb);
      allItem.appendChild(allTxt);
      valueWrap.appendChild(allItem);

      const divider = document.createElement("div");
      divider.className = "excel-menu-divider";
      valueWrap.appendChild(divider);

      function refreshAllCheckboxState() {
        allCb.checked = st.selected.size === st.allValues.length;
        allCb.indeterminate = st.selected.size > 0 && st.selected.size < st.allValues.length;
      }

      values.forEach((v) => {
        const item = document.createElement("label");
        item.className = "excel-menu-item";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = st.selected.has(v);
        cb.onclick = (evt) => evt.stopPropagation();
        cb.onchange = () => {
          if (cb.checked) st.selected.add(v);
          else st.selected.delete(v);
          refreshAllCheckboxState();
          apply();
        };
        const txt = document.createElement("span");
        txt.textContent = v === "" ? "(빈값)" : v;
        item.appendChild(cb);
        item.appendChild(txt);
        valueWrap.appendChild(item);
        valueCheckboxes.push({ value: v, cb });
      });

      const clearBtn = document.createElement("button");
      clearBtn.type = "button";
      clearBtn.className = "excel-menu-btn";
      clearBtn.textContent = "값필터해제";
      clearBtn.onclick = (evt) => {
        evt.preventDefault();
        evt.stopPropagation();
        st.selected = new Set(st.allValues);
        valueCheckboxes.forEach((x) => {
          x.cb.checked = true;
        });
        refreshAllCheckboxState();
        apply();
        closeMenu();
      };

      allCb.onclick = (evt) => evt.stopPropagation();
      allCb.onchange = () => {
        if (allCb.checked) {
          st.selected = new Set(st.allValues);
          valueCheckboxes.forEach((x) => {
            x.cb.checked = true;
          });
        } else {
          st.selected = new Set();
          valueCheckboxes.forEach((x) => {
            x.cb.checked = false;
          });
        }
        refreshAllCheckboxState();
        apply();
      };

      menu.appendChild(sortAsc);
      menu.appendChild(sortDesc);
      menu.appendChild(resetSort);
      menu.appendChild(clearBtn);
      menu.appendChild(valueWrap);

      document.body.appendChild(menu);
      const rect = trigger.getBoundingClientRect();
      menu.style.left = "0px";
      menu.style.top = "0px";
      const margin = 8;
      const menuW = menu.offsetWidth || 220;
      const menuH = menu.offsetHeight || 260;
      const maxLeft = window.innerWidth - menuW - margin;
      const maxTop = window.innerHeight - menuH - margin;
      const left = Math.min(Math.max(margin, rect.left), Math.max(margin, maxLeft));
      const top = Math.min(Math.max(margin, rect.bottom + 6), Math.max(margin, maxTop));
      menu.style.left = `${left}px`;
      menu.style.top = `${top}px`;
      openedMenu = menu;
    };

    th.appendChild(label);
    th.appendChild(trigger);
  });

  if (!table.dataset.filterOutsideBound) {
    document.addEventListener("click", () => {
      closeMenu();
    });
    table.dataset.filterOutsideBound = "1";
  }

  apply();
}

async function api(path, options = {}) {
  const res = await fetch(path, {
    headers: { "Content-Type": "application/json" },
    ...options
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "요청 실패");
  return data;
}

function switchView(name) {
  if (!views.includes(name)) return;
  if (name !== "products") delete TABLE_FILTER_MEMORY["#products-table"];
  for (const v of views) qs(`#view-${v}`).classList.add("hidden");
  qs(`#view-${name}`).classList.remove("hidden");
  document.querySelectorAll(".sidebar button[data-view]").forEach((b) => {
    b.classList.toggle("active", b.dataset.view === name);
  });
  // 해당 버튼이 속한 그룹을 자동으로 열기
  const activeBtn = document.querySelector(`.sidebar button[data-view="${name}"]`);
  if (activeBtn) {
    const group = activeBtn.closest(".sidebar-group");
    if (group && !group.classList.contains("open")) group.classList.add("open");
  }
  try {
    localStorage.setItem(LAST_VIEW_KEY, name);
  } catch (_) {}
}

function initSidebarAccordion() {
  document.querySelectorAll(".sidebar-group-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      const group = btn.closest(".sidebar-group");
      if (!group) return;
      const isOpen = group.classList.contains("open");
      // 다른 그룹 닫기
      document.querySelectorAll(".sidebar-group.open").forEach((g) => g.classList.remove("open"));
      if (!isOpen) group.classList.add("open");
    });
  });
}

function parseSheet(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(rows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseSheetMatrix(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        resolve(matrix);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

/** 첫 시트를 2차원 배열로 (발주서현황 업로드용) */
function parseSheetAsMatrix(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        resolve({ matrix, sourceFileName: file.name || "upload.xlsx" });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function applyProductHiddenColumnStyles() {
  let hidden = [];
  try {
    hidden = JSON.parse(localStorage.getItem(PRODUCT_HIDDEN_COLS_KEY) || "[]");
  } catch (_) {}
  const id = "wms-product-col-hide-style";
  let el = document.getElementById(id);
  if (!el) {
    el = document.createElement("style");
    el.id = id;
    document.head.appendChild(el);
  }
  if (!Array.isArray(hidden) || !hidden.length) {
    el.textContent = "";
    return;
  }
  el.textContent = hidden
    .map(
      (n) =>
        `#products-table thead th:nth-child(${n}), #products-table tbody td:nth-child(${n}) { display: none !important; }`
    )
    .join("\n");
}

async function uploadProductBulkFromFile(file) {
  const rows2 = await parseSheet(file);
  const normalized = normalizeProductRows(rows2);
  if (!normalized.length) {
    throw new Error("업로드 가능한 상품행이 없습니다. 헤더명(품목코드(이카운트), 품목명(이카운트))과 데이터 유무를 확인하세요.");
  }
  await api("/api/products/bulk", { method: "POST", body: JSON.stringify({ rows: normalized }) });
  await refreshCommon();
  return normalized.length;
}

async function uploadProductUpdateFromFile(file) {
  const rows2 = await parseSheet(file);
  const normalized = normalizeProductRows(rows2);
  if (!normalized.length) throw new Error("업데이엄트 가능한 상품행이 없습니다.");
  await api("/api/products/bulk", { method: "POST", body: JSON.stringify({ rows: normalized }) });
  await refreshCommon();
  return normalized.length;
}

function setupDropZone(zoneId, fileInputId, onFilePicked) {
  const zone = qs(`#${zoneId}`);
  const fileInput = qs(`#${fileInputId}`);
  if (!zone || !fileInput) return;
  zone.onclick = () => fileInput.click();
  zone.ondragover = (e) => {
    e.preventDefault();
    zone.classList.add("drop-over");
  };
  zone.ondragleave = () => zone.classList.remove("drop-over");
  zone.ondrop = (e) => {
    e.preventDefault();
    zone.classList.remove("drop-over");
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!f) return;
    try {
      fileInput.files = e.dataTransfer.files;
    } catch (_) {
      /* some browsers */
    }
    if (typeof onFilePicked === "function") onFilePicked(f);
  };
}

function downloadGuide(type) {
  const guides = {
    product: [
      {
        "품목코드(이카운트)": "EC-1001",
        "바코드(SKU)": "8800000000012",
        "바코드(중포)": "8800000000012-M",
        "바코드(카톤)": "L-8800000000012",
        "품목명(이카운트)": "샘플상품",
        "상태": "판매중",
        "판매처": "브랜드디자인, 김윤환",
        "판매처관리코드": "VD-001",
        "판매처 품목명": "샘플 납품품목명",
        "규격": "500ml",
        "구매처": "테스트구매처",
        "수급형태": "무역(수입)",
        "발주부서": "유통사업부",
        "발주담당자": "담당자A, 담당자B",
        "구매처 품목코드": "PV-001",
        "구매처 품목명": "구매처 샘플명",
        "창고그룹(이카운트)": "기본창고",
        "사용창고": "유통사업부, 다이소",
        "구분": "[상품]",
        "카테고리": "보통, 시즌"
      }
    ],
    IN: [{ "상품코드": "P-001", "수량": 10, "창고": "유통사업부", "구매처": "테스트구매처", "담당자": "admin", "메모": "초도 입고" }],
    OUT: [{ "상품코드": "P-001", "수량": 3, "창고": "유통사업부", "출고처": "출고처A", "담당자": "admin", "메모": "샘플 출고" }]
  };
  const ws = XLSX.utils.json_to_sheet(guides[type] || []);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "guide");
  XLSX.writeFile(wb, `${type}-upload-guide.xlsx`);
}

function getByKeys(row, keys) {
  const entries = Object.entries(row || {});
  const normalizeKey = (v) => String(v || "").replace(/[\s()[\]{}_\-./\\]/g, "").toLowerCase();
  for (const k of keys) {
    if (row[k] !== undefined && row[k] !== null && row[k] !== "") return row[k];
    const nk = normalizeKey(k);
    const hit = entries.find(([rk, rv]) => normalizeKey(rk) === nk && rv !== undefined && rv !== null && rv !== "");
    if (hit) return hit[1];
  }
  return undefined;
}

function normalizeProductRows(rows) {
  return rows
    .map((r) => {
      const ecountCode = getByKeys(r, ["품목코드(이카운트)", "품목코드", "ecountCode", "ECOUNT_CODE", "상품코드", "code", "Code"]);
      const ecountName = getByKeys(r, ["품목명(이카운트)", "품목명", "상품명", "name", "Name"]);
      return {
        code: String(ecountCode || "").trim(),
        name: String(ecountName || "").trim(),
        ecountCode: String(ecountCode || "").trim(),
        ecountName: String(ecountName || "").trim(),
        barcode: String(getByKeys(r, ["바코드(SKU)", "바코드", "barcode"]) || "").trim(),
        middleBarcode: String(
          getByKeys(r, ["바코드(중포)", "중포바코드", "중포 바코드", "중포", "middleBarcode", "middle_barcode"]) || ""
        ).trim(),
        logisticsBarcode: String(getByKeys(r, ["바코드(카톤)", "물류바코드", "logisticsBarcode"]) || "").trim(),
        status: String(getByKeys(r, ["상태", "status"]) || "판매중").trim(),
        deliveryVendors: toTagList(getByKeys(r, ["판매처", "deliveryVendors", "salesVendor"])),
        deliveryVendorCode: String(getByKeys(r, ["판매처관리코드", "deliveryVendorCode"]) || "").trim(),
        deliveryItemName: String(getByKeys(r, ["판매처 품목명", "deliveryItemName"]) || "").trim(),
        spec: String(getByKeys(r, ["규격", "spec"]) || "").trim(),
        purchaseVendor: String(getByKeys(r, ["구매처", "purchaseVendor"]) || "").trim(),
        supplyType: String(getByKeys(r, ["수급형태", "supplyType"]) || "").trim(),
        orderDept: String(getByKeys(r, ["발주부서", "orderDept"]) || "").trim(),
        orderManagers: toTagList(getByKeys(r, ["발주담당자", "orderManagers"])),
        purchaseItemCode: String(getByKeys(r, ["구매처 품목코드", "purchaseItemCode"]) || "").trim(),
        purchaseItemName: String(getByKeys(r, ["구매처 품목명", "purchaseItemName"]) || "").trim(),
        warehouseGroup: String(getByKeys(r, ["창고그룹(이카운트)", "창고그룹", "warehouseGroup"]) || "").trim(),
        usedWarehouses: toTagList(getByKeys(r, ["사용창고", "usedWarehouses"])),
        itemType: String(getByKeys(r, ["구분", "itemType"]) || "").trim(),
        categories: toTagList(getByKeys(r, ["카테고리", "categories", "category"])),
        unit: String(getByKeys(r, ["단위", "unit"]) || "EA").trim(),
        safetyStock: Number(getByKeys(r, ["안전재고", "safetyStock"]) || 0),
        optimalStock: Number(getByKeys(r, ["적정재고", "optimalStock"]) || 0)
      };
    })
    .filter((p) => p.ecountCode && p.ecountName);
}

function findProductForMovementImport(productCode) {
  const c = String(productCode || "").trim();
  if (!c) return null;
  const list = state.products || [];
  return (
    list.find((p) => String(p.code || "").trim() === c) ||
    list.find((p) => String(p.ecountCode || "").trim() === c) ||
    list.find((p) => String(p.barcode || "").trim() === c) ||
    list.find((p) => String(p.middleBarcode || "").trim() === c) ||
    null
  );
}

function fillInboundPartnerFromProduct(row) {
  const partner = String(row.partner || "").trim();
  if (partner) return { ...row, partner };
  const p = findProductForMovementImport(row.productCode);
  const fromMaster = p ? String(p.purchaseVendor || "").trim() : "";
  return { ...row, partner: fromMaster };
}

function normalizeMovementRows(rows, type) {
  const partnerKeys = type === "IN" ? ["구매처", "입고처", "partner"] : ["출고처", "partner"];
  const qtyKeys = ["수량", "qty", "Qty"];
  const mapped = rows
    .map((r) => {
      const productCode = getByKeys(r, ["상품코드", "productCode", "품목코드(이카운트)"]);
      return {
        productCode: String(productCode || "").trim(),
        qty: Number(getByKeys(r, qtyKeys) || 0),
        warehouse: String(getByKeys(r, ["창고", "warehouse"]) || "").trim(),
        partner: String(getByKeys(r, partnerKeys) || "").trim(),
        user: String(getByKeys(r, ["담당자", "user"]) || "").trim(),
        memo: String(getByKeys(r, ["메모", "memo"]) || "").trim()
      };
    })
    .map((x) => (type === "IN" ? fillInboundPartnerFromProduct(x) : x));

  if (type === "IN") {
    const needPartner = mapped.filter(
      (x) => x.productCode && x.qty !== 0 && x.user && x.warehouse && !String(x.partner || "").trim()
    );
    if (needPartner.length) {
      const sample = needPartner
        .map((x) => x.productCode)
        .filter(Boolean)
        .slice(0, 8)
        .join(", ");
      throw new Error(
        `구매처가 비어 있고 상품 기본정보에도 구매처가 없는 행이 ${needPartner.length}건 있습니다. (예: ${sample || "-"})`
      );
    }
  }

  return mapped.filter((x) => {
    if (!(x.productCode && x.qty !== 0 && x.user && x.warehouse)) return false;
    if (type === "IN") return Boolean(String(x.partner || "").trim());
    return Boolean(x.partner);
  });
}

async function refreshCommon() {
  const [productsRes, stockRes, partnersRes, managersRes, warehousesRes, optionRes] = await Promise.all([
    api("/api/products"),
    api("/api/stock"),
    api("/api/partners"),
    api("/api/managers"),
    api("/api/warehouses"),
    api("/api/product-options")
  ]);
  state.products = productsRes.items;
  state.stock = stockRes.items;
  state.partners = partnersRes.items || { inbound: [], outbound: [], purchase: [] };
  state.managers = managersRes.items || [];
  state.warehouses = warehousesRes.items || [];
  state.productOptions = optionRes.items || state.productOptions;
}

async function renderDashboard() {
  const d = await api("/api/dashboard");
  qs("#view-dashboard").innerHTML = `
    <div class="card"><h2>대시보드</h2><p class="muted">기본 지표(향후 업그레이드 예정)</p></div>
    <div class="grid3">
      <div class="card"><h3>총 상품수</h3><strong>${d.totalProducts}</strong></div>
      <div class="card"><h3>안전재고 이하</h3><strong>${d.lowStock}</strong></div>
      <div class="card"><h3>오늘 입출고 건수</h3><strong>${d.todayMoves}</strong></div>
    </div>
  `;
}

function renderMaster() {
  const productOptions = state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join("");
  const listTags = (arr) => arr.map((v) => `<span class="tag">${esc(v)}</span>`).join("") || "<span class='muted'>없음</span>";
  const partnerChip = (type, name) => `
    <span class="tag-wrap">
      <span class="tag">${esc(name)}</span>
      <button type="button" class="cancel-btn del-small partner-del" data-type="${type}" data-name="${encodeURIComponent(name)}">삭제</button>
    </span>
  `;
  const managerChip = (name) => `
    <span class="tag-wrap">
      <span class="tag">${esc(name)}</span>
      <button type="button" class="cancel-btn del-small manager-del" data-name="${encodeURIComponent(name)}">삭제</button>
    </span>
  `;
  const partnerRows = [
    ...state.partners.purchase.map((v) => ({ type: "purchase", label: "구매처", name: v })),
    ...state.partners.outbound.map((v) => ({ type: "outbound", label: "판매처", name: v }))
  ];
  const managerRows = state.managers.map((v) => ({ name: v }));

  qs("#view-master").innerHTML = `
    <div class="card"><h2>기본 정보</h2><p class="muted">거래처/담당자 등록</p></div>
    <div class="grid3">
      <div class="card">
        <h3>거래처 마스터 등록</h3>
        <form id="partner-master-form">
          <div><label>구분</label>
            <select name="type" required>
              <option value="purchase">구매처</option>
              <option value="outbound">판매처</option>
            </select>
          </div>
          <div><label>거래처명</label><input name="name" required /></div>
          <div><button class="primary" type="submit">등록</button></div>
        </form>
        <table class="mini-table">
          <thead><tr><th>구분</th><th>거래처</th><th>삭제</th></tr></thead>
          <tbody>
            ${partnerRows.length
              ? partnerRows
                  .map(
                    (r) =>
                      `<tr>
                        <td>${esc(r.label)}</td>
                        <td>${esc(r.name)}</td>
                        <td>
                          <button type="button" class="cancel-btn del-small partner-del" data-type="${r.type}" data-name="${encodeURIComponent(r.name)}">삭제</button>
                        </td>
                      </tr>`
                  )
                  .join("")
              : `<tr><td colspan="3"><span class="muted">없음</span></td></tr>`}
          </tbody>
        </table>
      </div>
      <div class="card">
        <h3>담당자 등록</h3>
        <form id="manager-form">
          <div><label>담당자명</label><input name="name" required /></div>
          <div><button class="primary" type="submit">등록</button></div>
        </form>
        <table class="mini-table">
          <thead><tr><th>담당자</th><th>삭제</th></tr></thead>
          <tbody>
            ${managerRows.length
              ? managerRows
                  .map(
                    (r) =>
                      `<tr>
                        <td>${esc(r.name)}</td>
                        <td><button type="button" class="cancel-btn del-small manager-del" data-name="${encodeURIComponent(r.name)}">삭제</button></td>
                      </tr>`
                  )
                  .join("")
              : `<tr><td colspan="2"><span class="muted">없음</span></td></tr>`}
          </tbody>
        </table>
      </div>
      <div class="card">
        <h3>창고 등록</h3>
        <form id="warehouse-form">
          <div><label>창고명</label><input name="name" required /></div>
          <div><button class="primary" type="submit">등록</button></div>
        </form>
        <table class="mini-table" id="warehouse-table">
          <thead><tr><th>창고명</th><th>수정</th><th>삭제</th></tr></thead>
          <tbody>
            ${(state.warehouses || []).length
              ? state.warehouses
                  .map(
                    (w) =>
                      `<tr>
                        <td>${esc(w)}</td>
                        <td><button type="button" class="primary del-small warehouse-rename" data-name="${encodeURIComponent(w)}">수정</button></td>
                        <td>${
                          w === "유통사업부"
                            ? `<span class="muted">기본창고</span>`
                            : `<button type="button" class="cancel-btn del-small warehouse-del" data-name="${encodeURIComponent(w)}">삭제</button>`
                        }</td>
                      </tr>`
                  )
                  .join("")
              : `<tr><td colspan="3"><span class="muted">없음</span></td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;

  preventEnterSubmit(qs("#partner-master-form"));
  preventEnterSubmit(qs("#manager-form"));
  preventEnterSubmit(qs("#warehouse-form"));

  qs("#partner-master-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    await api("/api/partners", { method: "POST", body: JSON.stringify(data) });
    await refreshCommon();
    renderMaster();
    renderProducts();
    renderInbound();
    renderOutbound();
    alert("거래처 등록 완료");
  };
  qs("#manager-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    await api("/api/managers", { method: "POST", body: JSON.stringify(data) });
    await refreshCommon();
    renderMaster();
    renderInbound();
    renderOutbound();
    renderAdjust();
    alert("담당자 등록 완료");
  };
  qs("#warehouse-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    await api("/api/warehouses", { method: "POST", body: JSON.stringify(data) });
    await refreshCommon();
    renderMaster();
    renderStock();
    renderInbound();
    renderOutbound();
    renderAdjust();
    alert("창고 등록 완료");
  };
  // 안전/적정재고는 기본상품정보 화면에서 수정합니다.

  document.querySelectorAll(".partner-del").forEach((btn) => {
    btn.onclick = async () => {
      const type = btn.dataset.type;
      const name = decodeURIComponent(btn.dataset.name || "");
      if (!type || !name) return;
      if (!confirm(`거래처를 삭제할까요? (${name})`)) return;
      try {
        await api("/api/partners", { method: "DELETE", body: JSON.stringify({ type, name }) });
        await refreshCommon();
        renderMaster();
        renderProducts();
        renderInbound();
        renderOutbound();
      } catch (err) {
        alert(err.message);
      }
    };
  });

  document.querySelectorAll(".manager-del").forEach((btn) => {
    btn.onclick = async () => {
      const name = decodeURIComponent(btn.dataset.name || "");
      if (!name) return;
      if (!confirm(`담당자를 삭제할까요? (${name})`)) return;
      try {
        await api("/api/managers", { method: "DELETE", body: JSON.stringify({ name }) });
        await refreshCommon();
        renderMaster();
        renderInbound();
        renderOutbound();
        renderAdjust();
      } catch (err) {
        alert(err.message);
      }
    };
  });

  const miniTables = document.querySelectorAll(".mini-table");
  if (miniTables[0]) miniTables[0].id = "partner-table";
  if (miniTables[1]) miniTables[1].id = "manager-table";
  if (miniTables[2]) miniTables[2].id = "warehouse-table";
  applyExcelLikeFilter("#partner-table");
  applyExcelLikeFilter("#manager-table");
  applyExcelLikeFilter("#warehouse-table");

  document.querySelectorAll(".warehouse-del").forEach((btn) => {
    btn.onclick = async () => {
      const name = decodeURIComponent(btn.dataset.name || "");
      if (!name) return;
      const first = confirm(`창고를 삭제하시겠습니까? (${name})`);
      if (!first) return;
      const approver = prompt("담당자명을 입력하세요. (삭제 권한: 박유정)");
      if (!approver) return;
      const approverName = String(approver).trim();
      if (!approverName) return;
      const second = confirm(`정말 삭제할까요? 이 작업은 되돌릴 수 없습니다. (${name})`);
      if (!second) return;
      try {
        await api("/api/warehouses", {
          method: "DELETE",
          body: JSON.stringify({ name, approver: approverName })
        });
        await refreshCommon();
        renderMaster();
        renderStock();
        renderInbound();
        renderOutbound();
        renderAdjust();
      } catch (err) {
        alert(err.message);
      }
    };
  });

  document.querySelectorAll(".warehouse-rename").forEach((btn) => {
    btn.onclick = async () => {
      const oldName = decodeURIComponent(btn.dataset.name || "");
      if (!oldName) return;
      const newName = prompt(`변경할 창고명을 입력하세요.`, oldName);
      if (!newName) return;
      const cleaned = newName.trim();
      if (!cleaned || cleaned === oldName) return;
      try {
        await api("/api/warehouses", {
          method: "PUT",
          body: JSON.stringify({ oldName, newName: cleaned })
        });
        await refreshCommon();
        renderMaster();
        renderStock();
        renderInbound();
        renderOutbound();
        renderAdjust();
      } catch (err) {
        alert(err.message);
      }
    };
  });
}

function renderProducts() {
  productListPageIndex = 0;
  const opts = state.productOptions || {};
  const datalist = (id, arr) => `<datalist id="${id}">${(arr || []).map((v) => `<option value="${esc(v)}"></option>`).join("")}</datalist>`;
  const selectItems = (arr) => (arr || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const selectOptions = (arr, current) =>
    (arr || [])
      .map((v) => `<option value="${esc(v)}" ${String(current || "") === String(v || "") ? "selected" : ""}>${esc(v)}</option>`)
      .join("");
  const optionMeta = [
    { key: "status", label: "상태" },
    { key: "deliveryVendors", label: "판매처" },
    { key: "orderDept", label: "발주부서" },
    { key: "orderManagers", label: "발주담당자" },
    { key: "supplyType", label: "수급형태" },
    { key: "warehouseGroup", label: "창고그룹(이카운트)" },
    { key: "itemType", label: "구분" },
    { key: "categories", label: "카테고리" }
  ];
  const rows = state.products
    .map((p) => {
      const searchText = `${p.ecountCode || p.code || ""} ${p.code || ""} ${p.ecountName || p.name || ""}`.toLowerCase();
      return `<tr data-search="${esc(searchText)}">
      <td><input type="checkbox" class="product-row-check" data-code="${esc(p.code)}" /></td>
      <td>${esc(p.ecountCode || p.code)}</td>
      <td>${esc(p.barcode)}</td>
      <td>${esc(p.middleBarcode || "")}</td>
      <td>${esc(p.logisticsBarcode)}</td>
      <td>${esc(p.ecountName || p.name)}</td>
      <td>${esc(p.status || "")}</td>
      <td>${renderTagChips(p.deliveryVendors)}</td>
      <td>${esc(p.deliveryVendorCode || "")}</td>
      <td>${esc(p.deliveryItemName || "")}</td>
      <td>${esc(p.spec || "")}</td>
      <td>${esc(p.purchaseVendor || "")}</td>
      <td>${esc(p.supplyType || "")}</td>
      <td>${esc(p.orderDept || "")}</td>
      <td>${renderTagChips(p.orderManagers)}</td>
      <td>${esc(p.purchaseItemCode || "")}</td>
      <td>${esc(p.purchaseItemName || "")}</td>
      <td>${esc(p.warehouseGroup || "")}</td>
      <td>${renderTagChips(p.usedWarehouses)}</td>
      <td>${esc(p.itemType || "")}</td>
      <td>${renderTagChips(p.categories)}</td>
    </tr>`;
    })
    .join("");

  qs("#view-products").innerHTML = `
    <div id="products-list-view" class="products-bh">
      <div class="card products-bh-card">
        <div class="products-bh-toolbar">
          <h2 class="products-bh-title">기본상품정보</h2>
          <div class="products-bh-actions">
            <button id="open-product-popup" class="bh-btn bh-btn-primary products-bh-main-btn" type="button">+ 상품 추가</button>
            <div class="bh-dropdown">
              <button type="button" class="bh-btn bh-btn-outline bh-excel-btn products-bh-main-btn" id="excel-import-toggle" aria-expanded="false">
                <span class="bh-mini-xls">XLS</span>
                엑셀 가져오기
                <span class="bh-caret">▾</span>
              </button>
              <div class="bh-dropdown-menu hidden" id="excel-import-menu">
                <button type="button" class="bh-menu-item" id="menu-product-guide">가이드 다운로드</button>
                <button type="button" class="bh-menu-item" id="menu-bulk-new">신규 상품 일괄 등록</button>
                <button type="button" class="bh-menu-item" id="menu-bulk-update">특정상품 정보 업데이트</button>
              </div>
            </div>
            <button type="button" class="bh-btn bh-btn-outline products-bh-main-btn" id="product-column-settings">컬럼 설정</button>
            <button type="button" class="bh-btn bh-btn-outline products-bh-main-btn" id="product-export-all">엑셀 내보내기</button>
          </div>
        </div>
        <div class="products-bh-search-row">
          <div class="products-bh-search-wrap">
            <span class="bh-search-icon" aria-hidden="true">🔍</span>
            <input type="text" id="product-search-q" class="products-bh-search" placeholder="이카운트 / 상품코드 / 품목명 검색" autocomplete="off" />
            <button type="button" class="bh-search-go" id="product-search-btn">조회</button>
          </div>
          <span class="products-bh-search-actions">
            <button type="button" id="product-edit-selected" class="bh-btn bh-btn-sm">선택 수정</button>
            <button type="button" id="product-delete-selected" class="bh-btn bh-btn-sm bh-btn-danger-outline">선택 삭제</button>
          </span>
        </div>
        <p id="product-search-result" class="muted products-bh-result">상품을 검색하세요.</p>
        <input type="file" id="product-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <input type="file" id="product-update-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <div class="products-bh-table-outer">
          <div class="table-scroll-x products-bh-y-scroll">
            <table id="products-table">
              <thead><tr><th><input id="product-check-all" type="checkbox" /></th><th>품목코드(이카운트)</th><th>바코드(SKU)</th><th>바코드(중포)</th><th>바코드(카톤)</th><th>품목명(이카운트)</th><th>상태</th><th>판매처</th><th>판매처관리코드</th><th>판매처 품목명</th><th>규격</th><th>구매처</th><th>수급형태</th><th>발주부서</th><th>발주담당자</th><th>구매처 품목코드</th><th>구매처 품목명</th><th>창고그룹(이카운트)</th><th>사용창고</th><th>구분</th><th>카테고리</th></tr></thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </div>
        <div class="products-bh-pagination">
          <label>보기
            <select id="product-page-size">
              <option value="50">50</option>
              <option value="100" selected>100</option>
              <option value="200">200</option>
              <option value="99999">전체</option>
            </select>
          </label>
          <span id="product-page-info">0 - 0 / 0</span>
          <button type="button" id="product-page-prev" class="bh-btn bh-btn-sm" title="이전">‹</button>
          <button type="button" id="product-page-next" class="bh-btn bh-btn-sm" title="다음">›</button>
        </div>
      </div>
    </div>

    <div id="product-modal-overlay" class="modal-overlay hidden">
      <div class="modal modal-product-full">
        <div class="modal-header product-modal-header-row">
          <h3 id="product-modal-title">개별상품등록</h3>
          <div class="product-modal-header-actions">
            <button type="button" id="product-form-reset" class="bh-btn bh-btn-sm bh-btn-outline">초기화</button>
            <button id="product-modal-close" class="bh-btn bh-btn-sm bh-btn-outline" type="button">닫기</button>
          </div>
        </div>
        <form id="product-popup-form" class="modal-form">
          <div><label>품목코드(이카운트)</label><input name="ecountCode" required /></div>
          <div><label>바코드(SKU)</label><input name="barcode" /></div>
          <div><label>바코드(중포)</label><input name="middleBarcode" /></div>
          <div><label>바코드(카톤)</label><input name="logisticsBarcode" /></div>
          <div><label>품목명(이카운트)</label><input name="ecountName" required /></div>
          <div><label>상태</label><select name="status">${selectOptions(opts.status, "판매중")}</select></div>
          <div class="multi-select-field">
            <label>판매처</label>
            <div class="row product-search-row multi-select-row">
              <select id="select-deliveryVendors"><option value="">옵션 선택</option>${selectItems(opts.deliveryVendors)}</select>
              <button type="button" id="add-deliveryVendors">추가</button>
            </div>
            <input type="hidden" name="deliveryVendors" />
          </div>
          <div id="preview-deliveryVendors" class="tagline multi-tags"></div>
          <div><label>판매처관리코드</label><input name="deliveryVendorCode" /></div>
          <div><label>판매처 품목명</label><input name="deliveryItemName" /></div>
          <div><label>규격</label><input name="spec" /></div>
          <div><label>구매처</label><input name="purchaseVendor" /></div>
          <div><label>수급형태</label><select name="supplyType"><option value=""></option>${selectOptions(opts.supplyType, "")}</select></div>
          <div><label>발주부서</label><select name="orderDept"><option value=""></option>${selectOptions(opts.orderDept, "")}</select></div>
          <div class="multi-select-field">
            <label>발주담당자</label>
            <div class="row product-search-row multi-select-row">
              <select id="select-orderManagers"><option value="">옵션 선택</option>${selectItems(opts.orderManagers)}</select>
              <button type="button" id="add-orderManagers">추가</button>
            </div>
            <input type="hidden" name="orderManagers" />
          </div>
          <div id="preview-orderManagers" class="tagline multi-tags"></div>
          <div><label>구매처 품목코드</label><input name="purchaseItemCode" /></div>
          <div><label>구매처 품목명</label><input name="purchaseItemName" /></div>
          <div><label>창고그룹(이카운트)</label><select name="warehouseGroup"><option value=""></option>${selectOptions(opts.warehouseGroup, "")}</select></div>
          <div class="multi-select-field">
            <label>사용창고</label>
            <div class="row product-search-row multi-select-row">
              <select id="select-usedWarehouses"><option value="">옵션 선택</option>${selectItems(state.warehouses || [])}</select>
              <button type="button" id="add-usedWarehouses">추가</button>
            </div>
            <input type="hidden" name="usedWarehouses" />
          </div>
          <div id="preview-usedWarehouses" class="tagline multi-tags"></div>
          <div><label>구분</label><select name="itemType"><option value=""></option>${selectOptions(opts.itemType, "")}</select></div>
          <div class="multi-select-field">
            <label>카테고리</label>
            <div class="row product-search-row multi-select-row">
              <select id="select-categories"><option value="">옵션 선택</option>${selectItems(opts.categories)}</select>
              <button type="button" id="add-categories">추가</button>
            </div>
            <input type="hidden" name="categories" />
          </div>
          <div id="preview-categories" class="tagline multi-tags"></div>
          <div><label>안전재고(선택)</label><input name="safetyStock" type="number" value="0" /></div>
          <div><label>적정재고(선택)</label><input name="optimalStock" type="number" value="0" /></div>
          <div><button class="primary" type="submit">입력 완료</button></div>
        </form>
      </div>
    </div>

    <div id="option-modal-overlay" class="modal-overlay hidden">
      <div class="modal">
        <div class="modal-header">
          <h3>옵션등록</h3>
          <button id="option-modal-close" class="cancel-btn del-small" type="button">닫기</button>
        </div>
        <form id="option-modal-form" class="modal-form">
          ${optionMeta
            .map(
              (m) => `
            <div><label>${m.label}</label><input name="${m.key}" value="${esc((opts[m.key] || []).join(", "))}" placeholder="${m.label} 옵션을 쉼표로 입력" /></div>
            <div class="tagline">${renderTagChips(opts[m.key], "tag tag-orange")}</div>
          `
            )
            .join("")}
          <div><button class="primary" type="submit">옵션 저장</button></div>
        </form>
      </div>
    </div>

    <div id="product-columns-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(420px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3>컬럼 설정</h3>
          <button type="button" id="product-columns-close" class="cancel-btn del-small">닫기</button>
        </div>
        <p class="muted" style="margin: 0 0 8px; font-size: 12px;">체크 해제 시 해당 열을 숨깁니다. (체크박스는 항상 표시)</p>
        <div id="product-columns-body" class="column-settings-list"></div>
        <button type="button" class="primary" id="product-columns-save" style="width: auto;">적용</button>
      </div>
    </div>

    ${datalist("opt-status", opts.status)}
    ${datalist("opt-deliveryVendors", opts.deliveryVendors)}
    ${datalist("opt-orderDept", opts.orderDept)}
    ${datalist("opt-orderManagers", opts.orderManagers)}
    ${datalist("opt-supplyType", opts.supplyType)}
    ${datalist("opt-warehouseGroup", opts.warehouseGroup)}
    ${datalist("opt-itemType", opts.itemType)}
    ${datalist("opt-categories", opts.categories)}

    <div id="stock-edit-modal-overlay" class="modal-overlay hidden">
      <div class="modal">
        <div class="modal-header">
          <h3>안전 / 적정재고 변경</h3>
          <button id="stock-edit-modal-close" class="cancel-btn" type="button">닫기</button>
        </div>
        <form id="stock-edit-form" class="modal-form">
          <div style="grid-column: span 2;"><label>상품코드</label><input id="stock-edit-code" name="code" required readonly /></div>
          <div style="grid-column: span 2;"><label>상품명</label><input id="stock-edit-name" name="name" required readonly /></div>
          <div style="grid-column: span 2;"><label>안전재고</label><input id="stock-edit-safety" name="safetyStock" type="number" required /></div>
          <div style="grid-column: span 2;"><label>적정재고</label><input id="stock-edit-optimal" name="optimalStock" type="number" required /></div>
          <div style="grid-column: span 6; margin-top:6px;"><button class="primary" type="submit">저장</button></div>
        </form>
      </div>
    </div>
  `;

  const modalOverlay = qs("#product-modal-overlay");
  const openBtn = qs("#open-product-popup");
  const closeBtn = qs("#product-modal-close");
  const productModalTitle = qs("#product-modal-title");
  const optionOverlay = qs("#option-modal-overlay");
  const optionOpenBtn = qs("#open-option-popup");
  const optionCloseBtn = qs("#option-modal-close");
  let editingCode = "";
  if (openBtn && modalOverlay) {
    openBtn.onclick = () => {
      editingCode = "";
      qs("#product-popup-form")?.reset();
      if (productModalTitle) productModalTitle.textContent = "개별상품등록";
      multiControllers.deliveryVendors?.setValues([]);
      multiControllers.orderManagers?.setValues([]);
      multiControllers.categories?.setValues([]);
      multiControllers.usedWarehouses?.setValues([]);
      modalOverlay.classList.remove("hidden");
    };
  }
  if (closeBtn && modalOverlay) closeBtn.onclick = () => modalOverlay.classList.add("hidden");
  if (modalOverlay) {
    modalOverlay.onclick = (e) => {
      if (e.target === modalOverlay) modalOverlay.classList.add("hidden");
    };
  }
  if (optionOpenBtn && optionOverlay) optionOpenBtn.onclick = () => optionOverlay.classList.remove("hidden");
  if (optionCloseBtn && optionOverlay) optionCloseBtn.onclick = () => optionOverlay.classList.add("hidden");
  if (optionOverlay) {
    optionOverlay.onclick = (e) => {
      if (e.target === optionOverlay) optionOverlay.classList.add("hidden");
    };
  }

  // 안전/적정재고 변경 모달
  const stockOverlay = qs("#stock-edit-modal-overlay");
  const stockCloseBtn = qs("#stock-edit-modal-close");
  const stockForm = qs("#stock-edit-form");
  const stockEditCode = qs("#stock-edit-code");
  const stockEditName = qs("#stock-edit-name");
  const stockEditSafety = qs("#stock-edit-safety");
  const stockEditOptimal = qs("#stock-edit-optimal");

  function openStockModal(p) {
    if (!stockOverlay) return;
    if (!p) return;
    stockEditCode.value = p.code;
    stockEditName.value = p.name;
    stockEditSafety.value = p.safetyStock ?? 0;
    stockEditOptimal.value = p.optimalStock ?? 0;
    stockOverlay.classList.remove("hidden");
  }
  function closeStockModal() {
    if (!stockOverlay) return;
    stockOverlay.classList.add("hidden");
  }

  if (stockCloseBtn && stockOverlay) stockCloseBtn.onclick = closeStockModal;
  if (stockOverlay) {
    stockOverlay.onclick = (e) => {
      if (e.target === stockOverlay) closeStockModal();
    };
  }

  if (stockForm) preventEnterSubmit(stockForm);
  if (stockForm) {
    stockForm.onsubmit = async (e) => {
      e.preventDefault();
      try {
        const code = stockEditCode.value;
        const name = stockEditName.value;
        const safetyStock = Number(stockEditSafety.value);
        const optimalStock = Number(stockEditOptimal.value);
        const current = state.products.find((p) => p.code === code);
        if (!current) throw new Error("상품을 찾을 수 없습니다.");
        await api("/api/products", {
          method: "POST",
          body: JSON.stringify({ ...current, safetyStock, optimalStock, name })
        });
        await refreshCommon();
        renderProducts();
        renderStock();
        renderMaster();
        await renderDashboard();
        closeStockModal();
        alert("안전/적정재고 저장 완료");
      } catch (err) {
        alert(err.message);
      }
    };
  }

  // 상품 조회 + 페이지네이션
  const searchInput = qs("#product-search-q");
  const searchBtn = qs("#product-search-btn");
  const resultEl = qs("#product-search-result");
  const pageSizeEl = qs("#product-page-size");
  const pageInfoEl = qs("#product-page-info");
  const pagePrev = qs("#product-page-prev");
  const pageNext = qs("#product-page-next");

  const applyProductListFilters = () => {
    const q = (searchInput?.value || "").trim().toLowerCase();
    const allRows = Array.from(document.querySelectorAll("#products-table tbody tr"));
    const matched = allRows.filter(
      (tr) =>
        tr.dataset.wmsExcelVisible !== "0" && (!q || String(tr.dataset.search || "").includes(q))
    );
    const size = Math.min(99999, Math.max(1, parseInt(pageSizeEl?.value || "100", 10) || 100));
    const total = matched.length;
    const pages = Math.max(1, Math.ceil(total / size));
    if (productListPageIndex >= pages) productListPageIndex = pages - 1;
    if (productListPageIndex < 0) productListPageIndex = 0;
    const page = productListPageIndex;
    const start = page * size;
    let mi = 0;
    allRows.forEach((tr) => {
      if (tr.dataset.wmsExcelVisible === "0") {
        tr.style.display = "none";
        return;
      }
      const isMatch = !q || String(tr.dataset.search || "").includes(q);
      if (!isMatch) {
        tr.style.display = "none";
        return;
      }
      const vis = mi >= start && mi < start + size;
      tr.style.display = vis ? "" : "none";
      mi += 1;
    });
    if (pageInfoEl) {
      if (total === 0) pageInfoEl.textContent = "0 - 0 / 0";
      else pageInfoEl.textContent = `${start + 1} - ${Math.min(total, start + size)} / ${total}`;
    }
    if (resultEl) resultEl.textContent = q ? `\uac80\uc0c9 \uacb0\uacfc: ${total}\uac74` : `\uc804\uccb4 ${allRows.length}\uac74`;
  };

  if (searchBtn && searchInput && resultEl) {
    preventEnterSubmit(qs(".products-bh-search-wrap"));
    searchBtn.onclick = () => {
      productListPageIndex = 0;
      applyProductListFilters();
    };
    searchInput.onkeydown = (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        productListPageIndex = 0;
        applyProductListFilters();
      }
    };
    pageSizeEl?.addEventListener("change", () => {
      productListPageIndex = 0;
      applyProductListFilters();
    });
    pagePrev?.addEventListener("click", () => {
      productListPageIndex = Math.max(0, productListPageIndex - 1);
      applyProductListFilters();
    });
    pageNext?.addEventListener("click", () => {
      const q2 = (searchInput?.value || "").trim().toLowerCase();
      const allRows2 = Array.from(document.querySelectorAll("#products-table tbody tr"));
      const total2 = allRows2.filter(
        (tr) =>
          tr.dataset.wmsExcelVisible !== "0" && (!q2 || String(tr.dataset.search || "").includes(q2))
      ).length;
      const size2 = Math.min(99999, Math.max(1, parseInt(pageSizeEl?.value || "100", 10) || 100));
      const pages2 = Math.max(1, Math.ceil(total2 / size2));
      productListPageIndex = Math.min(pages2 - 1, productListPageIndex + 1);
      applyProductListFilters();
    });
    applyProductListFilters();
  }

  if (!window.__wmsProductUiInit) {
    window.__wmsProductUiInit = true;
    document.addEventListener("click", () => {
      qs("#excel-import-menu")?.classList.add("hidden");
      qs("#more-menu")?.classList.add("hidden");
    });
  }
  qs("#excel-import-toggle")?.addEventListener("click", (e) => {
    e.stopPropagation();
    const m = qs("#excel-import-menu");
    const wasHidden = m?.classList.contains("hidden");
    qs("#excel-import-menu")?.classList.add("hidden");
    qs("#more-menu")?.classList.add("hidden");
    if (wasHidden) m?.classList.remove("hidden");
  });
  qs("#more-menu-toggle")?.addEventListener("click", (e) => {
    e.stopPropagation();
    const m = qs("#more-menu");
    const wasHidden = m?.classList.contains("hidden");
    qs("#excel-import-menu")?.classList.add("hidden");
    qs("#more-menu")?.classList.add("hidden");
    if (wasHidden) m?.classList.remove("hidden");
  });
  qs("#excel-import-menu")?.addEventListener("click", (e) => e.stopPropagation());
  qs("#more-menu")?.addEventListener("click", (e) => e.stopPropagation());
  qs("#menu-product-guide")?.addEventListener("click", () => {
    downloadGuide("product");
    qs("#excel-import-menu")?.classList.add("hidden");
  });
  qs("#menu-bulk-new")?.addEventListener("click", () => {
    qs("#product-file")?.click();
    qs("#excel-import-menu")?.classList.add("hidden");
  });
  qs("#menu-bulk-update")?.addEventListener("click", () => {
    qs("#product-update-file")?.click();
    qs("#excel-import-menu")?.classList.add("hidden");
  });

  const colOverlay = qs("#product-columns-overlay");
  qs("#product-column-settings")?.addEventListener("click", () => {
    const body = qs("#product-columns-body");
    if (!body || !colOverlay) return;
    const heads = Array.from(document.querySelectorAll("#products-table thead th"));
    let hidden = [];
    try {
      hidden = JSON.parse(localStorage.getItem(PRODUCT_HIDDEN_COLS_KEY) || "[]");
    } catch (_) {
      hidden = [];
    }
    body.innerHTML = heads
      .map((th, i) => {
        const nth = i + 1;
        if (nth === 1) return "";
        const lab = th.textContent.trim() || `\uc5f4 ${nth}`;
        const checked = !hidden.includes(nth);
        return `<label><input type="checkbox" data-nth="${nth}" ${checked ? "checked" : ""} /> ${esc(lab)}</label>`;
      })
      .join("");
    colOverlay.classList.remove("hidden");
  });
  qs("#product-columns-close")?.addEventListener("click", () => colOverlay?.classList.add("hidden"));
  if (colOverlay) {
    colOverlay.addEventListener("click", (e) => {
      if (e.target === colOverlay) colOverlay.classList.add("hidden");
    });
  }
  qs("#product-columns-save")?.addEventListener("click", () => {
    const body = qs("#product-columns-body");
    const hidden = [];
    body?.querySelectorAll('input[type="checkbox"]').forEach((inp) => {
      const nth = parseInt(inp.dataset.nth || "0", 10);
      if (!nth || nth === 1) return;
      if (!inp.checked) hidden.push(nth);
    });
    try {
      localStorage.setItem(PRODUCT_HIDDEN_COLS_KEY, JSON.stringify(hidden));
    } catch (_) {
      /* ignore */
    }
    applyProductHiddenColumnStyles();
    colOverlay?.classList.add("hidden");
  });

  preventEnterSubmit(qs("#product-popup-form"));
  const productForm = qs("#product-popup-form");
  const multiControllers = {};
  const setupMultiSelect = (field, selectId, addId, previewId) => {
    const hidden = productForm?.querySelector(`[name="${field}"]`);
    const selectEl = qs(`#${selectId}`);
    const addBtn = qs(`#${addId}`);
    const previewEl = qs(`#${previewId}`);
    if (!hidden || !selectEl || !addBtn || !previewEl) return { setValues: () => {} };
    let selected = [];
    const draw = () => {
      hidden.value = selected.join(", ");
      previewEl.innerHTML = selected.length
        ? selected
            .map(
              (v) =>
                `<span class="tag tag-orange tag-removable">${esc(v)} <button type="button" class="chip-remove" data-value="${encodeURIComponent(v)}">x</button></span>`
            )
            .join("")
        : "<span class='muted'>선택한 항목 없음</span>";
    };
    const addValue = (v) => {
      const value = String(v || "").trim();
      if (!value) return;
      if (!selected.includes(value)) selected.push(value);
      draw();
    };
    addBtn.onclick = () => {
      addValue(selectEl.value);
      selectEl.value = "";
    };
    previewEl.onclick = (e) => {
      const btn = e.target.closest(".chip-remove");
      if (!btn) return;
      const value = decodeURIComponent(btn.dataset.value || "");
      selected = selected.filter((x) => x !== value);
      draw();
    };
    draw();
    return {
      setValues(values) {
        selected = toTagList(values);
        draw();
      }
    };
  };
  multiControllers.deliveryVendors = setupMultiSelect("deliveryVendors", "select-deliveryVendors", "add-deliveryVendors", "preview-deliveryVendors");
  multiControllers.orderManagers = setupMultiSelect("orderManagers", "select-orderManagers", "add-orderManagers", "preview-orderManagers");
  multiControllers.categories = setupMultiSelect("categories", "select-categories", "add-categories", "preview-categories");
  multiControllers.usedWarehouses = setupMultiSelect("usedWarehouses", "select-usedWarehouses", "add-usedWarehouses", "preview-usedWarehouses");

  qs("#product-form-reset")?.addEventListener("click", () => {
    editingCode = "";
    qs("#product-popup-form")?.reset();
    if (productModalTitle) productModalTitle.textContent = "개별상품등록";
    multiControllers.deliveryVendors?.setValues([]);
    multiControllers.orderManagers?.setValues([]);
    multiControllers.categories?.setValues([]);
    multiControllers.usedWarehouses?.setValues([]);
  });

  const optionForm = qs("#option-modal-form");
  if (optionForm) {
    optionForm.onsubmit = async (e) => {
      e.preventDefault();
      try {
        const fd = new FormData(optionForm);
        for (const m of optionMeta) {
          const values = toTagList(fd.get(m.key));
          await api("/api/product-options", { method: "POST", body: JSON.stringify({ field: m.key, values, replace: true }) });
        }
        await refreshCommon();
        renderProducts();
        if (optionOverlay) optionOverlay.classList.add("hidden");
        alert("옵션 저장 완료");
      } catch (err) {
        alert(err.message);
      }
    };
  }

  const getSelectedCodes = () =>
    Array.from(document.querySelectorAll(".product-row-check:checked")).map((el) => String(el.dataset.code || ""));
  const checkAllEl = qs("#product-check-all");
  if (checkAllEl) {
    checkAllEl.onchange = () => {
      document.querySelectorAll(".product-row-check").forEach((el) => {
        el.checked = checkAllEl.checked;
      });
    };
  }

  const editSelectedBtn = qs("#product-edit-selected");
  if (editSelectedBtn) {
    editSelectedBtn.onclick = () => {
      const codes = getSelectedCodes();
      if (codes.length !== 1) return alert("수정은 1개만 선택하세요.");
      const p = state.products.find((x) => String(x.code) === String(codes[0]));
      if (!p || !productForm || !modalOverlay) return;
      editingCode = String(p.code || "");
      if (productModalTitle) productModalTitle.textContent = "상품등록정보 수정";
      const setVal = (name, val) => {
        const el = productForm.querySelector(`[name="${name}"]`);
        if (el) el.value = val ?? "";
      };
      setVal("ecountCode", p.ecountCode || p.code);
      setVal("barcode", p.barcode);
      setVal("middleBarcode", p.middleBarcode);
      setVal("logisticsBarcode", p.logisticsBarcode);
      setVal("ecountName", p.ecountName || p.name);
      setVal("status", p.status);
      multiControllers.deliveryVendors?.setValues(p.deliveryVendors);
      setVal("deliveryVendorCode", p.deliveryVendorCode);
      setVal("deliveryItemName", p.deliveryItemName);
      setVal("spec", p.spec);
      setVal("purchaseVendor", p.purchaseVendor);
      setVal("supplyType", p.supplyType);
      setVal("orderDept", p.orderDept);
      multiControllers.orderManagers?.setValues(p.orderManagers);
      setVal("purchaseItemCode", p.purchaseItemCode);
      setVal("purchaseItemName", p.purchaseItemName);
      setVal("warehouseGroup", p.warehouseGroup);
      multiControllers.usedWarehouses?.setValues(p.usedWarehouses);
      setVal("itemType", p.itemType);
      multiControllers.categories?.setValues(p.categories);
      setVal("safetyStock", p.safetyStock ?? 0);
      setVal("optimalStock", p.optimalStock ?? 0);
      modalOverlay.classList.remove("hidden");
    };
  }

  const deleteSelectedBtn = qs("#product-delete-selected");
  if (deleteSelectedBtn) {
    deleteSelectedBtn.onclick = async () => {
      const codes = getSelectedCodes();
      if (!codes.length) return alert("삭제할 상품을 선택하세요.");
      if (!confirm(`선택한 ${codes.length}개 상품을 삭제할까요?`)) return;
      try {
        await api("/api/products", { method: "DELETE", body: JSON.stringify({ codes }) });
        await refreshCommon();
        renderProducts();
        renderStock();
        renderMaster();
        await renderDashboard();
        alert("선택 상품 삭제 완료");
      } catch (err) {
        alert(err.message);
      }
    };
  }

  qs("#product-popup-form").onsubmit = async (e) => {
    e.preventDefault();
    const data = Object.fromEntries(new FormData(e.target));
    data.ecountCode = String(data.ecountCode || "").trim();
    data.ecountName = String(data.ecountName || "").trim();
    data.status = toTagList(data.status)[0] || "판매중";
    data.supplyType = toTagList(data.supplyType)[0] || "";
    data.orderDept = toTagList(data.orderDept)[0] || "";
    data.warehouseGroup = toTagList(data.warehouseGroup)[0] || "";
    data.itemType = toTagList(data.itemType)[0] || "";
    data.code = editingCode || data.ecountCode;
    data.name = data.ecountName;
    data.deliveryVendors = toTagList(data.deliveryVendors);
    data.orderManagers = toTagList(data.orderManagers);
    data.categories = toTagList(data.categories);
    data.usedWarehouses = toTagList(data.usedWarehouses);
    await api("/api/products", { method: "POST", body: JSON.stringify(data) });
    await refreshCommon();
    editingCode = "";
    if (modalOverlay) modalOverlay.classList.add("hidden");
    renderProducts();
    renderStock();
    renderMaster();
    await renderDashboard();
    alert("저장되었습니다.");
  };

  const productFileInput = qs("#product-file");
  const productUpdateInput = qs("#product-update-file");

  const afterBulkProductImport = async (label) => {
    renderProducts();
    renderStock();
    renderMaster();
    await renderDashboard();
    alert(label);
  };

  productFileInput?.addEventListener("change", async () => {
    const file = productFileInput.files[0];
    if (!file) return;
    try {
      const n = await uploadProductBulkFromFile(file);
      productFileInput.value = "";
      await afterBulkProductImport(`엑셀 업로드 완료 (${n}건)`);
    } catch (e2) {
      alert(e2.message);
      productFileInput.value = "";
    }
  });

  productUpdateInput?.addEventListener("change", async () => {
    const file = productUpdateInput.files[0];
    if (!file) return;
    try {
      const n = await uploadProductUpdateFromFile(file);
      productUpdateInput.value = "";
      await afterBulkProductImport(`상품정보 업데이엄트 완료 (${n}건)`);
    } catch (err) {
      alert(err.message);
      productUpdateInput.value = "";
    }
  });

  qs("#product-export-all").onclick = async () => {
    try {
      const latest = await api("/api/products");
      const items = Array.isArray(latest.items) ? latest.items : [];
      if (!items.length) {
        alert("다운로드할 상품정보가 없습니다.");
        return;
      }
      const rowsForExport = items.map((p) => ({
        "품목코드(이카운트)": p.ecountCode || p.code || "",
        "바코드(SKU)": p.barcode || "",
        "바코드(중포)": p.middleBarcode || "",
        "바코드(카톤)": p.logisticsBarcode || "",
        "품목명(이카운트)": p.ecountName || p.name || "",
        "상태": p.status || "",
        "판매처": toTagList(p.deliveryVendors).join(", "),
        "판매처관리코드": p.deliveryVendorCode || "",
        "판매처 품목명": p.deliveryItemName || "",
        "규격": p.spec || "",
        "구매처": p.purchaseVendor || "",
        "수급형태": p.supplyType || "",
        "발주부서": p.orderDept || "",
        "발주담당자": toTagList(p.orderManagers).join(", "),
        "구매처 품목코드": p.purchaseItemCode || "",
        "구매처 품목명": p.purchaseItemName || "",
        "창고그룹(이카운트)": p.warehouseGroup || "",
        "사용창고": toTagList(p.usedWarehouses).join(", "),
        "구분": p.itemType || "",
        "카테고리": toTagList(p.categories).join(", ")
      }));
      const ws = XLSX.utils.json_to_sheet(rowsForExport);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "products");
      XLSX.writeFile(wb, `products-all-${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
      alert(err.message || "전체 다운로드 실패");
    }
  };
  applyProductHiddenColumnStyles();
  applyExcelLikeFilter("#products-table", applyProductListFilters);
}

function renderStock() {
  const all = state.stock;
  const totalLow = all.filter((s) => Number(s.stock) <= Number(s.safetyStock || 0)).length;
  const totalOk = all.filter((s) => Number(s.stock) > Number(s.safetyStock || 0) && Number(s.stock) <= Number(s.optimalStock || Infinity)).length;
  const totalOver = all.filter((s) => Number(s.optimalStock || 0) > 0 && Number(s.stock) > Number(s.optimalStock || 0)).length;

  const warehouseOptions = [`<option value="ALL">창고 전체</option>`, ...(state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`)].join("");

  qs("#view-stock").innerHTML = `
    <div class="stock-page">
      <div class="stock-header">
        <div class="stock-header-title">
          <h2>재고현황</h2>
          <span class="stock-total-badge">${all.length}개 상품</span>
        </div>
        <div class="stock-summary-cards">
          <div class="stock-summary-card stock-sc-all" data-filter="ALL">
            <div class="stock-sc-label">전체</div>
            <div class="stock-sc-value">${all.length}</div>
          </div>
          <div class="stock-summary-card stock-sc-ok" data-filter="OK">
            <div class="stock-sc-label">정상</div>
            <div class="stock-sc-value">${totalOk}</div>
          </div>
          <div class="stock-summary-card stock-sc-low" data-filter="LOW">
            <div class="stock-sc-label">재고부족</div>
            <div class="stock-sc-value">${totalLow}</div>
          </div>
          <div class="stock-summary-card stock-sc-over" data-filter="OVER">
            <div class="stock-sc-label">적정초과</div>
            <div class="stock-sc-value">${totalOver}</div>
          </div>
        </div>
      </div>

      <div class="stock-toolbar">
        <div class="stock-search-wrap">
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input id="stock-q" placeholder="상품코드, 상품명, 바코드 검색" />
        </div>
        <select id="stock-warehouse">${warehouseOptions}</select>
        <select id="stock-status">
          <option value="ALL">상태 전체</option>
          <option value="LOW">재고부족</option>
          <option value="OK">정상</option>
          <option value="OVER">적정재고 초과</option>
        </select>
        <span id="stock-count-label" class="stock-count-label">표시: 0건</span>
      </div>

      <div class="stock-table-wrap">
        <table id="stock-table" class="stock-table draggable-table">
          <thead>
            <tr>
              <th draggable="true">상태</th>
              <th draggable="true">창고</th>
              <th draggable="true">상품코드</th>
              <th draggable="true">상품명</th>
              <th draggable="true">카테고리</th>
              <th draggable="true">판매처</th>
              <th draggable="true">구매처</th>
              <th draggable="true">단위</th>
              <th draggable="true">안전재고</th>
              <th draggable="true">적정재고</th>
              <th draggable="true">현재고</th>
              <th draggable="true">재고수준</th>
            </tr>
          </thead>
          <tbody id="stock-tbody"></tbody>
        </table>
      </div>
    </div>
  `;

  const draw = () => {
    const q = (qs("#stock-q")?.value || "").trim().toLowerCase();
    const status = (qs("#stock-status")?.value || "ALL").trim();
    const warehouse = (qs("#stock-warehouse")?.value || "ALL").trim();

    const filtered = state.stock.filter((s) => {
      if (warehouse !== "ALL" && String(s.warehouse || "") !== warehouse) return false;
      if (q) {
        const hay = `${s.code} ${s.name} ${s.ecountCode || ""} ${s.barcode || ""} ${s.warehouse || ""}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      if (status === "LOW") return Number(s.stock) <= Number(s.safetyStock || 0);
      if (status === "OK") return Number(s.stock) > Number(s.safetyStock || 0) && (Number(s.optimalStock || 0) === 0 || Number(s.stock) <= Number(s.optimalStock || 0));
      if (status === "OVER") return Number(s.optimalStock || 0) > 0 && Number(s.stock) > Number(s.optimalStock || 0);
      return true;
    });

    const countLabel = qs("#stock-count-label");
    if (countLabel) countLabel.textContent = `표시: ${filtered.length}건`;

    const tbody = qs("#stock-tbody");
    if (!tbody) return;

    tbody.innerHTML = filtered.map((s) => {
      const stock = Number(s.stock || 0);
      const safety = Number(s.safetyStock || 0);
      const optimal = Number(s.optimalStock || 0);
      const isLow = stock <= safety;
      const isOver = optimal > 0 && stock > optimal;

      let statusBadge, rowClass;
      if (isLow) {
        statusBadge = `<span class="stock-badge stock-badge-low">부족</span>`;
        rowClass = "stock-row-low";
      } else if (isOver) {
        statusBadge = `<span class="stock-badge stock-badge-over">초과</span>`;
        rowClass = "stock-row-over";
      } else {
        statusBadge = `<span class="stock-badge stock-badge-ok">정상</span>`;
        rowClass = "";
      }

      let barHtml = "-";
      if (optimal > 0) {
        const pct = Math.min((stock / optimal) * 100, 100).toFixed(0);
        const barColor = isLow ? "var(--red)" : isOver ? "var(--accent-2)" : "var(--green)";
        barHtml = `
          <div class="stock-bar-wrap">
            <div class="stock-bar-track">
              <div class="stock-bar-fill" style="width:${pct}%;background:${barColor};"></div>
            </div>
            <span class="stock-bar-pct">${pct}%</span>
          </div>`;
      }

      return `<tr class="${rowClass}" data-code="${esc(s.code)}">
        <td>${statusBadge}</td>
        <td>${esc(s.warehouse || "-")}</td>
        <td class="stock-code">${esc(s.code)}</td>
        <td class="stock-name">${esc(s.name)}</td>
        <td>${esc(s.category || "-")}</td>
        <td>${esc(s.salesVendor || "-")}</td>
        <td>${esc(s.purchaseVendor || "-")}</td>
        <td>${esc(s.unit || "-")}</td>
        <td class="num">${safety}</td>
        <td class="num">${optimal || "-"}</td>
        <td class="num stock-qty ${isLow ? "stock-qty-low" : ""}">${stock}</td>
        <td>${barHtml}</td>
      </tr>`;
    }).join("");

    applyExcelLikeFilter("#stock-table");
    enableColumnDrag("#stock-table");

    // summary card click filter
    qs(".stock-page").querySelectorAll(".stock-summary-card").forEach((card) => {
      card.onclick = () => {
        const f = card.dataset.filter;
        const stEl2 = qs("#stock-status");
        if (stEl2) { stEl2.value = f; draw(); }
      };
    });
  };

  draw();

  const qEl = qs("#stock-q");
  const wEl = qs("#stock-warehouse");
  const stEl = qs("#stock-status");
  if (qEl) qEl.oninput = () => draw();
  if (qEl) qEl.onkeydown = (e) => { if (e.key === "Enter") { e.preventDefault(); draw(); } };
  if (wEl) wEl.onchange = () => draw();
  if (stEl) stEl.onchange = () => draw();
}

function enableColumnDrag(tableSelector) {
  const table = document.querySelector(tableSelector);
  if (!table) return;
  const headRow = table.tHead && table.tHead.rows && table.tHead.rows[0] ? table.tHead.rows[0] : null;
  if (!headRow) return;

  const ths = Array.from(headRow.children);
  if (!ths.length) return;

  let dragIndex = null;

  ths.forEach((th, idx) => {
    th.classList.add("draggable-th");
    th.ondragstart = (e) => {
      dragIndex = idx;
      e.dataTransfer.effectAllowed = "move";
    };
    th.ondragover = (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
    };
    th.ondrop = (e) => {
      e.preventDefault();
      if (dragIndex === null || dragIndex === idx) return;

      // Swap all header rows (title row + excel filter row)
      const headRows = Array.from(table.tHead.rows || []);
      for (const hr of headRows) {
        const a = hr.children[dragIndex];
        const b = hr.children[idx];
        if (!a || !b) continue;
        const tmp = a.innerHTML;
        a.innerHTML = b.innerHTML;
        b.innerHTML = tmp;
      }

      // Swap body cells content
      const tbodyRows = Array.from(table.tBodies[0]?.rows || []);
      for (const r of tbodyRows) {
        const ca = r.children[dragIndex];
        const cb = r.children[idx];
        if (!ca || !cb) continue;
        const tmp2 = ca.innerHTML;
        ca.innerHTML = cb.innerHTML;
        cb.innerHTML = tmp2;
      }

      dragIndex = null;
      applyExcelLikeFilter(tableSelector);
    };
  });
}

function movementFormHtml(type, title, hint, partnerType, partnerLabel, partnerOptional = false) {
  const productOptions = state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join("");
  const partnerOptions = (state.partners[partnerType] || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const warehouseOptions = (state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");
  const partnerHint =
    type === "IN"
      ? `<span class="muted" style="font-size:12px;display:block;margin-top:4px;">미입력 시 해당 상품 기본정보의 구매처가 자동 적용됩니다.</span>`
      : "";
  return `
    <div class="card">
      <h2>${title}</h2>
      <form id="${type}-form">
        <div><label>상품 검색(코드)</label><input name="productCode" list="${type}-product-list" required /></div>
        <datalist id="${type}-product-list">${productOptions}</datalist>
        <div><label>수량</label><input name="qty" type="number" required /></div>
        <div><label>창고</label><input name="warehouse" list="${type}-warehouse-list" required /></div>
        <datalist id="${type}-warehouse-list">${warehouseOptions}</datalist>
        <div><label>${partnerLabel}</label><input name="partner" list="${type}-partner-list" ${partnerOptional ? "" : "required"} />${partnerHint}</div>
        <datalist id="${type}-partner-list">${partnerOptions}</datalist>
        <div><label>담당자</label><input name="user" list="${type}-manager-list" required /></div>
        <datalist id="${type}-manager-list">${managerOptions}</datalist>
        <div><label>메모</label><input name="memo" /></div>
        <div><button id="${type}-submit" class="primary" type="button">등록</button></div>
      </form>
      <p class="muted">${hint}</p>
    </div>
    <div class="card movement-upload-card">
      <h3>엑셀 업로드</h3>
      <div class="row">
        <button id="${type}-guide" type="button">가이드 다운로드</button>
        <button id="${type}-upload" class="primary" type="button">업로드</button>
      </div>
      <div id="${type}-dropzone" class="dropzone">파일을 여기로 드래그하거나 클릭해서 선택하세요</div>
      <input type="file" id="${type}-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
      <p class="muted">컬럼명: 상품코드, 수량, 창고, ${partnerLabel}(생략 시 상품 기본정보 구매처), 담당자, 메모</p>
    </div>
    <div class="card">
      <h3>최근 등록 내역</h3>
      <div id="recent-${type}" class="muted">최근 내역을 불러오는 중...</div>
    </div>
  `;
}

/** 입고 화면: 상단은 버튼만, 단건·엑셀은 각각 팝업에서 처리 */
function movementInboundPageHtml() {
  const partnerLabel = "구매처";
  const hint = "등록 버튼 클릭시에만 등록됩니다.";
  const productOptions = state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join("");
  const partnerOptions = (state.partners.purchase || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const warehouseOptions = (state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");
  const partnerHint = `<span class="muted" style="font-size:12px;display:block;margin-top:4px;">미입력 시 해당 상품 기본정보의 구매처가 자동 적용됩니다.</span>`;
  return `
    <div class="card inbound-toolbar-card">
      <h2>입고</h2>
      <p class="muted inbound-toolbar-hint">한 품목만 등록할 때는 <strong>단건입고</strong>, 여러 행은 <strong>엑셀 업로드</strong>를 사용하세요.</p>
      <div class="inbound-quick-actions">
        <button type="button" class="primary" id="inbound-open-single">단건입고</button>
        <button type="button" class="cancel-btn inbound-excel-open-btn" id="inbound-open-excel">엑셀 업로드</button>
      </div>
    </div>

    <div id="inbound-single-modal" class="modal-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="inbound-single-title">
      <div class="modal inbound-movement-modal">
        <div class="modal-header">
          <h3 id="inbound-single-title">단건입고</h3>
          <button type="button" class="cancel-btn del-small" id="inbound-single-close">닫기</button>
        </div>
        <form id="IN-form" class="modal-form">
          <div><label>상품 검색(코드)</label><input name="productCode" list="IN-product-list" required /></div>
          <datalist id="IN-product-list">${productOptions}</datalist>
          <div><label>수량</label><input name="qty" type="number" required /></div>
          <div><label>창고</label><input name="warehouse" list="IN-warehouse-list" required /></div>
          <datalist id="IN-warehouse-list">${warehouseOptions}</datalist>
          <div><label>${partnerLabel}</label><input name="partner" list="IN-partner-list" />${partnerHint}</div>
          <datalist id="IN-partner-list">${partnerOptions}</datalist>
          <div><label>담당자</label><input name="user" list="IN-manager-list" required /></div>
          <datalist id="IN-manager-list">${managerOptions}</datalist>
          <div><label>메모</label><input name="memo" /></div>
          <div><button id="IN-submit" class="primary" type="button">등록</button></div>
        </form>
        <p class="muted">${hint}</p>
      </div>
    </div>

    <div id="inbound-excel-modal" class="modal-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="inbound-excel-title">
      <div class="modal inbound-movement-modal">
        <div class="modal-header">
          <h3 id="inbound-excel-title">엑셀 업로드</h3>
          <button type="button" class="cancel-btn del-small" id="inbound-excel-close">닫기</button>
        </div>
        <div class="movement-upload-card inbound-modal-upload">
          <div class="row">
            <button id="IN-guide" type="button">가이드 다운로드</button>
            <button id="IN-upload" class="primary" type="button">업로드</button>
          </div>
          <div id="IN-dropzone" class="dropzone">파일을 여기로 드래그하거나 클릭해서 선택하세요</div>
          <input type="file" id="IN-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
          <p class="muted">컬럼명: 상품코드, 수량, 창고, ${partnerLabel}(생략 시 상품 기본정보 구매처), 담당자, 메모</p>
        </div>
      </div>
    </div>

    <div class="card">
      <h3>최근 등록 내역</h3>
      <div id="recent-IN" class="muted">최근 내역을 불러오는 중...</div>
    </div>
  `;
}

function setupInboundModals() {
  const single = qs("#inbound-single-modal");
  const excel = qs("#inbound-excel-modal");
  const closeSingle = () => single?.classList.add("hidden");
  const closeExcel = () => excel?.classList.add("hidden");
  qs("#inbound-open-single")?.addEventListener("click", () => {
    single?.classList.remove("hidden");
  });
  qs("#inbound-open-excel")?.addEventListener("click", () => {
    excel?.classList.remove("hidden");
  });
  qs("#inbound-single-close")?.addEventListener("click", closeSingle);
  qs("#inbound-excel-close")?.addEventListener("click", closeExcel);
  single?.addEventListener("click", (e) => {
    if (e.target === single) closeSingle();
  });
  excel?.addEventListener("click", (e) => {
    if (e.target === excel) closeExcel();
  });
}

/** 출고 화면: 입고와 동일 — 상단 버튼만, 단건·엑셀은 팝업 */
function movementOutboundPageHtml() {
  const partnerLabel = "출고처";
  const hint = "등록 버튼 클릭시에만 등록됩니다.";
  const productOptions = state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join("");
  const partnerOptions = (state.partners.outbound || []).map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const warehouseOptions = (state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");
  return `
    <div class="card outbound-toolbar-card">
      <h2>출고</h2>
      <p class="muted outbound-toolbar-hint">한 품목만 등록할 때는 <strong>단건출고</strong>, 여러 행은 <strong>엑셀 업로드</strong>를 사용하세요.</p>
      <div class="outbound-quick-actions">
        <button type="button" class="primary" id="outbound-open-single">단건출고</button>
        <button type="button" class="cancel-btn outbound-excel-open-btn" id="outbound-open-excel">엑셀 업로드</button>
      </div>
    </div>

    <div id="outbound-single-modal" class="modal-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="outbound-single-title">
      <div class="modal inbound-movement-modal">
        <div class="modal-header">
          <h3 id="outbound-single-title">단건출고</h3>
          <button type="button" class="cancel-btn del-small" id="outbound-single-close">닫기</button>
        </div>
        <form id="OUT-form" class="modal-form">
          <div><label>상품 검색(코드)</label><input name="productCode" list="OUT-product-list" required /></div>
          <datalist id="OUT-product-list">${productOptions}</datalist>
          <div><label>수량</label><input name="qty" type="number" required /></div>
          <div><label>창고</label><input name="warehouse" list="OUT-warehouse-list" required /></div>
          <datalist id="OUT-warehouse-list">${warehouseOptions}</datalist>
          <div><label>${partnerLabel}</label><input name="partner" list="OUT-partner-list" required /></div>
          <datalist id="OUT-partner-list">${partnerOptions}</datalist>
          <div><label>담당자</label><input name="user" list="OUT-manager-list" required /></div>
          <datalist id="OUT-manager-list">${managerOptions}</datalist>
          <div><label>메모</label><input name="memo" /></div>
          <div><button id="OUT-submit" class="primary" type="button">등록</button></div>
        </form>
        <p class="muted">${hint}</p>
      </div>
    </div>

    <div id="outbound-excel-modal" class="modal-overlay hidden" role="dialog" aria-modal="true" aria-labelledby="outbound-excel-title">
      <div class="modal inbound-movement-modal">
        <div class="modal-header">
          <h3 id="outbound-excel-title">엑셀 업로드</h3>
          <button type="button" class="cancel-btn del-small" id="outbound-excel-close">닫기</button>
        </div>
        <div class="movement-upload-card inbound-modal-upload">
          <div class="row">
            <button id="OUT-guide" type="button">가이드 다운로드</button>
            <button id="OUT-upload" class="primary" type="button">업로드</button>
          </div>
          <div id="OUT-dropzone" class="dropzone">파일을 여기로 드래그하거나 클릭해서 선택하세요</div>
          <input type="file" id="OUT-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
          <p class="muted">컬럼명: 상품코드, 수량, 창고, ${partnerLabel}, 담당자, 메모</p>
        </div>
      </div>
    </div>

    <div class="card">
      <h3>최근 등록 내역</h3>
      <div id="recent-OUT" class="muted">최근 내역을 불러오는 중...</div>
    </div>
  `;
}

function setupOutboundModals() {
  const single = qs("#outbound-single-modal");
  const excel = qs("#outbound-excel-modal");
  const closeSingle = () => single?.classList.add("hidden");
  const closeExcel = () => excel?.classList.add("hidden");
  qs("#outbound-open-single")?.addEventListener("click", () => {
    single?.classList.remove("hidden");
  });
  qs("#outbound-open-excel")?.addEventListener("click", () => {
    excel?.classList.remove("hidden");
  });
  qs("#outbound-single-close")?.addEventListener("click", closeSingle);
  qs("#outbound-excel-close")?.addEventListener("click", closeExcel);
  single?.addEventListener("click", (e) => {
    if (e.target === single) closeSingle();
  });
  excel?.addEventListener("click", (e) => {
    if (e.target === excel) closeExcel();
  });
}

function movementPayload(type) {
  const form = qs(`#${type}-form`);
  return { ...Object.fromEntries(new FormData(form)), type };
}

function bindMovement(type) {
  const form = qs(`#${type}-form`);
  preventEnterSubmit(form);
  qs(`#${type}-submit`).onclick = async () => {
    try {
      let payload = movementPayload(type);
      if (type === "IN") {
        payload = fillInboundPartnerFromProduct(payload);
        if (!String(payload.partner || "").trim()) {
          alert("구매처가 없습니다. 입력하거나 상품 기본정보에 구매처를 등록하세요.");
          return;
        }
      }
      await api("/api/movements", { method: "POST", body: JSON.stringify(payload) });
      await afterMovementDone();
      alert("등록되었습니다.");
      if (type === "IN") qs("#inbound-single-modal")?.classList.add("hidden");
      if (type === "OUT") qs("#outbound-single-modal")?.classList.add("hidden");
    } catch (err) {
      alert(err.message);
    }
  };

  setupDropZone(`${type}-dropzone`, `${type}-file`);
  qs(`#${type}-guide`).onclick = () => downloadGuide(type);
  qs(`#${type}-upload`).onclick = async () => {
    const file = qs(`#${type}-file`).files[0];
    if (!file) return alert("파일을 선택하세요.");
    try {
      const rows = await parseSheet(file);
      const normalized = normalizeMovementRows(rows, type).map((r) => ({ ...r, type }));
      await api("/api/movements/bulk", { method: "POST", body: JSON.stringify({ rows: normalized }) });
      await afterMovementDone();
      alert("업로드 완료");
      if (type === "IN") qs("#inbound-excel-modal")?.classList.add("hidden");
      if (type === "OUT") qs("#outbound-excel-modal")?.classList.add("hidden");
    } catch (e) {
      alert(e.message);
    }
  };
}

function renderInbound() {
  qs("#view-inbound").innerHTML = movementInboundPageHtml();
  bindMovement("IN");
  setupInboundModals();
  renderRecentMovements("IN");
}

async function renderInboundPlan() {
  const wrap = qs("#view-inbound-plan");
  if (!wrap) return;
  const today = new Date();
  const from = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30).toISOString().slice(0, 10);
  const to = today.toISOString().slice(0, 10);
  wrap.innerHTML = `
    <div class="card">
      <h2>입고예정목록 1</h2>
      <p class="muted">이카운트 OAPI로 발주 데이터를 읽기 전용 조회합니다.</p>
      <div class="row" style="grid-template-columns: 180px 180px 100px 1fr;">
        <input id="inbound-plan-from" type="date" value="${from}" />
        <input id="inbound-plan-to" type="date" value="${to}" />
        <button id="inbound-plan-search" class="primary" type="button">조회</button>
        <div id="inbound-plan-status" class="muted" style="align-self:center;">대기 중</div>
      </div>
    </div>
    <div class="card">
      <div id="inbound-plan-table-wrap" class="muted">조회 버튼을 눌러주세요.</div>
    </div>
    <div id="inbound-plan-detail-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(760px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3 id="inbound-plan-detail-title">전표 상세</h3>
          <button type="button" id="inbound-plan-detail-close" class="cancel-btn del-small">닫기</button>
        </div>
        <div id="inbound-plan-detail-body" style="max-height: 65vh; overflow: auto;"></div>
      </div>
    </div>
  `;

  let inboundPlanQuerySeq = 0;
  let inboundPlanItems = [];
  const detailOverlay = qs("#inbound-plan-detail-overlay");
  const detailBody = qs("#inbound-plan-detail-body");
  const detailTitle = qs("#inbound-plan-detail-title");
  const closeDetail = () => detailOverlay?.classList.add("hidden");
  qs("#inbound-plan-detail-close")?.addEventListener("click", closeDetail);
  detailOverlay?.addEventListener("click", (e) => {
    if (e.target === detailOverlay) closeDetail();
  });

  const renderDetail = async (item) => {
    if (!item || !detailBody) return;
    const slipNo = String(item.poNo || "");
    if (!slipNo) return;
    const fromDate = String(qs("#inbound-plan-from")?.value || "");
    const toDate = String(qs("#inbound-plan-to")?.value || "");
    if (detailTitle) detailTitle.textContent = `전표 상세 - ${slipNo || "-"}`;
    detailBody.innerHTML = `<p class="muted">전표 상세를 조회 중입니다...</p>`;
    detailOverlay?.classList.remove("hidden");

    let lines = [];
    let head = item;
    let detailRes = null;
    let fetchErrorHtml = "";
    const fallbackLines = inboundPlanItems.filter((x) => String(x.poNo || "") === slipNo);
    const detailTimeoutMs = 120000;
    try {
      const q = `slipNo=${encodeURIComponent(slipNo)}&from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}`;
      const ac = new AbortController();
      const t = setTimeout(() => ac.abort(), detailTimeoutMs);
      try {
        detailRes = await api(`/api/inbound-plans/detail?${q}`, { signal: ac.signal });
      } finally {
        clearTimeout(t);
      }
      lines = Array.isArray(detailRes.items) ? detailRes.items : [];
      head = detailRes.item || lines[0] || item;
    } catch (err) {
      // ECOUNT can rate-limit repeated requests(412). Fall back to current list rows.
      lines = fallbackLines;
      head = lines[0] || item;
      const aborted = err && (err.name === "AbortError" || /aborted/i.test(String(err.message || "")));
      const errMsg = esc(
        aborted
          ? `응답 시간 초과(${Math.round(detailTimeoutMs / 1000)}초). 서버·이카운트 부하를 줄이고 다시 시도하세요.`
          : err.message || "오류"
      );
      fetchErrorHtml = `<p class="muted" style="margin-bottom:8px;">상세 API 재조회 실패로 현재 목록 데이터로 표시합니다. (${errMsg})</p>`;
    }
    if (detailTitle) detailTitle.textContent = `전표 상세 - ${slipNo || "-"}`;

    const qtyText = (v) => {
      const n = Number(String(v ?? "").replace(/,/g, ""));
      return Number.isFinite(n) ? n.toLocaleString("ko-KR", { maximumFractionDigits: 0 }) : String(v || "");
    };
    const lineRows = lines
      .map(
        (r, idx) => `<tr>
          <td>${idx + 1}</td>
          <td>${esc(r.itemCode || "")}</td>
          <td>${esc(r.barcode || "")}</td>
          <td>${esc(r.itemName || "")}</td>
          <td>${esc(r.spec || "")}</td>
          <td>${esc(qtyText(r.boxQty || ""))}</td>
          <td>${esc(qtyText(r.qty || ""))}</td>
          <td>${esc(r.remark || "")}</td>
          <td>${esc(r.note || "")}</td>
        </tr>`
      )
      .join("");

    const dbg = detailRes?.debug;
    const summaryLine =
      lines.length === 1 &&
      /외\s*\d+건/.test(String(lines[0]?.itemName || "")) &&
      !String(lines[0]?.itemCode || "").trim();
    const summaryHint = summaryLine
      ? `<p class="muted" style="margin-bottom:10px;font-size:13px;">품목명이 「…외 N건」 형태면 이카운트 발주 목록만 반영된 요약입니다. 품목별 행·코드는 구매/입고 라인 API 연동 또는 <code style="font-size:12px;">ECOUNT_INBOUND_LINE_API_PATHS</code> 설정 후 재조회가 필요합니다.</p>`
      : "";

    let debugHtml = "";
    if (dbg && typeof dbg === "object") {
      const ep = Array.isArray(dbg.endpointErrors) ? dbg.endpointErrors : [];
      const epHtml = ep.length
        ? `<ul style="margin:6px 0 0;padding-left:18px;font-size:12px;">${ep.map((e) => `<li>${esc(String(e))}</li>`).join("")}</ul>`
        : "";
      debugHtml = `
        <details style="margin-top:14px;" class="muted">
          <summary style="cursor:pointer;font-size:13px;">이카운트 상세 조회 디버그 (usedEndpoint / endpointErrors)</summary>
          <div style="margin-top:8px;font-size:12px;line-height:1.45;">
            <div><strong>usedEndpoint</strong>: ${esc(String(dbg.usedEndpoint ?? "(없음)"))}</div>
            <div><strong>lineSource</strong>: ${esc(String(dbg.lineSource ?? ""))}</div>
            <div><strong>nestedLineCount</strong>: ${esc(String(dbg.nestedLineCount ?? ""))}</div>
            ${
              dbg.endpointErrorAttempts != null
                ? `<div><strong>endpointErrorAttempts</strong>: ${esc(String(dbg.endpointErrorAttempts))} (압축 전 시도 횟수)</div>`
                : ""
            }
            ${
              dbg.skipBuiltinLineApis != null
                ? `<div><strong>skipBuiltinLineApis</strong>: ${esc(String(dbg.skipBuiltinLineApis))} — ${
                    dbg.skipBuiltinLineApis
                      ? "내장 후보 API 자동 호출 안 함(기본). 켜려면 .env에 ECOUNT_SKIP_BUILTIN_LINE_APIS=0 후 서버 재시작."
                      : "내장 후보 API 자동 호출 중(.env ECOUNT_SKIP_BUILTIN_LINE_APIS=0). 끄려면 해당 줄을 제거 후 재시작."
                  }</div>`
                : ""
            }
            ${dbg.hint ? `<div style="margin-top:6px;"><strong>hint</strong>: ${esc(String(dbg.hint))}</div>` : ""}
            ${ep.length ? `<div style="margin-top:6px;"><strong>endpointErrors</strong>:</div>${epHtml}` : `<div style="margin-top:6px;">endpointErrors: 없음</div>`}
          </div>
        </details>
      `;
    }

    detailBody.innerHTML = `${fetchErrorHtml}${summaryHint}
      <div class="row" style="grid-template-columns: repeat(4, minmax(140px, 1fr)); gap: 8px; margin-bottom: 12px;">
        <div><strong>전표번호</strong><div>${esc(slipNo || "-")}</div></div>
        <div><strong>일자</strong><div>${esc(formatYmd(head.poDate || ""))}</div></div>
        <div><strong>거래처</strong><div>${esc(head.vendor || "")}</div></div>
        <div><strong>담당자</strong><div>${esc(head.manager || "")}</div></div>
        <div><strong>입고창고</strong><div>${esc(head.whName || "")}</div></div>
        <div><strong>통화</strong><div>${esc(head.currency || "내자")}</div></div>
        <div><strong>납기일자</strong><div>${esc(formatYmd(head.dueDate || ""))}</div></div>
        <div><strong>상태</strong><div>${esc(head.status || "")}</div></div>
      </div>
      <table>
        <thead>
          <tr>
            <th>순번</th>
            <th>품목코드</th>
            <th>바코드</th>
            <th>품목명</th>
            <th>규격</th>
            <th>수량(BOX)</th>
            <th>수량</th>
            <th>적요</th>
            <th>비고</th>
          </tr>
        </thead>
        <tbody>${lineRows || `<tr><td colspan="9" class="muted">품목 라인 정보가 없습니다.</td></tr>`}</tbody>
      </table>
    ${debugHtml}`;
    detailOverlay?.classList.remove("hidden");
  };

  const draw = async () => {
    const currentSeq = ++inboundPlanQuerySeq;
    const fromEl = qs("#inbound-plan-from");
    const toEl = qs("#inbound-plan-to");
    const statusEl = qs("#inbound-plan-status");
    const tableWrap = qs("#inbound-plan-table-wrap");
    const searchBtn = qs("#inbound-plan-search");
    const fromDate = fromEl?.value || "";
    const toDate = toEl?.value || "";
    try {
      if (searchBtn) searchBtn.disabled = true;
      if (statusEl) statusEl.textContent = "조회 중...";
      const qsPart = `from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}`;
      const res = await api(`/api/inbound-plans?${qsPart}`);
      // Ignore stale responses from earlier requests.
      if (currentSeq !== inboundPlanQuerySeq) return;
      const items = Array.isArray(res.items) ? res.items : [];
      inboundPlanItems = items;
      if (statusEl) {
        const extra = res?.debug?.countRaw ? ` (원본 ${res.debug.countRaw})` : "";
        statusEl.textContent = `총 ${items.length}건${extra}`;
      }
      if (!items.length) {
        const code = res?.debug?.code ? ` / code: ${esc(res.debug.code)}` : "";
        const msg = res?.debug?.message ? ` / msg: ${esc(res.debug.message)}` : "";
        const keys = res?.debug?.topKeys?.length ? ` / keys: ${esc(res.debug.topKeys.join(", "))}` : "";
        const dkeys = res?.debug?.dataKeys?.length ? ` / dataKeys: ${esc(res.debug.dataKeys.join(", "))}` : "";
        const dtype = res?.debug?.dataType ? ` / dataType: ${esc(res.debug.dataType)}` : "";
        const adj = res?.debug?.adjustedDateRange
          ? ` / adjustedRange: ${esc(res.debug.adjustedDateRange.from)}~${esc(res.debug.adjustedDateRange.to)}`
          : "";
        const preview = res?.debug?.dataPreview ? `<br /><small class="muted">dataPreview: ${esc(res.debug.dataPreview)}</small>` : "";
        const errPreview = res?.debug?.errorPreview ? `<br /><small class="muted">error: ${esc(res.debug.errorPreview)}</small>` : "";
        const errsPreview = res?.debug?.errorsPreview ? `<br /><small class="muted">errors: ${esc(res.debug.errorsPreview)}</small>` : "";
        tableWrap.innerHTML = `<span class="muted">조회 결과가 없습니다.${code}${msg}${keys}${dkeys}${dtype}${adj}</span>${preview}${errPreview}${errsPreview}`;
        return;
      }
      const rows = items
        .map(
          (x, idx) => {
            const qtyNum = Number(String(x.qty ?? "").replace(/,/g, ""));
            const qtyText = Number.isFinite(qtyNum) ? qtyNum.toLocaleString("ko-KR", { maximumFractionDigits: 0 }) : String(x.qty || "");
            return `<tr>
            <td>${idx + 1}</td>
            <td><button type="button" class="bh-link-btn inbound-plan-open-detail" data-idx="${idx}">${esc(x.poNo || "")}</button></td>
            <td>${esc(formatYmd(x.poDate || ""))}</td>
            <td>${esc(x.vendor || "")}</td>
            <td>${esc(x.manager || "")}</td>
            <td>${esc(x.itemName || "")}</td>
            <td>${esc(qtyText)}</td>
            <td>${esc(formatYmd(x.dueDate || ""))}</td>
            <td>${esc(x.whName || "")}</td>
            <td>${esc(x.status || "")}</td>
          </tr>`;
          }
        )
        .join("");
      tableWrap.innerHTML = `
        <table id="inbound-plan-table">
          <thead>
            <tr>
              <th>순번</th>
              <th>발주번호</th>
              <th>발주일자</th>
              <th>거래처</th>
              <th>담당자</th>
              <th>품목명</th>
              <th>발주수량</th>
              <th>납기일(입고예정일)</th>
              <th>창고</th>
              <th>상태</th>
            </tr>
          </thead>
          <tbody>${rows}</tbody>
        </table>
      `;
      delete TABLE_FILTER_MEMORY["#inbound-plan-table"];
      applyExcelLikeFilter("#inbound-plan-table");
      tableWrap.querySelectorAll(".inbound-plan-open-detail").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const idx = Number(btn.getAttribute("data-idx"));
          if (!Number.isInteger(idx) || idx < 0) return;
          await renderDetail(inboundPlanItems[idx]);
        });
      });
    } catch (err) {
      if (currentSeq !== inboundPlanQuerySeq) return;
      if (statusEl) statusEl.textContent = "조회 실패";
      tableWrap.innerHTML = `<span class="muted">${esc(err.message || "조회 실패")}</span>`;
    } finally {
      if (currentSeq === inboundPlanQuerySeq && searchBtn) searchBtn.disabled = false;
    }
  };

  qs("#inbound-plan-search").onclick = () => draw();
}

async function renderInboundPlan2() {
  const wrap = qs("#view-inbound-plan-2");
  if (!wrap) return;
  const today = new Date();
  const defaultApiTo = localDateYmd(today);
  const defaultApiFrom = localDateYmd(new Date(today.getFullYear(), today.getMonth(), today.getDate() - 30));
  let apiFrom = defaultApiFrom;
  let apiTo = defaultApiTo;
  try {
    const sF = localStorage.getItem("inbound-plan-2-api-from");
    const sT = localStorage.getItem("inbound-plan-2-api-to");
    if (sF && /^\d{4}-\d{2}-\d{2}$/.test(sF)) apiFrom = sF;
    if (sT && /^\d{4}-\d{2}-\d{2}$/.test(sT)) apiTo = sT;
  } catch (_) {
    /* ignore */
  }
  wrap.innerHTML = `
    <div class="card">
      <h2>입고예정목록 2 <span id="inbound-plan-2-new-badge"></span></h2>
      <p class="muted">이카운트 ERP에서 「발주서현황」을 엑셀 저장한 파일(.xlsx)을 올리면 품목 라인까지 목록·전표 상세로 볼 수 있습니다. 이미 저장된 <strong>발주번호(일자-No.)</strong>와 같은 전표는 덮어쓰지 않고 건너뛰며, <strong>신규 발주만</strong> 추가됩니다.</p>
      <div class="row" style="grid-template-columns: 1fr auto auto auto auto; gap: 10px; align-items: center;">
        <input type="file" id="inbound-plan-2-file" accept=".xlsx,.xls" />
        <button type="button" id="inbound-plan-2-upload" class="primary">업로드·반영</button>
        <button type="button" id="inbound-plan-2-refresh" class="cancel-btn">목록 새로고침</button>
        <button type="button" id="inbound-plan-2-check-new" class="cancel-btn">API 신규 확인</button>
        <button type="button" id="inbound-plan-2-delete-selected" class="cancel-btn">선택 삭제</button>
      </div>
      <div class="row" style="display: grid; grid-template-columns: auto 170px auto 170px 1fr; align-items: center; gap: 8px; margin-top: 10px;">
        <span class="muted" style="white-space: nowrap;">신규 확인(API) 조회 기간</span>
        <input type="date" id="inbound-plan-2-api-from" value="${esc(apiFrom)}" style="width:170px;" />
        <span class="muted">~</span>
        <input type="date" id="inbound-plan-2-api-to" value="${esc(apiTo)}" style="width:170px;" />
        <span class="muted" style="font-size: 12px; white-space: nowrap;">이카운트 발주서 조회 구간(BASE_DATE) · <strong>입고예정목록 1</strong>과 별도입니다.</span>
      </div>
      <div id="inbound-plan-2-meta" class="muted" style="margin-top:8px;"></div>
      <div id="inbound-plan-2-api-status" class="muted" style="margin-top:4px;"></div>
    </div>
    <div class="card">
      <div id="inbound-plan-2-table-wrap" class="muted">파일을 업로드하거나 새로고침 하세요.</div>
    </div>
    <div id="inbound-plan-2-detail-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(920px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3 id="inbound-plan-2-detail-title">전표 상세 (업로드)</h3>
          <button type="button" id="inbound-plan-2-detail-close" class="cancel-btn del-small">닫기</button>
        </div>
        <div id="inbound-plan-2-detail-body" style="max-height: 65vh; overflow: auto;"></div>
      </div>
    </div>
  `;

  let inboundPlan2Items = [];
  let selectedSlipKeys = new Set();
  const detailOverlay = qs("#inbound-plan-2-detail-overlay");
  const detailBody = qs("#inbound-plan-2-detail-body");
  const detailTitle = qs("#inbound-plan-2-detail-title");
  const metaEl = qs("#inbound-plan-2-meta");
  const tableWrap = qs("#inbound-plan-2-table-wrap");
  const apiStatusEl = qs("#inbound-plan-2-api-status");
  const newBadgeEl = qs("#inbound-plan-2-new-badge");

  const closeDetail = () => detailOverlay?.classList.add("hidden");
  qs("#inbound-plan-2-detail-close")?.addEventListener("click", closeDetail);
  detailOverlay?.addEventListener("click", (e) => {
    if (e.target === detailOverlay) closeDetail();
  });

  const qtyText = (v) => {
    const n = Number(String(v ?? "").replace(/,/g, ""));
    return Number.isFinite(n) ? n.toLocaleString("ko-KR", { maximumFractionDigits: 0 }) : String(v || "");
  };

  const collectDetailLineQtyMap = () => {
    const map = {};
    detailBody?.querySelectorAll(".inbound-plan-2-detail-qty").forEach((el) => {
      const idx = el.getAttribute("data-line-idx");
      if (idx == null) return;
      map[idx] = el.value;
    });
    return map;
  };

  const renderDetail = async (listRow) => {
    if (!listRow || !detailBody) return;
    const slipNo = String(listRow.poNo || "");
    if (!slipNo) return;
    if (detailTitle) detailTitle.textContent = `전표 상세 - ${slipNo}`;
    detailBody.innerHTML = `<p class="muted">불러오는 중…</p>`;
    detailOverlay?.classList.remove("hidden");
    let lines = [];
    let head = listRow;
    let workflow = { status: "draft", lineQtyByIndex: {} };
    try {
      const res = await api(`/api/inbound-plan-upload/detail?slipNo=${encodeURIComponent(slipNo)}`);
      lines = Array.isArray(res.items) ? res.items : [];
      head = res.item || lines[0] || listRow;
      workflow = res.workflow && typeof res.workflow === "object" ? res.workflow : workflow;
    } catch (err) {
      detailBody.innerHTML = `<p class="muted">${esc(err.message || "오류")}</p>`;
      return;
    }
    const confirmed = workflow.status === "confirmed";
    const purchaseComplete = Boolean(workflow.ecountPurchaseSavedAt && workflow.wmsStockInboundAt);
    const statusLabel = confirmed ? "입고 확정" : "초안(수정 가능)";
    const qtyDisabled = confirmed ? "disabled" : "";
    const orderQtyVal = (r) => {
      const o = r.orderQty != null ? r.orderQty : r.qty;
      return qtyText(o ?? "");
    };
    const inboundCheckVal = (r) => {
      const v = r.inboundCheckQty != null ? r.inboundCheckQty : r.orderQty != null ? r.orderQty : r.qty;
      return String(v ?? "").replace(/,/g, "");
    };
    const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
    let stockUserDefault = "";
    try {
      stockUserDefault = localStorage.getItem("inbound-plan-2-stock-user") || "";
    } catch (_) {
      /* ignore */
    }
    const lineRows = lines
      .map(
        (r, idx) => `<tr>
          <td>${idx + 1}</td>
          <td>${esc(r.itemCode || "")}</td>
          <td>${esc(r.barcode || "")}</td>
          <td>${esc(r.itemName || "")}</td>
          <td>${esc(r.spec || "")}</td>
          <td>${esc(qtyText(r.boxQty || ""))}</td>
          <td>${esc(orderQtyVal(r))}</td>
          <td><input type="number" min="0" step="1" class="inbound-plan-2-detail-qty" data-line-idx="${idx}" value="${esc(
            inboundCheckVal(r)
          )}" style="width:88px;" ${qtyDisabled} /></td>
          <td>${esc(qtyText(r.unitPrice ?? ""))}</td>
          <td>${esc(qtyText(r.supplyAmount ?? ""))}</td>
          <td>${esc(r.vendorCode || "")}</td>
          <td>${esc(r.remark || "")}</td>
          <td>${esc(r.note || "")}</td>
        </tr>`
      )
      .join("");
    detailBody.innerHTML = `
      <p class="muted" style="margin-bottom:10px;"><strong>수량(발주)</strong>는 엑셀 그대로이며 수정할 수 없습니다. 실제 입고 예정 수량은 <strong>수량(입고확인)</strong>에 입력한 뒤 <strong>수량(입고확인) 중간 저장</strong> → <strong>입고 확정</strong> → <strong>구매입력 전송</strong>(이카운트 + WMS 입고) 순서로 진행하세요.</p>
      <div class="row" style="display:flex; flex-wrap:wrap; align-items:center; gap:10px; margin-bottom:12px;">
        <span><strong>입고 상태</strong>: <span id="inbound-plan-2-wf-status-label">${esc(statusLabel)}</span></span>
        <span id="inbound-plan-2-purchase-done-badge" class="muted" style="font-size:12px;"></span>
        <label class="muted" style="font-size:13px; display:flex; align-items:center; gap:6px;">WMS 입고 담당자
          <select id="inbound-plan-2-stock-user" style="min-width:120px;" ${purchaseComplete ? "disabled" : ""}>
            <option value="">선택…</option>
            ${managerOptions}
          </select>
        </label>
        <button type="button" class="cancel-btn" id="inbound-plan-2-save-qty" ${confirmed ? "disabled" : ""}>수량(입고확인) 중간 저장</button>
        <button type="button" class="primary" id="inbound-plan-2-confirm-slip" ${confirmed ? "disabled" : ""}>입고 확정</button>
        <button type="button" class="cancel-btn" id="inbound-plan-2-reopen-slip" ${confirmed ? "" : "disabled"}>확정 취소</button>
        <button type="button" class="primary" id="inbound-plan-2-send-purchase" ${confirmed && !purchaseComplete ? "" : "disabled"} style="${confirmed && !purchaseComplete ? "" : "opacity:0.5;"}">구매입력 전송</button>
        <span id="inbound-plan-2-purchase-env-hint" class="muted" style="font-size:12px;"></span>
      </div>
      <div id="inbound-plan-2-detail-msg" class="muted" style="margin-bottom:8px; min-height:1.2em;"></div>
      <div class="row" style="grid-template-columns: repeat(4, minmax(140px, 1fr)); gap: 8px; margin-bottom: 12px;">
        <div><strong>전표번호</strong><div>${esc(slipNo || "-")}</div></div>
        <div><strong>일자</strong><div>${esc(formatYmdLoose(head.poDate || ""))}</div></div>
        <div><strong>거래처</strong><div>${esc(head.vendor || "")}</div></div>
        <div><strong>거래처코드</strong><div>${esc(head.vendorCode || "")}</div></div>
        <div><strong>담당자</strong><div>${esc(head.manager || "")}</div></div>
        <div><strong>입고창고</strong><div>${esc(head.whName || "")}</div></div>
        <div><strong>납기일자</strong><div>${esc(formatYmdLoose(head.dueDate || ""))}</div></div>
        <div><strong>최종수정일자</strong><div>${esc(head.lastModifiedAt || "")}</div></div>
        <div><strong>품목 행 수</strong><div>${lines.length}</div></div>
      </div>
      <table>
        <thead>
          <tr>
            <th>순번</th>
            <th>품목코드</th>
            <th>바코드</th>
            <th>품목명</th>
            <th>규격</th>
            <th>수량(BOX)</th>
            <th>수량(발주)</th>
            <th>수량(입고확인)</th>
            <th>단가</th>
            <th>공급가액</th>
            <th>거래처코드</th>
            <th>적요</th>
            <th>비고</th>
          </tr>
        </thead>
        <tbody>${lineRows || `<tr><td colspan="13" class="muted">품목 라인이 없습니다.</td></tr>`}</tbody>
      </table>
    `;

    const msgEl = qs("#inbound-plan-2-detail-msg");
    const setMsg = (t, isErr) => {
      if (!msgEl) return;
      msgEl.textContent = t || "";
      msgEl.style.color = isErr ? "#b42318" : "";
    };

    const envHintEl = qs("#inbound-plan-2-purchase-env-hint");
    const sendPurchaseBtn = qs("#inbound-plan-2-send-purchase");
    const purchaseDoneBadge = qs("#inbound-plan-2-purchase-done-badge");
    const stockUserSel = qs("#inbound-plan-2-stock-user");
    if (stockUserSel && stockUserDefault && state.managers.includes(stockUserDefault)) {
      stockUserSel.value = stockUserDefault;
    }
    const applyPurchaseEnvHint = async () => {
      if (!envHintEl) return;
      envHintEl.style.color = "";
      if (purchaseDoneBadge) {
        purchaseDoneBadge.textContent = purchaseComplete ? "이카운트·WMS 입고 완료" : "";
        purchaseDoneBadge.style.color = purchaseComplete ? "#137333" : "";
      }
      if (!confirmed) {
        envHintEl.textContent = "입고 확정 시 실제 WMS 입고가 반영되며, 구매입력 전송은 확정 후에만 사용할 수 있습니다.";
        return;
      }
      if (purchaseComplete) {
        envHintEl.textContent = "이 전표는 처리가 끝났습니다. 재전송하지 않습니다.";
        return;
      }
      envHintEl.textContent = "서버 설정 확인 중…";
      try {
        const st = await api("/api/ecount/save-purchases-env");
        const ok =
          st.testSavePurchases && st.defaultWhCdConfigured && st.comCodeConfigured;
        if (!st.testSavePurchases) {
          envHintEl.textContent =
            "구매입력 전송은 서버에서 꺼져 있습니다. .env에 ECOUNT_TEST_SAVE_PURCHASES=1 을 넣고 서버(Node)를 재시작하세요.";
          envHintEl.style.color = "#b42318";
        } else if (!st.defaultWhCdConfigured) {
          envHintEl.textContent =
            "창고코드가 없습니다. .env에 ECOUNT_DEFAULT_WH_CD=이카운트창고코드 를 넣고 재시작하세요.";
          envHintEl.style.color = "#b42318";
        } else if (!st.comCodeConfigured) {
          envHintEl.textContent = "ECOUNT_COM_CODE 가 비어 있습니다. .env를 확인한 뒤 재시작하세요.";
          envHintEl.style.color = "#b42318";
        } else {
          envHintEl.textContent =
            "입고 확정 시 실제 WMS 입고가 반영됩니다. 이 버튼은 이카운트 SavePurchases(구매입력) 전송용입니다.";
        }
        if (sendPurchaseBtn) {
          sendPurchaseBtn.disabled = !ok;
          sendPurchaseBtn.style.opacity = ok ? "" : "0.5";
        }
      } catch (_) {
        envHintEl.textContent = "서버 설정(/api/ecount/save-purchases-env)을 확인하지 못했습니다. 서버를 최신 코드로 실행 중인지 확인하세요.";
        envHintEl.style.color = "#b42318";
        if (sendPurchaseBtn) {
          sendPurchaseBtn.disabled = true;
          sendPurchaseBtn.style.opacity = "0.5";
        }
      }
    };
    void applyPurchaseEnvHint();

    const refreshAfterWorkflow = async () => {
      await loadList();
      const row = inboundPlan2Items.find((x) => normalizeSlipNoText(x.poNo) === normalizeSlipNoText(slipNo));
      await renderDetail(row || { poNo: slipNo });
      await renderRecentMovements("IN");
    };

    qs("#inbound-plan-2-save-qty")?.addEventListener("click", async () => {
      try {
        setMsg("");
        await api("/api/inbound-plan-upload/slip-workflow", {
          method: "POST",
          body: JSON.stringify({ slipNo, action: "save-qty", lineQtyByIndex: collectDetailLineQtyMap() })
        });
        setMsg("입고확인 수량을 저장했습니다.");
        await refreshAfterWorkflow();
      } catch (e) {
        setMsg(e.message || "오류", true);
      }
    });
    qs("#inbound-plan-2-confirm-slip")?.addEventListener("click", async () => {
      try {
        const su = String(stockUserSel?.value || "").trim();
        if (!su) {
          setMsg("입고 확정 전에 WMS 입고 담당자를 선택하세요.", true);
          return;
        }
        setMsg("");
        try {
          localStorage.setItem("inbound-plan-2-stock-user", su);
        } catch (_) {
          /* ignore */
        }
        await api("/api/inbound-plan-upload/slip-workflow", {
          method: "POST",
          body: JSON.stringify({ slipNo, action: "confirm", lineQtyByIndex: collectDetailLineQtyMap(), stockUser: su })
        });
        setMsg("입고 확정과 동시에 실제 WMS 입고가 반영되었습니다. 이제 구매입력 전송을 진행하세요.");
        await refreshAfterWorkflow();
      } catch (e) {
        setMsg(e.message || "오류", true);
      }
    });
    qs("#inbound-plan-2-reopen-slip")?.addEventListener("click", async () => {
      try {
        setMsg("");
        const ok = confirm("확정 취소 시 입고 확정으로 반영된 재고가 자동 원복됩니다. 진행할까요?");
        if (!ok) return;
        let allowReopenAfterEcountCancel = false;
        if (workflow.ecountPurchaseSavedAt) {
          const ok2 = confirm(
            "이 전표는 구매입력 전송 이력이 있습니다.\n이카운트에서 구매입력 전표를 이미 삭제한 것이 맞다면 '확인'을 누르세요."
          );
          if (!ok2) return;
          allowReopenAfterEcountCancel = true;
        }
        await api("/api/inbound-plan-upload/slip-workflow", {
          method: "POST",
          body: JSON.stringify({ slipNo, action: "reopen", allowReopenAfterEcountCancel })
        });
        setMsg("확정을 취소했습니다. 재고 원복(재고조정 -) 이력이 입출고 이력에 기록되었습니다.");
        alert("확정 취소가 완료되었습니다. 재고 원복 내역이 입출고 이력에 반영되었습니다.");
        await refreshAfterWorkflow();
      } catch (e) {
        setMsg(e.message || "오류", true);
      }
    });
    qs("#inbound-plan-2-send-purchase")?.addEventListener("click", async () => {
      if (!confirmed || purchaseComplete) return;
      const ok = confirm(
        "이카운트 구매입력(SavePurchases) 전송을 진행할까요?\n실제 구매 전표가 생성됩니다."
      );
      if (!ok) return;
      try {
        setMsg("");
        const res = await api("/api/ecount/save-purchases-from-upload-slip", {
          method: "POST",
          body: JSON.stringify({ slipNo })
        });
        if (res.alreadyComplete) {
          setMsg(res.message || "이미 처리된 전표입니다.");
        } else {
          const ec = res.ecount || {};
          setMsg(`이카운트 구매입력 전송 완료. SuccessCnt=${ec.SuccessCnt ?? "?"} FailCnt=${ec.FailCnt ?? "?"}`);
        }
        await refreshAfterWorkflow();
      } catch (e) {
        setMsg(e.message || "오류", true);
      }
    });
  };

  const drawTable = () => {
    if (!tableWrap) return;
    if (!inboundPlan2Items.length) {
      tableWrap.innerHTML = `<span class="muted">표시할 전표가 없습니다. 엑셀을 업로드하세요.</span>`;
      return;
    }
    const rows = inboundPlan2Items
      .map(
        (x, idx) => {
          const qtyNum = Number(String(x.qty ?? "").replace(/,/g, ""));
          const qtyDisp = Number.isFinite(qtyNum) ? qtyNum.toLocaleString("ko-KR", { maximumFractionDigits: 0 }) : String(x.qty || "");
          const wfSt = x.workflowStatus === "confirmed" ? "확정" : "초안";
          return `<tr>
            <td><input type="checkbox" class="inbound-plan-2-row-check" data-slip="${esc(x.poNo || "")}" ${
              selectedSlipKeys.has(normalizeSlipNoText(x.poNo)) ? "checked" : ""
            } /></td>
            <td>${idx + 1}</td>
            <td><button type="button" class="bh-link-btn inbound-plan-2-open-detail" data-idx="${idx}">${esc(x.poNo || "")}</button></td>
            <td>${esc(formatYmdLoose(x.poDate || ""))}</td>
            <td>${esc(x.vendor || "")}</td>
            <td>${esc(x.manager || "")}</td>
            <td>${esc(x.itemName || "")}</td>
            <td>${esc(qtyDisp)}</td>
            <td>${esc(formatYmdLoose(x.dueDate || ""))}</td>
            <td>${esc(x.whName || "")}</td>
            <td>${esc(x.lastModifiedAt || "")}</td>
            <td>${esc(wfSt)}</td>
            <td>${esc(String(x.lineCount || ""))}</td>
          </tr>`;
        }
      )
      .join("");
    tableWrap.innerHTML = `
      <table id="inbound-plan-2-table">
        <thead>
          <tr>
            <th><input type="checkbox" id="inbound-plan-2-check-all" /></th>
            <th>순번</th>
            <th>발주번호</th>
            <th>발주일자</th>
            <th>거래처</th>
            <th>담당자</th>
            <th>품목요약</th>
            <th>수량합계</th>
            <th>납기일</th>
            <th>창고</th>
            <th>최종수정일자</th>
            <th>입고</th>
            <th>품목행수</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
    const checkAll = qs("#inbound-plan-2-check-all");
    const rowChecks = Array.from(tableWrap.querySelectorAll(".inbound-plan-2-row-check"));
    if (checkAll) {
      const allChecked = rowChecks.length > 0 && rowChecks.every((el) => el.checked);
      checkAll.checked = allChecked;
      checkAll.addEventListener("change", () => {
        selectedSlipKeys = new Set();
        rowChecks.forEach((el) => {
          el.checked = checkAll.checked;
          if (checkAll.checked) selectedSlipKeys.add(normalizeSlipNoText(el.getAttribute("data-slip")));
        });
      });
    }
    rowChecks.forEach((el) => {
      el.addEventListener("change", () => {
        const k = normalizeSlipNoText(el.getAttribute("data-slip"));
        if (!k) return;
        if (el.checked) selectedSlipKeys.add(k);
        else selectedSlipKeys.delete(k);
        if (checkAll) checkAll.checked = rowChecks.length > 0 && rowChecks.every((x) => x.checked);
      });
    });
    delete TABLE_FILTER_MEMORY["#inbound-plan-2-table"];
    applyExcelLikeFilter("#inbound-plan-2-table");
    tableWrap.querySelectorAll(".inbound-plan-2-open-detail").forEach((btn) => {
      btn.addEventListener("click", async () => {
        const idx = Number(btn.getAttribute("data-idx"));
        if (!Number.isInteger(idx) || idx < 0) return;
        await renderDetail(inboundPlan2Items[idx]);
      });
    });
  };

  const loadList = async () => {
    try {
      const res = await api("/api/inbound-plan-upload");
      inboundPlan2Items = Array.isArray(res.items) ? res.items : [];
      selectedSlipKeys = new Set(
        Array.from(selectedSlipKeys).filter((k) => inboundPlan2Items.some((x) => normalizeSlipNoText(x.poNo) === k))
      );
      if (metaEl) {
        const fn = res.sourceFileName ? esc(res.sourceFileName) : "-";
        const at = res.uploadedAt ? esc(res.uploadedAt.slice(0, 19).replace("T", " ")) : "-";
        metaEl.innerHTML = `저장된 파일: <strong>${fn}</strong> · 반영 시각: ${at} · 품목 행 <strong>${Number(res.lineCount || 0)}</strong> · 전표 <strong>${inboundPlan2Items.length}</strong>`;
      }
      drawTable();
    } catch (e) {
      if (metaEl) metaEl.textContent = "";
      tableWrap.innerHTML = `<span class="muted">${esc(e.message || "목록 조회 실패")}</span>`;
    }
  };

  const checkApiNewSlips = async () => {
    const fromEl = qs("#inbound-plan-2-api-from");
    const toEl = qs("#inbound-plan-2-api-to");
    let from = String(fromEl?.value || "").trim();
    let to = String(toEl?.value || "").trim();
    if (!from || !to) {
      alert("신규 확인(API) 조회 기간의 시작일·종료일을 모두 선택하세요.");
      return;
    }
    if (from > to) {
      const tmp = from;
      from = to;
      to = tmp;
      if (fromEl) fromEl.value = from;
      if (toEl) toEl.value = to;
    }
    try {
      localStorage.setItem("inbound-plan-2-api-from", from);
      localStorage.setItem("inbound-plan-2-api-to", to);
    } catch (_) {
      /* ignore */
    }

    if (apiStatusEl) apiStatusEl.textContent = `이카운트 발주 API 조회 중… (${from} ~ ${to})`;
    if (newBadgeEl) newBadgeEl.innerHTML = "";
    try {
      const res = await api(`/api/inbound-plans?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`);
      const apiItems = Array.isArray(res.items) ? res.items : [];
      const adj = res.debug && res.debug.adjustedDateRange;
      const adjNote =
        adj && adj.from && adj.to
          ? ` · 서버 적용 구간 ${formatYmd(adj.from)} ~ ${formatYmd(adj.to)}(최대 30일로 자동 축소됨)`
          : "";
      const uploadedMap = new Map();
      inboundPlan2Items.forEach((x) => {
        const k = normalizeSlipNoText(x.poNo);
        if (!k) return;
        uploadedMap.set(k, {
          poNo: x.poNo || "",
          lastModifiedKey: normalizeDateTimeDigits(x.lastModifiedAt),
          lastModifiedAt: String(x.lastModifiedAt || "")
        });
      });
      const newItems = [];
      const modifiedItems = [];
      apiItems.forEach((x) => {
        const k = normalizeSlipNoText(x.poNo);
        if (!k) return;
        const apiLastModifiedKey = normalizeDateTimeDigits(x.lastModifiedAt);
        const uploaded = uploadedMap.get(k);
        if (!uploaded) {
          newItems.push(x);
          return;
        }
        if (uploaded.lastModifiedKey && apiLastModifiedKey && apiLastModifiedKey > uploaded.lastModifiedKey) {
          modifiedItems.push({
            poNo: x.poNo || uploaded.poNo || "",
            apiLastModifiedAt: x.lastModifiedAt || "",
            uploadedLastModifiedAt: uploaded.lastModifiedAt || ""
          });
        }
      });
      if (newItems.length || modifiedItems.length) {
        const newSample = newItems
          .slice(0, 5)
          .map((x) => x.poNo)
          .filter(Boolean)
          .join(", ");
        const modifiedSample = modifiedItems
          .slice(0, 5)
          .map((x) => x.poNo)
          .filter(Boolean)
          .join(", ");
        const parts = [];
        if (newItems.length) {
          parts.push(
            `<strong style="color:#b42318;">신규 발주서 ${newItems.length}건</strong> (입고예정목록 2 업로드 목록에 없음)${
              newSample ? ` · 예: ${esc(newSample)}${newItems.length > 5 ? " …" : ""}` : ""
            }`
          );
        }
        if (modifiedItems.length) {
          parts.push(
            `<strong style="color:#1d4ed8;">수정 발주서 ${modifiedItems.length}건</strong> (동일 발주번호, API 최종수정일자 최신)${
              modifiedSample ? ` · 예: ${esc(modifiedSample)}${modifiedItems.length > 5 ? " …" : ""}` : ""
            }`
          );
        }
        if (apiStatusEl) {
          apiStatusEl.innerHTML = `${parts.join(" · ")}<span class="muted" style="font-weight:normal;font-size:12px;"> · 조회 ${esc(
            from
          )} ~ ${esc(to)}</span>${esc(adjNote)}`;
        }
        if (newBadgeEl) {
          const badges = [];
          if (newItems.length) {
            badges.push(
              '<span style="display:inline-block; margin-left:8px; padding:2px 8px; border-radius:999px; background:#fee4e2; color:#b42318; font-size:12px; vertical-align:middle;">신규 ' +
                newItems.length +
                "건</span>"
            );
          }
          if (modifiedItems.length) {
            badges.push(
              '<span style="display:inline-block; margin-left:8px; padding:2px 8px; border-radius:999px; background:#dbeafe; color:#1d4ed8; font-size:12px; vertical-align:middle;">수정 ' +
                modifiedItems.length +
                "건</span>"
            );
          }
          newBadgeEl.innerHTML = badges.join("");
        }
      } else if (apiStatusEl) {
        apiStatusEl.innerHTML = `API 기준 신규/수정 발주서 없음 (업로드 목록이 해당 구간을 반영한 상태).<span class="muted" style="font-weight:normal;font-size:12px;"> · 조회 ${esc(
          from
        )} ~ ${esc(to)}</span>${esc(adjNote)}`;
      }
    } catch (e) {
      if (apiStatusEl) apiStatusEl.textContent = `API 신규 확인 실패: ${e.message || "오류"}`;
    }
  };

  qs("#inbound-plan-2-refresh")?.addEventListener("click", () => loadList());
  qs("#inbound-plan-2-check-new")?.addEventListener("click", () => checkApiNewSlips());
  qs("#inbound-plan-2-delete-selected")?.addEventListener("click", async () => {
    const selected = inboundPlan2Items.filter((x) => selectedSlipKeys.has(normalizeSlipNoText(x.poNo)));
    if (!selected.length) {
      alert("삭제할 전표를 먼저 체크하세요.");
      return;
    }
    const ok = confirm(`선택한 전표 ${selected.length}건을 삭제할까요?\n(전표별 품목 행 전체가 삭제됩니다.)`);
    if (!ok) return;
    try {
      const res = await api("/api/inbound-plan-upload/delete", {
        method: "POST",
        body: JSON.stringify({ slipNos: selected.map((x) => x.poNo) })
      });
      inboundPlan2Items = Array.isArray(res.items) ? res.items : [];
      selectedSlipKeys = new Set();
      if (metaEl) {
        const fn = res.sourceFileName ? esc(res.sourceFileName) : "-";
        const at = res.uploadedAt ? esc(res.uploadedAt.slice(0, 19).replace("T", " ")) : "-";
        metaEl.innerHTML = `저장된 파일: <strong>${fn}</strong> · 반영 시각: ${at} · 누적 품목 행 <strong>${Number(
          res.lineCount || 0
        )}</strong> · 누적 전표 <strong>${res.slipCount ?? inboundPlan2Items.length}</strong>`;
      }
      drawTable();
      await checkApiNewSlips();
      if (!Number(res.deletedLineCount || 0)) {
        alert("삭제된 항목이 없습니다. 체크한 발주번호 형식이 저장 데이터와 다른지 확인해 주세요.");
      } else {
        alert(`삭제 완료: 전표 ${res.deletedSlipCount ?? selected.length}건, 품목 행 ${res.deletedLineCount ?? 0}줄`);
      }
    } catch (e) {
      alert(e.message || "선택 삭제 실패");
    }
  });
  qs("#inbound-plan-2-upload")?.addEventListener("click", async () => {
    const input = qs("#inbound-plan-2-file");
    const file = input?.files?.[0];
    if (!file) {
      alert("엑셀 파일을 선택하세요.");
      return;
    }
    try {
      const { matrix, sourceFileName } = await parseSheetAsMatrix(file);
      const res = await api("/api/inbound-plan-upload", {
        method: "POST",
        body: JSON.stringify({ matrix, sourceFileName })
      });
      inboundPlan2Items = Array.isArray(res.items) ? res.items : [];
      if (metaEl) {
        const fn = esc(res.sourceFileName || sourceFileName);
        const at = res.uploadedAt ? esc(res.uploadedAt.slice(0, 19).replace("T", " ")) : "";
        const skip =
          res.skippedLineCount > 0
            ? ` · 이번 업로드에서 생략(기존 발주와 동일) 품목 행 <strong>${res.skippedLineCount}</strong> · 해당 전표 수 <strong>${res.skippedExistingSlipCount ?? 0}</strong>`
            : "";
        const add =
          res.addedLineCount != null
            ? ` · 이번에 추가된 품목 행 <strong>${res.addedLineCount}</strong> · 신규 전표 <strong>${res.newSlipCount ?? 0}</strong>`
            : "";
        metaEl.innerHTML = `저장된 파일: <strong>${fn}</strong> · 반영 시각: ${at} · 누적 품목 행 <strong>${Number(res.lineCount || 0)}</strong> · 누적 전표 <strong>${res.slipCount ?? inboundPlan2Items.length}</strong>${add}${skip}`;
      }
      drawTable();
      const skipMsg =
        res.skippedLineCount > 0
          ? `\n건너뜀: 동일 발주번호 ${res.skippedExistingSlipCount ?? 0}건(품목 행 ${res.skippedLineCount}줄). 기존 데이터 유지.`
          : "";
      const addMsg =
        res.addedLineCount != null
          ? `\n추가: 신규 전표 ${res.newSlipCount ?? 0}건, 품목 행 ${res.addedLineCount}줄.`
          : "";
      alert(
        `반영 완료.\n누적 전표 ${res.slipCount ?? inboundPlan2Items.length}건, 누적 품목 행 ${res.lineCount ?? 0}건.${addMsg}${skipMsg}`
      );
      await checkApiNewSlips();
    } catch (e) {
      alert(e.message || "업로드 실패");
    }
  });

  await loadList();
  await checkApiNewSlips();
}

function renderOutbound() {
  qs("#view-outbound").innerHTML = movementOutboundPageHtml();
  bindMovement("OUT");
  setupOutboundModals();
  renderRecentMovements("OUT");
}

function renderOutboundPlanMenu() {
  const wrap = qs("#view-outbound-plan");
  if (!wrap) return;
  wrap.innerHTML = `
    <div class="card">
      <h2>출고예정관리</h2>
      <p class="muted">판매처별 발주서 업로드 메뉴를 선택하세요.</p>
      <div class="row" style="grid-template-columns: repeat(3, minmax(160px, 1fr)); gap: 10px;">
        <button type="button" class="primary" id="go-outbound-daiso">다이소</button>
        <button type="button" class="primary" id="go-outbound-emart">이마트</button>
        <button type="button" class="primary" id="go-outbound-lotte">롯데마트</button>
      </div>
    </div>
  `;
  qs("#go-outbound-daiso")?.addEventListener("click", async () => {
    switchView("outbound-upload-daiso");
    await renderOutboundPartnerUpload("다이소", "daiso");
  });
  qs("#go-outbound-emart")?.addEventListener("click", async () => {
    switchView("outbound-upload-emart");
    await renderOutboundPartnerUpload("이마트", "emart");
  });
  qs("#go-outbound-lotte")?.addEventListener("click", async () => {
    switchView("outbound-upload-lotte");
    await renderOutboundPartnerUpload("롯데마트", "lotte");
  });
}

async function renderOutboundPartnerUpload(partner, key) {
  const wrap = qs(`#view-outbound-upload-${key}`);
  if (!wrap) return;
  wrap.innerHTML = `
    <!-- 컴팩트 헤더 -->
    <div class="ob-topbar card">
      <div class="ob-topbar-left">
        <span class="ob-partner-badge">${esc(partner)}</span>
        <span class="ob-topbar-title">판매처별 업로드</span>
        <div class="ob-pipeline-inline">
          <span>업로드</span><span class="ob-pi-arrow">›</span>
          <span>검증</span><span class="ob-pi-arrow">›</span>
          <span>전표생성</span><span class="ob-pi-arrow">›</span>
          <span>미출계산</span><span class="ob-pi-arrow">›</span>
          <span>출고확정</span><span class="ob-pi-arrow">›</span>
          <span>판매입력 전송</span>
        </div>
      </div>
      <span id="outbound-upload-status-${key}" class="ob-status-badge ob-status-idle">대기 중</span>
    </div>

    <!-- 탭 + 패널 -->
    <div class="card ob-main-card">
      <div class="ob-tab-bar">
        <div class="seg-tabs">
          <button type="button" class="seg-tab active" id="outbound-tab-orders-${key}">주문서</button>
          <button type="button" class="seg-tab" id="outbound-tab-master-${key}">마스터</button>
          <button type="button" class="seg-tab" id="outbound-tab-slip-${key}">전표관리</button>
          <button type="button" class="seg-tab" id="outbound-tab-confirmlist-${key}">확정리스트</button>
          <button type="button" class="seg-tab" id="outbound-tab-unship-${key}">미출내역</button>
          ${key === "emart" ? `<button type="button" class="seg-tab" id="outbound-tab-compare-${key}">미출 대조</button>` : ""}
        </div>
      </div>

      <!-- 주문서 패널 -->
      <div id="outbound-orders-panel-${key}" class="ob-panel">
        <!-- 통합 툴바 -->
        <div class="ob-toolbar">
          <label class="ob-file-label">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            파일 선택
            <input type="file" id="outbound-upload-file-${key}" accept=".xlsx,.xls,.csv" class="hidden-file" />
          </label>
          <button type="button" class="bh-btn bh-btn-primary bh-btn-sm" id="outbound-upload-run-${key}">업로드</button>
          <div class="ob-divider"></div>
          <button type="button" class="bh-btn bh-btn-sm" id="outbound-upload-refresh-${key}">새로고침</button>
          <button type="button" class="bh-btn bh-btn-sm" id="outbound-upload-template-${key}">엑셀 템플릿</button>
          <button type="button" class="bh-btn bh-btn-sm" id="outbound-upload-batch-open-${key}">파일 리스트</button>
          <div class="ob-divider"></div>
          <button type="button" class="bh-btn bh-btn-sm bh-btn-danger-outline" id="outbound-batch-delete-selected-${key}">선택 삭제</button>
          <button type="button" class="bh-btn bh-btn-sm bh-btn-danger-outline" id="outbound-batch-delete-all-${key}">전체 삭제</button>
        </div>
        <div id="outbound-upload-batches-${key}" class="hidden ob-batches-wrap"></div>
        <div id="outbound-line-tabs-${key}" class="ob-line-tabs"></div>
        <div id="outbound-line-tools-${key}" class="ob-line-tools"></div>
        <div id="outbound-line-table-${key}" class="ob-table-wrap"></div>
      </div>

      <!-- 마스터 패널 -->
      <div id="outbound-master-panel-${key}" class="ob-panel hidden">
        <div class="ob-toolbar">
          <label class="ob-file-label">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            파일 선택
            <input type="file" id="outbound-master-file-${key}" accept=".xlsx,.xls,.csv" class="hidden-file" />
          </label>
          <button type="button" class="bh-btn bh-btn-primary bh-btn-sm" id="outbound-master-run-${key}">업로드</button>
          <span id="outbound-master-status-${key}" class="ob-status-badge ob-status-idle">대기 중</span>
          <span class="ob-toolbar-hint">상품코드 매핑용 코드마스터 파일을 업로드합니다.</span>
        </div>
        <div id="outbound-master-table-${key}" class="ob-table-wrap"></div>
      </div>

      <!-- 전표관리 패널 -->
      <div id="outbound-slip-panel-${key}" class="ob-panel hidden">
        <div class="ob-toolbar">
          <span class="ob-toolbar-hint">전표 목록 조회·선택 후 <strong>확정리스트</strong> 탭에서 출고확정·판매입력 전송을 진행합니다.</span>
        </div>
        <div id="outbound-upload-table-${key}" class="ob-table-wrap"></div>
      </div>

      <!-- 확정리스트 패널 -->
      <div id="outbound-confirmlist-panel-${key}" class="ob-panel hidden">
        <div class="ob-toolbar">
          <button type="button" class="bh-btn bh-btn-primary bh-btn-sm" id="outbound-confirm-all-${key}">전체 출고확정</button>
          <span class="ob-toolbar-hint" id="outbound-confirm-all-hint-${key}">선택 전표(없으면 전체) 기준으로 출고확정합니다. · <strong>센터 통합 전표</strong> 기준 (센터·저장일·점입점일)</span>
          <button type="button" class="bh-btn bh-btn-sm" style="margin-left:auto;" id="outbound-confirmlist-refresh-${key}">새로고침</button>
          <span class="muted" id="outbound-confirmlist-status-${key}"></span>
        </div>
        <div id="outbound-confirmlist-table-${key}" class="ob-table-wrap"></div>
      </div>

      <!-- 미출내역 패널 -->
      <div id="outbound-unship-panel-${key}" class="ob-panel hidden">
        <div class="ob-toolbar">
          <span class="ob-toolbar-hint">출고 상태가 <strong>미출</strong>인 라인만 표시합니다.</span>
          <button type="button" class="bh-btn bh-btn-sm" style="margin-left:auto;" id="outbound-unship-refresh-${key}">새로고침</button>
          <span class="muted" id="outbound-unship-status-${key}"></span>
        </div>
        <div id="outbound-unship-tabs-${key}" class="ob-line-tabs"></div>
        <div id="outbound-unship-tools-${key}" class="ob-line-tools"></div>
        <div id="outbound-unship-table-${key}" class="ob-table-wrap"></div>
      </div>

      ${key === "emart" ? `
      <!-- 미출 대조 패널 (이마트 전용) -->
      <div id="outbound-compare-panel-${key}" class="ob-panel hidden">

        <!-- eCvan 자동수집 섹션 -->
        <div class="ecvan-section">
          <div class="ecvan-section-head">
            <span class="ecvan-logo">eCvan</span>
            <span class="ecvan-label">자동 미출 수집</span>
            <button type="button" class="bh-btn bh-btn-primary bh-btn-sm" id="ecvan-start-${key}">로그인 &amp; 수집 시작</button>
            <div id="ecvan-status-${key}" class="ecvan-status-bar hidden"></div>
          </div>
        </div>

        <!-- eCvan 로그인/OTP 모달 -->
        <div id="ecvan-modal-${key}" class="ecvan-modal hidden">
          <div class="ecvan-modal-box">
            <div class="ecvan-modal-head">
              <span class="ecvan-logo">eCvan</span>&nbsp;로그인
              <button type="button" class="ecvan-modal-close" id="ecvan-close-${key}">✕</button>
            </div>
            <div id="ecvan-login-step-${key}">
              <div class="ecvan-field"><label>아이디</label><input type="text" id="ecvan-id-${key}" placeholder="eCvan 아이디" autocomplete="username" /></div>
              <div class="ecvan-field"><label>비밀번호</label><input type="password" id="ecvan-pw-${key}" placeholder="비밀번호" autocomplete="current-password" /></div>
              <div style="display:flex;gap:8px;margin-top:8px;">
                <button type="button" class="bh-btn bh-btn-secondary" style="flex:0 0 auto;" id="ecvan-save-btn-${key}">저장</button>
                <button type="button" class="bh-btn bh-btn-primary" style="flex:1;" id="ecvan-login-btn-${key}">로그인 &amp; SMS 발송</button>
              </div>
            </div>
            <div id="ecvan-otp-step-${key}" class="hidden">
              <p class="ecvan-otp-hint">휴대폰으로 발송된 인증번호를 입력하세요.</p>
              <div class="ecvan-field"><label>인증번호</label><input type="text" id="ecvan-otp-${key}" placeholder="6자리 인증번호" maxlength="6" inputmode="numeric" /></div>
              <button type="button" class="bh-btn bh-btn-primary" style="width:100%;margin-top:8px;" id="ecvan-otp-btn-${key}">확인 &amp; 수집 시작</button>
            </div>
            <div id="ecvan-progress-${key}" class="ecvan-progress hidden">
              <div class="ecvan-spinner"></div>
              <div id="ecvan-progress-text-${key}" class="ecvan-progress-text">처리 중...</div>
            </div>
          </div>
        </div>

        <!-- 수동 업로드 툴바 -->
        <div class="ob-toolbar">
          <span class="ob-toolbar-hint" style="color:var(--t3);font-size:11.5px;">또는 직접 업로드:</span>
          <label class="ob-file-label">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
            이마트 미출 파일
            <input type="file" id="uc-file-${key}" accept=".xlsx,.xls,.csv" class="hidden-file" />
          </label>
          <button type="button" class="bh-btn bh-btn-sm" id="uc-upload-${key}">업로드</button>
          <span class="ob-toolbar-hint" id="uc-upload-hint-${key}"></span>
          <div class="ob-divider"></div>
          <input type="date" class="uc-date-input" id="uc-start-${key}" />
          <span class="ob-pi-arrow">~</span>
          <input type="date" class="uc-date-input" id="uc-end-${key}" />
          <button type="button" class="bh-btn bh-btn-sm" id="uc-search-${key}">조회</button>
          <button type="button" class="bh-btn bh-btn-sm bh-btn-danger-outline" style="margin-left:auto;" id="uc-clear-${key}">초기화</button>
        </div>

        <!-- 컬럼 매핑 안내 -->
        <div class="uc-guide" id="uc-guide-${key}">
          <div class="uc-guide-title">📋 엑셀 컬럼 매핑 안내</div>
          <div class="uc-guide-cols">
            <span><strong>납기일</strong> 납기일자 / 납기일 / DueDate</span>
            <span><strong>상품코드</strong> 상품코드 / ItemCode / 품목코드</span>
            <span><strong>상품명</strong> 상품명 / ItemName / 품목명</span>
            <span><strong>센터</strong> 센터 / 물류센터 / 센터명 / Center</span>
            <span><strong>미출수량</strong> 미출수량 / 미납수량 / UnshipQty / 미출</span>
          </div>
        </div>

        <!-- 대조 테이블 -->
        <div id="uc-table-${key}"></div>
      </div>
      ` : ""}
    </div>

    <!-- 모달: 전표 상세 -->
    <div id="outbound-upload-detail-overlay-${key}" class="modal-overlay hidden">
      <div class="modal large">
        <div class="modal-header">
          <h3 id="outbound-upload-detail-title-${key}">전표 상세</h3>
          <button type="button" class="cancel-btn del-small" id="outbound-upload-detail-close-${key}">닫기</button>
        </div>
        <div id="outbound-upload-detail-body-${key}" style="max-height:65vh; overflow:auto;"></div>
      </div>
    </div>

    <!-- 모달: 업로드 파일 리스트 -->
    <div id="outbound-batch-overlay-${key}" class="modal-overlay hidden">
      <div class="modal large">
        <div class="modal-header">
          <h3>업로드 파일 리스트</h3>
          <button type="button" class="cancel-btn del-small" id="outbound-batch-overlay-close-${key}">닫기</button>
        </div>
        <div id="outbound-batch-overlay-body-${key}" style="max-height:60vh; overflow:auto;"></div>
      </div>
    </div>

    <!-- 모달: 확정 목록 상세 -->
    <div id="outbound-confirmlist-detail-overlay-${key}" class="modal-overlay hidden" role="dialog" aria-modal="true">
      <div class="modal large">
        <div class="modal-header">
          <h3 id="outbound-confirmlist-detail-title-${key}">확정 목록</h3>
          <button type="button" class="cancel-btn del-small" id="outbound-confirmlist-detail-close-${key}">닫기</button>
        </div>
        <div id="outbound-confirmlist-detail-body-${key}" style="max-height:65vh; overflow:auto;"></div>
      </div>
    </div>
  `;
  const tableEl = qs(`#outbound-upload-table-${key}`);
  const statusEl = qs(`#outbound-upload-status-${key}`);
  const masterStatusEl = qs(`#outbound-master-status-${key}`);
  const batchesEl = qs(`#outbound-upload-batches-${key}`);
  const masterTableEl = qs(`#outbound-master-table-${key}`);
  const lineTabsEl = qs(`#outbound-line-tabs-${key}`);
  const lineToolsEl = qs(`#outbound-line-tools-${key}`);
  const lineTableEl = qs(`#outbound-line-table-${key}`);
  const confirmListPanel = qs(`#outbound-confirmlist-panel-${key}`);
  const overlay = qs(`#outbound-upload-detail-overlay-${key}`);
  const batchOverlay = qs(`#outbound-batch-overlay-${key}`);
  const batchOverlayBody = qs(`#outbound-batch-overlay-body-${key}`);
  const detailBody = qs(`#outbound-upload-detail-body-${key}`);
  const detailTitle = qs(`#outbound-upload-detail-title-${key}`);
  const ordersPanel = qs(`#outbound-orders-panel-${key}`);
  const masterPanel = qs(`#outbound-master-panel-${key}`);
  const slipPanel = qs(`#outbound-slip-panel-${key}`);
  const confirmListTableEl = qs(`#outbound-confirmlist-table-${key}`);
  const confirmListStatusEl = qs(`#outbound-confirmlist-status-${key}`);
  const confirmListRefreshBtn = qs(`#outbound-confirmlist-refresh-${key}`);
  const unshipPanel = qs(`#outbound-unship-panel-${key}`);
  const unshipTableEl = qs(`#outbound-unship-table-${key}`);
  const unshipTabsEl = qs(`#outbound-unship-tabs-${key}`);
  const unshipToolsEl = qs(`#outbound-unship-tools-${key}`);
  const unshipStatusEl = qs(`#outbound-unship-status-${key}`);
  const unshipRefreshBtn = qs(`#outbound-unship-refresh-${key}`);
  const confirmListDetailOverlay = qs(`#outbound-confirmlist-detail-overlay-${key}`);
  const confirmListDetailBody = qs(`#outbound-confirmlist-detail-body-${key}`);
  const confirmListDetailTitle = qs(`#outbound-confirmlist-detail-title-${key}`);
  const closeConfirmListDetail = () => confirmListDetailOverlay?.classList.add("hidden");
  qs(`#outbound-confirmlist-detail-close-${key}`)?.addEventListener("click", closeConfirmListDetail);
  confirmListDetailOverlay?.addEventListener("click", (e) => {
    if (e.target === confirmListDetailOverlay) closeConfirmListDetail();
  });
  const tabOrdersBtn = qs(`#outbound-tab-orders-${key}`);
  const tabMasterBtn = qs(`#outbound-tab-master-${key}`);
  const tabSlipBtn = qs(`#outbound-tab-slip-${key}`);
  const tabConfirmListBtn = qs(`#outbound-tab-confirmlist-${key}`);
  const tabUnshipBtn = qs(`#outbound-tab-unship-${key}`);
  const tabCompareBtn = qs(`#outbound-tab-compare-${key}`);
  const comparePanel = qs(`#outbound-compare-panel-${key}`);
  const confirmAllBtn = qs(`#outbound-confirm-all-${key}`);
  let slipItems = [];
  let selectedSlipKeys = new Set();
  let selectedBatchIds = new Set();
  let currentBatches = [];
  let appliedBatchId = "";
  let centerTab = "all";
  let lineSearchQ = "";
  let selectedLineKeys = new Set();
  let unshipCenterTab = "all";
  let unshipSearchQ = "";

  const switchTopTab = (tab) => {
    const isOrders = tab === "orders";
    const isMaster = tab === "master";
    const isSlip = tab === "slip";
    const isConfirmList = tab === "confirmlist";
    const isUnship = tab === "unship";
    const isCompare = tab === "compare";
    ordersPanel?.classList.toggle("hidden", !isOrders);
    masterPanel?.classList.toggle("hidden", !isMaster);
    slipPanel?.classList.toggle("hidden", !isSlip);
    confirmListPanel?.classList.toggle("hidden", !isConfirmList);
    unshipPanel?.classList.toggle("hidden", !isUnship);
    comparePanel?.classList.toggle("hidden", !isCompare);
    if (tabOrdersBtn) tabOrdersBtn.className = isOrders ? "seg-tab active" : "seg-tab";
    if (tabMasterBtn) tabMasterBtn.className = isMaster ? "seg-tab active" : "seg-tab";
    if (tabSlipBtn) tabSlipBtn.className = isSlip ? "seg-tab active" : "seg-tab";
    if (tabConfirmListBtn) tabConfirmListBtn.className = isConfirmList ? "seg-tab active" : "seg-tab";
    if (tabUnshipBtn) tabUnshipBtn.className = isUnship ? "seg-tab active" : "seg-tab";
    if (tabCompareBtn) tabCompareBtn.className = isCompare ? "seg-tab active" : "seg-tab";
  };
  tabOrdersBtn?.addEventListener("click", () => switchTopTab("orders"));
  tabMasterBtn?.addEventListener("click", async () => {
    switchTopTab("master");
    await loadMaster();
  });
  tabSlipBtn?.addEventListener("click", () => {
    switchTopTab("slip");
  });
  tabConfirmListBtn?.addEventListener("click", async () => {
    switchTopTab("confirmlist");
    await loadList();
    await loadConfirmList();
  });
  tabUnshipBtn?.addEventListener("click", async () => {
    switchTopTab("unship");
    await loadUnshipLines();
  });
  unshipRefreshBtn?.addEventListener("click", async () => {
    await loadUnshipLines();
  });

  // ── 미출 대조 (이마트 전용) ──────────────────────────────────────
  if (key === "emart" && tabCompareBtn) {
    const ucFile = qs(`#uc-file-${key}`);
    const ucUploadBtn = qs(`#uc-upload-${key}`);
    const ucUploadHint = qs(`#uc-upload-hint-${key}`);
    const ucStartInput = qs(`#uc-start-${key}`);
    const ucEndInput = qs(`#uc-end-${key}`);
    const ucSearchBtn = qs(`#uc-search-${key}`);
    const ucClearBtn = qs(`#uc-clear-${key}`);
    const ucTableEl = qs(`#uc-table-${key}`);
    const ucGuide = qs(`#uc-guide-${key}`);

    // 날짜 기본값: 오늘 기준 -14일 ~ +14일
    const today = new Date();
    const fmt = (d) => d.toISOString().slice(0, 10);
    const d14 = new Date(today); d14.setDate(today.getDate() - 14);
    if (ucStartInput) ucStartInput.value = fmt(d14);
    if (ucEndInput) ucEndInput.value = fmt(today);

    const ucActionLabels = { modify_qty: "수량변경", add_product: "추가", confirmed: "확인완료", none: "보류" };

    const renderCompareTable = (rows) => {
      if (!ucTableEl) return;
      if (!rows.length) {
        ucTableEl.innerHTML = `<div class="uc-empty">조회된 대조 데이터가 없습니다. 날짜 범위를 확인하거나 이마트 미출 파일을 업로드해주세요.</div>`;
        return;
      }
      // 납기일 그룹핑
      const byDate = new Map();
      for (const r of rows) {
        if (!byDate.has(r.dueDate)) byDate.set(r.dueDate, []);
        byDate.get(r.dueDate).push(r);
      }
      let html = `<div class="uc-table-wrap"><table class="uc-table"><thead><tr>
        <th>납기일</th><th>상품코드</th><th>상품명</th><th>센터</th>
        <th class="num">WMS 미출</th><th class="num">이마트 미출</th>
        <th class="num uc-diff-col">차이</th><th>처리상태</th><th>액션</th>
      </tr></thead><tbody>`;
      for (const [dueDate, dateRows] of byDate) {
        for (let i = 0; i < dateRows.length; i++) {
          const r = dateRows[i];
          const diff = r.diff;
          const diffClass = diff > 0 ? "uc-diff-plus" : diff < 0 ? "uc-diff-minus" : "uc-diff-zero";
          const diffText = diff > 0 ? `+${diff}` : String(diff);
          const adj = r.adjustment;
          const adjLabel = adj ? (ucActionLabels[adj.action] || adj.action) : "";
          const adjClass = adj ? `uc-adj-${adj.action}` : "";
          const rowKey = `${esc(r.dueDate)}|||${esc(r.productCode)}|||${esc(r.centerName)}`;
          html += `<tr data-rowkey="${rowKey}" class="${diff === 0 ? "" : "uc-row-diff"}">
            ${i === 0 ? `<td rowspan="${dateRows.length}" class="uc-date-cell">${esc(dueDate)}</td>` : ""}
            <td class="mono">${esc(r.productCode)}</td>
            <td>${esc(r.productName)}</td>
            <td>${esc(r.centerName)}</td>
            <td class="num">${r.wmsQty.toLocaleString()}</td>
            <td class="num">${r.officialQty > 0 ? r.officialQty.toLocaleString() : '<span class="muted">-</span>'}</td>
            <td class="num ${diffClass}">${r.officialQty > 0 || r.wmsQty > 0 ? diffText : '<span class="muted">-</span>'}</td>
            <td>${adj ? `<span class="uc-adj-badge ${adjClass}">${esc(adjLabel)}</span><br><span class="muted" style="font-size:11px;">${adj.newQty != null ? adj.newQty + "개" : ""}</span>` : '<span class="muted">미처리</span>'}</td>
            <td class="uc-actions">
              ${diff < 0 ? `<button class="bh-btn bh-btn-sm" data-act="modify_qty" data-row='${JSON.stringify({dueDate:r.dueDate,productCode:r.productCode,productName:r.productName,centerName:r.centerName,wmsQty:r.wmsQty,officialQty:r.officialQty})}'>수량변경</button>` : ""}
              ${r.wmsQty === 0 && r.officialQty > 0 ? `<button class="bh-btn bh-btn-sm bh-btn-primary" data-act="add_product" data-row='${JSON.stringify({dueDate:r.dueDate,productCode:r.productCode,productName:r.productName,centerName:r.centerName,officialQty:r.officialQty})}'>추가</button>` : ""}
              ${diff > 0 ? `<button class="bh-btn bh-btn-sm bh-btn-primary" data-act="modify_qty" data-row='${JSON.stringify({dueDate:r.dueDate,productCode:r.productCode,productName:r.productName,centerName:r.centerName,wmsQty:r.wmsQty,officialQty:r.officialQty})}'>수량변경</button>` : ""}
              <button class="bh-btn bh-btn-sm" data-act="confirmed" data-row='${JSON.stringify({dueDate:r.dueDate,productCode:r.productCode,productName:r.productName,centerName:r.centerName})}'>확인</button>
            </td>
          </tr>`;
        }
      }
      html += `</tbody></table></div>`;
      ucTableEl.innerHTML = html;

      // 액션 버튼 이벤트
      ucTableEl.querySelectorAll("button[data-act]").forEach((btn) => {
        btn.addEventListener("click", async () => {
          const act = btn.dataset.act;
          let rowData;
          try { rowData = JSON.parse(btn.dataset.row); } catch { return; }
          let newQty = null;
          let note = "";
          if (act === "modify_qty") {
            const v = prompt(`[${rowData.productCode}] ${rowData.productName}\n센터: ${rowData.centerName}\n납기일: ${rowData.dueDate}\n\nWMS 미출: ${rowData.wmsQty}개 / 이마트 미출: ${rowData.officialQty}개\n\n조정할 수량을 입력하세요:`);
            if (v === null) return;
            newQty = parseInt(v, 10);
            if (!Number.isFinite(newQty) || newQty < 0) { alert("올바른 수량을 입력해주세요."); return; }
          } else if (act === "add_product") {
            const v = prompt(`[${rowData.productCode}] ${rowData.productName}\n센터: ${rowData.centerName}\n납기일: ${rowData.dueDate}\n\n이마트 미출 수량: ${rowData.officialQty}개\n\n추가할 수량을 입력하세요:`);
            if (v === null) return;
            newQty = parseInt(v, 10);
            if (!Number.isFinite(newQty) || newQty <= 0) { alert("올바른 수량을 입력해주세요."); return; }
          }
          try {
            await api("/api/unship-compare/adjust", "POST", {
              partnerType: key,
              dueDate: rowData.dueDate, productCode: rowData.productCode,
              productName: rowData.productName, centerName: rowData.centerName,
              action: act, newQty, note
            });
            await loadCompare();
          } catch (e) { alert(e.message); }
        });
      });
    };

    const loadCompare = async () => {
      if (!ucTableEl) return;
      const start = ucStartInput?.value || "";
      const end = ucEndInput?.value || "";
      try {
        ucTableEl.innerHTML = `<div class="uc-empty">불러오는 중...</div>`;
        const res = await api(`/api/unship-compare?partnerType=${key}&startDate=${start}&endDate=${end}`);
        if (ucGuide) ucGuide.classList.toggle("hidden", res.hasOfficial);
        if (ucUploadHint) {
          ucUploadHint.textContent = res.hasOfficial
            ? `이마트 공식 미출 ${res.officialRowCount}건 로드됨`
            : "이마트에서 받은 공식 미출 내역 파일을 업로드합니다.";
        }
        renderCompareTable(res.rows || []);
      } catch (e) {
        ucTableEl.innerHTML = `<div class="uc-empty" style="color:var(--red);">${esc(e.message)}</div>`;
      }
    };

    tabCompareBtn.addEventListener("click", async () => {
      switchTopTab("compare");
      await loadCompare();
    });

    ucSearchBtn?.addEventListener("click", loadCompare);
    ucStartInput?.addEventListener("keydown", (e) => { if (e.key === "Enter") loadCompare(); });
    ucEndInput?.addEventListener("keydown", (e) => { if (e.key === "Enter") loadCompare(); });

    ucUploadBtn?.addEventListener("click", async () => {
      const file = ucFile?.files?.[0];
      if (!file) { alert("파일을 선택해주세요."); return; }
      try {
        const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!raw.length) throw new Error("파일에 데이터가 없습니다.");
        const headers = Object.keys(raw[0]);
        const findCol = (...aliases) => headers.find((h) => aliases.some((a) => String(h).includes(a))) || null;
        const colDueDate = findCol("납기일자", "납기일", "DueDate", "dueDate", "납기");
        const colCode = findCol("상품코드", "품목코드", "ItemCode", "itemcode", "코드");
        const colName = findCol("상품명", "품목명", "ItemName", "itemname", "상품");
        const colCenter = findCol("물류센터", "센터명", "센터", "Center", "center");
        const colQty = findCol("미출수량", "미납수량", "UnshipQty", "미출", "미납", "수량");
        if (!colCode || !colQty) throw new Error(`필수 컬럼을 찾을 수 없습니다.\n상품코드: ${colCode||"없음"} / 미출수량: ${colQty||"없음"}`);
        const rows = raw.map((r) => ({
          dueDate: String(r[colDueDate] ?? "").trim(),
          productCode: String(r[colCode] ?? "").trim(),
          productName: String(r[colName] ?? "").trim(),
          centerName: String(r[colCenter] ?? "").trim(),
          officialQty: Number(String(r[colQty] ?? "0").replace(/,/g, ""))
        })).filter((r) => r.productCode);
        const res = await api(`/api/unship-compare/upload?partnerType=${key}`, "POST", { rows });
        alert(`${res.count}건 업로드 완료`);
        if (ucFile) ucFile.value = "";
        await loadCompare();
      } catch (e) { alert(`업로드 오류: ${e.message}`); }
    });

    ucClearBtn?.addEventListener("click", async () => {
      if (!confirm("이마트 공식 미출 데이터와 조정 기록을 모두 초기화합니다. 계속하시겠습니까?")) return;
      await api("/api/unship-compare/clear", "POST", { partnerType: key });
      await loadCompare();
    });

    // ── eCvan 자동수집 ────────────────────────────────────────────
    const ecvanStartBtn = qs(`#ecvan-start-${key}`);
    const ecvanModal = qs(`#ecvan-modal-${key}`);
    const ecvanCloseBtn = qs(`#ecvan-close-${key}`);
    const ecvanLoginStep = qs(`#ecvan-login-step-${key}`);
    const ecvanOtpStep = qs(`#ecvan-otp-step-${key}`);
    const ecvanProgressEl = qs(`#ecvan-progress-${key}`);
    const ecvanProgressText = qs(`#ecvan-progress-text-${key}`);
    const ecvanSaveBtn = qs(`#ecvan-save-btn-${key}`);
    const ecvanLoginBtn = qs(`#ecvan-login-btn-${key}`);
    const ecvanOtpBtn = qs(`#ecvan-otp-btn-${key}`);
    const ecvanStatusBar = qs(`#ecvan-status-${key}`);
    let ecvanSessionId = null;
    let ecvanPollTimer = null;

    const ecvanShowStep = (step) => {
      ecvanLoginStep?.classList.toggle("hidden", step !== "login");
      ecvanOtpStep?.classList.toggle("hidden", step !== "otp");
      ecvanProgressEl?.classList.toggle("hidden", step !== "progress");
    };

    const ecvanPollStatus = () => {
      if (!ecvanSessionId) return;
      ecvanPollTimer = setInterval(async () => {
        try {
          const res = await api(`/api/ecvan/status?sessionId=${ecvanSessionId}`);
          if (ecvanProgressText) ecvanProgressText.textContent = res.progress || res.status;
          if (ecvanStatusBar) {
            ecvanStatusBar.classList.remove("hidden");
            ecvanStatusBar.textContent = res.progress || res.status;
          }
          if (res.status === "waiting_otp") {
            clearInterval(ecvanPollTimer);
            ecvanShowStep("otp");
          } else if (res.status === "done") {
            clearInterval(ecvanPollTimer);
            ecvanModal?.classList.add("hidden");
            if (ecvanStatusBar) ecvanStatusBar.textContent = `✅ ${res.progress}`;
            await loadCompare();
            switchTopTab("compare");
          } else if (res.status === "error") {
            clearInterval(ecvanPollTimer);
            ecvanShowStep("login");
            alert(`오류: ${res.error}`);
          }
        } catch (e) { /* 네트워크 오류 무시 */ }
      }, 1500);
    };

    ecvanStartBtn?.addEventListener("click", async () => {
      ecvanShowStep("login");
      ecvanModal?.classList.remove("hidden");
      try {
        const creds = await api("/api/ecvan/credentials");
        const idEl = qs(`#ecvan-id-${key}`);
        const pwEl = qs(`#ecvan-pw-${key}`);
        if (idEl && creds.id) idEl.value = creds.id;
        if (pwEl && creds.hasPw) {
          pwEl.placeholder = "저장된 비밀번호 사용 (변경 시 입력)";
          pwEl.dataset.hasSaved = "1";
        }
      } catch {}
    });
    ecvanCloseBtn?.addEventListener("click", () => {
      ecvanModal?.classList.add("hidden");
      clearInterval(ecvanPollTimer);
    });

    ecvanSaveBtn?.addEventListener("click", async () => {
      const id = qs(`#ecvan-id-${key}`)?.value.trim();
      const pw = qs(`#ecvan-pw-${key}`)?.value.trim();
      if (!id) { alert("아이디를 입력해주세요."); return; }
      try {
        await api("/api/ecvan/credentials", "POST", { id, pw: pw || undefined });
        const pwEl = qs(`#ecvan-pw-${key}`);
        if (pwEl) {
          pwEl.value = "";
          pwEl.placeholder = "저장된 비밀번호 사용 (변경 시 입력)";
          pwEl.dataset.hasSaved = "1";
        }
        alert("저장됐습니다.");
      } catch (e) {
        alert(`저장 실패: ${e.message}`);
      }
    });

    ecvanLoginBtn?.addEventListener("click", async () => {
      const idEl = qs(`#ecvan-id-${key}`);
      const pwEl = qs(`#ecvan-pw-${key}`);
      const id = idEl?.value.trim();
      const pw = pwEl?.value.trim();
      const useSaved = !pw && pwEl?.dataset.hasSaved === "1";
      if (!id) { alert("아이디를 입력해주세요."); return; }
      if (!pw && !useSaved) { alert("비밀번호를 입력해주세요."); return; }
      ecvanShowStep("progress");
      if (ecvanProgressText) ecvanProgressText.textContent = "로그인 중... (브라우저 자동 실행)";
      try {
        const body = useSaved ? { id } : { id, pw };
        const res = await api("/api/ecvan/start", "POST", body);
        ecvanSessionId = res.sessionId;
        ecvanPollStatus();
      } catch (e) {
        ecvanShowStep("login");
        alert(`로그인 실패: ${e.message}`);
      }
    });

    ecvanOtpBtn?.addEventListener("click", async () => {
      const otp = qs(`#ecvan-otp-${key}`)?.value.trim();
      if (!otp) { alert("인증번호를 입력해주세요."); return; }
      ecvanShowStep("progress");
      if (ecvanProgressText) ecvanProgressText.textContent = "인증 후 데이터 수집 중...";
      try {
        await api("/api/ecvan/otp", "POST", { sessionId: ecvanSessionId, otp });
        ecvanPollStatus();
      } catch (e) {
        ecvanShowStep("otp");
        alert(`인증 실패: ${e.message}`);
      }
    });
  }

  switchTopTab("orders");

  const renderBatchList = (batches = [], appliedId = "") => {
    if (!batchesEl && !batchOverlayBody) return;
    appliedBatchId = String(appliedId || "").trim();
    currentBatches = [...batches].sort((a, b) => {
      const ta = new Date(a?.uploadedAt || 0).getTime();
      const tb = new Date(b?.uploadedAt || 0).getTime();
      return tb - ta;
    });
    selectedBatchIds = new Set(
      Array.from(selectedBatchIds).filter((id) => currentBatches.some((b) => String(b.uploadBatchId || "") === id))
    );
    const rows = currentBatches
      .map(
        (x, idx) => `<tr>
          <td><input type="checkbox" class="outbound-batch-check outbound-batch-check-${key}" data-batch="${esc(x.uploadBatchId || "")}" ${
            selectedBatchIds.has(String(x.uploadBatchId || "")) ? "checked" : ""
          } /></td>
          <td>${idx + 1}</td>
          <td>${esc(formatDateTimeDisplay(x.uploadedAt || ""))}</td>
          <td>${esc(x.sourceFileName || "")}${appliedBatchId === String(x.uploadBatchId || "") ? ` <span class="muted">(현재 적용)</span>` : ""}</td>
          <td>${esc(String(x.okCount || 0))}</td>
          <td>${esc(String(x.errorCount || 0))}</td>
          <td>
            <button type="button" class="primary del-small outbound-batch-apply-${key}" data-batch="${esc(x.uploadBatchId || "")}">적용</button>
            <button type="button" class="cancel-btn del-small outbound-batch-delete-${key}" data-batch="${esc(
              x.uploadBatchId || ""
            )}">업로드 삭제</button>
          </td>
        </tr>`
      )
      .join("");
    const appliedFile = currentBatches.find((b) => String(b.uploadBatchId || "") === appliedBatchId);
    const html = `<div class="muted" style="margin-bottom:6px;">업로드 이력(업로드 후 적용 버튼을 눌러야 전표 목록에 반영됩니다)</div>
      <div class="muted" style="margin-bottom:6px;">현재 적용 파일: ${appliedFile ? esc(appliedFile.sourceFileName || appliedBatchId) : "없음"}</div>
      <table id="outbound-batch-table-${key}">
      <thead><tr><th><input type="checkbox" class="outbound-batch-check-all" id="outbound-batch-check-all-${key}" /></th><th>순번</th><th>업로드일시</th><th>파일명</th><th>성공</th><th>실패</th><th>액션</th></tr></thead>
      <tbody>${rows || `<tr><td colspan="7" class="muted">업로드 이력 없음</td></tr>`}</tbody></table>`;
    const root = batchOverlayBody || batchesEl;
    root.innerHTML = html;
    const batchChecks = Array.from(root.querySelectorAll(`.outbound-batch-check-${key}`));
    const batchCheckAll = qs(`#outbound-batch-check-all-${key}`);
    if (batchCheckAll) {
      batchCheckAll.checked = batchChecks.length > 0 && batchChecks.every((x) => x.checked);
      batchCheckAll.addEventListener("change", () => {
        selectedBatchIds = new Set();
        batchChecks.forEach((el) => {
          el.checked = batchCheckAll.checked;
          if (batchCheckAll.checked) selectedBatchIds.add(String(el.getAttribute("data-batch") || ""));
        });
      });
    }
    batchChecks.forEach((el) =>
      el.addEventListener("change", () => {
        const id = String(el.getAttribute("data-batch") || "");
        if (!id) return;
        if (el.checked) selectedBatchIds.add(id);
        else selectedBatchIds.delete(id);
        if (batchCheckAll) batchCheckAll.checked = batchChecks.length > 0 && batchChecks.every((x) => x.checked);
      })
    );
    root.querySelectorAll(`.outbound-batch-delete-${key}`).forEach((btn) =>
      btn.addEventListener("click", async () => {
        const batchId = String(btn.getAttribute("data-batch") || "").trim();
        if (!batchId) return;
        const ok = confirm("이 업로드 배치를 삭제할까요? (해당 업로드로 생성된 주문 라인/전표도 함께 정리됩니다)");
        if (!ok) return;
        try {
          await api("/api/outbound-order-upload/delete-batches", {
            method: "POST",
            body: JSON.stringify({ partnerType: key, uploadBatchIds: [batchId] })
          });
          await loadList();
          await loadLines();
          alert("업로드 데이터를 삭제했습니다.");
        } catch (e) {
          alert(e.message || "삭제 실패");
        }
      })
    );
    root.querySelectorAll(`.outbound-batch-apply-${key}`).forEach((btn) =>
      btn.addEventListener("click", async () => {
        const batchId = String(btn.getAttribute("data-batch") || "").trim();
        if (!batchId) return;
        try {
          await api("/api/outbound-order-upload/apply-batch", {
            method: "POST",
            body: JSON.stringify({ partnerType: key, uploadBatchId: batchId })
          });
          await loadList();
          await loadLines();
          alert("선택한 업로드 파일을 현재 주문서 데이터로 적용했습니다.");
        } catch (e) {
          alert(e.message || "적용 실패");
        }
      })
    );
  };

  const loadList = async () => {
    if (statusEl) statusEl.textContent = "조회 중...";
    try {
      const res = await api(`/api/outbound-order-upload?partnerType=${encodeURIComponent(key)}`);
      slipItems = Array.isArray(res.items) ? res.items : [];
      const rows = slipItems
        .map(
          (x, idx) => `<tr>
            <td><input type="checkbox" class="outbound-slip-check-${key}" data-slip="${esc(x.slipNo || "")}" ${
              selectedSlipKeys.has(normalizeSlipNoText(x.slipNo)) ? "checked" : ""
            } /></td>
            <td>${idx + 1}</td>
            <td><button type="button" class="bh-link-btn outbound-slip-open" data-slip="${esc(x.slipNo || "")}">${esc(x.slipNo || "")}</button></td>
            <td>${esc(x.orderNo || "")}</td>
            <td>${esc(formatYmdLoose(x.orderDate || ""))}</td>
            <td>${esc(x.partner || "")}</td>
            <td>${esc(String(x.totalOrderQty || 0))}</td>
            <td>${esc(String(x.totalConfirmQty || 0))}</td>
            <td>${esc(x.status || "draft")}</td>
            <td>${esc(formatDateTimeDisplay(x.confirmedAt || ""))}</td>
            <td>${esc(formatDateTimeDisplay(x.sentAt || ""))}</td>
          </tr>`
        )
        .join("");
      if (tableEl) {
        tableEl.innerHTML = `<table id="outbound-upload-slip-table-${key}">
          <thead><tr><th><input type="checkbox" id="outbound-slip-check-all-${key}" /></th><th>순번</th><th>전표번호</th><th>주문번호</th><th>주문일</th><th>판매처</th><th>주문수량</th><th>출고확인수량</th><th>상태</th><th>출고확정일시</th><th>판매입력전송일시</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="11" class="muted">전표가 없습니다.</td></tr>`}</tbody>
        </table>`;
        applyExcelLikeFilter(`#outbound-upload-slip-table-${key}`);
        const checks = Array.from(tableEl.querySelectorAll(`.outbound-slip-check-${key}`));
        const checkAllEl = qs(`#outbound-slip-check-all-${key}`);
        if (checkAllEl) {
          checkAllEl.checked = checks.length > 0 && checks.every((x) => x.checked);
          checkAllEl.addEventListener("change", () => {
            selectedSlipKeys = new Set();
            checks.forEach((el) => {
              el.checked = checkAllEl.checked;
              if (checkAllEl.checked) selectedSlipKeys.add(normalizeSlipNoText(el.getAttribute("data-slip")));
            });
          });
        }
        checks.forEach((el) =>
          el.addEventListener("change", () => {
            const k = normalizeSlipNoText(el.getAttribute("data-slip"));
            if (!k) return;
            if (el.checked) selectedSlipKeys.add(k);
            else selectedSlipKeys.delete(k);
            if (checkAllEl) checkAllEl.checked = checks.length > 0 && checks.every((x) => x.checked);
          })
        );
        tableEl.querySelectorAll(".outbound-slip-open").forEach((btn) =>
          btn.addEventListener("click", () => renderDetail(btn.getAttribute("data-slip"), { fullWorkflow: true }))
        );
      }
      const batches = Array.isArray(res.batches) ? res.batches : [];
      renderBatchList(batches, res.appliedBatchId || "");
      if (statusEl) statusEl.textContent = `조회 완료 (${slipItems.length}건)`;
    } catch (e) {
      if (statusEl) statusEl.textContent = e.message || "조회 실패";
      if (tableEl) tableEl.innerHTML = `<span class="muted">${esc(e.message || "오류")}</span>`;
    }
  };

  const runOutboundBulkConfirmAll = async () => {
    try {
      await loadList();
      const slipMap = new Map(slipItems.map((x) => [normalizeSlipNoText(x.slipNo), x]));
      const selectedKeys = Array.from(selectedSlipKeys || []);
      const candidateKeys =
        selectedKeys.length > 0
          ? selectedKeys
          : slipItems.map((x) => normalizeSlipNoText(x.slipNo)).filter(Boolean);

      const targets = candidateKeys
        .map((k) => slipMap.get(k))
        .filter((x) => x && String(x.status || "draft") === "draft");

      if (!targets.length) {
        alert("확정할 전표가 없습니다. (이미 확정/전송된 전표만 있거나 선택이 없습니다)");
        return;
      }

      const manager = prompt("출고 담당자명을 입력하세요");
      if (!manager) return;

      try {
        await api("/api/outbound-order-upload/workflow/bulk-confirm", {
          method: "POST",
          body: JSON.stringify({ slipNos: targets.map((t) => t.slipNo), stockUser: manager.trim() })
        });
      } catch (e) {
        alert(e.message || "전체 출고확정 실패(전표 단위로 오류가 나면 한 건도 확정되지 않습니다)");
        return;
      }

      await loadList();
      await loadLines();
      if (confirmListPanel && !confirmListPanel.classList.contains("hidden")) {
        await loadConfirmList();
      }
      alert(`전체 출고확정 완료: ${targets.length}건`);
    } catch (e) {
      alert(e.message || "전체 출고확정 실패");
    }
  };

  const loadConfirmList = async () => {
    if (!confirmListTableEl) return;
    try {
      if (confirmListStatusEl) confirmListStatusEl.textContent = "확정 저장 내역 불러오는 중...";
      const res = await api(`/api/outbound-order-upload/confirm-list?partnerType=${encodeURIComponent(key)}`);
      const items = Array.isArray(res.items) ? res.items : [];

      if (!items.length) {
        confirmListTableEl.innerHTML = `<span class="muted">확정 저장 내역이 없습니다.</span>`;
        if (confirmListStatusEl) confirmListStatusEl.textContent = "0건";
        return;
      }

      const rows = items
        .map((x) => {
          const saveDateDisp = formatYmd(String(x.saveDateYmd || ""));
          const storeInDateDisp = formatYmd(String(x.storeInDateYmd || ""));
          const listKeyAttr = esc(x.key || "");
          const slipCountNum = Number(x.slipCount || 0);
          const sentCountNum = Number(x.sentCount || 0);
          const sentLabel = x.allSent ? "전송됨" : slipCountNum > 0 ? `${sentCountNum}/${slipCountNum}` : "0/0";
          const sendLabel = x.allSent ? "재전송" : "판매입력 전송";
          const sendDisabled = slipCountNum <= 0 ? "disabled" : "";
          return `<tr>
            <td>${esc(saveDateDisp)}</td>
            <td>${esc(x.centerName || "")}</td>
            <td>${esc(storeInDateDisp)}</td>
            <td>${esc(String(x.slipCount || 0))}</td>
            <td>${esc(String(x.lineCount || 0))}</td>
            <td>${esc(String(x.productCount || 0))}</td>
            <td>${esc(String(x.totalQty || 0))}</td>
            <td>${esc(sentLabel)}</td>
            <td><button type="button" class="bh-link-btn outbound-confirmlist-key-${key}" data-list-key="${listKeyAttr}" title="목록 보기">${esc(
            x.key || ""
          )}</button></td>
            <td>
              <button type="button" class="primary del-small outbound-confirmlist-send-${key}" data-list-key="${listKeyAttr}" ${sendDisabled}>${sendLabel}</button>
              <button type="button" class="cancel-btn del-small outbound-confirmlist-cancel-${key}" data-list-key="${listKeyAttr}">확정취소</button>
            </td>
          </tr>`;
        })
        .join("");

      confirmListTableEl.innerHTML = `<div class="outbound-list-scroll" style="max-height:690px;">
        <table id="outbound-confirm-list-table-${key}">
          <thead><tr>
            <th>저장날짜</th><th>센터명</th><th>점입점일자</th><th>전표수</th><th>라인수</th><th>품목수</th><th>전체수량</th><th>전송상태</th><th>센터 통합 전표번호</th><th>액션</th>
          </tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>`;

      confirmListTableEl.querySelectorAll(`.outbound-confirmlist-key-${key}`).forEach((btn) => {
        btn.addEventListener("click", async () => {
          const listKey = String(btn.dataset.listKey || "").trim();
          if (!listKey || !confirmListDetailBody || !confirmListDetailTitle) return;
          confirmListDetailBody.innerHTML = `<p class="muted">불러오는 중...</p>`;
          confirmListDetailOverlay?.classList.remove("hidden");
          confirmListDetailTitle.textContent = `센터 통합 전표 — ${listKey}`;
          try {
            const res = await api(
              `/api/outbound-order-upload/confirm-list/detail?partnerType=${encodeURIComponent(key)}&listKey=${encodeURIComponent(listKey)}`
            );
            const item = res.item;
            if (!item) {
              confirmListDetailBody.innerHTML = `<span class="muted">데이터가 없습니다.</span>`;
              return;
            }
            const slipRows = (item.slipNos || [])
              .map((s, i) => `<tr><td>${i + 1}</td><td>${esc(s)}</td></tr>`)
              .join("");
            const lineRows = (item.lines || [])
              .map(
                (ln, i) => `<tr>
              <td>${i + 1}</td>
              <td>${esc(ln.slipNo || "")}</td>
              <td>${esc(ln.sourceSlipNo || "")}</td>
              <td>${esc(ln.productCode || "")}</td>
              <td>${esc(ln.productName || "")}</td>
              <td>${esc(String(ln.qty ?? ""))}</td>
              <td>${esc(ln.warehouse || "")}</td>
              <td>${esc(ln.lot || "")}</td>
              <td>${esc(ln.orderQty || "")}</td>
              <td>${esc(ln.fixedQty || "")}</td>
              <td>${esc(ln.unshipStatus || "")}</td>
            </tr>`
              )
              .join("");
            confirmListDetailTitle.textContent = `센터 통합 전표 — ${item.centerName || ""} / ${formatYmd(String(item.storeInDateYmd || ""))}`;
            confirmListDetailBody.innerHTML = `
              <p class="muted" style="margin-bottom:10px;">센터 통합 전표번호 <code style="font-size:12px;">${esc(
                item.centerMergedSlipNo || item.key || listKey
              )}</code>
                · 저장일 ${esc(formatYmd(String(item.saveDateYmd || "")))}
                · 갱신 ${esc(formatDateTimeDisplay(item.updatedAt || ""))}</p>
              <h4 style="margin:12px 0 6px; font-size:14px;">포함 전표</h4>
              <div class="outbound-list-scroll" style="max-height:220px; margin-bottom:14px;">
                <table class="outbound-nested-table"><thead><tr><th>#</th><th>전표번호</th></tr></thead><tbody>${
                  slipRows || `<tr><td colspan="2" class="muted">없음</td></tr>`
                }</tbody></table>
              </div>
              <h4 style="margin:12px 0 6px; font-size:14px;">품목 라인</h4>
              <div class="outbound-list-scroll" style="max-height:min(420px,50vh);">
                <table class="outbound-nested-table"><thead><tr>
                  <th>#</th><th>센터 통합 전표</th><th>원본 전표</th><th>상품코드</th><th>품명</th><th>확정수량</th><th>창고</th><th>LOT</th><th>주문수량</th><th>확인수량</th><th>미출</th>
                </tr></thead><tbody>${
                  lineRows || `<tr><td colspan="11" class="muted">없음</td></tr>`
                }</tbody></table>
              </div>`;
          } catch (e) {
            confirmListDetailBody.innerHTML = `<span class="muted">${esc(e.message || "조회 실패")}</span>`;
          }
        });
      });

      confirmListTableEl.querySelectorAll(`.outbound-confirmlist-cancel-${key}`).forEach((btn) => {
        btn.addEventListener("click", async () => {
          const listKey = String(btn.dataset.listKey || "").trim();
          if (!listKey) return;
          if (
            !confirm(
              "이 확정 묶음을 취소할까요? 포함된 전표는 출고확정이 취소되고, 출고분만큼 재고가 원복(조정 입고)됩니다. 판매입력 전송이 끝난 전표가 있으면 취소할 수 없습니다."
            )
          )
            return;
          try {
            await api("/api/outbound-order-upload/confirm-list/cancel", {
              method: "POST",
              body: JSON.stringify({ partnerType: key, listKey })
            });
            await refreshCommon();
            closeConfirmListDetail();
            await loadConfirmList();
            await loadList();
            await loadLines();
            alert("확정 취소 및 재고 원복을 반영했습니다.");
          } catch (e) {
            alert(e.message || "처리 실패");
          }
        });
      });

      confirmListTableEl.querySelectorAll(`.outbound-confirmlist-send-${key}`).forEach((btn) => {
        btn.addEventListener("click", async () => {
          const listKey = String(btn.dataset.listKey || "").trim();
          if (!listKey) return;
          try {
            await api("/api/outbound-order-upload/confirm-list/send-sales", {
              method: "POST",
              body: JSON.stringify({ partnerType: key, listKey })
            });
            await loadList();
            await loadConfirmList();
            alert("판매입력 전송을 반영했습니다.");
          } catch (e) {
            alert(e.message || "전송 실패");
          }
        });
      });

      if (confirmListStatusEl) confirmListStatusEl.textContent = `확정 저장 ${items.length}건`;
    } catch (e) {
      if (confirmListStatusEl) confirmListStatusEl.textContent = "불러오기 실패";
      confirmListTableEl.innerHTML = `<span class="muted">${esc(e.message || "오류")}</span>`;
    }
  };

  confirmListRefreshBtn?.addEventListener("click", async () => {
    await loadList();
    await loadConfirmList();
  });

  confirmAllBtn?.addEventListener("click", async () => {
    try {
      const slipMap = new Map(slipItems.map((x) => [normalizeSlipNoText(x.slipNo), x]));
      const selectedKeys = Array.from(selectedSlipKeys || []);
      const candidateKeys =
        selectedKeys.length > 0
          ? selectedKeys
          : slipItems.map((x) => normalizeSlipNoText(x.slipNo)).filter(Boolean);

      const targets = candidateKeys
        .map((k) => slipMap.get(k))
        .filter((x) => x && String(x.status || "draft") === "draft");

      if (!targets.length) return alert("확정할 전표가 없습니다. (이미 확정/전송된 전표만 있거나 선택이 없습니다)");

      const manager = prompt("출고 담당자명을 입력하세요");
      if (!manager) return;

      try {
        await api("/api/outbound-order-upload/workflow/bulk-confirm", {
          method: "POST",
          body: JSON.stringify({ slipNos: targets.map((t) => t.slipNo), stockUser: manager.trim() })
        });
      } catch (e) {
        alert(e.message || "전체 출고확정 실패(전표 단위로 오류가 나면 한 건도 확정되지 않습니다)");
        return;
      }

      await loadList();
      await loadLines();
      if (confirmListPanel && !confirmListPanel.classList.contains("hidden")) {
        await loadConfirmList();
      }
      alert(`전체 출고확정 완료: ${targets.length}건`);
    } catch (e) {
      alert(e.message || "전체 출고확정 실패");
    }
  });

  const loadUnshipLines = async () => {
    if (!unshipTableEl) return;
    try {
      if (unshipStatusEl) unshipStatusEl.textContent = "미출 라인 불러오는 중...";
      const res = await api(
        `/api/outbound-order-upload/unship-lines?partnerType=${encodeURIComponent(key)}&centerTab=${encodeURIComponent(unshipCenterTab)}`
      );
      const itemsRaw = Array.isArray(res.items) ? res.items : [];
      const qlow = String(unshipSearchQ || "").trim().toLowerCase();
      const items = !qlow
        ? itemsRaw
        : itemsRaw.filter((x) => {
            const hay = [
              x.slipNo,
              x.productCode,
              x.sourceProductCode,
              x.productName,
              x.storeName,
              x.centerName,
              x.centerCode,
              x.lot,
              x.warehouse
            ]
              .map((s) => String(s || "").toLowerCase())
              .join(" ");
            return hay.includes(qlow);
          });
      const counts = res.counts && typeof res.counts === "object" ? res.counts : { all: itemsRaw.length };

      if (unshipTabsEl) {
        const tabs = key === "emart" ? ["all", "여주", "대구", "시화"] : ["all"];
        unshipTabsEl.className = "seg-tabs";
        unshipTabsEl.innerHTML = tabs
          .map((t) => {
            const label = t === "all" ? "전체" : `${t}센터`;
            const n = Number(counts[t] ?? 0);
            return `<button type="button" class="${t === unshipCenterTab ? "seg-tab active" : "seg-tab"} outbound-unship-tab-${key}" data-tab="${esc(
              t
            )}">${label}(${n})</button>`;
          })
          .join("");
        unshipTabsEl.querySelectorAll(`.outbound-unship-tab-${key}`).forEach((btn) =>
          btn.addEventListener("click", async () => {
            unshipCenterTab = String(btn.getAttribute("data-tab") || "all");
            await loadUnshipLines();
          })
        );
      }

      if (unshipToolsEl && !qs(`#outbound-unship-search-${key}`)) {
        unshipToolsEl.innerHTML = `
          <span class="muted" id="outbound-unship-count-${key}" style="align-self:center;">표시 0건</span>
          <input type="text" id="outbound-unship-search-${key}" placeholder="전표·상품·점포·센터 검색" />
        `;
        const searchEl = qs(`#outbound-unship-search-${key}`);
        if (searchEl) {
          searchEl.value = unshipSearchQ;
          searchEl.addEventListener("input", async (e) => {
            unshipSearchQ = String(e.target?.value || "");
            await loadUnshipLines();
          });
        }
      }
      const unshipCountEl = qs(`#outbound-unship-count-${key}`);
      if (unshipCountEl) unshipCountEl.textContent = `현재 탭·검색 기준 ${items.length}건`;

      const rows = items
        .map((x, idx) => {
          const dueDisp = formatYmdLoose(x.dueDate) || formatYmdLoose(x.storeInDate) || "";
          const poDisp = formatYmdLoose(x.poDate) || formatYmdLoose(x.orderDate) || "";
          const storeInDisp = formatYmdLoose(x.storeInDate) || "";
          return `<tr>
            <td>${idx + 1}</td>
            <td>${esc(dueDisp)}</td>
            <td>${esc(x.centerName || "")}</td>
            <td>${esc(x.centerCode || "")}</td>
            <td><button type="button" class="bh-link-btn outbound-unship-slip-open-${key}" data-slip="${esc(x.slipNo || "")}">${esc(
            x.slipNo || ""
          )}</button></td>
            <td>${esc(x.storeName || "")}</td>
            <td>${esc(x.sourceProductCode || x.productCode || "")}</td>
            <td>${esc(x.productName || "")}</td>
            <td>${esc(String(x.orderQty ?? ""))}</td>
            <td>${esc(String(x.fixedQty ?? "0"))}</td>
            <td>${esc(storeInDisp)}</td>
            <td>${esc(poDisp)}</td>
            <td>${esc(x.lot || "")}</td>
            <td>${esc(x.warehouse || "")}</td>
          </tr>`;
        })
        .join("");

      unshipTableEl.innerHTML = `<div class="outbound-list-scroll" style="max-height:min(720px, calc(100vh - 220px));">
        <table id="outbound-unship-table-list-${key}">
          <thead><tr>
            <th>순번</th><th>납기일</th><th>센터명</th><th>센터코드</th><th>전표번호</th><th>점포명</th>
            <th>상품코드</th><th>상품명</th><th>주문수량</th><th>확정수량</th><th>점입점일자</th><th>발주일자</th><th>LOT</th><th>창고</th>
          </tr></thead>
          <tbody>${rows || `<tr><td colspan="14" class="muted">미출 라인이 없습니다.</td></tr>`}</tbody>
        </table>
      </div>`;
      delete TABLE_FILTER_MEMORY[`#outbound-unship-table-list-${key}`];
      applyExcelLikeFilter(`#outbound-unship-table-list-${key}`);
      unshipTableEl.querySelectorAll(`.outbound-unship-slip-open-${key}`).forEach((btn) =>
        btn.addEventListener("click", () => renderDetail(btn.getAttribute("data-slip"), { fullWorkflow: true }))
      );
      if (unshipStatusEl) unshipStatusEl.textContent = `미출 ${itemsRaw.length}건 (탭·검색 후 ${items.length}건 표시)`;
    } catch (e) {
      if (unshipStatusEl) unshipStatusEl.textContent = "불러오기 실패";
      unshipTableEl.innerHTML = `<span class="muted">${esc(e.message || "오류")}</span>`;
    }
  };

  const loadMaster = async () => {
    if (!masterTableEl) return;
    try {
      const res = await api(`/api/outbound-code-master?partnerType=${encodeURIComponent(key)}`);
      const items = Array.isArray(res.items) ? res.items : [];
      const rows = items
        .map(
          (x, idx) => {
            const fallbackName = (() => {
              const code = String(x.ecountCode || "").trim();
              if (!code) return "";
              const p = (state.products || []).find((it) => String(it.ecountCode || it.code || "").trim() === code);
              return p ? String(p.ecountName || p.name || "").trim() : "";
            })();
            return `<tr><td>${idx + 1}</td><td>${esc(x.sourceCode || "")}</td><td>${esc(x.ecountCode || "")}</td><td>${esc(
              x.name || fallbackName || "-"
          )}</td><td>${esc(formatDateTimeDisplay(x.updatedAt || ""))}</td></tr>`
          }
        )
        .join("");
      masterTableEl.innerHTML = `<div class="outbound-list-scroll" style="max-height:690px;">
        <table id="outbound-master-table-list-${key}">
          <thead><tr><th>순번</th><th>바코드 SKU</th><th>이카운트코드</th><th>품명 및 규격(이카운트)</th><th>수정일시</th></tr></thead>
          <tbody>${rows || `<tr><td colspan="5" class="muted">등록된 코드마스터 없음</td></tr>`}</tbody>
        </table>
      </div>`;
      applyExcelLikeFilter(`#outbound-master-table-list-${key}`);
      if (masterStatusEl) masterStatusEl.textContent = `코드마스터 ${items.length}건`;
    } catch (e) {
      if (masterStatusEl) masterStatusEl.textContent = e.message || "코드마스터 조회 실패";
      masterTableEl.innerHTML = `<span class="muted">${esc(e.message || "오류")}</span>`;
    }
  };

  const saveOutboundLineUpdates = async (updates) => {
    const payload = (Array.isArray(updates) ? updates : []).filter(
      (x) => x && String(x.uploadBatchId || "").trim() && Number.isFinite(Number(x.sourceRowNo))
    );
    if (!payload.length) return;
    await api("/api/outbound-order-upload/lines/update", {
      method: "POST",
      body: JSON.stringify({ partnerType: key, updates: payload })
    });
  };

  const loadLines = async () => {
    if (!lineTableEl) return;
    try {
      const res = await api(
        `/api/outbound-order-upload/lines?partnerType=${encodeURIComponent(key)}&centerTab=${encodeURIComponent(centerTab)}`
      );
      const itemsRaw = Array.isArray(res.items) ? res.items : [];
      const q = String(lineSearchQ || "").trim().toLowerCase();
      const items = !q
        ? itemsRaw
        : itemsRaw.filter((x) => {
            const code = String(x.sourceProductCode || x.productCode || "").toLowerCase();
            const name = String(x.productName || "").toLowerCase();
            return code.includes(q) || name.includes(q);
          });
      const counts = res.counts && typeof res.counts === "object" ? res.counts : { all: items.length };
      selectedLineKeys = new Set(
        Array.from(selectedLineKeys).filter((k) =>
          items.some((x) => `${String(x.uploadBatchId || "")}|${String(x.sourceRowNo || "")}` === k)
        )
      );
      if (lineTabsEl) {
        const tabs = key === "emart" ? ["all", "여주", "대구", "시화"] : ["all"];
        lineTabsEl.className = "seg-tabs";
        lineTabsEl.innerHTML = tabs
          .map((t) => {
            const label = t === "all" ? "전체" : t;
            const n = Number(counts[t] ?? 0);
            return `<button type="button" class="${t === centerTab ? "seg-tab active" : "seg-tab"} outbound-line-tab-${key}" data-tab="${esc(
              t
            )}">${label}(${n})</button>`;
          })
          .join("");
        lineTabsEl.querySelectorAll(`.outbound-line-tab-${key}`).forEach((btn) =>
          btn.addEventListener("click", async () => {
            centerTab = String(btn.getAttribute("data-tab") || "all");
            await loadLines();
          })
        );
      }
      if (lineToolsEl && !qs(`#outbound-line-search-${key}`)) {
        lineToolsEl.innerHTML = `
          <span class="muted" id="outbound-line-count-${key}" style="align-self:center;">현재 탭 기준 ${items.length}건</span>
          <input type="text" id="outbound-line-search-${key}" placeholder="선택 탭 내 상품 검색" />
          <select id="outbound-line-bulk-status-${key}">
            <option value="">일괄 출고상태 변경</option>
            <option value="정상">정상</option>
            <option value="미출">미출</option>
          </select>
          <button type="button" class="cancel-btn" id="outbound-line-bulk-apply-${key}">일괄 적용</button>
          <button type="button" class="primary" id="outbound-line-confirm-all-${key}">전체 출고 확정</button>
        `;
        const createdSearchEl = qs(`#outbound-line-search-${key}`);
        if (createdSearchEl) {
          createdSearchEl.value = lineSearchQ;
          createdSearchEl.addEventListener("input", async (e) => {
            lineSearchQ = String(e.target?.value || "");
            await loadLines();
          });
        }
        qs(`#outbound-line-bulk-apply-${key}`)?.addEventListener("click", async () => {
          const bulkStatus = String(qs(`#outbound-line-bulk-status-${key}`)?.value || "");
          if (!bulkStatus) {
            alert("일괄 출고상태를 선택하세요.");
            return;
          }
          const selectedRows = lineTableEl.querySelectorAll(`.outbound-line-check-${key}:checked`);
          if (!selectedRows.length) {
            alert("체크된 항목이 없습니다.");
            return;
          }
          const updates = [];
          selectedRows.forEach((chk) => {
            const tr = chk.closest("tr");
            const sel = tr?.querySelector(`.outbound-status-select-${key}`);
            if (!sel) return;
            sel.value = bulkStatus;
            const row = sel.getAttribute("data-row");
            const orderQty = Number(sel.getAttribute("data-order-qty") || "0");
            const fixedEl = lineTableEl.querySelector(`.outbound-fixed-qty-${key}[data-row="${row}"]`);
            if (fixedEl) {
              if (bulkStatus === "미출") {
                fixedEl.value = "0";
                tr?.classList.add("outbound-unship-row");
              } else {
                fixedEl.value = String(Number.isFinite(orderQty) ? Math.max(0, Math.round(orderQty)) : 0);
                tr?.classList.remove("outbound-unship-row");
              }
            }
            const uploadBatchId = String(sel.getAttribute("data-upload-batch-id") || "").trim();
            const sourceRowNo = Number(sel.getAttribute("data-source-row-no") || "0");
            const fixedQty = Number(fixedEl?.value || "0");
            if (uploadBatchId && Number.isFinite(sourceRowNo) && sourceRowNo > 0) {
              updates.push({ uploadBatchId, sourceRowNo, unshipStatus: bulkStatus, fixedQty });
            }
          });
          try {
            await saveOutboundLineUpdates(updates);
            selectedLineKeys = new Set();
            const lineChecks = Array.from(lineTableEl.querySelectorAll(`.outbound-line-check-${key}`));
            lineChecks.forEach((el) => {
              el.checked = false;
            });
            const lineCheckAll = qs(`#outbound-line-check-all-${key}`);
            if (lineCheckAll) lineCheckAll.checked = false;
            await loadLines();
          } catch (e) {
            alert(e.message || "일괄 적용 저장 실패");
          }
        });
        qs(`#outbound-line-confirm-all-${key}`)?.addEventListener("click", () => void runOutboundBulkConfirmAll());
      }
      const lineCountEl = qs(`#outbound-line-count-${key}`);
      if (lineCountEl) lineCountEl.textContent = `현재 탭 기준 ${items.length}건`;
      const centerSeqMap = {};
      const rows = items
        .map((x, idx) => {
          const rowKey = `${String(x.uploadBatchId || "")}|${String(x.sourceRowNo || "")}`;
          const centerKey = String(x.centerName || "").trim() || "-";
          centerSeqMap[centerKey] = (centerSeqMap[centerKey] || 0) + 1;
          const centerSeq = centerTab === "all" ? idx + 1 : centerSeqMap[centerKey];
          const orderQtyNum = Number(String(x.orderQty || "").replace(/,/g, ""));
          const orderQty = Number.isFinite(orderQtyNum) ? Math.max(0, Math.round(orderQtyNum)) : 0;
          const initialStatus = String(x.unshipStatus || "").trim() === "미출" ? "미출" : "정상";
          const initialFixedRaw =
            x.fixedQty != null && String(x.fixedQty).trim() !== "" ? Number(String(x.fixedQty).replace(/,/g, "")) : orderQty;
          const initialFixed = initialStatus === "미출" ? 0 : Number.isFinite(initialFixedRaw) ? Math.max(0, Math.round(initialFixedRaw)) : orderQty;
          return `<tr class="${initialStatus === "미출" ? "outbound-unship-row" : ""}" data-line-key="${esc(rowKey)}">
            <td><input type="checkbox" class="outbound-line-check outbound-line-check-${key}" data-line-key="${esc(
              rowKey
            )}" ${selectedLineKeys.has(rowKey) ? "checked" : ""} /></td>
            <td>${centerSeq}</td><td>${esc(x.storeInDate || x.dueDate || "")}</td><td>${esc(x.poDate || x.orderDate || "")}</td>
            <td>${esc(x.storeName || "")}</td><td>${esc(x.sourceProductCode || x.productCode || "")}</td><td>${esc(x.productName || "")}</td>
            <td>${esc(x.orderUnit || "")}</td><td>${esc(x.lot || "")}</td><td>${esc(String(x.orderQty || ""))}</td>
            <td>
              <select class="outbound-status-select-${key}" data-row="${idx}" data-order-qty="${orderQty}" data-upload-batch-id="${esc(
                x.uploadBatchId || ""
              )}" data-source-row-no="${esc(String(x.sourceRowNo || ""))}">
                <option value="정상" ${initialStatus === "정상" ? "selected" : ""}>정상</option>
                <option value="미출" ${initialStatus === "미출" ? "selected" : ""}>미출</option>
              </select>
            </td>
            <td><input type="number" min="0" step="1" class="outbound-fixed-qty-${key}" data-row="${idx}" value="${esc(
              String(initialFixed)
            )}" style="width:90px;" readonly /></td><td>${esc(x.orderAmount || "")}</td>
            <td>${esc(x.centerInDate || "")}</td><td>${esc(x.centerCode || "")}</td><td>${esc(x.centerName || "")}</td>
          </tr>`;
        })
        .join("");
      lineTableEl.innerHTML = `<div class="outbound-list-scroll"><table id="outbound-line-table-list-${key}">
        <thead><tr><th><input type="checkbox" class="outbound-line-check-all" id="outbound-line-check-all-${key}" /></th><th>순번</th><th>점입점일자</th><th>발주일자</th><th>점포명</th><th>상품코드</th><th>상품명</th><th>발주단위</th><th>LOT</th><th>수량</th><th>출고상태</th><th>확정수량</th><th>발주금액</th><th>센터입하일자</th><th>센터코드</th><th>센터이름</th></tr></thead>
        <tbody>${rows || `<tr><td colspan="16" class="muted">업로드된 라인이 없습니다.</td></tr>`}</tbody>
      </table></div>`;
      delete TABLE_FILTER_MEMORY[`#outbound-line-table-list-${key}`];
      applyExcelLikeFilter(`#outbound-line-table-list-${key}`);
      const lineChecks = Array.from(lineTableEl.querySelectorAll(`.outbound-line-check-${key}`));
      const lineCheckAll = qs(`#outbound-line-check-all-${key}`);
      if (lineCheckAll) {
        lineCheckAll.checked = lineChecks.length > 0 && lineChecks.every((x) => x.checked);
        lineCheckAll.addEventListener("change", () => {
          selectedLineKeys = new Set();
          lineChecks.forEach((el) => {
            el.checked = lineCheckAll.checked;
            const rowKey = String(el.getAttribute("data-line-key") || "");
            if (lineCheckAll.checked && rowKey) selectedLineKeys.add(rowKey);
          });
        });
      }
      lineChecks.forEach((el) =>
        el.addEventListener("change", () => {
          const rowKey = String(el.getAttribute("data-line-key") || "");
          if (!rowKey) return;
          if (el.checked) selectedLineKeys.add(rowKey);
          else selectedLineKeys.delete(rowKey);
          if (lineCheckAll) lineCheckAll.checked = lineChecks.length > 0 && lineChecks.every((x) => x.checked);
        })
      );
      lineTableEl.querySelectorAll(`.outbound-status-select-${key}`).forEach((sel) => {
        sel.addEventListener("change", async () => {
          const row = sel.getAttribute("data-row");
          const orderQty = Number(sel.getAttribute("data-order-qty") || "0");
          const fixedEl = lineTableEl.querySelector(`.outbound-fixed-qty-${key}[data-row="${row}"]`);
          const tr = sel.closest("tr");
          if (!fixedEl) return;
          if (sel.value === "미출") {
            fixedEl.value = "0";
            tr?.classList.add("outbound-unship-row");
          } else {
            fixedEl.value = String(Number.isFinite(orderQty) ? Math.max(0, Math.round(orderQty)) : 0);
            tr?.classList.remove("outbound-unship-row");
          }
          try {
            const uploadBatchId = String(sel.getAttribute("data-upload-batch-id") || "").trim();
            const sourceRowNo = Number(sel.getAttribute("data-source-row-no") || "0");
            const fixedQty = Number(fixedEl.value || "0");
            await saveOutboundLineUpdates([
              {
                uploadBatchId,
                sourceRowNo,
                unshipStatus: sel.value === "미출" ? "미출" : "정상",
                fixedQty
              }
            ]);
          } catch (e) {
            alert(e.message || "출고상태 저장 실패");
          }
        });
      });
      if (unshipPanel && !unshipPanel.classList.contains("hidden")) {
        await loadUnshipLines();
      }
    } catch (e) {
      lineTableEl.innerHTML = `<span class="muted">${esc(e.message || "라인 조회 실패")}</span>`;
    }
  };

  const renderDetail = async (slipNo, opts = {}) => {
    if (!detailBody || !slipNo) return;
    const fullWorkflow = opts.fullWorkflow === true;
    overlay?.classList.remove("hidden");
    detailBody.innerHTML = `<p class="muted">불러오는 중...</p>`;
    if (detailTitle) detailTitle.textContent = `전표 상세 - ${slipNo}`;
    try {
      const res = await api(`/api/outbound-order-upload/detail?slipNo=${encodeURIComponent(slipNo)}`);
      const items = Array.isArray(res.items) ? res.items : [];
      const wf = res.workflow || { status: "draft", lineQtyByIndex: {} };
      const confirmed = wf.status === "confirmed" || wf.status === "sent";
      const sent = wf.status === "sent";
      const rows = items
        .map(
          (x, idx) => `<tr>
            <td>${idx + 1}</td><td>${esc(x.productCode || "")}</td><td>${esc(x.productName || "")}</td>
            <td>${esc(String(x.orderQty || 0))}</td>
            <td><input type="number" min="0" step="1" class="outbound-confirm-qty-${key}" data-idx="${idx}" value="${esc(String(
              x.confirmQty ?? x.orderQty ?? ""
            ))}" ${confirmed ? "disabled" : ""} style="width:90px;" /></td>
            <td>${esc(x.warehouse || "")}</td>
            <td>${esc(x.remark || "")}</td>
          </tr>`
        )
        .join("");
      const wfBtns = fullWorkflow
        ? `<button type="button" class="primary" id="outbound-confirm-${key}" ${confirmed ? "disabled" : ""}>출고확정</button>
          <button type="button" class="primary" id="outbound-send-${key}" ${confirmed && !sent ? "" : "disabled"}>판매입력 전송</button>`
        : "";
      detailBody.innerHTML = `
        <div class="row" style="display:flex; gap:8px; align-items:center; margin-bottom:8px; flex-wrap:wrap;">
          <span><strong>상태:</strong> ${esc(wf.status || "draft")}</span>
          <button type="button" class="cancel-btn" id="outbound-save-${key}" ${confirmed ? "disabled" : ""}>중간저장</button>
          ${wfBtns}
          <button type="button" class="cancel-btn" id="outbound-reopen-${key}" ${confirmed && !sent ? "" : "disabled"}>확정취소</button>
        </div>
        <div id="outbound-detail-msg-${key}" class="muted" style="margin-bottom:8px;"></div>
        <table><thead><tr><th>순번</th><th>상품코드</th><th>상품명</th><th>수량(주문)</th><th>수량(출고확인)</th><th>창고</th><th>비고</th></tr></thead>
        <tbody>${rows || `<tr><td colspan="7" class="muted">품목 없음</td></tr>`}</tbody></table>
      `;
      const msgEl = qs(`#outbound-detail-msg-${key}`);
      const setMsg = (t, err) => {
        if (!msgEl) return;
        msgEl.textContent = t || "";
        msgEl.style.color = err ? "#b42318" : "";
      };
      const qtyMap = () => {
        const out = {};
        detailBody.querySelectorAll(`.outbound-confirm-qty-${key}`).forEach((el) => {
          const i = el.getAttribute("data-idx");
          out[i] = el.value;
        });
        return out;
      };
      qs(`#outbound-save-${key}`)?.addEventListener("click", async () => {
        try {
          await api("/api/outbound-order-upload/workflow", {
            method: "POST",
            body: JSON.stringify({ slipNo, action: "save-check", lineQtyByIndex: qtyMap() })
          });
          setMsg("중간저장 완료");
          await loadList();
          await renderDetail(slipNo, { fullWorkflow });
        } catch (e) {
          setMsg(e.message || "오류", true);
        }
      });
      if (fullWorkflow) {
        qs(`#outbound-confirm-${key}`)?.addEventListener("click", async () => {
          try {
            const manager = prompt("출고 담당자명을 입력하세요");
            if (!manager) return;
            await api("/api/outbound-order-upload/workflow", {
              method: "POST",
              body: JSON.stringify({ slipNo, action: "confirm", stockUser: manager.trim(), lineQtyByIndex: qtyMap() })
            });
            setMsg("출고확정 완료(재고 반영)");
            await loadList();
            if (confirmListPanel && !confirmListPanel.classList.contains("hidden")) await loadConfirmList();
            await renderDetail(slipNo, { fullWorkflow });
          } catch (e) {
            setMsg(e.message || "오류", true);
          }
        });
      }
      qs(`#outbound-reopen-${key}`)?.addEventListener("click", async () => {
        try {
          const ok = confirm("출고확정을 취소하고 재고를 원복할까요?");
          if (!ok) return;
          await api("/api/outbound-order-upload/workflow", {
            method: "POST",
            body: JSON.stringify({ slipNo, action: "reopen" })
          });
          alert("확정취소 완료");
          await loadList();
          if (confirmListPanel && !confirmListPanel.classList.contains("hidden")) await loadConfirmList();
          await renderDetail(slipNo, { fullWorkflow });
        } catch (e) {
          setMsg(e.message || "오류", true);
        }
      });
      if (fullWorkflow) {
        qs(`#outbound-send-${key}`)?.addEventListener("click", async () => {
          try {
            await api("/api/outbound-order-upload/workflow", {
              method: "POST",
              body: JSON.stringify({ slipNo, action: "send-sales" })
            });
            setMsg("판매입력 전송 상태를 기록했습니다.");
            await loadList();
            if (confirmListPanel && !confirmListPanel.classList.contains("hidden")) await loadConfirmList();
            await renderDetail(slipNo, { fullWorkflow });
          } catch (e) {
            setMsg(e.message || "오류", true);
          }
        });
      }
    } catch (e) {
      detailBody.innerHTML = `<p class="muted">${esc(e.message || "오류")}</p>`;
    }
  };

  qs(`#outbound-upload-detail-close-${key}`)?.addEventListener("click", () => overlay?.classList.add("hidden"));
  overlay?.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.classList.add("hidden");
  });

  qs(`#outbound-upload-run-${key}`)?.addEventListener("click", async () => {
    const file = qs(`#outbound-upload-file-${key}`)?.files?.[0];
    if (!file) return alert("파일을 선택하세요.");
    try {
      if (statusEl) statusEl.textContent = "파일 파싱 중...";
      const matrix = await parseSheetMatrix(file);
      const res = await api("/api/outbound-order-upload", {
        method: "POST",
        body: JSON.stringify({ partnerType: key, sourceFileName: file.name, matrix })
      });
      const errCount = Array.isArray(res.errorRows) ? res.errorRows.length : 0;
      if (statusEl) statusEl.textContent = `업로드 완료 (성공 ${res.okCount || 0}, 실패 ${errCount}) - 업로드 파일 리스트에서 적용 버튼을 눌러 반영하세요`;
      if (errCount) alert(`실패 행 ${errCount}건이 있습니다. (응답 errorRows 확인)`);
      await loadList();
      await loadLines();
    } catch (e) {
      if (statusEl) statusEl.textContent = e.message || "업로드 실패";
    }
  });
  qs(`#outbound-master-run-${key}`)?.addEventListener("click", async () => {
    const file = qs(`#outbound-master-file-${key}`)?.files?.[0];
    if (!file) return alert("코드마스터 파일을 선택하세요.");
    try {
      if (masterStatusEl) masterStatusEl.textContent = "코드마스터 업로드 중...";
      const matrix = await parseSheetMatrix(file);
      const res = await api("/api/outbound-code-master/upload", {
        method: "POST",
        body: JSON.stringify({ partnerType: key, sourceFileName: file.name, matrix })
      });
      const errCount = Array.isArray(res.errorRows) ? res.errorRows.length : 0;
      if (masterStatusEl) masterStatusEl.textContent = `코드마스터 업로드 완료 (${res.count || 0}건)`;
      if (errCount) alert(`코드마스터 실패 행 ${errCount}건`);
      await loadMaster();
    } catch (e) {
      if (masterStatusEl) masterStatusEl.textContent = e.message || "코드마스터 업로드 실패";
    }
  });
  qs(`#outbound-upload-refresh-${key}`)?.addEventListener("click", () => loadList());
  qs(`#outbound-upload-batch-open-${key}`)?.addEventListener("click", async () => {
    await loadList();
    batchOverlay?.classList.remove("hidden");
  });
  qs(`#outbound-batch-overlay-close-${key}`)?.addEventListener("click", () => batchOverlay?.classList.add("hidden"));
  batchOverlay?.addEventListener("click", (e) => {
    if (e.target === batchOverlay) batchOverlay.classList.add("hidden");
  });
  qs(`#outbound-batch-delete-selected-${key}`)?.addEventListener("click", async () => {
    try {
      const uploadBatchIds = Array.from(selectedBatchIds).filter(Boolean);
      if (!uploadBatchIds.length) return alert("삭제할 업로드 이력을 선택하세요.");
      const ok = confirm(`선택 업로드 ${uploadBatchIds.length}건을 삭제할까요?`);
      if (!ok) return;
      await api("/api/outbound-order-upload/delete-batches", {
        method: "POST",
        body: JSON.stringify({ partnerType: key, uploadBatchIds })
      });
      selectedBatchIds = new Set();
      await loadList();
      await loadLines();
      alert("선택 업로드를 삭제했습니다.");
    } catch (e) {
      alert(e.message || "삭제 실패");
    }
  });
  qs(`#outbound-batch-delete-all-${key}`)?.addEventListener("click", async () => {
    try {
      const ok = confirm("해당 판매처의 주문서 업로드 데이터를 전체 삭제할까요?");
      if (!ok) return;
      await api("/api/outbound-order-upload/delete", {
        method: "POST",
        body: JSON.stringify({ partnerType: key, all: true })
      });
      selectedBatchIds = new Set();
      await loadList();
      await loadLines();
      alert("주문서 업로드 데이터를 전체 삭제했습니다.");
    } catch (e) {
      alert(e.message || "삭제 실패");
    }
  });
  qs(`#outbound-upload-template-${key}`)?.addEventListener("click", () => {
    window.location.href = `/api/outbound-order-upload/template?partnerType=${encodeURIComponent(key)}`;
  });
  await loadList();
  await loadMaster();
  await loadLines();
}

async function renderRecentMovements(type) {
  const box = qs(`#recent-${type}`);
  if (!box) return;
  box.innerHTML = "최근 내역을 불러오는 중...";

  const partnerLabel = type === "IN" ? "구매처" : "출고처";
  const recentLimit = 120;
  try {
    const res = await api(`/api/recent?type=${type}&limit=${recentLimit}`);
    const items = res.items || [];
    if (!items.length) {
      box.innerHTML = `<span class="muted">등록 내역이 없습니다.</span>`;
      return;
    }

    const displayItems =
      type === "OUT"
        ? (() => {
            const grouped = new Map();
            for (const x of items) {
              const slipKey = String(x.slipNoDisplay || x.slipNo || "").trim() || "NOSLIP";
              const productKey = String(x.productCode || "").trim() || "NOPROD";
              const gk = `${slipKey}|${productKey}`;
              if (!grouped.has(gk)) {
                grouped.set(gk, {
                  ...x,
                  _qtySum: 0,
                  _lineCount: 0
                });
              }
              const g = grouped.get(gk);
              g._qtySum += Math.abs(Number(x.qty || 0));
              g._lineCount += 1;
              if (new Date(x.createdAt || 0).getTime() > new Date(g.createdAt || 0).getTime()) {
                g.createdAt = x.createdAt;
              }
            }
            return Array.from(grouped.values()).sort((a, b) => new Date(b.createdAt || 0) - new Date(a.createdAt || 0));
          })()
        : items;

    const rows = displayItems
      .map((x) => {
        const value = type === "IN" ? x.partner : x.partner;
        const idDisp = x.id != null && Number.isFinite(Number(x.id)) ? x.id : "—";
        const idCell = type === "IN" ? "" : `<td>${idDisp}</td>`;
        const slipCell =
          type === "IN"
            ? `<td>${esc(x.slipNo || "")}${x.slipCancelled ? ` <span class="badge gray">취소됨</span>` : ""}</td>`
            : `<td>${esc(x.slipNoDisplay || x.slipNo || "")}</td>`;
        const canCancel =
          !x.slipConfirm &&
          x.id != null &&
          Number.isFinite(Number(x.id)) &&
          String(x.user || "").trim() !== "" &&
          !String(x.slipNo || "").trim();
        const cancelCell = canCancel
          ? `<button type="button" class="cancel-btn del-small recent-cancel-btn" data-id="${x.id}" data-user="${esc(x.user)}">취소</button>`
          : `<span class="muted">—</span>`;
        return `<tr>
          ${idCell}
          ${slipCell}
          <td>${esc(x.productCode)}</td>
          <td>${esc(x.ecountCode || "")}</td>
          <td>${esc(x.productName || "")}</td>
          <td>${
            x.type === "OUT"
              ? esc(String(x.qtyDisplay ?? (Number.isFinite(Number(x._qtySum)) ? Number(x._qtySum) : Math.abs(Number(x.qty || 0)))))
              : x.qty
          }</td>
          <td>${esc(x.warehouse || "")}</td>
          <td>${esc(value || "")}</td>
          <td>${esc(x.user)}</td>
          <td>${
            x.type === "OUT" && Number(x._lineCount || 0) > 1
              ? `${esc(x.memo || "")}${x.memo ? " / " : ""}센터 통합 합산(${Number(x._lineCount)}건)`
              : esc(x.memo || "")
          }</td>
          <td>${esc(formatDateTimeDisplay(x.createdAt))}</td>
          <td>${cancelCell}</td>
        </tr>`;
      })
      .join("");

    const idHead = type === "IN" ? "" : `<th>이력번호</th>`;
    const slipHead = `<th>전표번호</th>`;
    const timeHead = type === "IN" ? "입고일시" : "일시";
    const tableHtml = `
      <table id="recent-${type}-table">
        <thead>
          <tr>
            ${idHead}
            ${slipHead}
            <th>상품코드</th>
            <th>품목코드</th>
            <th>상품명</th>
            <th>수량</th>
            <th>창고</th>
            <th>${partnerLabel}</th>
            <th>담당</th>
            <th>메모</th>
            <th>${timeHead}</th>
            <th>액션</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
    `;
    box.innerHTML = `<div class="recent-movements-wrap"><div class="recent-movements-scroll">${tableHtml}</div></div>`;

    box.querySelectorAll(".recent-cancel-btn").forEach((btn) => {
      btn.onclick = async () => {
        const id = Number(btn.dataset.id);
        const manager = String(btn.dataset.user || "").trim();
        if (!Number.isFinite(id)) return;
        if (!manager) return alert("등록자 정보가 없어 취소할 수 없습니다.");
        try {
          await api("/api/movements/cancel", {
            method: "POST",
            body: JSON.stringify({ id, user: manager })
          });
          await afterMovementDone();
        } catch (err) {
          alert(err.message);
        }
      };
    });

    applyExcelLikeFilter(`#recent-${type}-table`);
  } catch (e) {
    box.innerHTML = `<span class="muted">최근 내역을 불러오지 못했습니다.</span>`;
  }
}

function renderAdjust() {
  const productOptions = state.products.map((p) => `<option value="${esc(p.code)}">${esc(p.code)} - ${esc(p.name)}</option>`).join("");
  const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const warehouseOptions = (state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");
  qs("#view-adjust").innerHTML = `
    <div class="card">
      <h2>재고 조정 / 이동</h2>
      <form id="ADJUST-form">
        <div><label>상품 검색(코드)</label><input name="productCode" required /></div>
        <div><label>창고</label><input name="warehouse" list="ADJUST-warehouse-list" required /></div>
        <datalist id="ADJUST-warehouse-list">${warehouseOptions}</datalist>
        <div><label>조정수량(+/-)</label><input name="qty" type="number" required /></div>
        <div><label>담당자</label><input name="user" list="ADJUST-manager-list" required /></div>
        <datalist id="ADJUST-manager-list">${managerOptions}</datalist>
        <div><label>메모</label><input name="memo" placeholder="파손/오차 등" /></div>
        <div><button id="ADJUST-submit" class="primary" type="button">조정 등록</button></div>
      </form>
    </div>
    <div class="card">
      <h3>창고 간 이동</h3>
      <form id="TRANSFER-form">
        <div><label>상품 검색(코드)</label><input name="productCode" required /></div>
        <div><label>출발창고</label><input name="warehouse" list="TRANSFER-from-list" required /></div>
        <datalist id="TRANSFER-from-list">${warehouseOptions}</datalist>
        <div><label>도착창고</label><input name="toWarehouse" list="TRANSFER-to-list" required /></div>
        <datalist id="TRANSFER-to-list">${warehouseOptions}</datalist>
        <div><label>이동수량</label><input name="qty" type="number" required /></div>
        <div><label>담당자</label><input name="user" list="TRANSFER-manager-list" required /></div>
        <datalist id="TRANSFER-manager-list">${managerOptions}</datalist>
        <div><label>메모</label><input name="memo" placeholder="창고 이동" /></div>
        <div><button id="TRANSFER-submit" class="primary" type="button">이동 등록</button></div>
      </form>
    </div>
  `;
  const form = qs("#ADJUST-form");
  preventEnterSubmit(form);
  qs("#ADJUST-submit").onclick = async () => {
    try {
      const data = Object.fromEntries(new FormData(form));
      await api("/api/movements", { method: "POST", body: JSON.stringify({ ...data, type: "ADJUST" }) });
      await afterMovementDone();
      alert("조정 등록 완료");
    } catch (err) {
      alert(err.message);
    }
  };
  const tForm = qs("#TRANSFER-form");
  preventEnterSubmit(tForm);
  qs("#TRANSFER-submit").onclick = async () => {
    try {
      const data = Object.fromEntries(new FormData(tForm));
      await api("/api/movements", { method: "POST", body: JSON.stringify({ ...data, type: "TRANSFER" }) });
      await afterMovementDone();
      alert("이동 등록 완료");
    } catch (err) {
      alert(err.message);
    }
  };
}

function typeBadge(type, cancelled, memo, originType) {
  if (cancelled) return `<span class="badge gray">취소됨</span>`;
  if (type === "IN") return `<span class="badge green">입고</span>`;
  if (type === "OUT") return `<span class="badge red">출고</span>`;
  if (type === "ADJUST") {
    const isInboundReopenRollback = String(memo || "").includes("입고확정 취소 원복");
    if (isInboundReopenRollback) return `<span class="badge blue">입고확정취소</span>`;
    return `<span class="badge blue">재고조정</span>`;
  }
  if (type === "TRANSFER") return `<span class="badge blue">이동</span>`;
  if (type === "CANCEL") {
    const originLabel =
      originType === "IN"
        ? "입고"
        : originType === "OUT"
          ? "출고"
          : originType === "ADJUST"
            ? "재고조정"
            : originType === "TRANSFER"
              ? "이동"
              : "";
    return `<span class="badge gray">취소반영${originLabel ? `(${originLabel})` : ""}</span>`;
  }
  return `<span class="badge gray">${esc(type)}</span>`;
}

async function renderHistory() {
  const wrap = qs("#view-history");
  const today = localDateYmd(new Date());
  const managerOptions = state.managers.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`).join("");
  const warehouseOptions = (state.warehouses || []).map((w) => `<option value="${esc(w)}">${esc(w)}</option>`).join("");

  wrap.innerHTML = `
    <div class="card history-search-card">
      <div class="history-search-header">
        <h2>입출고 이력</h2>
        <p class="muted">입고·출고·재고조정·이동 내역을 조회합니다.</p>
      </div>
      <div class="history-filter-row">
        <div class="history-search-input-wrap">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
          <input id="history-q" class="history-q-input" placeholder="상품코드 · 상품명 검색" />
        </div>
        <select id="history-type-filter">
          <option value="ALL">구분 전체</option>
          <option value="IN">입고</option>
          <option value="OUT">출고</option>
          <option value="ADJUST">재고조정</option>
          <option value="TRANSFER">이동</option>
          <option value="CANCEL">취소반영</option>
        </select>
        <select id="history-warehouse-filter">
          <option value="">창고 전체</option>
          ${warehouseOptions}
        </select>
        <select id="history-manager-filter">
          <option value="">담당자 전체</option>
          ${managerOptions}
        </select>
      </div>
      <div class="history-filter-row">
        <div class="history-date-wrap">
          <input type="date" id="history-date-from" value="${today}" />
          <span class="history-date-sep">~</span>
          <input type="date" id="history-date-to" value="${today}" />
        </div>
        <div class="history-quick-dates">
          <button type="button" class="bh-btn bh-btn-sm history-quick-btn" data-range="today">오늘</button>
          <button type="button" class="bh-btn bh-btn-sm history-quick-btn" data-range="week">이번 주</button>
          <button type="button" class="bh-btn bh-btn-sm history-quick-btn" data-range="month">이번 달</button>
          <button type="button" class="bh-btn bh-btn-sm history-quick-btn" data-range="all">전체</button>
        </div>
        <div class="history-search-actions">
          <button type="button" class="bh-btn" id="history-reset">초기화</button>
          <button type="button" class="bh-btn bh-btn-primary" id="history-search">조회</button>
        </div>
      </div>
    </div>
    <div class="card" id="history-result-card" style="display:none;">
      <div class="history-result-header">
        <span class="history-result-count">조회 결과 <strong id="history-count">0</strong>건</span>
        <span class="muted" id="history-stock-info" style="font-size:12px;"></span>
        <button type="button" class="bh-btn bh-btn-sm" id="history-excel-btn" style="margin-left:auto;">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          엑셀 다운로드
        </button>
      </div>
      <div id="history-table">
        <p class="muted" style="text-align:center; padding:32px 0;">조회 버튼을 눌러 이력을 불러오세요.</p>
      </div>
    </div>
    <div class="card" id="history-empty-card">
      <p class="muted" style="text-align:center; padding:32px 0;">조회 버튼을 눌러 이력을 불러오세요.</p>
    </div>
  `;

  let lastDisplayItems = [];

  async function draw() {
    const qTrim = (qs("#history-q")?.value || "").trim();
    const typeFilter = qs("#history-type-filter")?.value || "ALL";
    const warehouseFilter = qs("#history-warehouse-filter")?.value || "";
    const managerFilter = qs("#history-manager-filter")?.value || "";
    const dateFrom = qs("#history-date-from")?.value || "";
    const dateTo = qs("#history-date-to")?.value || "";

    qs("#history-table").innerHTML = `<p class="muted" style="text-align:center; padding:32px 0;">조회 중...</p>`;
    qs("#history-result-card").style.display = "block";
    qs("#history-empty-card").style.display = "none";

    const [res, stockRes] = await Promise.all([
      api(`/api/history?q=${encodeURIComponent(qTrim)}`),
      api("/api/stock")
    ]);

    const matched = qTrim
      ? stockRes.items.filter((s) => String(s.code).includes(qTrim) || String(s.name).includes(qTrim)).map((s) => `${s.code}: ${s.stock}`)
      : [];
    const matchedStocks = matched.length ? matched.join(" / ") : "";

    let items = res.items;
    if (typeFilter !== "ALL") items = items.filter((x) => x.type === typeFilter);
    if (warehouseFilter) items = items.filter((x) => x.warehouse === warehouseFilter || x.toWarehouse === warehouseFilter);
    if (managerFilter) items = items.filter((x) => x.user === managerFilter);
    if (dateFrom || dateTo) {
      items = items.filter((x) => {
        const d = (x.createdAt || "").slice(0, 10);
        if (dateFrom && d < dateFrom) return false;
        if (dateTo && d > dateTo) return false;
        return true;
      });
    }

    const displayItems = (() => {
      const outGrouped = new Map();
      const out = [];
      for (const x of items) {
        if (x.type !== "OUT") { out.push(x); continue; }
        const slipKey = String(x.slipNoDisplay || x.slipNo || "").trim();
        const productKey = String(x.productCode || "").trim();
        const gk = `${slipKey}|${productKey}`;
        if (!gk) { out.push(x); continue; }
        if (!outGrouped.has(gk)) {
          outGrouped.set(gk, { ...x, _qtySum: 0, _lineCount: 0 });
          out.push(outGrouped.get(gk));
        }
        const g = outGrouped.get(gk);
        g._qtySum += Math.abs(Number(x.qty || 0));
        g._lineCount += 1;
        if (new Date(x.createdAt || 0).getTime() > new Date(g.createdAt || 0).getTime()) {
          g.createdAt = x.createdAt; g.id = x.id; g.user = x.user; g.memo = x.memo; g.stockAfter = x.stockAfter;
        }
      }
      return out;
    })();

    lastDisplayItems = displayItems;
    qs("#history-count").textContent = displayItems.length.toLocaleString();
    qs("#history-stock-info").textContent = matchedStocks ? `현재고: ${matchedStocks}` : "";

    const rows = displayItems.map((x) => `<tr>
      <td>${x.id}</td>
      <td>${typeBadge(x.type, x.cancelled, x.memo, x.originType)}</td>
      <td>${esc(x.slipNoDisplay || x.slipNo || "")}</td>
      <td>${esc(x.productCode)}</td>
      <td>${esc(x.ecountCode || "")}</td>
      <td>${esc(x.productName)}</td>
      <td class="num">${x.type === "OUT" ? `-${Number.isFinite(Number(x._qtySum)) ? Number(x._qtySum) : Math.abs(Number(x.qty || 0))}` : x.qty}</td>
      <td>${esc(x.warehouse || "-")}</td>
      <td>${esc(x.toWarehouse || "-")}</td>
      <td>${x.type === "IN" ? esc(x.partner) : x.type === "CANCEL" && x.originType === "IN" ? esc(x.partner) : "-"}</td>
      <td>${x.type === "OUT" ? esc(x.partner) : x.type === "CANCEL" && x.originType === "OUT" ? esc(x.partner) : "-"}</td>
      <td>${esc(x.user)}</td>
      <td>${x.type === "OUT" && Number(x._lineCount || 0) > 1 ? `${esc(x.memo || "")}${x.memo ? " / " : ""}센터 통합 합산(${Number(x._lineCount)}건)` : esc(x.memo)}</td>
      <td>${esc(formatDateTimeDisplay(x.createdAt))}</td>
      <td class="num">${x.stockAfter ?? "-"}</td>
      <td>${(!x.cancelled && ["IN", "OUT", "ADJUST"].includes(x.type) && !String(x.slipNo || "").trim()) ? `<button class="cancel-btn history-cancel-btn" data-id="${x.id}">등록취소</button>` : "-"}</td>
    </tr>`).join("");

    qs("#history-table").innerHTML = `
      <div class="history-table-outer">
        <div class="history-table-scroll" id="history-table-scroll">
          <table id="history-table-list">
            <thead><tr><th>ID</th><th>구분</th><th>전표번호</th><th>상품코드</th><th>품목코드</th><th>상품명</th><th>수량</th><th>창고</th><th>이동창고</th><th>구매처</th><th>판매처</th><th>담당</th><th>메모</th><th>일시</th><th>시점재고</th><th>액션</th></tr></thead>
            <tbody>${rows || `<tr><td colspan="16" style="text-align:center; padding:32px; color:var(--t3);">조회된 이력이 없습니다.</td></tr>`}</tbody>
          </table>
        </div>
        <div class="history-scroll-proxy-wrap" id="history-scroll-proxy-wrap">
          <div class="history-scroll-proxy-inner" id="history-scroll-proxy-inner"></div>
        </div>
      </div>
    `;

    const tblScroll = qs("#history-table-scroll");
    const proxyWrap = qs("#history-scroll-proxy-wrap");
    const proxyInner = qs("#history-scroll-proxy-inner");
    if (tblScroll && proxyWrap && proxyInner) {
      const tbl = tblScroll.querySelector("table");
      if (tbl) proxyInner.style.width = tbl.scrollWidth + "px";
      tblScroll.addEventListener("scroll", () => { proxyWrap.scrollLeft = tblScroll.scrollLeft; });
      proxyWrap.addEventListener("scroll", () => { tblScroll.scrollLeft = proxyWrap.scrollLeft; });
    }

    applyExcelLikeFilter("#history-table-list");

    document.querySelectorAll(".history-cancel-btn").forEach((btn) => {
      btn.onclick = async () => {
        const manager = prompt("취소 담당자명을 입력하세요 (기본정보>담당자 등록 필요)");
        if (!manager) return;
        try {
          await api("/api/movements/cancel", { method: "POST", body: JSON.stringify({ id: Number(btn.dataset.id), user: manager.trim() }) });
          await afterMovementDone();
        } catch (err) { alert(err.message); }
      };
    });
  }

  /* 빠른 날짜 버튼 */
  wrap.querySelectorAll(".history-quick-btn").forEach((btn) => {
    btn.onclick = () => {
      const now = new Date();
      const t = localDateYmd(now);
      if (btn.dataset.range === "today") {
        qs("#history-date-from").value = t;
        qs("#history-date-to").value = t;
      } else if (btn.dataset.range === "week") {
        const mon = new Date(now);
        mon.setDate(now.getDate() - ((now.getDay() + 6) % 7));
        qs("#history-date-from").value = localDateYmd(mon);
        qs("#history-date-to").value = t;
      } else if (btn.dataset.range === "month") {
        qs("#history-date-from").value = localDateYmd(new Date(now.getFullYear(), now.getMonth(), 1));
        qs("#history-date-to").value = t;
      } else if (btn.dataset.range === "all") {
        qs("#history-date-from").value = "";
        qs("#history-date-to").value = "";
      }
    };
  });

  /* 초기화 */
  qs("#history-reset").onclick = () => {
    const t = localDateYmd(new Date());
    qs("#history-q").value = "";
    qs("#history-type-filter").value = "ALL";
    qs("#history-warehouse-filter").value = "";
    qs("#history-manager-filter").value = "";
    qs("#history-date-from").value = t;
    qs("#history-date-to").value = t;
  };

  /* 엑셀 다운로드 */
  qs("#history-excel-btn").onclick = () => {
    if (!lastDisplayItems.length) { alert("조회 결과가 없습니다."); return; }
    const headers = ["ID","구분","전표번호","상품코드","품목코드","상품명","수량","창고","이동창고","구매처","판매처","담당","메모","일시","시점재고"];
    const data = lastDisplayItems.map((x) => [
      x.id,
      x.type + (x.cancelled ? "(취소)" : ""),
      x.slipNoDisplay || x.slipNo || "",
      x.productCode, x.ecountCode || "", x.productName,
      x.type === "OUT" ? -(Number.isFinite(Number(x._qtySum)) ? Number(x._qtySum) : Math.abs(Number(x.qty || 0))) : x.qty,
      x.warehouse || "", x.toWarehouse || "",
      x.type === "IN" ? x.partner : "",
      x.type === "OUT" ? x.partner : "",
      x.user, x.memo, formatDateTimeDisplay(x.createdAt), x.stockAfter ?? ""
    ]);
    const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "입출고이력");
    XLSX.writeFile(wb, `입출고이력_${localDateYmd(new Date())}.xlsx`);
  };

  /* 조회 */
  qs("#history-search").onclick = () => draw();
  qs("#history-q").addEventListener("keydown", (e) => { if (e.key === "Enter") draw(); });
}

function renderAlert() {
  qs("#view-alert").innerHTML = `
    <div class="card">
      <h2>알림</h2>
      <p class="muted">추후 업그레이드 예정 메뉴입니다.</p>
    </div>
  `;
}

async function afterMovementDone() {
  await refreshCommon();
  renderStock();
  renderProducts();
  renderMaster();
  await renderDashboard();
  await renderHistory();
  const active = document.querySelector(".main .view:not(.hidden)");
  if (active) {
    if (active.id === "view-inbound") renderInbound();
    if (active.id === "view-outbound") renderOutbound();
    if (active.id === "view-adjust") renderAdjust();
  }
}

async function init() {
  initSidebarAccordion();

  document.querySelectorAll(".sidebar button[data-view]").forEach((btn) => {
    btn.onclick = async () => {
      const v = btn.dataset.view;
      switchView(v);
      if (v === "dashboard") await renderDashboard();
      if (v === "master") renderMaster();
      if (v === "products") renderProducts();
      if (v === "stock") renderStock();
      if (v === "inbound") renderInbound();
      if (v === "inbound-plan") await renderInboundPlan();
      if (v === "inbound-plan-2") await renderInboundPlan2();
      if (v === "outbound") renderOutbound();
      if (v === "outbound-plan") renderOutboundPlanMenu();
      if (v === "outbound-upload-daiso") await renderOutboundPartnerUpload("다이소", "daiso");
      if (v === "outbound-upload-emart") await renderOutboundPartnerUpload("이마트", "emart");
      if (v === "outbound-upload-lotte") await renderOutboundPartnerUpload("롯데마트", "lotte");
      if (v === "adjust") renderAdjust();
      if (v === "history") await renderHistory();
      if (v === "alert") renderAlert();
    };
  });

  await refreshCommon();
  await renderDashboard();
  renderMaster();
  renderProducts();
  renderStock();
  renderInbound();
  await renderInboundPlan();
  await renderInboundPlan2();
  renderOutbound();
  renderOutboundPlanMenu();
  await renderOutboundPartnerUpload("다이소", "daiso");
  await renderOutboundPartnerUpload("이마트", "emart");
  await renderOutboundPartnerUpload("롯데마트", "lotte");
  renderAdjust();
  await renderHistory();
  renderAlert();
  let initialView = "dashboard";
  try {
    const saved = localStorage.getItem(LAST_VIEW_KEY);
    if (saved && views.includes(saved)) initialView = saved;
  } catch (_) {
    // ignore storage errors
  }
  switchView(initialView);
}

init().catch((e) => {
  alert(`초기화 실패: ${e.message}`);
});
