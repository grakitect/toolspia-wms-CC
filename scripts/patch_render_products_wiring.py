# -*- coding: utf-8 -*-
"""Replace product list search block, fix upload error strings, submit button label."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
path = ROOT / "public" / "app.js"
lines = path.read_text(encoding="utf-8").splitlines(keepends=True)

# JS fragment; Korean via unicode escapes to keep this file robust
NEW_SEARCH_BLOCK = r"""  // 상품 조회 + 페이지네이션
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
    const matched = allRows.filter((tr) => !q || String(tr.dataset.search || "").includes(q));
    const size = Math.min(99999, Math.max(1, parseInt(pageSizeEl?.value || "100", 10) || 100));
    const total = matched.length;
    const pages = Math.max(1, Math.ceil(total / size));
    if (productListPageIndex >= pages) productListPageIndex = pages - 1;
    if (productListPageIndex < 0) productListPageIndex = 0;
    const page = productListPageIndex;
    const start = page * size;
    let mi = 0;
    allRows.forEach((tr) => {
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
      const total2 = allRows2.filter((tr) => !q2 || String(tr.dataset.search || "").includes(q2)).length;
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
"""


def main():
    start = None
    for i, line in enumerate(lines):
        if "아래 리스트 필터" in line and line.strip().startswith("// 상품 조회"):
            start = i
            break
    if start is None:
        raise SystemExit("start anchor not found")

    end = None
    for j in range(start + 1, len(lines)):
        if lines[j].strip() == 'preventEnterSubmit(qs("#product-popup-form"));':
            if j + 1 < len(lines) and "const productForm = qs(" in lines[j + 1]:
                end = j
                break
    if end is None:
        raise SystemExit("end anchor not found")

    new_lines = NEW_SEARCH_BLOCK.splitlines(keepends=True)
    if not new_lines[-1].endswith("\n"):
        new_lines[-1] += "\n"

    text = "".join(lines[:start] + new_lines + lines[end + 1 :])

    import re

    bulk_msg = (
        "throw new Error(\"\uc5c5\ub85c\ub4dc \uac00\ub2a5\ud55c \uc0c1\ud488\ud589\uc774 \uc5c6\uc2b5\ub2c8\ub2e4. "
        "\ud5e4\ub354\uba85(\ud488\ubaa9\ucf54\ub4dc(\uc774\uce74\uc6b4\ud2b8), \ud488\ubaa9\uba85(\uc774\uce74\uc6b4\ud2b8))"
        "\uacfc \ub370\uc774\ud130 \uc720\ubb34\ub97c \ud655\uc778\ud558\uc138\uc694.\");"
    )
    text = re.sub(
        r'throw new Error\("업로드 가능한 상품행이 없습니다\.[^"]*"\);',
        bulk_msg,
        text,
        count=1,
    )
    upd_msg = (
        'if (!normalized.length) throw new Error("\uc5c5\ud37c\uc774\ud2b8 \uac00\ub2a5\ud55c '
        '\uc0c1\ud488\ud589\uc774 \uc5c6\uc2b5\ub2c8\ub2e4.");'
    )
    text = re.sub(
        r"if \(!normalized\.length\) throw new Error\(\"[^\"]*\"\);",
        upd_msg,
        text,
        count=1,
    )

    text = re.sub(
        r'(<div><button class="primary" type="submit">)([^<]+)(</button></div>\s*</form>\s*</div>\s*</div>\s*\n\s*<div id="option-modal-overlay")',
        lambda m: m.group(1) + "\uc785\ub825 \uc644\ub8cc" + m.group(3),
        text,
        count=1,
    )

    path.write_text(text, encoding="utf-8")
    print("ok: search/pagination, menus, columns, upload messages, submit label")


if __name__ == "__main__":
    main()
