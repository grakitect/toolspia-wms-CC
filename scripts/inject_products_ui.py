# -*- coding: utf-8 -*-
"""One-off: replace view-products innerHTML in public/app.js."""
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent.parent
path = ROOT / "public" / "app.js"
t = path.read_text(encoding="utf-8")

m = re.search(r"(<form id=\"product-popup-form\" class=\"modal-form\">[\s\S]*?</form>)", t)
if not m:
    raise SystemExit("form not found")
form_inner = m.group(1)
form_inner = re.sub(
    r"\s*<div><button class=\"primary\" type=\"submit\">등록/수정</button></div>\s*",
    "",
    form_inner,
)

option_block = re.search(
    r"(<div id=\"option-modal-overlay\"[\s\S]*?</div>\s*</div>\s*)", t
).group(1)

tail = re.search(
    r"(\$\{datalist\(\"opt-status\"[\s\S]*?<div id=\"stock-edit-modal-overlay\"[\s\S]*?</div>\s*</div>\s*)",
    t,
).group(1)

rows_marker = "__ROWS__"

list_view = f"""    <div id="products-list-view" class="products-bh">
      <div class="card products-bh-card">
        <div class="products-bh-toolbar">
          <h2 class="products-bh-title">기본상품정보</h2>
          <div class="products-bh-actions">
            <button type="button" id="open-product-page" class="bh-btn bh-btn-primary">+ 상품 추가</button>
            <div class="bh-dropdown">
              <button type="button" class="bh-btn bh-btn-outline bh-excel-btn" id="excel-import-toggle" aria-expanded="false">
                <span class="bh-mini-xls">XLS</span>
� 가져오기
                <span class="bh�</span>
              </button>
              <div class="bh-dropdown-menu hidden" id="excel-import-menu">
                <button type="button" class="bh-menu-item" id="menu-product-guide">가이드 다운로드</button>
                <button type="button" class="bh-menu-item" id="menu-bulk-new">신���� 등록</button>
                <button type="button" class="bh-menu-item" id="menu-b�정상품 정보 업��이트</button>
              </div>
            </div>
            <div class="bh-dropdown">
              <button type="button" class="bh-btn bh-btn-icon" id="more-menu-toggle" title="더보기">⋯</button>
              <div class="bh-dropdown-menu hidden" id="more-menu">
                <button type="button" class="bh-menu-item" id="open-option-popup">��션등록</button>
              </div>
            </div>
          </div>
        </div>
        <div class="products-bh-subbar">
          <button type="button" class="bh-link-btn" id="product�기</button>
          <button type="button" class="bh-link-btn" id="product-column-settings">���� 설정</button>
          <span class="products-bh-bulk">
            <button type="button" id="product-edit-selected" class="bh-btn bh-btn-sm">선택 수정</button>
            <button type="button" id="product-delete-selected" class="bh-btn bh-btn-sm bh-btn-danger-outline">선택 ��제</button>
          </span>
        </div>
        <div class="products-bh-search-wrap">
          <span class="bh-search-icon" aria-hidden�</span>
          <input type="text" id="product-search-q" class="products-bh-search" placeholder="이카운트 / 상품코드 / ���목명 검색" autocomplete="off" />
          <button type="button" class="bh-search-go" id="product-search-btn">조회</button>
        </div>
        <p id="product-search-result" class="muted products-bh-result">상품을 검색하세요.</p>
        <div id="product-dropzone" class="dropzone bh-dropzone��� 파일을 여기에 ��래그하거나 ����하여 선택 (신�� 등록)</div>
        <input type="file" id="product-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <input type="file" id="product-update-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <div class="products-bh-table-outer">
          <div class="table-scroll-proxy-wrap"><div id="products-scroll-proxy" class="table-scroll-proxy-inner"></div></div>
          <div class="table-scroll-x products-bh-y-scroll">
            <table id="products-table">
              <thead><tr><th><input id="product-check-all" type="checkbox" /></th><th>품목코드(이카운트)</th><th>바코드</th><th>물류바코드</th><th>품목명(이카운트)</th><th>상태</th><th>판매처</th><th>판매처관리코드</th><th>판�목명</th><th>��격</th><th>구매처</th><th>수��형태</th><th>발주부서</th><th>발주��당자</th><th>구매처 ���목코드</th><th>��목명</th><th>��고그��(이카운트)</th><th�고</th><th>구분</th><th>카테고리</th></tr></thead>
              <tbody>{rows_marker}</tbody>
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
          <button type="button" id="product-page-prev" class="bh-btn bh-btn-sm" title="이전">��</button>
          <button type="button" id="product-page-next" class="bh-btn bh-btn-sm" title="다음">›</button>
        </div>
      </div>
    </div>

    <div id="products-add-view" class="products-bh products-add-page hidden">
      <div class="card products-bh-card products-add-card">
        <div class="products-add-header">
          <h2 id="product-page-title">개별상품등록</h2>
          <div class="products-add-header-actions">
            <button type="button" id="product-page-reset" class="bh-btn bh-btn-sm bh-btn-outline">���기화</button>
            <button type="button" id="product-page-back" class="bh-btn bh-btn-sm bh-btn-outline">목록으로</button>
          </div>
        </div>
        <div class="products-add-scroll">
"""

add_view_mid = (
    form_inner.replace(
        '<form id="product-popup-form" class="modal-form">',
        '<form id="product-popup-form" class="modal-form">\n'
        '            <div class="product-sec-title">품목 기본 정보</div>\n',
        1,
    )
    .replace(
        '<div><label>판매처</label>',
        '<div class="product-sec-title">판매·발주</div>\n            <div><label>판매처</label>',
        1,
    )
    .replace(
        '<div><label>��고그��(이카운트)</label>',
        '<div class="product-sec-title">��고·분류</div>\n            <div><label>��고그��(이카운트)</label>',
        1,
    )
    .replace(
        '<div><label>안전재고(선택)</label>',
        '<div class="product-sec-title">재고 ��션</div>\n            <div><label>안전재고(선택)</label>',
        1,
    )
)

add_view_end = """
        </div>
        <div class="products-add-footer">
          <button type="submit" form="product-popup-form" class="bh-btn bh-btn-primary">입력 ��료</button>
          <button type="button" id="product-page-cancel" class="bh-btn bh��소</button>
        </div>
      </div>
    </div>

"""

columns_modal = """    <div id="product-columns-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(420px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3>���� 설정</h3>
          <button type="button" id="product-columns-close" class="cancel-btn del-small">��기</button>
        </div>
        <p class="muted" style="margin: 0 0 8px; font-size: 12px;">체크 해제 시 해당���니다. (체크�은 ��상 표시)</p>
        <div id="product-columns-body" class="column-settings-list"></div>
        <button type="button" class="primary" id="product-columns-save" style="width: auto;">적용</button>
      </div>
    </div>

"""

new_inner = (
    list_view.replace(rows_marker, "${rows}")
    + add_view_mid
    + add_view_end
    + option_block
    + columns_modal
    + tail
)

# splicea = t.index('  qs("#view-products").innerHTML = `')
b = t.index("\n  `;", a) + len("\n  `;")
old_assign = t[a:b]
if "modalOverlay" not in t[b : b + 80]:
    raise SystemExit("unexpected content after template")
new_assign = '  qs("#view-products").innerHTML = `\n' + new_inner + "\n  `;"
t = t[:a] + new_assign + t[b:]

path.write_text(t, encoding="utf-8")
print("injected ok, len", len(new_assign))
