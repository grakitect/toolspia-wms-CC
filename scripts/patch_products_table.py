# -*- coding: utf-8 -*-
from pathlib import Path

p = Path(__file__).resolve().parent.parent / "public" / "app.js"
lines = p.read_text(encoding="utf-8").splitlines(keepends=True)
text = "".join(lines)

# Insert table before closing list-view (after 다중 ��그 line)
old_close = (
    '        <p class="muted" style="margin:0;font-size:11px;">다중 ��그는 ���표(,) 구분 업로드</p>\r\n'
    "      </div>\r\n"
    "    </div>\r\n"
    "\r\n"
    '    <div id="product-modal-overlay"'
)
old_close_n = old_close.replace("\r\n", "\n")
idx = text.find(old_close_n)
if idx < 0:
    old_close = (
        '        <p class="muted" style="margin:0;font-size:11px;">다중 ���표(,) 구분 업로드</p>\n'
        "      </div>\n"
        "    </div>\n"
        "\n"
        '    <div id="product-modal-overlay"'
    )
    idx = text.find(old_close)
if idx < 0:
    raise SystemExit("close block not found")

insert = (
    '        <p class="muted" style="margin:0;font-size:11px;">다중 ��그는 ���표(,) 구분 업로드</p>\n'
    '        <div class="products-bh-table-outer">\n'
    '          <div class="table-scroll-proxy-wrap"><div id="products-scroll-proxy" class="table-scroll-proxy-inner"></div></div>\n'
    '          <div class="table-scroll-x products-bh-y-scroll">\n'
    '            <table id="products-table">\n'
    '              <thead><tr><th><input id="product-check-all" type="checkbox" /></th><th>품목코드(이카운트)</th><th>바코드</th><th>물류바코드</th><th>품목명(이카운트)</th><th>상태</th><th>판매처</th><th>판매처관리코드</th><th>판�목명</th><th>��격</th><th>구매처</th><th>수��형태</th><th>발주부서</th><th>발주��당자</th><th>구매처 ���목코드</th><th>��목명</th><th>��고그��(이카운트)</th><th�고</th><th>구분</th><th>카테고리</th></tr></thead>\n'
    "              <tbody>${rows}</tbody>\n"
    "            </table>\n"
    "          </div>\n"
    "        </div>\n"
    '        <div class="products-bh-pagination">\n'
    "          <label>보기\n"
    '            <select id="product-page-size">\n'
    '              <option value="50">50</option>\n'
    '              <option value="100" selected>100</option>\n'
    '              <option value="200">200</option>\n'
    '              <option value="99999">전체</option>\n'
    "            </select>\n"
    "          </label>\n"
    '          <span id="product-page-info">0 - 0 / 0</span>\n'
    '          <button type="button" id="product-page-prev" class="bh-btn bh-btn-sm" title="이전">��</button>\n'
    '          <button type="button" id="product-page-next" class="bh-btn bh-btn-sm" title="다음">›</button>\n'
    "        </div>\n"
    "      </div>\n"
    "    </div>\n"
    "\n"
    '    <div id="product-modal-overlay"'
)

if old_close_n in text:
    text = text.replace(old_close_n, insert, 1)
else:
    text = text.replace(old_close, insert, 1)

dup = """    <div class="card products-left">
      <div class="row" style="grid-template-columns: 1fr 420px; align-items: start; margin-bottom: 10px;">
        <h3 style="margin: 0;">현재 등록된 상품정보</h3>
        <div>
          <h3 style="margin: 0 0 8px;">상품 조회</h3>
          <div class="row product-search-row" style="margin-bottom: 10px;">
            <input id="product-search-q" class="product-search-q" placeholder="이카운트/상품코드 검색" />
            <button class="primary" id="product-search-btn">조회</button>
          </div>
          <div id="product-search-result" class="muted" style="margin-bottom:0;">
            상품을 검색하세요.
          </div>
        </div>
      </div>
        <div class="row product-search-row" style="margin-bottom:10px; grid-template-columns: 120px 120px;">
          <button id="product-edit-selected" type="button">선택 수정</button>
          <button id="product-delete-selected" class="cancel-btn" type="button">선택 ��제</button>
        </div>
        <div class="table-scroll-proxy-wrap"><div id="products-scroll-proxy" class="table-scroll-proxy-inner"></div></div>
        <div class="table-scroll-x">
          <table id="products-table">
            <thead><tr><th><input id="product-check-all" type="checkbox" /></th><th>품목코드(이카운트)</th><th>바코드</th><th>물류바코드</th><th>품목명(이카운트)</th><th>상태</th><th>판매처</th><th>판매처관리코드</th><th>��목명</th><th>��격</th><th>구매처</th><th>수��형태</th><th>발주부서</th><th>발주��당자</th><th>구�목코드</th><th>구매처 ���목명</th><th>��고그��(이카운트)</th><th>사용��고</th><th>구분</th><th>카테고리</th></tr></thead>
            <tbody>${rows}</tbody>
          </table>
        </div>
      </div>

"""

if dup not in text:
    raise SystemExit("dup card not found")
text = text.replace(dup, "", 1)

col_modal = """    <div id="product-columns-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(420px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3� 설정</h3>
          <button type="button" id="product-columns-close" class="cancel-btn del-small">��기</button>
        </div>
        <p class="muted" style="margin: 0 0 8px; font-size: 12px;">체크 해제 시 해당 ��을 �����니다. (체크��스�상 표시)</p>
        <div id="product-columns-body" class="column-settings-list"></div>
        <button type="button" class="primary" id="product-columns-save" style="width: auto;">적용</button>
      </div>
    </div>

"""

needle = '    ${datalist("opt-status", opts.status)}'
if needle not in text:
    raise SystemExit("datalist needle missing")
text = text.replace(needle, col_modal + needle, 1)

p.write_text(text, encoding="utf-8")
print("ok")
