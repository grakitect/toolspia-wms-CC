# -*- coding: utf-8 -*-
import codecs
import re
from pathlib import Path

p = Path(__file__).resolve().parent.parent / "public" / "app.js"
s = p.read_text(encoding="utf-8")

m_th = re.search(r'<table id="products-table">\s*\n\s*(<thead>[\s\S]*?</thead>)', s)
if not m_th:
    raise SystemExit("thead not found")
thead_only = m_th.group(1)

pat = re.compile(
    r'(        <p class="muted" style="margin:0;font-size:11px;">[^<]+</p>)\n      </div>\n    </div>\n\n(    <div id="product-modal-overlay" class="modal-overlay hidden">)'
)
m = pat.search(s)
if not m:
    raise SystemExit("patch point not found")

body = (
    m.group(1)
    + """
        <div class="products-bh-table-outer">
          <div class="table-scroll-proxy-wrap"><div id="products-scroll-proxy" class="table-scroll-proxy-inner"></div></div>
          <div class="table-scroll-x products-bh-y-scroll">
            <table id="products-table">
              """
    + thead_only
    + """
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
          <button type="button" id="product-page-prev" class="bh-btn bh-btn-sm" title="\uc774\uc804">\u2039</button>
          <button type="button" id="product-page-next" class="bh-btn bh-btn-sm" title="\ub2e4\uc74c">\u203a</button>
        </div>
      </div>
    </div>

"""
    + m.group(2)
)

s = s[: m.start()] + body + s[m.end() :]

a = s.find("\n    <div class=\"card products-left\">")
b = s.find('    ${datalist("opt-status"', a)
if a < 0 or b < 0:
    raise SystemExit("dup card not found")
s = s[:a] + "\n" + s[b:]

if "product-columns-overlay" not in s:
    col = codecs.decode(
        r"""    <div id="product-columns-overlay" class="modal-overlay hidden">
      <div class="modal" style="width: min(420px, calc(100vw - 24px));">
        <div class="modal-header">
          <h3>\uceec\ub7fc \uc124\uc815</h3>
          <button type="button" id="product-columns-close" class="cancel-btn del-small">\ub2eb\uae30</button>
        </div>
        <p class="muted" style="margin: 0 0 8px; font-size: 12px;">\uccb4\ud06c \ud574\uc81c \uc2dc \ud574\ub2f9 \uc5f4\uc744 \uc228\uae41\ub2c8\ub2e4. (\uccb4\ud06c\ubc15\uc2a4\ub294 \ud56d\uc0c1 \ud45c\uc2dc)</p>
        <div id="product-columns-body" class="column-settings-list"></div>
        <button type="button" class="primary" id="product-columns-save" style="width: auto;">\uc801\uc6a9</button>
      </div>
    </div>

""",
        "unicode_escape",
    )
    s = s.replace('    ${datalist("opt-status"', col + '    ${datalist("opt-status"', 1)

p.write_text(s, encoding="utf-8")
print("ok")
