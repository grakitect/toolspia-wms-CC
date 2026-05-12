# -*- coding: utf-8 -*-
import codecs
from pathlib import Path

root = Path(__file__).resolve().parent.parent
p = root / "public" / "app.js"
s = p.read_text(encoding="utf-8")

needle_start = "    <div class=\"card\">\n      <h2>\uae30\ubcf8\uc0c1\ud488\uc815\ubcf4</h2>"
needle_end = "        <form id=\"product-popup-form\""

a = s.find(needle_start)
b = s.find(needle_end, a)
if a < 0 or b < 0:
    raise SystemExit("markers not found: %s %s" % (a, b))

raw = r"""    <div id="products-list-view" class="products-bh">
      <div class="card products-bh-card">
        <div class="products-bh-toolbar">
          <h2 class="products-bh-title">\uae30\ubcf8\uc0c1\ud488\uc815\ubcf4</h2>
          <div class="products-bh-actions">
            <button id="open-product-popup" class="bh-btn bh-btn-primary" type="button">+ \uc0c1\ud488 \ucd94\uac00</button>
            <div class="bh-dropdown">
              <button type="button" class="bh-btn bh-btn-outline bh-excel-btn" id="excel-import-toggle" aria-expanded="false">
                <span class="bh-mini-xls">XLS</span>
                \uc5d1\uc140 \uac00\uc838\uc624\uae30
                <span class="bh-caret">\u25be</span>
              </button>
              <div class="bh-dropdown-menu hidden" id="excel-import-menu">
                <button type="button" class="bh-menu-item" id="menu-product-guide">\uac00\uc774\ub4dc \ub2e4\uc6b4\ub85c\ub4dc</button>
                <button type="button" class="bh-menu-item" id="menu-bulk-new">\uc2e0\uaddc \uc0c1\ud488 \uc77c\uad04 \ub4f1\ub85d</button>
                <button type="button" class="bh-menu-item" id="menu-bulk-update">\ud2b9\uc815\uc0c1\ud488 \uc815\ubcf4 \uc5c5\ub370\uc774\ud2b8</button>
              </div>
            </div>
            <div class="bh-dropdown">
              <button type="button" class="bh-btn bh-btn-icon" id="more-menu-toggle" title="\ub354\ubcf4\uae30">\u22ef</button>
              <div class="bh-dropdown-menu hidden" id="more-menu">
                <button type="button" class="bh-menu-item" id="open-option-popup">\uc635\uc158\ub4f1\ub85d</button>
              </div>
            </div>
          </div>
        </div>
        <div class="products-bh-subbar">
          <button type="button" class="bh-link-btn" id="product-export-all">\uc5d1\uc140 \ub0b4\ubcf4\ub0b4\uae30</button>
          <button type="button" class="bh-link-btn" id="product-column-settings">\uceec\ub7fc \uc124\uc815</button>
          <span class="products-bh-bulk">
            <button type="button" id="product-edit-selected" class="bh-btn bh-btn-sm">\uc120\ud0dd \uc218\uc815</button>
            <button type="button" id="product-delete-selected" class="bh-btn bh-btn-sm bh-btn-danger-outline">\uc120\ud0dd \uc0ad\uc81c</button>
          </span>
        </div>
        <div class="products-bh-search-wrap">
          <span class="bh-search-icon" aria-hidden="true">\U0001f50d</span>
          <input type="text" id="product-search-q" class="products-bh-search" placeholder="\uc774\uce74\uc6b4\ud2b8 / \uc0c1\ud488\ucf54\ub4dc / \ud488\ubaa9\uba85 \uac80\uc0c9" autocomplete="off" />
          <button type="button" class="bh-search-go" id="product-search-btn">\uc870\ud68c</button>
        </div>
        <p id="product-search-result" class="muted products-bh-result">\uc0c1\ud488\uc744 \uac80\uc0c9\ud558\uc138\uc694.</p>
        <div id="product-dropzone" class="dropzone bh-dropzone-compact">\uc5d1\uc140 \ud30c\uc77c\uc744 \ub4dc\ub798\uadf8 \ud558\uac70\ub098 \ud074\ub9ad\ud558\uc5ec \uc120\ud0dd (\uc2e0\uaddc \uc77c\uad04 \ub4f1\ub85d)</div>
        <input type="file" id="product-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <input type="file" id="product-update-file" accept=".xlsx,.xls,.csv" class="hidden-file" />
        <p class="muted" style="margin:0;font-size:11px;">\ub2e4\uc911 \ud0dc\uadf8\ub294 \uc27c\ud45c(,) \uad6c\ubd84 \uc5c5\ub85c\ub4dc</p>
      </div>
    </div>

    <div id="product-modal-overlay" class="modal-overlay hidden">
      <div class="modal modal-product-full">
        <div class="modal-header product-modal-header-row">
          <h3 id="product-modal-title">\uac1c\ubcc4\uc0c1\ud488\ub4f1\ub85d</h3>
          <div class="product-modal-header-actions">
            <button type="button" id="product-form-reset" class="bh-btn bh-btn-sm bh-btn-outline">\ucd08\uae30\ud654</button>
            <button id="product-modal-close" class="bh-btn bh-btn-sm bh-btn-outline" type="button">\ub2eb\uae30</button>
          </div>
        </div>
"""
new_block = codecs.decode(raw, "unicode_escape")

s = s[:a] + new_block + s[b:]
p.write_text(s, encoding="utf-8")
print("ok", a, b)
