# -*- coding: utf-8 -*-
"""Fix corrupted bulk-upload handlers and remove legacy #product-upload DOM wiring."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
path = ROOT / "public" / "app.js"
text = path.read_text(encoding="utf-8")

text = text.replace("\uc5c5\ud37c\uc774\ud2b8 \uac00\ub2a5\ud55c", "\uc5c5\ub370\uc774\uc5c4\ud2b8 \uac00\ub2a5\ud55c")

start = text.find('  const productFileInput = qs("#product-file");')
if start == -1:
    raise SystemExit("productFileInput block not found")

end_export = text.find('  qs("#product-export-all").onclick = async () => {', start)
if end_export == -1:
    raise SystemExit("export anchor not found")

msg_xls = (
    "await afterBulkProductImport(`"
    "\uC5D1\uC14C \uC5C5\uB85C\uB4DC \uC644\uB8CC (${n}\uAC74)`);"
)
msg_upd = (
    "await afterBulkProductImport(`"
    "\uC0C1\uD488\uC815\uBCF4 \uC5C5\uB370\uC774\uC5C4\uD2B8 \uC644\uB8CC (${n}\uAC74)`);"
)

new_mid = f"""  const productFileInput = qs("#product-file");
  const productUpdateInput = qs("#product-update-file");

  const afterBulkProductImport = async (label) => {{
    renderProducts();
    renderStock();
    renderMaster();
    await renderDashboard();
    alert(label);
  }};

  setupDropZone("product-dropzone", "product-file", async (f) => {{
    if (!f) return;
    try {{
      const n = await uploadProductBulkFromFile(f);
      if (productFileInput) productFileInput.value = "";
      {msg_xls}
    }} catch (e2) {{
      alert(e2.message);
      if (productFileInput) productFileInput.value = "";
    }}
  }});

  productFileInput?.addEventListener("change", async () => {{
    const file = productFileInput.files[0];
    if (!file) return;
    try {{
      const n = await uploadProductBulkFromFile(file);
      productFileInput.value = "";
      {msg_xls}
    }} catch (e2) {{
      alert(e2.message);
      productFileInput.value = "";
    }}
  }});

  productUpdateInput?.addEventListener("change", async () => {{
    const file = productUpdateInput.files[0];
    if (!file) return;
    try {{
      const n = await uploadProductUpdateFromFile(file);
      productUpdateInput.value = "";
      {msg_upd}
    }} catch (err) {{
      alert(err.message);
      productUpdateInput.value = "";
    }}
  }});

"""

text = text[:start] + new_mid + text[end_export:]

dup_start = text.find('  qs("#product-upload").onclick = async () => {')
if dup_start == -1:
    raise SystemExit("legacy product-upload not found")
excel_start = text.find('  applyExcelLikeFilter("#products-table");', dup_start)
if excel_start == -1:
    raise SystemExit("applyExcelLikeFilter not found")

text = text[:dup_start] + "  applyProductHiddenColumnStyles();\n" + text[excel_start:]

path.write_text(text, encoding="utf-8")
print("ok")
