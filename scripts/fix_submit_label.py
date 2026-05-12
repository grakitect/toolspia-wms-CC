import re
from pathlib import Path

p = Path(__file__).resolve().parent.parent / "public" / "app.js"
t = p.read_text(encoding="utf-8")
label = "\uc785\ub825 \uc644\ub8cc".encode("utf-8").decode("unicode_escape")
t2, n = re.subn(
    r'(<form id="product-popup-form"[\s\S]*?<div><button class="primary" type="submit">)[^<]*(</button></div>)',
    r"\1" + label + r"\2",
    t,
    count=1,
)
if n != 1:
    raise SystemExit("replace count %s" % n)
p.write_text(t2, encoding="utf-8")
print("ok")
