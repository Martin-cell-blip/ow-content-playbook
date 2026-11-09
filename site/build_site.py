# -*- coding: utf-8 -*-
"""把 template.html 与 data.json 合成为单文件 site/index.html（离线可开、可打印）。"""
import json
from pathlib import Path
HERE = Path(__file__).resolve().parent
data = (HERE / "data.json").read_text(encoding="utf-8")
tpl = (HERE / "template.html").read_text(encoding="utf-8")
html = tpl.replace("/*__DATA__*/", data.replace("</", "<\\/"))
(HERE / "index.html").write_text(html, encoding="utf-8")
print("index.html", len(html))
