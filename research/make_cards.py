# -*- coding: utf-8 -*-
"""把 research/sample_seven_cards.json 排成 7 张 1080×1440 文字卡（PNG），输出到 deliverables/七图文字样张/。
纯文字自制排版，不使用任何官方或他人图片；字体为 Windows 自带微软雅黑。"""
import json
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "deliverables" / "七图文字样张"; OUT.mkdir(parents=True, exist_ok=True)
data = json.load((ROOT / "research" / "sample_seven_cards.json").open(encoding="utf-8"))
F = "C:/Windows/Fonts/"
fb = lambda sz: ImageFont.truetype(F + "msyhbd.ttc", sz)
fr = lambda sz: ImageFont.truetype(F + "msyh.ttc", sz)
W, H = 1080, 1440
BG, RED, RED2, INK, GREY, LINE = (247, 243, 236), (122, 31, 43), (163, 48, 63), (34, 30, 28), (100, 96, 92), (222, 214, 202)

def wrap(draw, text, font, maxw):
    lines, cur = [], ""
    for ch in text:
        if ch == "\n": lines.append(cur); cur = ""; continue
        if draw.textlength(cur + ch, font=font) > maxw and cur:
            lines.append(cur); cur = ch
        else:
            cur += ch
    if cur: lines.append(cur)
    return lines

for c in data["cards"]:
    im = Image.new("RGB", (W, H), BG); d = ImageDraw.Draw(im)
    d.rectangle([0, 0, 28, H], fill=RED)
    d.text((96, 92), f"回坑说明书 · 选题一 · {c['n']}/7", font=fr(30), fill=RED)
    d.text((96, 150), f"{c['n']:02d}", font=fb(150), fill=RED)
    d.text((96, 330), c["title"], font=fb(74), fill=INK)
    y = 450; d.line([(96, y), (W - 96, y)], fill=LINE, width=3); y += 50
    for ln in wrap(d, c["quote"], fb(58), W - 192):
        d.text((96, y), ln, font=fb(58), fill=RED2); y += 84
    y += 30
    for ln in wrap(d, c["body"], fr(42), W - 192):
        d.text((96, y), ln, font=fr(42), fill=INK); y += 68
    d.line([(96, H - 230), (W - 96, H - 230)], fill=LINE, width=3)
    d.text((96, H - 200), f"来源：{c['src']}", font=fr(30), fill=GREY)
    d.text((96, H - 150), "事实以国服官网公告为准，发布前重新核对日期", font=fr(28), fill=GREY)
    d.text((96, H - 100), "未发布样张 · 2026.11 · 自制排版", font=fr(28), fill=GREY)
    im.save(OUT / f"{c['n']:02d}.png", optimize=True)
    print("saved", c["n"], "body lines", len(wrap(d, c["body"], fr(42), W - 192)))
