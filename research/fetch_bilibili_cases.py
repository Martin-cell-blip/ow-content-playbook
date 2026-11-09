# -*- coding: utf-8 -*-
"""
不保存 UP 主 mid 或评论者 UID（2026-11-09 审阅后删除该字段）。B 站真实案例采集（公开 web 搜索 API 与评论 API，匿名态；观察日期写入每条记录）。

用途：为「回坑说明书」内容方案收集 2026-10-07（第5赛季开启）以来、与回流玩家/版本变化/总决赛相关的 B 站视频
及其可见互动数据（播放/点赞/收藏/评论/弹幕/时长/发布时间）与首页热评，作为平台案例分析的原始材料。
边界：搜索结果是检索样本不是全量；评论只取匿名态首页热评；不存用户名/UID。
输出：research/bilibili_cases.jsonl, research/bilibili_case_comments.jsonl, research/bilibili_fetch_log.json
"""
import json, re, time, sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
import requests

sys.stdout.reconfigure(encoding="utf-8")
HERE = Path(__file__).resolve().parent
BJT = timezone(timedelta(hours=8))
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
BASE = "https://api.bilibili.com/x/web-interface/search/type"
TAG_RE = re.compile(r"<[^>]+>")
KEYWORDS = ["守望先锋 回坑", "守望先锋 回归 老玩家", "守望先锋 血律", "守望先锋 黑影 支援",
            "守望先锋 往期战令", "守望先锋 新手 教学", "守望先锋 总决赛 广州", "守望先锋 伊芙计划"]
SINCE = int(datetime(2026, 10, 7, tzinfo=BJT).timestamp())
PAGES = 3
COMMENT_TOP = 16


def now():
    return datetime.now(BJT).isoformat(timespec="seconds")


def clean(s):
    return TAG_RE.sub("", s or "")


def main():
    sess = requests.Session()
    sess.headers.update({"User-Agent": UA, "Referer": "https://www.bilibili.com/"})
    sess.get("https://www.bilibili.com/", timeout=15)
    log = {"started": now(), "keywords": KEYWORDS, "since": "2026-10-07", "pages_per_kw": PAGES, "runs": []}
    hits = {}
    for kw in KEYWORDS:
        for page in range(1, PAGES + 1):
            params = {"search_type": "video", "keyword": kw, "page": page, "pubtime_begin_s": SINCE}
            r = None
            for wait in (0, 30, 90):
                time.sleep(wait)
                r = sess.get(BASE, params=params, timeout=20)
                if r.status_code == 200:
                    break
            body = r.json() if r.status_code == 200 else {"code": f"http_{r.status_code}"}
            if body.get("code") != 0:
                log["runs"].append({"kw": kw, "page": page, "code": body.get("code")})
                break
            res = body.get("data", {}).get("result") or []
            log["runs"].append({"kw": kw, "page": page, "n": len(res)})
            for rank, v in enumerate(res, 1):
                if v.get("type") != "video":
                    continue
                b = v.get("bvid")
                rec = hits.setdefault(b, {
                    "bvid": b, "aid": v.get("id"), "url": f"https://www.bilibili.com/video/{b}",
                    "title": clean(v.get("title")), "author": v.get("author"), 
                    "typename": v.get("typename"), "play": v.get("play"), "danmaku": v.get("danmaku"),
                    "like": v.get("like"), "favorites": v.get("favorites"), "review": v.get("review"),
                    "duration": v.get("duration"), "pubdate": v.get("pubdate"),
                    "pubdate_bjt": datetime.fromtimestamp(v.get("pubdate") or 0, BJT).strftime("%Y-%m-%d"),
                    "tag": v.get("tag"), "description": clean(v.get("description"))[:200],
                    "observed_at": now(), "hit_by": []})
                rec["hit_by"].append({"kw": kw, "page": page, "rank": rank})
            if not res:
                break
            time.sleep(1.2)
    rows = sorted(hits.values(), key=lambda x: -(x["play"] or 0))
    with (HERE / "bilibili_cases.jsonl").open("w", encoding="utf-8") as f:
        for x in rows:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    print("videos:", len(rows))

    # 热评：按主题词优先 + 播放 TOP 补足
    theme = re.compile(r"回坑|回归|老玩家|退坑|血律|黑影|战令|新手|总决赛|广州|伊芙")
    targets = [x for x in rows if theme.search(x["title"])][:COMMENT_TOP]
    for x in rows:
        if len(targets) >= COMMENT_TOP:
            break
        if x not in targets:
            targets.append(x)
    n_c = 0
    with (HERE / "bilibili_case_comments.jsonl").open("w", encoding="utf-8") as f:
        for v in targets:
            try:
                r = sess.get("https://api.bilibili.com/x/v2/reply",
                             params={"type": 1, "oid": v["aid"], "pn": 1, "sort": 1}, timeout=20)
            except requests.exceptions.ConnectionError as e:
                log["runs"].append({"reply": v["bvid"], "code": type(e).__name__}); time.sleep(8); continue
            body = r.json() if r.status_code == 200 else {"code": f"http_{r.status_code}"}
            if body.get("code") != 0:
                log["runs"].append({"reply": v["bvid"], "code": body.get("code")})
                time.sleep(30 if r.status_code == 412 else 1.5); continue
            for rank, rep in enumerate((body.get("data") or {}).get("replies") or [], 1):
                msg = (rep.get("content") or {}).get("message", "").strip()
                if msg:
                    f.write(json.dumps({"bvid": v["bvid"], "video_title": v["title"], "rank": rank,
                                        "comment": re.sub(r"\s+", " ", msg)[:400], "like": rep.get("like"),
                                        "rcount": rep.get("rcount"), "observed_at": now()}, ensure_ascii=False) + "\n")
                    n_c += 1
            time.sleep(1.5)
    log["finished"] = now(); log["videos"] = len(rows); log["comments"] = n_c
    (HERE / "bilibili_fetch_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
    print("comments:", n_c)


if __name__ == "__main__":
    main()
