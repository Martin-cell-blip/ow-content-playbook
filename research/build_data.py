# -*- coding: utf-8 -*-
"""把三平台原始观察整理成：① deliverables/02_真实案例分析表.xlsx ② site/data.json（展示页内联数据）。
全部数字只来自 research/ 下的采集文件；本脚本不生成任何估计值。"""
import json, re, statistics
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
R = ROOT / "research"
OUT_X = ROOT / "deliverables" / "02_真实案例分析表.xlsx"
OUT_J = ROOT / "site" / "data.json"
(ROOT / "deliverables").mkdir(exist_ok=True)
(ROOT / "site").mkdir(exist_ok=True)

OBS = "2026-11-09"
HERO = "血律|黑影|路霸|源氏|天使|雾子|安娜|莫伊拉|D.Va|DVA|铁拳|斩仇|西格玛|温斯顿|卡西迪|死神|狂鼠|朱诺|拉玛刹|破坏球|哈蒙德|查莉娅|索杰恩|骇灾|死怨|金驭|瑞稀|弗蕾娅|芙蕾雅|黑百合|士兵76|艾什|秩序之光|猎空|半藏|堡垒|巴蒂斯特|布里吉塔|卢西奥|安燃|D.Mon|埃姆雷|探奇|雪神|大锤|莱因哈特|生命之梭|伊拉锐|西拉|奥丽莎|托比昂|法老之鹰|回声|毛妹"
OW = re.compile(r"守望|overwatch|\bow\b|owcs|" + HERO, re.I)
NON_OW = re.compile(r"影之刃|英雄联盟|K/DA|无畏契约|黎明杀机|Steam|魔兽|APEX|笔记本|专升本|单机神作|无期迷途|剑星R|手游|星际战甲|和平精英|第五人格|使命召唤|绝区零|CODM|DNF|蔚蓝星球|瓦罗兰特|良心单机|史低", re.I)

def theme_of(t):
    if re.search(r"回坑|回归(?!首秀)|老玩家|退坑|复建|多年没玩|牢玩家", t): return "回归/老玩家"
    if re.search(r"新手|教学|入坑|入门|技巧|设置|思路", t): return "新手/教学"
    if re.search(r"补丁|平衡|削弱|加强|季中|版本|改动|数值", t): return "版本/补丁"
    if re.search(r"血律|黑影|路霸", t): return "新英雄/重做"
    if re.search(r"剑星|伊芙|皮肤|神话|特写|联动|战令|商城|返场", t): return "外观/战令/联动"
    if re.search(r"总决赛|广州|OWCS|对阵|赛后|采访|冠军|季后赛|WBG|AG", t, re.I): return "赛事"
    if re.search(r"封号|代练|炸鱼|陪玩|严打|治理|卜天才|无糖可乐|主播", t): return "环境治理/主播事件"
    if re.search(r"语音|动画|故事|剧情|短片|彩蛋", t): return "剧情/语音"
    return "其他/整活"

rows = [json.loads(l) for l in (R / "bilibili_cases.jsonl").open(encoding="utf-8")]
recent = [r for r in rows if r["pubdate_bjt"] >= "2026-10-07"]
ow = []
for r in recent:
    blob = f'{r["title"]} {r.get("tag","")} {r.get("description","")}'
    if NON_OW.search(r["title"]) and not re.search(r"守望|OW", r["title"], re.I):
        continue
    if OW.search(blob):
        r = dict(r); r["theme"] = theme_of(r["title"]); ow.append(r)
ow.sort(key=lambda x: -(x["play"] or 0))

def med(xs): return round(statistics.median(xs), 4) if xs else None
themes = {}
for r in ow:
    t = themes.setdefault(r["theme"], {"n": 0, "play": [], "like_r": [], "fav_r": [], "rev_r": []})
    t["n"] += 1; p = r["play"] or 0
    t["play"].append(p)
    if p:
        t["like_r"].append((r["like"] or 0) / p); t["fav_r"].append((r["favorites"] or 0) / p); t["rev_r"].append((r["review"] or 0) / p)
theme_rows = sorted([{"theme": k, "n": v["n"], "median_play": int(med(v["play"])), "median_like_rate": med(v["like_r"]),
                      "median_fav_rate": med(v["fav_r"]), "median_reply_rate": med(v["rev_r"])} for k, v in themes.items()], key=lambda x: -x["n"])

comments = [json.loads(l) for l in (R / "bilibili_case_comments.jsonl").open(encoding="utf-8")]
manual = json.load((R / "manual_cases_2026-11-09.json").open(encoding="utf-8"))

# 判断依据 → 对设计的影响（人工撰写，逐条指向可查的案例 ID）
judgements = [
    ("J01", "B 站样本", "2026-10-07 后 OW 相关样本中「回归/老玩家」主题条数少、而「新手/教学」「版本/补丁」多（见「B站主题统计」表）；10-15 研究 3,760 条里明确面向回归者的只有 18 条", "案例观察", "两批检索样本都显示面向回归者的内容很少：做「回坑说明书」是补样本里缺的一类内容；样本不是全量，需求强度留给试读验证，不写「全平台空白」。"),
    ("J02", "B 站 BV1gn4b6iEno / 抖音 DY09", "同一作者「卡破伏岓」同主题补丁分析：B 站 8 分钟版评论 543、收藏 175；抖音 8 分钟版评论 93、分享 210", "案例观察", "B 站承载长讨论与收藏，适合放完整版说明书视频；抖音更像转发传播，不作为本方案主阵地。"),
    ("J03", "B 站 BV1hE5RzfE2D 热评", "《穷哥们》「10-14 战令自选优先级」热评是具体的兑换顺序建议（1226 / 740 / 487 赞）", "案例观察", "回归者对「先领什么、先补哪期」有真实决策需求 → 选题二做成可照着做的清单，而不是泛讲福利。"),
    ("J04", "B 站 BV1erc6ecEW4 热评", "「守望国服回归怎么把皮肤领齐」热评第一是 695 赞的「省流」分步说明", "案例观察", "评论区会自发产出「省流版」；我们在正文里先给省流，再给细节，并在置顶评论放可复制清单。"),
    ("J05", "小红书 XHS01", "回坑一个月的玩家提问「从没看过比赛建议去看 owcs 总决赛吗」（10-13，95 赞），最佳回答是「先认识哪些队和选手」的清单", "案例观察", "选题三的形态定为「回归者视角的观赛入门清单」，而非赛事战报。"),
    ("J06", "小红书 搜索「守望先锋回坑」列表", "保存的 19 条里 8 条是找固玩/找搭子（观察时前 30 条中数到 9 条，完整列表未保存），多数 0–20 赞、评论区空；联想词前列是「回坑找人 / 找队友 / 组队」", "案例观察", "「和谁玩」是小红书上一类明显的需求线索 → 每篇小红书样稿末尾固定一个「留下段位+时段」的组队互动位，试一试评论区能不能接住。"),
    ("J07", "小红书 XHS04 vs XHS06", "同一补丁节点：「再见了，所有安娜玩家。」278 赞 / 296 评论；完整《补丁说明》84 赞 / 55 评论；按英雄拆解的《英雄增强削弱》223 赞", "案例观察（不做因果归因）", "小红书版本内容按「我的本命怎么了」切，而不是整份补丁；标题只提一个英雄或一个结论。"),
    ("J08", "小红书 XHS05 / XHS07（官方账号）", "官方公告「向违规行为重拳出击」293 赞 / 381 评论；「疯狂谜影即将结束」96 赞 / 108 评论，热评「我要这个模式常驻」", "案例观察", "官方已在小红书运营公告类内容，评论多于赞；我们的内容不重复公告，而是把公告翻译成「对回归者意味着什么」。"),
    ("J09", "抖音 搜索「守望先锋 回坑」列表", "保存的 13 条里以回坑/回流为题 5 条，其中最近一个月的两条（DY02 25 秒短片、DY03 3 分半解说）点赞均不足 100；高赞是 18 秒梗片（DY01 2.2 万）与入坑指南合集（1.7 万）", "案例观察", "两条样本不足以判断抖音回坑内容的普遍表现，本期不做是制作资源取舍；如果日后做抖音版，形态应是 15–30 秒梗向开头 + 评论区放说明书链接，本期不展开。"),
    ("J10", "抖音 DY05 / DY06（官方账号）", "官方 PV 1.5 万赞 / 8714 分享（分享低于点赞）；神话皮预览 2.4 万赞 / 3.5 万分享（分享高于点赞），热评是平衡吐槽（1772 赞）", "案例观察", "两条官方案例方向相反，不概括成平台规律；说明书里保留一句「发给回坑朋友」的分享话术，作为待验证的设计。"),
    ("J11", "抖音 DY04 / DY08", "评论区高赞多为具体问题（「竞逐打满还要自己领特签吗」48 赞）或作者逐条回复（童老六）", "案例观察", "样稿的互动设计把「作者按可投入时间回复、先回具体问题」写进执行项；FAQ 从评论区问题生成。"),
    ("J12", "小红书 XHS06 标注", "资讯号补丁图文标注「含AI生成内容」且带 #vibecoding 标签，仍获 84 赞 / 12 收藏", "案例观察", "一个案例说明带 AI 声明的图文也能获得互动，不代表平台政策；本项目若生成环节用到 AI，按发布页当时的规则声明。"),
    ("J13", "官网 24303010（11-04 季中更新）", "争霸之战 11-04 至 12-08：24 级活动战令，免费线靠游玩与完成挑战推进、推到位解锁史诗西格玛皮肤；狂鼠补给狩猎 11-21 至 24 双倍经验；OWCS 总决赛 12-02 至 06 广州", "官方规则/公告", "排期锚点：11-10 起 7 周到总决赛；选题二在 11-21 双倍经验前发，选题三在 12-02 前一周发。"),
    ("J14", "官网 S09 / S10", "返场战令可从当时进度继续、免费进度继承；新手福利限注册不足 60 天的新账号", "官方规则/公告", "说明书明确告诉回归者「哪些福利你有、哪些没有」，不制造虚假期待；核心卖点是「进度没丢」而非「送东西」。"),
]

facts = [
    ("F01", "2026-08-12", "第 4 赛季「釜山群英」上线，新重装 D.Mon；战令改版为可自选顺序的升级路线", "官网 24295381 / 24296053（S01、S02）", "官方"),
    ("F02", "2026-09-13 至 10-06", "暴雪嘉年华期间完成一局得神话代券，可换第 1–14 赛季任一神话礼包（已截止）", "官网 24294376（S07）", "官方"),
    ("F03", "2026-09-23", "国服累计禁赛 30,776 账号、回收奖励 38,757 账号（官方自报）；代练双方与炸鱼关联大号连坐", "官网 20260923/41367_1314781（S08）", "官方"),
    ("F04", "2026-10-07", "第 5 赛季「暗红之律」：新支援「血律」；黑影由输出改为支援；多年来首张新运载地图；第 1–9、15 赛季往期战令返场，可从当时进度继续，免费进度继承", "官网 24303008（S09）", "官方"),
    ("F05", "2026-10-07", "「圣邪颂歌」登录送 3 抽；注册不足 60 天的新玩家完成新手任务免费得满级神话皮肤——回归老账号不适用", "官网 269312113（S10）", "官方"),
    ("F06", "2026-11-04", "季中更新：争霸之战活动战令 11-04 至 12-08，24 级；免费线「通过游玩和完成挑战逐步赢得」、最终解锁史诗西格玛换色皮肤，不是登录即得；1000 金币升级开第二条路线，最终解锁传奇末日铁拳皮肤 / 斩仇传奇武器；伊芙计划（剑星）联动 11-04 至 11-18；死灵骑士路霸神话皮肤；轰轰钻机挖宝 11-04 至 12-08；狂鼠补给狩猎 11-21 至 24 双倍经验；竞逐 12-04 至 08；黑色星期五促销 11-04 至 12-08", "官网 24303010《黑爪之治第5赛季季中更新》", "官方"),
    ("F07", "2026-11-04 00:00 生效", "游戏环境治理专项行动：六项举措，严打代练、演员、剧组、炸鱼、外挂", "官网 20261103/41367_1318539、20261104/41365_1318531；小红书官方 XHS05", "官方"),
    ("F08", "2026-11-04 凌晨结束", "万圣节限时模式「疯狂谜影」结束；UGC 多帖要求常驻（192 / 69 赞）", "小红书官方 XHS07 + UGC 列表", "官方 + 案例观察"),
    ("F09", "2026-12-02 至 12-06", "OWCS 2026 全球总决赛在广州：02/03 演播室阶段，04–06 广州体育馆；助威活动 12-03 至 07", "广州白云区官媒（S14）；官网 24303010", "官方/官媒"),
    ("F10", "2026-10-03 至 11-08", "OWCS 中国赛区第三赛段 8 队；11-07/08 B 站出现「WBG 零封 AG」「国区冠军采访」等 UGC 标题——晋级总决赛的中国队名单以官方公布为准，本项目标「待核验」", "Liquipedia（S15）+ B 站标题观察", "第三方 + 待核验"),
    ("F11", "2026-07（第 3 赛季）", "官方披露每日模式分布约数：5v5 预设休闲 54% / 竞技 37%；6v6 开放休闲 19% / 竞技 8%；分母与统计窗口官方未说明，只用于「6v6 参与比例明显低于 5v5」，不推算排队时长", "官网 24289101（S12）", "官方"),
    ("F12", "2026-08-12、10-07", "英雄总数 52 → 54；职责 输出 24→23、支援 14→16、重装 14→15；为 8 月与 10 月两个快照之差，不是长期离开者错过的全部", "OverFast API（S16）", "第三方数据接口"),
    ("F13", "检索日 2026-10-15", "B 站官方视频号 51.1 万粉、电竞官方号 40.2 万粉；窗口内官方 17 条新手英雄教学全部面向新手，无一条面向回归者", "B 站公开接口（S17）+ 10-15 样本", "案例观察"),
]

# ---------- xlsx ----------
wb = Workbook()
H = Font(bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="7A1F2B")
thin = Side(style="thin", color="DDDDDD"); B = Border(left=thin, right=thin, top=thin, bottom=thin)
def sheet(title, header, data, widths):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]: c.font = H; c.fill = HF; c.alignment = Alignment(vertical="center", wrap_text=True)
    for row in data: ws.append(list(row))
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for r in ws.iter_rows(min_row=2):
        for c in r: c.alignment = Alignment(vertical="top", wrap_text=True); c.border = B
    ws.freeze_panes = "A2"
    return ws

wb.remove(wb.active)
sheet("说明", ["项", "内容"], [
    ("项目", "《守望先锋》回坑说明书——面向回流玩家的跨平台内容方案（作品集项目，未发布）"),
    ("观察日期", f"{OBS}（北京时间）；B 站数据为公开搜索接口匿名态读取，抖音/小红书为本人登录态人工观察"),
    ("数据性质", "全部为平台页面可见的公开计数（播放/赞/收藏/评论/分享），是检索样本不是全量；只用于判断内容形态，不推断平台总体"),
    ("三类标注", "官方规则/公告＝官网或平台可查的规则；案例观察＝本表记录的可见数据与评论；推断＝本人判断，另行标出"),
    ("不做的事", "不把平台差异归因于算法；不把播放量归因于某个标题或剪辑技巧；不编造任何未观察到的数字"),
    ("B 站样本口径", f"8 组关键词 × 3 页搜索，返回 {len(rows)} 条；按发布日期 ≥2026-10-07 过滤得 {len(recent)} 条；再剔除非 OW 内容得 {len(ow)} 条进入统计"),
    ("抖音/小红书列表口径", "抖音搜索列表保存 13 条、小红书「回坑」列表保存 19 条；占比以保存条目为分母，观察时翻到的前 20 / 前 30 条不可从记录复算"),
    ("用户 ID", "B 站数据不含 UP 主 mid（2026-11-09 已从脚本与数据文件删除）；评论数据不含评论者 UID"),
    ("来源文件", "research/bilibili_cases.jsonl、bilibili_case_comments.jsonl、manual_cases_2026-11-09.json、bilibili_fetch_log.json、sample_seven_cards.json"),
], [16, 120])

sheet("判断依据→设计", ["编号", "来源", "观察到什么", "性质", "对内容设计的影响"], judgements, [7, 30, 70, 16, 70])
sheet("版本事实核验", ["编号", "时间", "事实", "来源", "性质"], facts, [7, 20, 90, 42, 14])

sheet("B站主题统计", ["主题", "条数", "播放中位数", "赞/播放 中位", "收藏/播放 中位", "评论/播放 中位"],
      [(t["theme"], t["n"], t["median_play"], t["median_like_rate"], t["median_fav_rate"], t["median_reply_rate"]) for t in theme_rows], [20, 8, 14, 14, 16, 16])
sheet("B站样本(10-07后)", ["发布日期", "主题", "标题", "UP 主", "时长", "播放", "点赞", "收藏", "评论", "弹幕", "链接", "命中关键词", "观察日期"],
      [(r["pubdate_bjt"], r["theme"], r["title"], r["author"], r["duration"], r["play"], r["like"], r["favorites"], r["review"], r["danmaku"], r["url"],
        "；".join(sorted({h["kw"] for h in r["hit_by"]})), OBS) for r in ow], [11, 14, 60, 16, 7, 10, 8, 8, 8, 8, 40, 30, 11])
sheet("B站热评", ["视频", "BV", "热评序", "评论", "赞", "回复数", "观察日期"],
      [(c["video_title"], c["bvid"], c["rank"], c["comment"], c["like"], c["rcount"], OBS[:10]) for c in comments], [36, 14, 7, 90, 8, 8, 11])

dy = manual["douyin"]
sheet("抖音观察", ["编号", "发布时间", "作者", "标题", "时长", "点赞", "评论", "收藏", "分享", "类型", "如何找到", "观察笔记", "链接", "观察日期"],
      [(d["id"], d["published"], d["author"], d["title"], d["duration"], d["like"], d["comment"], d["collect"], d["share"], d["type"], d["found_by"], d["note"], d["url"], OBS) for d in dy],
      [6, 17, 16, 48, 7, 8, 7, 7, 8, 18, 26, 70, 40, 11])
sheet("抖音搜索列表", ["标题", "作者", "页面显示时间", "点赞", "时长", "关键词", "观察日期"],
      [(i["title"], i["author"], i["age"], i["like"], i["duration"], manual["douyin_search_listing"]["keyword"], OBS) for i in manual["douyin_search_listing"]["items"]] +
      [("【整体阅读】" + manual["douyin_search_listing"]["reading"], "", "", "", "", "", OBS)], [60, 20, 12, 8, 7, 30, 11])
xs = manual["xiaohongshu"]
sheet("小红书笔记", ["编号", "发布", "作者", "标题", "类型", "赞", "收藏", "评论", "正文摘录", "热评", "观察笔记", "链接", "观察日期"],
      [(x["id"], x["published"], x["author"], x["title"], x["type"], x.get("like"), x.get("collect"), x.get("comment"), x.get("body_excerpt", ""), x.get("top_comment", ""), x["note"], x["url"], OBS) for x in xs],
      [7, 20, 16, 40, 16, 6, 6, 6, 60, 50, 60, 44, 11])
sheet("小红书搜索列表", ["关键词", "列表（标题，作者，日期，赞）", "联想搜索", "整体阅读", "观察日期"],
      [(x["keyword"], x["items"], x.get("related_searches", ""), x.get("reading", ""), OBS) for x in manual["xiaohongshu_search_listing"]], [22, 110, 40, 60, 11])
sheet("观察限制", ["限制"], [(l,) for l in manual["limits"]], [140])
wb.save(OUT_X)

# ---------- site data ----------
site = {
    "observed_at": OBS,
    "bilibili": {"n_total": len(rows), "n_recent": len(recent), "n_ow": len(ow), "themes": theme_rows,
                  "top": [{k: r[k] for k in ("pubdate_bjt", "theme", "title", "author", "duration", "play", "like", "favorites", "review", "url")} for r in ow[:40]],
                  "returning": [{k: r[k] for k in ("pubdate_bjt", "title", "author", "play", "like", "favorites", "review", "url")} for r in ow if r["theme"] == "回归/老玩家"]},
    "comments": comments,
    "douyin": dy, "douyin_listing": manual["douyin_search_listing"],
    "xhs": xs, "xhs_listing": manual["xiaohongshu_search_listing"],
    "judgements": [dict(zip(("id", "source", "obs", "kind", "impact"), j)) for j in judgements],
    "facts": [dict(zip(("id", "when", "fact", "source", "kind"), f)) for f in facts],
    "limits": manual["limits"],
    "sample7": json.load((R / "sample_seven_cards.json").open(encoding="utf-8")),
}
OUT_J.write_text(json.dumps(site, ensure_ascii=False, indent=1), encoding="utf-8")
print("ow recent:", len(ow)); print(json.dumps(theme_rows, ensure_ascii=False, indent=1))
print("returning:", [(r["title"][:30], r["play"]) for r in site["bilibili"]["returning"]])
