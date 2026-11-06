#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""试卷分析报告批量生成器（零第三方依赖，WorkBuddy skill 主脚本）

输入（同一个考试文件夹）：
  paper.json     卷子结构：题号 / 分值 / 模块 / 知识点 / 课本位置 / 学习建议（+ 可选班均）
  scores.csv     全班逐题得分：学号,姓名,[班级,教师,层级],Q1..Qn
  comments.json  可选：{"学号": "评语"}；不存在时用 paper.json 里的分档话术模板

输出：
  reports/NN_姓名.html / .pdf      每人 2 页 A4
  reports/班级总览.json             供 AI 写班级诊断（不用回读 PDF，省额度）
  reports/全部报告.zip
  台账.csv

用法：
  python build_reports.py --exam-dir "考试文件夹"
  可选：--avg manual|auto  覆盖 paper.json 的 avg_mode
        --grade pct|quantile  覆盖等级口径
        --no-pdf  只出 HTML（调试用）
        --only 学号1,学号2  只出指定学生
"""
import argparse
import csv
import html as H
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE.parent / "assets" / "report.html"
sys.path.insert(0, str(HERE))

import brand  # noqa: E402  同目录的品牌配置（Logo / 主题色）

# 配色（DESIGN.md：瑞士教研网格，3 色 + 灰阶）。主题色可由机构 Logo 派生，见 brand.py
PALETTE = brand.palette(brand.DEFAULT)
ACCENT = PALETTE["accent"]
CONNECTOR = PALETTE["connector"]
CRITICAL = PALETTE["critical"]
MUTED = "#68707C"
BAR_IDLE = "#C9CED5"


def set_theme(pal):
    """换主题：SVG 图表与模板 token 都读这几个全局"""
    global PALETTE, ACCENT, CONNECTOR, CRITICAL
    PALETTE = pal
    ACCENT = pal["accent"]
    CONNECTOR = pal["connector"]
    CRITICAL = pal["critical"]

# 第二页起每页可用高度（mm）：按每题文字量估算行高装页，题多就续页，既不缩字号也不裁字
DETAIL_BUDGET_MM = 226.0

EDGE_CANDIDATES = [
    os.environ.get("EDGE_PATH", ""),
    r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    r"C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    r"C:/Program Files/Google/Chrome/Application/chrome.exe",
    r"C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/microsoft-edge",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
]

# PATH 上的可执行名（which/where 兜底）：用户把 Edge/Chrome 装在非标准位置时的最后一道
EDGE_NAMES = ["msedge", "chrome", "google-chrome", "microsoft-edge", "chromium", "chromium-browser"]


# ---------------------------------------------------------------- 读取与校验
def load_json(path, default=None):
    p = Path(path)
    if not p.exists():
        return default
    return json.loads(p.read_text(encoding="utf-8-sig"))


def load_scores(path):
    p = Path(path)
    if p.suffix.lower() == ".json":
        data = json.loads(p.read_text(encoding="utf-8-sig"))
        rows = []
        for s in data.get("students", []):
            row = {"学号": str(s.get("学号", "")).strip(), "姓名": str(s.get("姓名", "")).strip(),
                   "班级": s.get("班级", ""), "教师": s.get("教师", ""), "层级": s.get("层级", ""),
                   "校区": s.get("校区", ""), "缺考": "1" if s.get("缺考") else "",
                   "备注": s.get("note", "") or s.get("备注", ""),
                   "__errors": json.dumps(s.get("errors") or {}, ensure_ascii=False)}
            for no, v in (s.get("scores") or {}).items():
                row[f"Q{no}"] = "" if v is None else str(v)
            rows.append(row)
        if not rows:
            raise SystemExit(f"[错误] {path} 里没有学生数据")
        return rows
    with open(p, encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        raise SystemExit(f"[错误] {path} 没有数据行")
    return rows


def load_manual_avg(path):
    """录分结果.json 里手工录入的班均（自动开关关闭时才有意义）"""
    p = Path(path)
    if p.suffix.lower() != ".json" or not p.exists():
        return {}
    data = json.loads(p.read_text(encoding="utf-8-sig"))
    out = {}
    for q in data.get("questions", []):
        if q.get("avg") is not None:
            out[str(q["no"])] = float(q["avg"])
    return out


def norm_key(row):
    """把学号/姓名列名做归一（容忍 学号/考号/姓名/学生 等写法）"""
    out = {}
    for k, v in row.items():
        if k is None:
            continue
        kk = k.strip()
        out[kk] = (v or "").strip()
    return out


def get_field(row, *names):
    for n in names:
        if n in row and row[n] != "":
            return row[n]
    return ""


def parse_qmap(paper):
    """题号 → 满分，保持卷面顺序"""
    qs = paper["questions"]
    qmap = {}
    for q in qs:
        no = str(q["no"]).strip()
        qmap[no] = float(q["full"])
    return qmap


# ---------------------------------------------------------------- 统计
def grade_of_pct(rate, thresholds):
    for g in ("A", "B", "C", "D"):
        if rate >= thresholds[g]:
            return g
    return "E"


def grade_of_quantile(rank_index, n):
    """按名次分位 10/20/40/20/10 → A/B/C/D/E"""
    cuts = [(0.10, "A"), (0.30, "B"), (0.70, "C"), (0.90, "D")]
    pos = (rank_index + 0.5) / max(n, 1)
    for c, g in cuts:
        if pos <= c:
            return g
    return "E"


def short_name(name, limit=6):
    """雷达轴名兜底：去掉"的X"后缀，仍超长则截断。优先用 paper.json 的 module_shorts 显式指定。"""
    n = re.sub(r"的(应用|探究|计算|理解|分析)$", "", name)
    if len(n) <= limit:
        return n
    return n[: limit - 2] + "…"


def compute(paper, rows, avg_mode, grade_mode, manual_avg=None):
    qmap = parse_qmap(paper)
    qnos = list(qmap.keys())
    thresholds = paper.get("grades", {"A": 0.90, "B": 0.70, "C": 0.30, "D": 0.10})

    students = []
    for raw in rows:
        r = norm_key(raw)
        sid = get_field(r, "学号", "考号", "编号", "ID")
        name = get_field(r, "姓名", "学生", "学生姓名")
        absent = get_field(r, "缺考", "absent").lower() in ("1", "y", "yes", "是", "true")
        scores = {}
        for no in qnos:
            v = get_field(r, f"Q{no}", f"题{no}", no)
            if v == "" or v in ("-", "缺", "/"):
                scores[no] = None
            else:
                scores[no] = float(v)
        raw_err = get_field(r, "__errors", "errors")
        try:
            own = json.loads(raw_err) if raw_err else {}
        except Exception:
            own = {}
        students.append({
            "sid": sid, "name": name, "absent": absent, "errors": own,
            "cls": get_field(r, "班级", "班型"), "teacher": get_field(r, "教师", "任课教师"),
            "level": get_field(r, "层级"), "campus": get_field(r, "校区"),
            "note": get_field(r, "备注", "note", "学生备注"),
            "scores": scores,
        })

    live = [s for s in students if not s["absent"]]
    if not live:
        raise SystemExit("[错误] 没有非缺考学生")

    total_full = sum(qmap.values())
    declared = paper.get("exam", {}).get("full")
    if declared and abs(float(declared) - total_full) > 1e-6:
        print(f"[警告] 题目分值合计 {total_full} ≠ 卷面满分 {declared}")

    # 班均：auto = 用本班实际得分算；manual = 用 paper.json 每题自带 avg
    q_avg = {}
    if avg_mode == "auto":
        for no in qnos:
            vals = [s["scores"][no] for s in live if s["scores"][no] is not None]
            q_avg[no] = round(sum(vals) / len(vals), 2) if vals else 0.0
    else:
        manual_avg = manual_avg or {}
        for q, no in zip(paper["questions"], qnos):
            v = manual_avg.get(no, q.get("avg"))
            if v is None:
                raise SystemExit(f"[错误] avg_mode=manual，但第 {no} 题既无手工班均也无 avg 字段")
            q_avg[no] = float(v)

    # 模块归属
    mods = []
    seen = {}
    mshorts = paper.get("module_shorts") or {}
    for q in paper["questions"]:
        m = q["module"]
        if m not in seen:
            seen[m] = {"name": m,
                       "short": q.get("module_short") or mshorts.get(m) or short_name(m),
                       "full": 0.0, "nos": []}
            mods.append(seen[m])
        seen[m]["full"] += float(q["full"])
        seen[m]["nos"].append(str(q["no"]))
    for m in mods:
        m["avg"] = round(sum(q_avg[n] for n in m["nos"]), 1)

    # 显示用模块组：试卷结构与模块对比图共用，≤4 个全留，超出按分值保留前 4 + 其他
    groups = group_modules(mods, 4)

    # 每生汇总
    for s in students:
        if s["absent"]:
            s["total"] = None
            continue
        s["total"] = round(sum(v for v in s["scores"].values() if v is not None), 1)
        s["rate"] = s["total"] / total_full if total_full else 0
        s["mods"] = []
        for m in mods:
            got = round(sum(s["scores"][n] or 0 for n in m["nos"]), 1)
            s["mods"].append({"name": m["name"], "short": m["short"], "full": m["full"], "avg": m["avg"],
                              "score": got, "gap": round(got - m["avg"], 1),
                              "rate": got / m["full"] if m["full"] else 0,
                              "avg_rate": m["avg"] / m["full"] if m["full"] else 0})
        s["lost"] = sorted(            [{"no": n, "lost": qmap[n] - (s["scores"][n] or 0), "got": s["scores"][n] or 0,
              "full": qmap[n],
              **{k: v for k, v in next(q for q in paper["questions"] if str(q["no"]) == n).items()
                 if k in ("module", "kp", "book", "advice", "skill")}}
             for n in qnos if (s["scores"][n] or 0) < qmap[n]],
            key=lambda x: -x["lost"])
        # 与首页「试卷结构」表同源的模块组（含「其他」合并行）
        s["groups"] = []
        for g in groups:
            got = round(sum(s["scores"][n] or 0 for n in g["nos"]), 1)
            avg = sum(q_avg[n] for n in g["nos"])
            s["groups"].append({
                "name": g["name"], "short": g["short"], "full": g["full"], "nos": g["nos"],
                "score": got, "avg": round(avg, 1),
                "rate": got / g["full"] if g["full"] else 0,
                "avg_rate": avg / g["full"] if g["full"] else 0,
            })

    # 等级
    if grade_mode == "quantile":
        ordered = sorted(live, key=lambda s: -s["total"])
        for i, s in enumerate(ordered):
            s["grade"] = grade_of_quantile(i, len(live))
    else:
        for s in live:
            s["grade"] = grade_of_pct(s["rate"], thresholds)
    for s in live:
        for m in s["mods"]:
            m["grade"] = grade_of_pct(m["rate"], thresholds)

    # 班级统计
    totals = [s["total"] for s in live]
    band_def = paper.get("bands") or [[0, 0.3], [0.3, 0.4], [0.4, 0.5], [0.5, 0.6],
                                      [0.6, 0.7], [0.7, 0.8], [0.8, 1.0]]
    # 上界 ≤1 视为"占满分比例"，换算成绝对分（换任何满分都对）
    if max(float(hi) for _, hi in band_def) <= 1.0 + 1e-9:
        band_def = [[round(float(lo) * total_full, 2), round(float(hi) * total_full, 2)]
                    for lo, hi in band_def]
    counts = []
    for lo, hi in band_def:
        counts.append(sum(1 for t in totals if lo <= t < hi or (hi == band_def[-1][1] and t == hi)))

    tag_labels, tag_defs = load_tags()
    # 每生：知识点掌握 + 错因构成
    for s_ in live:
        s_["kps"] = kp_mastery(paper, qmap, q_avg, lambda no, s_=s_: (s_["scores"].get(no) or 0))
        s_["errs"] = err_breakdown(paper, lambda no, s_=s_: (s_["scores"].get(no) or 0), tag_labels)
        # 教师逐题标的学生级错因（录分结果里可选带 errors）
        s_["own_errs"] = {str(k): v for k, v in (s_.get("errors") or {}).items()}

    # 班级：知识点掌握（用班均当"班级个体"，再看每题得分率）
    class_kps = kp_mastery(paper, qmap, q_avg, lambda no: q_avg.get(no, 0.0))
    diff_bands = {"easy": [], "medium": [], "hard": []}
    for no in qnos:
        lv = diff_level(q_avg[no] / qmap[no] if qmap[no] else 0)
        diff_bands[lv].append(no)
    q_rates = {no: (q_avg[no] / qmap[no] if qmap[no] else 0) for no in qnos}
    ordered = sorted(qnos, key=lambda n: q_rates[n])
    class_errs = err_breakdown(paper, lambda no: q_avg.get(no, 0.0), tag_labels)

    class_stat = {
        "考试": paper.get("exam", {}).get("name", ""),
        "人数": len(live),
        "缺考": [s["name"] for s in students if s["absent"]],
        "班均": round(sum(totals) / len(totals), 1),
        "最高": max(totals), "最低": min(totals),
        "满分": total_full,
        "等级分布": {g: sum(1 for s in live if s["grade"] == g) for g in "ABCDE"},
        "模块班均": {m["name"]: {"满分": m["full"], "班均": m["avg"],
                             "得分率": round(m["avg"] / m["full"] * 100, 1) if m["full"] else 0}
                 for m in mods},
        "模块最弱": min(mods, key=lambda m: (m["avg"] / m["full"]) if m["full"] else 1)["name"],
        "逐题班均": {no: q_avg[no] for no in qnos},
        "逐题难度": {no: round(q_avg[no] / qmap[no], 2) if qmap[no] else 0 for no in qnos},
        "知识点掌握": [{"kp": k["kp"], "题号": k["nos"], "满分": k["full"],
                    "班均": round(k["avg"], 1), "得分率": round(k["rate"] * 100, 1),
                    "水平": k["level"], "常见错因": [tag_labels.get(c, c) for c in k["errors"]]}
                   for k in class_kps],
        "难度分布": {DIFF_NAME[k]: v for k, v in diff_bands.items()},
        "最难题": [{"no": n, "得分率": round(q_rates[n] * 100, 1),
                  "kp": next((q.get("kp", "") for q in paper["questions"] if str(q["no"]) == n), "")}
                 for n in ordered[:3]],
        "最易题": [{"no": n, "得分率": round(q_rates[n] * 100, 1),
                  "kp": next((q.get("kp", "") for q in paper["questions"] if str(q["no"]) == n), "")}
                 for n in ordered[-3:][::-1]],
        "错因排行": [{"错因": e["label"], "涉及题": e["nos"], "涉及分值": e["lost"]} for e in class_errs],
    }
    return {"students": students, "live": live, "mods": mods, "groups": groups,
            "qmap": qmap, "qnos": qnos,
            "q_avg": q_avg, "total_full": total_full, "class": class_stat,
            "avg_mode": avg_mode, "grade_mode": grade_mode, "bands": band_def,
            "band_counts": counts, "tag_defs": tag_defs, "tag_labels": tag_labels}


# ---------------------------------------------------------------- 版式：容量规则
def format_question_nos(nos, max_chars=26):
    """题号先合并连续区间（1–6、8、11），超长再压成「1–6、8 等 16 题」"""
    texts = [str(n) for n in nos]
    if not texts:
        return "—"
    if all(t.isdigit() for t in texts):
        values = sorted(int(t) for t in texts)
        parts, start, prev = [], values[0], values[0]
        for v in values[1:] + [None]:
            if v is not None and v == prev + 1:
                prev = v
                continue
            parts.append(str(start) if start == prev else f"{start}–{prev}")
            if v is not None:
                start = prev = v
    else:
        parts = texts
    full = "、".join(parts)
    if len(full) <= max_chars:
        return full
    suffix = f"等{len(texts)}题"
    kept = []
    for part in parts:
        if len("、".join(kept + [part])) + 1 + len(suffix) > max_chars:
            break
        kept.append(part)
    return ("、".join(kept) + " " if kept else "") + suffix


def group_modules(mods, max_single=4):
    """试卷结构与模块图共用同一组模块；超过 4 个按分值保留前 4，其余并入「其他」"""
    if len(mods) <= max_single:
        return [dict(m) for m in mods]
    keep = {m["name"] for m in sorted(mods, key=lambda x: -x["full"])[:max_single]}
    shown = [dict(m) for m in mods if m["name"] in keep]
    rest = [m for m in mods if m["name"] not in keep]
    other = {
        "name": "其他", "short": "其他",
        "full": sum(m["full"] for m in rest),
        "avg": sum(m["avg"] for m in rest),
        "nos": [n for m in rest for n in m["nos"]],
    }
    return shown + [other]


# ---------------------------------------------------------------- SVG 图表
def svg_histogram(bands, w=340, h=420):
    """班级成绩分布直方图：「你在这里」用箭头+文字标注，不单靠颜色区分（黑白可打印）"""
    left, right, top, bottom = 30, 8, 22, 52
    plot_w, plot_h = w - left - right, h - top - bottom
    max_count = max(4, max(b["count"] for b in bands)) if bands else 4
    step = plot_w / max(len(bands), 1)
    bar_w = min(26, step * 0.55)
    parts = [f'<svg viewBox="0 0 {w} {h}" class="chart-svg" aria-label="班级成绩分布直方图">']
    parts.append(f'<text x="{left}" y="12" class="axis-name">人数</text>')
    for tick in range(0, max_count + 1, max(1, round(max_count / 4))):
        y = top + plot_h - plot_h * tick / max_count
        parts.append(f'<line x1="{left}" y1="{y:.1f}" x2="{w-right}" y2="{y:.1f}" class="grid-line"/>')
        parts.append(f'<text x="{left-7}" y="{y+4:.1f}" text-anchor="end" class="tick">{tick}</text>')
    for i, b in enumerate(bands):
        x = left + step * i + step / 2
        bh = plot_h * b["count"] / max_count
        y = top + plot_h - bh
        fill = ACCENT if b["active"] else BAR_IDLE
        parts.append(f'<rect x="{x-bar_w/2:.1f}" y="{y:.1f}" width="{bar_w:.1f}" height="{bh:.1f}" fill="{fill}"/>')
        parts.append(f'<text x="{x:.1f}" y="{y-6:.1f}" text-anchor="middle" class="bar-value">{b["count"]}</text>')
        parts.append(f'<text x="{x:.1f}" y="{top+plot_h+17:.1f}" text-anchor="middle" class="tick">{b["label"]}</text>')
        if b["active"]:
            parts.append(f'<path d="M {x-4:.1f} {top+plot_h+23:.1f} L {x+4:.1f} {top+plot_h+23:.1f} '
                         f'L {x:.1f} {top+plot_h+29:.1f} Z" fill="{ACCENT}"/>')
            parts.append(f'<text x="{x:.1f}" y="{top+plot_h+42:.1f}" text-anchor="middle" class="self-label">你在这里</text>')
    parts.append(f'<line x1="{left}" y1="{top+plot_h}" x2="{w-right}" y2="{top+plot_h}" class="axis-line"/>')
    parts.append("</svg>")
    return "".join(parts)


def svg_module_compare(groups, w=340, h=420):
    """模块得分对比点线图：实心点＝个人，空心点＝班均（形状区分，黑白也能分）"""
    x0, x1 = 96, 306
    top = 52
    # 行距按画布高度摊开，模块少时不至于挤在上半页
    row_gap = max(58, min(76, (h - top - 34) / max(1, len(groups) - 1)))
    scale = lambda v: x0 + (x1 - x0) * max(0.0, min(1.0, v))
    parts = [f'<svg viewBox="0 0 {w} {h}" class="chart-svg" aria-label="模块得分率个人与班均对比">']
    parts.append(f'<circle cx="150" cy="12" r="4" fill="{ACCENT}"/><text x="159" y="16" class="legend">个人</text>')
    parts.append(f'<circle cx="208" cy="12" r="4" fill="#FFFFFF" stroke="{MUTED}" stroke-width="1.5"/>'
                 f'<text x="217" y="16" class="legend">班均</text>')
    ax_y = top + row_gap * (len(groups) - 1) + 8
    for tick in (0, 50, 100):
        x = scale(tick / 100)
        parts.append(f'<line x1="{x:.1f}" y1="26" x2="{x:.1f}" y2="{ax_y:.1f}" class="grid-line"/>')
        parts.append(f'<text x="{x:.1f}" y="{ax_y + 20:.1f}" text-anchor="middle" class="tick">{tick}%</text>')
    for i, g in enumerate(groups):
        y = top + i * row_gap
        pr, cr = g["rate"], g["avg_rate"]
        px, cx = scale(pr), scale(cr)
        label = g["short"] if len(g["short"]) <= 6 else g["short"][:5] + "…"
        parts.append(f'<text x="0" y="{y+4}" class="row-label">{label}</text>')
        parts.append(f'<line x1="{min(px,cx):.1f}" y1="{y}" x2="{max(px,cx):.1f}" y2="{y}" '
                     f'stroke="{CONNECTOR}" stroke-width="3"/>')
        parts.append(f'<circle cx="{cx:.1f}" cy="{y}" r="5" fill="#FFFFFF" stroke="{MUTED}" stroke-width="1.6"/>')
        parts.append(f'<circle cx="{px:.1f}" cy="{y}" r="5.5" fill="{ACCENT}"/>')
        parts.append(f'<text x="{px:.1f}" y="{y-9:.1f}" text-anchor="middle" class="point-self">{pr*100:.0f}%</text>')
        parts.append(f'<text x="{cx:.1f}" y="{y+15:.1f}" text-anchor="middle" class="point-class">{cr*100:.0f}%</text>')
    parts.append("</svg>")
    return "".join(parts)


def logo_markup(info):
    """机构 Logo 转 data URI 内联（file:// 打印时不能依赖相对路径）"""
    path = (info or {}).get("logo")
    if not path:
        return ""
    import base64
    p = Path(path)
    try:
        raw = p.read_bytes()
    except OSError:
        return ""
    mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".svg": "image/svg+xml", ".webp": "image/webp", ".bmp": "image/bmp",
            ".gif": "image/gif"}.get(p.suffix.lower(), "image/png")
    b64 = base64.b64encode(raw).decode("ascii")
    return f'<img src="data:{mime};base64,{b64}" alt="机构 Logo">'


def band_view(st, total):
    """分段标签 + 人数 + 是否本人所在段"""
    bands = st["bands"]
    counts = st["band_counts"]
    full = st["total_full"]
    out = []
    for i, (lo, hi) in enumerate(bands):
        if i == 0:
            label = f"<{fmt(hi,0)}"
        elif i == len(bands) - 1:
            label = f"{fmt(lo,0)}–{fmt(full,0)}"
        else:
            label = f"{fmt(lo,0)}–{fmt(hi,0)}"
        active = lo <= total < hi or (i == len(bands) - 1 and total == hi)
        out.append({"label": label, "count": counts[i], "active": active})
    return out



# ---------------------------------------------------------------- 渲染
def rate_cell(rate, warn):
    pct = round(rate * 100)
    return (f'<div class="rate"><div class="track"><i class="{"warn" if warn else ""}" '
            f'style="width:{pct}%"></i></div><b>{pct}%</b></div>')


def fmt(x, nd=1):
    return f"{x:.{nd}f}"


def fmt_plain(x):
    """80.0 → 80；2.5 → 2.5（评语里不出现无意义小数）"""
    x = float(x)
    return f"{x:g}" if abs(x - round(x)) < 1e-9 else f"{x:.1f}"


DIFF_NAME = {"easy": "易", "medium": "中", "hard": "难"}


def lv_text(level, n_q):
    """水平列统一口径：≥3 题显示「档位·n题」，1–2 题只显示「n题」（样本不够判档）"""
    return level if level.endswith("题") else f"{level}·{n_q}题"


def mastery_level(rate, n_q):
    """掌握度分档（竞品通口径）：≥80% 掌握 / 60–80% 待巩固 / <60% 薄弱；题量 <3 标样本不足"""
    if n_q < 3:
        return f"{n_q}题"          # 单卷里一个知识点常只对应 1–2 题，硬套三档会满屏"样本不足"
    if rate >= 0.80:
        return "掌握"
    if rate >= 0.60:
        return "待巩固"
    return "薄弱"


def diff_level(rate):
    """难度分档（易 0.75–0.95 / 中 0.60–0.74 / 难 0.20–0.59），用班均得分率换算"""
    if rate >= 0.75:
        return "easy"
    if rate >= 0.60:
        return "medium"
    return "hard"


def kp_mastery(paper, qmap, q_avg, score_of):
    """按知识点聚合（同一 kp 多题合并）。score_of(no) 返回该生该题得分"""
    groups = {}
    for q in paper["questions"]:
        kp = (q.get("kp") or "").strip()
        if not kp:
            continue
        g = groups.setdefault(kp, {"kp": kp, "kp_id": q.get("kp_id", ""), "nos": [],
                                   "full": 0.0, "avg": 0.0, "score": 0.0,
                                   "errors": [], "book": q.get("book", "")})
        no = str(q["no"])
        g["nos"].append(no)
        g["full"] += float(q["full"])
        g["avg"] += q_avg.get(no, 0.0)
        g["score"] += score_of(no)
        for c in q.get("errors") or []:
            if c not in g["errors"]:
                g["errors"].append(c)
    out = []
    for g in groups.values():
        g["rate"] = g["score"] / g["full"] if g["full"] else 0
        g["avg_rate"] = g["avg"] / g["full"] if g["full"] else 0
        g["lost"] = round(g["full"] - g["score"], 1)
        g["level"] = mastery_level(g["rate"], len(g["nos"]))
        out.append(g)
    out.sort(key=lambda x: (x["rate"], -x["lost"]))
    return out


def err_breakdown(paper, by_no_score, tag_labels):
    """错因归因：把「未满分题目」的题级常见错因按丢分加权"""
    agg = {}
    for q in paper["questions"]:
        no = str(q["no"])
        lost = float(q["full"]) - by_no_score(no)
        if lost <= 1e-9:
            continue
        for code in q.get("errors") or []:
            a = agg.setdefault(code, {"code": code, "label": tag_labels.get(code, code),
                                      "lost": 0.0, "nos": []})
            a["lost"] = round(a["lost"] + lost, 1)
            if no not in a["nos"]:
                a["nos"].append(no)
    return sorted(agg.values(), key=lambda x: -x["lost"])


def load_tags():
    f = HERE.parent / "assets" / "error-tags.json"
    if not f.exists():
        return {}, {}
    tags = json.loads(f.read_text(encoding="utf-8"))["tags"]
    return {t["code"]: t["label"] for t in tags}, {t["code"]: t for t in tags}


def gap_pp(m):
    """得分率差（百分点）：唯一可跨模块比较的强弱口径——模块满分不同，绝对分差不具可比性"""
    return (m["rate"] - m["avg_rate"]) * 100


def strong_mods_of(s):
    """真优势：得分率领先班均 ≥5 个百分点，或得分率本身 ≥85%。差 0.x 分不算优势。"""
    return [m for m in s["mods"] if gap_pp(m) >= 5 or m["rate"] >= 0.85]


def weak_by_rate(s):
    return min(s["mods"], key=gap_pp)


def weak_by_points(s):
    """丢分最多（满分−得分，绝对分），与"补分效率"口径一致；不是与班均差"""
    return max(s["mods"], key=lambda m: m["full"] - m["score"])


def lost_of(m):
    return m["full"] - m["score"]


def plan_detail_pages(qs, st, s):
    """按每题实际要显示的字数估行高，再把题目均匀装进各页。

    估算口径：知识点 13 字/行、常见失分点 10 字/行、复习建议 21 字/行（按 210mm 宽、三个列宽
    比例折算）。行高基准 7.2mm，多一行加 3.8mm。先按总高算出页数，再按「每页目标高度」均分，
    避免最后一页只剩一两行；任何一页都不超过可用高度，宁可续页也不缩字号。
    """
    hs = []
    for q in qs:
        got = s["scores"].get(str(q["no"]))
        expired = got is None or got < float(q["full"])
        err = adv = ""
        if expired:
            err = (q.get("errors_note") or "").strip() or "、".join(
                st["tag_labels"].get(c, c) for c in (q.get("errors") or []))
            adv = "　".join(x for x in [(q.get("book") or "").strip(),
                                       (q.get("advice") or "").strip()] if x)
        lines = max(1,
                    math.ceil(len(q.get("kp") or "—") / 13),
                    math.ceil(len(err or "—") / 10),
                    math.ceil(len(adv or "—") / 21))
        hs.append(7.2 + 3.8 * (lines - 1))

    total = sum(hs)
    n = max(1, math.ceil(total / DETAIL_BUDGET_MM))
    target = total / n
    limit = min(DETAIL_BUDGET_MM, max(target * 1.06, max(hs) if hs else 0))

    pages, cur, used = [], [], 0.0
    for q, h in zip(qs, hs):
        if cur and used + h > limit and len(pages) < n - 1:
            pages.append(cur)
            cur, used = [], 0.0
        cur.append(q)
        used += h
    if cur:
        pages.append(cur)
    return pages or [[]]


def render_pages(paper, st, s, panel_text, comment, template_text, info=None):
    """第 1 页结论与图表，第 2 页起题目明细（按 DETAIL_ROWS 续页，页码动态）"""
    exam = paper.get("exam", {})
    cls = st["class"]
    total = s["total"]

    i0 = template_text.index("<!--DETAIL-->")
    i1 = template_text.index("<!--/DETAIL-->")
    head = template_text[:i0]
    dtail = template_text[i0 + len("<!--DETAIL-->"):i1]
    tail = template_text[i1 + len("<!--/DETAIL-->"):]

    # ---- 第 1 页 ----
    struct_rows = "".join(
        f'<tr><td>{H.escape(g["name"])}</td>'
        f'<td class="question-nos">{H.escape(format_question_nos(g["nos"]))}</td>'
        f'<td class="num">{fmt(g["full"], 0)}</td></tr>'
        for g in s["groups"])

    comment_text = (panel_text or "").strip()
    if not comment_text:
        comment_text = ("这次卷面上能看出的主要问题集中在少数几个模块，逐题得分与班均的差距见第 2 页。"
                        "先按失分最多的两道题把对应知识点重做一遍，再看整套题的错因有没有共性。")

    exam_line = "".join(
        f"<span>{label}：{H.escape(value)}</span>"
        for label, value in [
            ("年级", exam.get("grade", "")),
            ("班级", s.get("cls") or exam.get("class", "")),
            ("科目", exam.get("subject", "")),
            ("授课教师", s.get("teacher") or exam.get("teacher", "")),
            ("校区", exam.get("campus", "")),
            ("考试日期", exam.get("date", "")),
        ] if value)

    base = {
        "LOGO": logo_markup(info),
        "ACCENT": PALETTE["accent"],
        "ACCENT_SOFT": PALETTE["soft"],
        "ACCENT_LINE": PALETTE["line"],
        "ACCENT_BADGE": PALETTE["badge"],
        "ACCENT_CONNECTOR": PALETTE["connector"],
        "STUDENT_NAME": H.escape(s["name"]),
        # 明细页标题：老师填了试卷名称就用它，没填退回「年级科目 试卷分析报告」
        "EXAM_TITLE": H.escape(exam.get("title_given") or exam.get("title", "试卷分析报告")),
        # 页眉右侧：试卷名称只在填了的时候出现，空着就只留「试卷分析报告」
        "EXAM_HEAD": (f'<span class="exam-head">{H.escape(exam["title_given"])}</span>'
                      if exam.get("title_given") else ""),
        "EXAM_LINE": exam_line,
        "TOTAL": fmt(total, 0),
        "FULL": fmt(st["total_full"], 0),
        "GRADE": s["grade"],
        "COMMENT": H.escape(comment_text),
        "STRUCT_ROWS": struct_rows,
        "CLASS_COUNT": str(cls["人数"]),
        "CLASS_AVG": fmt(cls["班均"]),
        "CLASS_MAX": fmt(cls["最高"], 0),
        "CLASS_MIN": fmt(cls["最低"], 0),
        "BAND_SVG": svg_histogram(band_view(st, total)),
        "MODULE_SVG": svg_module_compare(s["groups"]),
    }

    # ---- 第 2 页起：题目明细 ----
    qs = paper["questions"]
    chunks = plan_detail_pages(qs, st, s)
    total_pages = 1 + len(chunks)

    first_repl = dict(base)
    first_repl.update({"PAGE_NO": "1", "PAGE_TOTAL": str(total_pages)})
    pages = [fill_tokens(head, first_repl)]

    for idx, chunk in enumerate(chunks):
        rows = []
        for q in chunk:
            no = str(q["no"])
            got = s["scores"].get(no)
            gotv = 0.0 if got is None else got
            full = float(q["full"])
            expired = gotv < full
            err_txt = ""
            advice = ""
            if expired:
                err_txt = (q.get("errors_note") or "").strip() or "、".join(
                    st["tag_labels"].get(c, c) for c in (q.get("errors") or []))
                advice = "　".join(x for x in [(q.get("book") or "").strip(),
                                              (q.get("advice") or "").strip()] if x)
            score_txt = fmt_plain(gotv)
            score_cell = (f'<span class="lost-score">{score_txt}</span>' if expired else score_txt) \
                + f" / {fmt_plain(full)}"
            advice_cell = f'<span class="book">{H.escape(advice)}</span>' if advice else "—"
            rows.append(
                f'<tr class="{"lost" if expired else ""}">'
                f'<td class="q-no">{H.escape(no)}</td>'
                f'<td class="num">{score_cell}</td>'
                f'<td class="num">{fmt(st["q_avg"][no], 2)}</td>'
                f'<td>{H.escape(q.get("kp") or "—")}</td>'
                f'<td>{H.escape(err_txt) if err_txt else "—"}</td>'
                f'<td>{advice_cell}</td>'
                f'</tr>')

        last = idx == len(chunks) - 1
        cmt_block = ""
        if comment and last:
            cmt_block = f'<div class="cmt"><h3>老师的话</h3><p>{H.escape(comment)}</p></div>'
        repl = dict(base)
        repl.update({
            "QT_ROWS": "".join(rows),
            "COMMENT_BLOCK": cmt_block,
            "PAGE_NO": str(idx + 2),
            "PAGE_TOTAL": str(total_pages),
        })
        pages.append(fill_tokens(dtail, repl))
    return "".join(pages) + fill_tokens(tail, {})


def fill_tokens(text, mapping):
    for k, v in mapping.items():
        text = text.replace(f"@@{k}@@", str(v))
    return text


# ---------------------------------------------------------------- PDF
def find_edge():
    for p in EDGE_CANDIDATES:
        if p and Path(p).exists():
            return p
    # 非标准安装位置：从 PATH 上找可执行名
    for name in EDGE_NAMES:
        hit = shutil.which(name)
        if hit:
            return hit
    raise SystemExit(
        "[错误] 找不到 Edge/Chrome，无法输出 PDF。\n"
        "  装一个 Edge（Windows/macOS 一般自带），或用环境变量指定绝对路径：\n"
        "    Windows: set EDGE_PATH=C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe\n"
        "    macOS/Linux: export EDGE_PATH=/Applications/Microsoft\\ Edge.app/Contents/MacOS/Microsoft\\ Edge\n"
        "  只想先看版式不外发：加 --no-pdf，只出 HTML。"
    )


def print_pdf(edge, html_path, pdf_path, profile_dir):
    # 先删旧文件：Edge 覆盖不了「正被 PDF 阅读器/预览窗口打开」的目标，会静默失败，
    # 留下上一次的 PDF —— 报告改了却看不到变化，就是这么来的。
    if pdf_path.exists():
        try:
            pdf_path.unlink()
        except OSError:
            raise SystemExit(
                f"[错误] {pdf_path.name} 正被其他程序占用（PDF 阅读器或预览窗口），"
                f"关掉它再重跑。不关的话打出来的是上一版报告。")
    cmd = [edge, "--headless=new", "--disable-gpu", "--no-first-run",
           f"--user-data-dir={profile_dir}",
           "--no-pdf-header-footer", "--virtual-time-budget=3000",
           f"--print-to-pdf={str(pdf_path)}", Path(html_path).resolve().as_uri()]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
    if not Path(pdf_path).exists() or Path(pdf_path).stat().st_size < 2000:
        raise SystemExit(f"[错误] PDF 生成失败：{pdf_path}\n{(r.stderr or '')[-800:]}")


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exam-dir", required=True, help="含 paper.json / scores.csv 的考试文件夹")
    ap.add_argument("--avg", choices=["auto", "manual"])
    ap.add_argument("--grade", choices=["pct", "quantile"])
    ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--only", default="")
    args = ap.parse_args()

    exam_dir = Path(args.exam_dir).resolve()
    paper = load_json(exam_dir / "paper.json")
    if not paper:
        raise SystemExit(f"[错误] 缺少 {exam_dir/'paper.json'}")
    if not paper.get("questions"):
        raise SystemExit("[错误] paper.json 里没有 questions")

    # 机构 Logo 与主题色：默认蓝；paper.json 里有 brand 就用它（brand.py 负责写入）
    pal, brand_info = brand.resolve(paper, exam_dir)
    set_theme(pal)
    for i, line in enumerate(brand.describe(brand_info).splitlines()):
        print(("[品牌] " if i == 0 else "        ") + line)

    # 录分来源：优先 录分结果.json（界面导出，自带班均开关状态），否则 scores.csv
    sheet = exam_dir / "录分结果.json"
    src = sheet if sheet.exists() else exam_dir / "scores.csv"
    if not src.exists():
        raise SystemExit(f"[错误] 找不到录分数据：{exam_dir/'录分结果.json'} 或 {exam_dir/'scores.csv'}")
    rows = load_scores(src)
    manual_avg = load_manual_avg(src)
    sheet_mode = None
    if src == sheet:
        sheet_mode = (json.loads(sheet.read_text(encoding="utf-8-sig")) or {}).get("avg_mode")

    avg_mode = args.avg or sheet_mode or paper.get("avg_mode", "auto")
    grade_mode = args.grade or paper.get("grade_mode", "pct")
    st = compute(paper, rows, avg_mode, grade_mode, manual_avg)
    print(f"[录分来源] {src.name}　[班均] {avg_mode}　[等级] {grade_mode}")

    comments = load_json(exam_dir / "comments.json", {}) or {}
    panels = load_json(exam_dir / "panels.json", {}) or {}
    st["panels"] = panels.get("students") or {}
    st["panel_class"] = panels.get("class") or {}
    template_text = TEMPLATE.read_text(encoding="utf-8")
    out_dir = exam_dir / "reports"
    out_dir.mkdir(exist_ok=True)

    only = {x.strip() for x in args.only.split(",") if x.strip()}
    edge = None if args.no_pdf else find_edge()
    profile = Path(tempfile.mkdtemp(prefix="wb_pdf_"))
    made, pages_of = [], {}
    for s in st["students"]:
        if s["absent"]:
            continue
        if only and s["sid"] not in only and s["name"] not in only:
            continue
        # 老师的话只在老师逐人写了才出现；否则不生成
        comment = comments.get(s["sid"]) or comments.get(s["name"]) or ""
        # 个人点评：panels.json → students.<学号>.comment（v0.6 起报告只有这一段 AI 文字）
        pn = (panels.get("students") or {}).get(s["sid"]) \
            or (panels.get("students") or {}).get(s["name"]) or {}
        panel_text = pn if isinstance(pn, str) else (pn.get("comment") or pn.get("personal") or "")
        html = render_pages(paper, st, s, panel_text, comment, template_text, brand_info)
        stem = f'{s["sid"]}_{s["name"]}'.strip("_") or s["name"]
        hp = out_dir / f"{stem}.html"
        hp.write_text(html, encoding="utf-8")
        if edge:
            print_pdf(edge, hp, out_dir / f"{stem}.pdf", profile)
        made.append(stem)
        pages_of[stem] = 1 + len(plan_detail_pages(paper["questions"], st, s))
        print(f"  生成 {stem}")

    shutil.rmtree(profile, ignore_errors=True)

    # 点评证据包（v0.6：AI 只能引用这里登记的事实，verify.py 拿它做数字白名单）
    ev_path = out_dir / "点评证据包.json"
    try:
        import evidence as EV
        ev, _ = EV.build(exam_dir)
        ev_path.write_text(json.dumps(ev, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception as exc:                       # 证据包失败不该挡住出报告
        print(f"[提示] 点评证据包生成失败：{exc}")

    # 班级总览 + 待写评语清单（供 AI 一次生成，省额度）
    (out_dir / "班级总览.json").write_text(
        json.dumps(st["class"], ensure_ascii=False, indent=2), encoding="utf-8")
    need = {s["sid"]: {"姓名": s["name"], "等级": s["grade"], "总分": s["total"],
                       "优势": [m["name"] for m in s["mods"] if m["gap"] > 0],
                       "最弱": min(s["mods"], key=lambda m: m["gap"])["name"],
                       "失分题": [x["no"] for x in s["lost"][:3]],
                       "学生备注": (s.get("note") or "").strip()}
            for s in st["live"] if s["sid"] not in comments and s["name"] not in comments}
    (out_dir / "评语待写.json").write_text(
        json.dumps(need, ensure_ascii=False, indent=2), encoding="utf-8")

    info_path = out_dir / "渲染信息.json"
    prev = {}
    if info_path.exists():                       # --only 只出几个人时不要把其他人的页数冲掉
        try:
            prev = json.loads(info_path.read_text(encoding="utf-8")).get("pages") or {}
        except Exception:
            prev = {}
    prev.update(pages_of)
    info_path.write_text(
        json.dumps({"pages": prev, "detail_budget_mm": DETAIL_BUDGET_MM}, ensure_ascii=False, indent=1),
        encoding="utf-8")

    # 清掉上一轮残留（学号被删、改名后的旧文件），否则会被打进 zip 发给老师。
    # --only 只出几个人时不动别人的文件。
    if not only:
        keep = {f"{m}{ext}" for m in made for ext in (".pdf", ".html")}
        keep |= {"卷面分析.pdf", "卷面分析.html"}      # 教师版单页，单独交付，不清也不进 zip
        keep |= {"卷面分析.pdf", "卷面分析.html"}      # 教师版单页，单独交付，不清也不进 zip
        for f in out_dir.iterdir():
            if f.suffix in (".pdf", ".html") and f.name not in keep:
                f.unlink()

    if made:
        # 只打包本轮生成的报告；卷面分析是教师版，单独给，不混进学生包里
        with zipfile.ZipFile(out_dir / "全部报告.zip", "w", zipfile.ZIP_DEFLATED) as z:
            for m in made:
                for ext in (".pdf", ".html"):
                    f = out_dir / f"{m}{ext}"
                    if f.exists():
                        z.write(f, f.name)

    with open(exam_dir / "台账.csv", "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    paper["exam"].get("name", ""), len(made),
                    f'avg={st["avg_mode"]}', f'grade={st["grade_mode"]}', str(out_dir)])
    print(f"\n完成：{len(made)} 份 · 班均 {st['class']['班均']} · 输出 {out_dir}")
    if st["class"]["缺考"]:
        print("缺考（已跳过，不计入班均）：" + "、".join(st["class"]["缺考"]))


if __name__ == "__main__":
    main()
