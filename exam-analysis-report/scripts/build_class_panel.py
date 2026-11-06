#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成「班级卷面分析」单页 A4（给老师讲评用；数据全来自录分与标注，文字来自 panels.json.class）

用法：python build_class_panel.py --exam-dir "考试文件夹"
输出：reports/卷面分析.pdf（+ 同名 html）
"""
import argparse
import html as H
import json
import tempfile
from pathlib import Path

import build_reports as B

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent
TPL = SKILL / "assets" / "class_panel.html"


def svg_qbars(paper, q_rates, qmap, w=700, h=120):
    """各题得分率条形图（全宽）：柱高=得分率，<60% 标红"""
    pad_l, pad_b, pad_t = 22, 22, 12
    n = len(q_rates)
    plot_w, plot_h = w - pad_l - 4, h - pad_b - pad_t
    step = plot_w / max(n, 1)
    bw = step * 0.62
    parts = [f'<svg viewBox="0 0 {w} {h}" style="width:100%;height:auto;display:block">']
    for i in (0.25, 0.5, 0.75, 1.0):
        y = pad_t + plot_h - plot_h * i
        parts.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{w-4}" y2="{y:.1f}" stroke="#e6eaee" stroke-width="0.7"/>')
        parts.append(f'<text x="{pad_l-4}" y="{y+3:.1f}" font-size="7" fill="#8d97a2" text-anchor="end">{int(i*100)}</text>')
    for i, (no, rate) in enumerate(q_rates):
        cx = pad_l + step * i + step / 2
        bh = plot_h * max(rate, 0.004)
        color = B.CRITICAL if rate < 0.6 else B.PALETTE["bar"]
        parts.append(f'<rect x="{cx-bw/2:.1f}" y="{pad_t+plot_h-bh:.1f}" width="{bw:.1f}" '
                     f'height="{bh:.1f}" fill="{color}" rx="1"/>')
        parts.append(f'<text x="{cx:.1f}" y="{pad_t+plot_h-bh-3:.1f}" font-size="6.6" fill="#3b4652" '
                     f'text-anchor="middle">{rate*100:.0f}</text>')
        parts.append(f'<text x="{cx:.1f}" y="{pad_t+plot_h+9:.1f}" font-size="7" fill="#4a5560" '
                     f'text-anchor="middle">{no}</text>')
        parts.append(f'<text x="{cx:.1f}" y="{pad_t+plot_h+18:.1f}" font-size="6" fill="#98a2ad" '
                     f'text-anchor="middle">{qmap.get(no, 0):g}分</text>')
    parts.append(f'<line x1="{pad_l}" y1="{pad_t+plot_h}" x2="{w-4}" y2="{pad_t+plot_h}" stroke="#b9c2cb" stroke-width="0.8"/>')
    parts.append("</svg>")
    return "".join(parts)


def lv_cls(level):
    return "weak" if level == "薄弱" else ("half" if level == "待巩固" else "ok")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exam-dir", required=True)
    ap.add_argument("--no-pdf", action="store_true")
    args = ap.parse_args()

    exam_dir = Path(args.exam_dir).resolve()
    paper = B.load_json(exam_dir / "paper.json")
    pal, binfo = B.brand.resolve(paper, exam_dir)
    B.set_theme(pal)
    for i, line in enumerate(B.brand.describe(binfo).splitlines()):
        print(("[品牌] " if i == 0 else "        ") + line)
    src = exam_dir / "录分结果.json"
    if not src.exists():
        src = exam_dir / "scores.csv"
    rows = B.load_scores(src)
    manual = B.load_manual_avg(src)
    st = B.compute(paper, rows, paper.get("avg_mode", "auto"), paper.get("grade_mode", "pct"), manual)
    cls = st["class"]
    panels = (B.load_json(exam_dir / "panels.json", {}) or {}).get("class") or {}

    kps = sorted(st["class"]["知识点掌握"], key=lambda k: (k["得分率"], -k["满分"]))
    kp_rows = []
    for k in kps:
        warn = " weak" if k["得分率"] < 60 else ""
        kp_rows.append(
            f'<tr><td>{H.escape(k["kp"])}</td>'
            f'<td class="num">{H.escape("/".join(k["题号"]))}</td>'
            f'<td class="num">{k["满分"]:g}</td>'
            f'<td><div class="bar"><i class="{warn.strip()}" style="width:{min(k["得分率"],100)}%"></i></div></td>'
            f'<td class="num">{k["得分率"]:g}%</td>'
            f'<td class="num">{len(k["题号"])} 题</td></tr>')
    if not kp_rows:
        kp_rows.append('<tr><td colspan="6" style="text-align:center;color:#8a94a0">'
                       '本题卷未标注知识点</td></tr>')

    diff_rows = []
    for name in ("易", "中", "难"):
        nos = cls["难度分布"].get(name, [])
        qs = "、".join(nos) if nos else "—"
        diff_rows.append(f'<tr><td class="num">{name}</td><td class="num">{len(nos)}</td>'
                         f'<td style="text-align:left">{qs}</td></tr>')

    err_tags = "".join(f'<span class="t">{H.escape(e["错因"])} · {e["涉及分值"]:g}分</span>'
                       for e in cls["错因排行"][:8]) or '<span class="t">未标注错因</span>'

    gd = cls["等级分布"]
    gd_txt = "/".join(str(gd.get(g, 0)) for g in "ABCDE")
    repl = {
        "ACCENT": pal["accent"], "ACCENT_SOFT": pal["soft"], "ACCENT_LINE": pal["line"],
        "ACCENT_BADGE": pal["badge"], "ACCENT_CONNECTOR": pal["connector"],
        "ACCENT_BAR": pal["bar"],
        "EXAM_NAME": H.escape(cls["考试"]),
        "GRADE_SUBJECT": H.escape(f'{paper.get("exam", {}).get("grade", "")}'
                                  f'{paper.get("exam", {}).get("subject", "")}'),
        "FULL": f'{st["total_full"]:g}',
        "DATE": H.escape(paper.get("exam", {}).get("date", "")),
        "N": str(cls["人数"]), "AVG": f'{cls["班均"]:g}',
        "MAX": f'{cls["最高"]:g}', "MIN": f'{cls["最低"]:g}', "GRADEDIST": gd_txt,
        "KP_ROWS": "".join(kp_rows), "DIFF_ROWS": "".join(diff_rows), "ERR_TAGS": err_tags,
        "QBAR_SVG": svg_qbars(paper, [(no, cls["逐题难度"].get(no, 0)) for no in st["qnos"]], st["qmap"]),
        "P_PAPER_BLOCK": (
            '<div class="ai"><h3>试卷与班级情况</h3><p>'
            + H.escape(panels.get("paper", "") or panels.get("overview", "")) + "</p></div>"
        ) if (panels.get("paper") or panels.get("overview")) else "",
    }
    out = TPL.read_text(encoding="utf-8")
    for k, v in repl.items():
        out = out.replace(f"@@{k}@@", str(v))

    rdir = exam_dir / "reports"
    rdir.mkdir(exist_ok=True)
    hp = rdir / "卷面分析.html"
    hp.write_text(out, encoding="utf-8")
    print(f"卷面分析：{hp}")
    if not args.no_pdf:
        B.print_pdf(B.find_edge(), hp, rdir / "卷面分析.pdf", tempfile.mkdtemp(prefix="wb_pdf_"))
        print(f"PDF：{rdir / '卷面分析.pdf'}")


if __name__ == "__main__":
    main()
