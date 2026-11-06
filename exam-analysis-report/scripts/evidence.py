#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成「点评证据包」：AI 写个人点评时唯一可用的事实来源

用法：python evidence.py --exam-dir "考试文件夹"
输出：reports/点评证据包.json

为什么要这一层：
  点评里出现任何数字、题号、知识点，都必须能回溯到本次输入或本次脚本计算。
  证据包把「可写的数字」提前登记成白名单，verify.py 拿它做拦截，而不是靠人眼核。

包含：
  · 考试公共信息 + 考试备注 + 原卷文件名（没有就是空，AI 不得声称看过原卷）
  · 逐题：分值、模块、知识点、常见错因、自由错因、题目备注、复习位置与建议
  · 班级事实：人数、班均、最高最低、等级分布、模块班均、最难/最易题、错因排行
  · 每人：成绩与模块得分率、失分题、知识点掌握、学生备注（只进本人）、数字白名单、
          已用开篇（相邻去重）、全班开篇清单
  · 写作约束：字数、题号上限、禁用词、禁复读总分
"""
import argparse
import json
import re
from pathlib import Path

import build_reports as B

# 与 comment-prompt.md 同步；verify.py 也用这份清单
BANNED = [
    "戒骄戒躁", "百尺竿头", "继续努力", "再接再厉", "不要气馁", "争取好成绩",
    "综上所述", "总而言之", "显而易见", "不难看出", "令人遗憾的是", "交出满意答卷",
    "粗心大意", "马虎", "基础不牢", "思维不灵活", "多做题", "多复习", "认真听讲", "查漏补缺",
    "拖后腿", "差得很远", "你必须", "必然提分", "水到渠成", "名列前茅", "无可挑剔",
]


def numbers_in(text):
    return re.findall(r"\d+(?:\.\d+)?", text or "")


def build(exam_dir):
    d = Path(exam_dir).resolve()
    paper = B.load_json(d / "paper.json")
    if not paper:
        raise SystemExit(f"[错误] 缺少 {d / 'paper.json'}")
    src = d / "录分结果.json"
    if not src.exists():
        raise SystemExit(f"[错误] 缺少 {d / '录分结果.json'}（先跑 ingest_excel.py）")
    rows = B.load_scores(src)
    st = B.compute(paper, rows, paper.get("avg_mode", "auto"),
                   paper.get("grade_mode", "pct"), B.load_manual_avg(src))
    exam = paper.get("exam", {})
    cls = st["class"]
    qnos = set(st["qnos"])

    questions = [{
        "题号": str(q["no"]),
        "分值": q["full"],
        "模块": q.get("module", ""),
        "知识点": q.get("kp", ""),
        "常见错因": [st["tag_labels"].get(c, c) for c in (q.get("errors") or [])],
        "自由错因": q.get("errors_note", ""),
        "题目备注": q.get("note", ""),
        "复习位置": q.get("book", ""),
        "学习建议": q.get("advice", ""),
        "题型": q.get("type", ""),
    } for q in paper["questions"]]

    class_facts = {
        "人数": cls["人数"],
        "班均": cls["班均"],
        "最高": cls["最高"],
        "最低": cls["最低"],
        "满分": st["total_full"],
        "题数": len(paper["questions"]),
        "等级分布": cls["等级分布"],
        "模块班均": cls["模块班均"],
        "最难三题": cls["最难题"],
        "最易三题": cls["最易题"],
        "错因排行": cls["错因排行"][:4],
        "缺考": cls["缺考"],
    }

    # 公共可引用数字：班级事实里出现过的数，加上满分、题数、人数
    common_nums = set()
    for v in [cls["班均"], cls["最高"], cls["最低"], st["total_full"], len(paper["questions"]), cls["人数"]]:
        common_nums.add(f"{float(v):.2f}".rstrip("0").rstrip("."))
    for m in cls["模块班均"].values():
        for k, v in m.items():
            common_nums.add(f"{float(v):.2f}".rstrip("0").rstrip("."))
    for t in cls["最难题"] + cls["最易题"]:
        common_nums.add(f"{float(t['no']):.2f}".rstrip("0").rstrip("."))
        common_nums.add(f"{float(t['得分率']):.2f}".rstrip("0").rstrip("."))

    students = []
    for s in st["live"]:
        mods = [{"模块": m["name"], "满分": m["full"], "得分": m["score"],
                 "班均": m["avg"], "得分率": round(m["rate"] * 100, 1),
                 "班均得分率": round(m["avg_rate"] * 100, 1), "差值百分点": round(B.gap_pp(m), 1)}
                for m in s["groups"]]
        lost = [{"题号": x["no"], "知识点": x["kp"], "模块": x["module"],
                 "得分": x["got"], "满分": x["full"], "丢分": x["lost"],
                 "常见失分点": (x.get("errors_note") or "、".join(
                     st["tag_labels"].get(c, c) for c in (x.get("errors") or []))),
                 "复习位置": x.get("book", ""), "学习建议": x.get("advice", "")}
                for x in s["lost"]]

        own_nums = set(common_nums)
        for v in [s["total"], s["rate"] * 100]:
            own_nums.add(f"{float(v):.2f}".rstrip("0").rstrip("."))
        for m in mods:
            for k in ("得分", "班均", "得分率", "班均得分率"):
                own_nums.add(f"{float(m[k]):.2f}".rstrip("0").rstrip("."))
            own_nums.add(f"{abs(float(m['差值百分点'])):.2f}".rstrip("0").rstrip("."))
        for x in lost:
            for k in ("得分", "满分", "丢分"):
                own_nums.add(f"{float(x[k]):.2f}".rstrip("0").rstrip("."))

        students.append({
            "学号": s["sid"],
            "姓名": s["name"],
            "班级": s.get("cls") or exam.get("class", ""),
            "学生备注": (s.get("note") or "").strip(),
            "总分": s["total"],
            "得分率": round(s["rate"] * 100, 1),
            "等级": s["grade"],
            "模块": mods,
            "失分题": lost[:6],
            "知识点掌握": [{"知识点": k["kp"], "题号": k["nos"], "得分率": round(k["rate"] * 100, 1),
                        "水平": k["level"]} for k in s["kps"][:6]],
            "数字白名单": sorted(own_nums, key=lambda x: float(x)),
        })

    evidence = {
        "考试": {
            "名称": exam.get("name", ""), "日期": exam.get("date", ""),
            "学科": exam.get("subject", ""), "年级": exam.get("grade", ""),
            "校区": exam.get("campus", ""), "教师": exam.get("teacher", ""),
            "班级": exam.get("class", ""), "满分": st["total_full"],
            "考试备注": exam.get("note", ""),
            "原卷": exam.get("paper_file", ""),
        },
        "题目": questions,
        "班级事实": class_facts,
        "学生": students,
        "写作约束": {
            "字数": "建议 120–180 字，硬上限 200 字（不含标点也算在内，脚本按去空白字符计）",
            "结构": "一个自然段，无标题、无分点、无 Markdown",
            "角色与口吻": "任课老师写给学生和家长：真诚、亲切、专业、有温度；像在办公室当面交底与指导，不说教、不冷漠、不居高临下",
            "内容结构": "先真诚肯定已有基础与闪光点（立足信心），再点明1个关键瓶颈（点透思维或方法卡点），最后给一个切实可行的复习抓手（教方法），结尾可附一句温暖期许",
            "表达原则": "重在学情诊断与学习方法指导，坚决不写成机械的题目操作步骤或答案解析；对薄弱生严禁使用打击自尊或冷嘲热讽的词句（如‘有些还记得/没完全丢/几乎没拿到分’等），看到每一点点滴进步与潜力",
            "推断边界": "基于得分事实与教师备注作客观分析；无该生明确作答记录时，不凭空脑补草稿涂改、漏看单位等具体过程细节。原卷是题目，不是学生答卷",
            "建议抓手": "将教师登记的学习建议转化为科学的复习方法（如回归课本原理、重做错题并讲清逻辑、圈画已知条件等）；不编造具体训练天数或虚假提分承诺",
            "禁止": BANNED,
            "禁复读": "不写总分、不写排名、不与同学横向比较拉踩",
            "题号上限": 2,
            "数字": "尽量少用数字；若用成绩数字，只能使用该生「数字白名单」里的数字；题号不受限。白名单不是训练量来源",
            "备注": "学生备注只用于本人；将老师的观察转化为得体温暖的指导，不得原样搬入隐私、家庭信息或负面情绪评价",
            "原卷": ("本次已归档原卷（题目语境）：" + exam.get("paper_file", "")) if exam.get("paper_file")
                    else "本次未提供原卷：不得引用题干细节，也不得声称看过原始作答",
            "开篇去重": "相邻两份前 6 字不得重复；因人施教，多角度自然切入",
            "文风自检": "写完务必自读一遍：语气像不像老师当面跟孩子和家长说话？读完是深受鼓舞、明确方向，还是感到刺耳或枯燥？",
        },
    }
    return evidence, st


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exam-dir", required=True)
    ap.add_argument("--out", default="")
    args = ap.parse_args()
    ev, st = build(args.exam_dir)
    out = Path(args.out) if args.out else Path(args.exam_dir).resolve() / "reports" / "点评证据包.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(ev, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"证据包：{out}")
    print(f"  考试：{ev['考试']['名称']}　题数 {len(ev['题目'])}　学生 {len(ev['学生'])}")
    print(f"  作业：让 AI 读这个文件，写 panels.json 的 students.<学号>.comment")
    print(f"  之后：python build_reports.py --exam-dir \"{Path(args.exam_dir).resolve()}\" && verify.py")


if __name__ == "__main__":
    main()
