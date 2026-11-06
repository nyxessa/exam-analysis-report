#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成空白《试卷分析模板.xlsx》（两个 Sheet：试卷信息 / 学生成绩）

用法：
  python make_template.py                          # 写 assets/试卷分析模板.xlsx（26 题 / 60 行学生）
  python make_template.py --questions 30 --students 80 --out "C:/tmp/模板.xlsx"

为什么模板要预置列：老师拿到就能填，不需要先生成"考试专用模板"；题号列预填 1..N，
学生成绩表按"第 k 题"命名，收表时按表头文字匹配，老师自己加列也不会错位。
"""
import argparse
from pathlib import Path

import xlsxlite as X

HERE = Path(__file__).resolve().parent
DEFAULT_OUT = HERE.parent / "assets" / "试卷分析模板.xlsx"

Q_HEADERS = ["题号", "题型", "本题分值", "对应模块", "具体知识点/解法",
             "常见错因", "复习页码", "具体复习题号", "后续学习建议", "题目备注"]

TYPES = "选择题,填空题,作图题,实验探究题,解答题,计算题"


def build(questions=26, students=60):
    # ---------------- Sheet 1 试卷信息 ----------------
    rows = []
    rows.append([("试卷信息", X.ST_TITLE)] + [("", 0)] * 9)
    rows.append([("考试名称", X.ST_HEAD), ("", X.ST_NOTE), ("", X.ST_NOTE),
                 ("考试日期", X.ST_HEAD), ("", X.ST_NOTE),
                 ("校区", X.ST_HEAD), ("", X.ST_NOTE),
                 ("年级", X.ST_HEAD), ("", X.ST_NOTE), ("", X.ST_NOTE)])
    rows.append([("科目", X.ST_HEAD), ("", X.ST_NOTE), ("", X.ST_NOTE),
                 ("授课教师", X.ST_HEAD), ("", X.ST_NOTE),
                 ("卷面满分", X.ST_HEAD), ("", X.ST_NOTE),
                 ("班级", X.ST_HEAD), ("", X.ST_NOTE), ("", X.ST_NOTE)])
    rows.append([("考试备注", X.ST_HEAD), ("", X.ST_NOTE)] + [("", X.ST_NOTE)] * 8)
    rows.append([(h, X.ST_HEAD) for h in Q_HEADERS])
    for i in range(1, questions + 1):
        rows.append([(i, X.ST_BODY)] + [("", X.ST_WRAP) for _ in range(9)])
    rows.append([("合计", X.ST_HEAD), ("", X.ST_HEAD), ("", X.ST_HEAD)] + [("", X.ST_HEAD)] * 7)
    rows.append([("填写说明", X.ST_TITLE)] + [("", 0)] * 9)
    for line in [
        "1. 只填黄色格子；蓝色表头不要改名、不要删列，脚本按表头文字识别。",
        "2. 「本题分值」必须是数字，且全部题目的分值合计要等于「卷面满分」，否则收表时会被拦下并指出行号。",
        "3. 题号可以写 1、2、15(1) 这类形式；题号必须与「学生成绩」表的列一一对应。",
        "4. 「常见错因」可从下列代码里选（可多选，用顿号分隔），也可以直接写自己的说法：",
        "    知识没掌握、方法没想到、审题失误、计算失误、表达不规范、不会迁移运用、时间不够、非智力失分",
        "5. 「题目备注」写命题意图、评分要点、课堂讲授情况等；「学生备注」写该生卷面表现。备注只作写作证据，不会原样公开。",
        "6. 三张表都不填也能出报告，但有备注的点评会更贴近实际。",
    ]:
        rows.append([(line, X.ST_TITLE)] + [("", 0)] * 9)

    sheet1 = {
        "name": "试卷信息",
        "rows": rows,
        "widths": [11, 12, 12, 11, 13, 12, 11, 12, 15, 30],
        "freeze": 5,
        "merges": ["A1:J1", "B2:C2", "B3:C3", "B4:J4",
                   f"A{7 + questions}:J{7 + questions}"],
        "heights": {1: 30, 4: 64, 5: 30},
        "validations": [(f"B{6}:B{5 + questions}", TYPES)],
    }

    # ---------------- Sheet 2 学生成绩 ----------------
    hdr = ["学号", "姓名", "缺考"] + [f"第{i}题" for i in range(1, questions + 1)] + ["学生备注"]
    srows = [[("学生成绩", X.ST_TITLE)] + [("", 0)] * (len(hdr) - 1),
             [(h, X.ST_HEAD) for h in hdr]]
    for _ in range(students):
        srows.append([("", X.ST_BODY), ("", X.ST_BODY), ("", X.ST_BODY)]
                     + [("", X.ST_NUM) for _ in range(questions)]
                     + [("", X.ST_WRAP)])
    last = f"{len(hdr)}"
    sheet2 = {
        "name": "学生成绩",
        "rows": srows,
        "widths": [8, 11, 7] + [7] * questions + [34],
        "freeze": 2,
        "merges": [f"A1:{X.col_letters(len(hdr))}1"],
        "heights": {1: 30, 2: 34},
        "validations": [(f"C3:C{2 + students}", "是,否")],
    }
    return [sheet1, sheet2]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--questions", type=int, default=26)
    ap.add_argument("--students", type=int, default=60)
    ap.add_argument("--out", default=str(DEFAULT_OUT))
    args = ap.parse_args()
    out = X.write_book(args.out, build(args.questions, args.students), title="试卷分析模板")
    print(f"已生成：{out}（{args.questions} 题 / {args.students} 行学生 / {out.stat().st_size} 字节）")


if __name__ == "__main__":
    main()
