#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出报告前的质检（全绿才算能交）

用法：python verify.py --exam-dir "考试文件夹"
退出码：0 全绿 / 1 有硬问题

十四项，分三层：
  数据自洽   ① 分值合计 ② 个人合计 ③ 班均≤满分 ④ 每人有分 ⑤ 缺题提示
  版面       ⑥ PDF 数量与页数（对渲染信息.json 逐份核）
  文字       ⑦ 数字清单 ⑧ 字数 ⑨ 禁用词 ⑩ 总分/排名复读
             ⑪ 举例题号数量 ⑫ 相邻开篇重复 ⑬ 数字来源（对证据包白名单）⑭ 备注隐私直抄
"""
import argparse
import json
import re
from pathlib import Path

from evidence import BANNED


def clen(t):
    return len(re.sub(r"\s+", "", t or ""))


def num_key(x):
    try:
        return round(float(x), 2)
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--exam-dir", required=True)
    args = ap.parse_args()
    d = Path(args.exam_dir).resolve()

    paper = json.loads((d / "paper.json").read_text(encoding="utf-8-sig"))
    sheet = d / "录分结果.json"
    if not sheet.exists():
        raise SystemExit("[错误] 还没有 录分结果.json（先跑 ingest_excel.py）")
    data = json.loads(sheet.read_text(encoding="utf-8-sig"))

    qfull = {str(q["no"]): float(q["full"]) for q in paper["questions"]}
    declared = paper.get("exam", {}).get("full")
    students = [s for s in data.get("students", []) if not s.get("缺考")]
    fails, warns = [], []

    print(f"考试：{paper.get('exam',{}).get('name','')}　人数：{len(students)}　"
          f"题数：{len(qfull)}　班均模式：{data.get('avg_mode')}")

    # ① Σ题目满分 = 卷面满分
    s1 = sum(qfull.values())
    if declared and abs(float(declared) - s1) > 1e-6:
        fails.append(f"① 题目分值合计 {s1:g} ≠ 卷面满分 {float(declared):g}")
    else:
        print(f"① 分值合计一致：{s1:g}")

    # ② Σ个人各题 = 总分
    bad = [f'{s.get("姓名")}：{sum(float(v) for v in (s.get("scores") or {}).values()):g} vs {s.get("总分")}'
           for s in students
           if "总分" in s and s.get("scores")
           and abs(sum(float(v) for v in s["scores"].values()) - float(s["总分"])) > 1e-6]
    if bad:
        fails.append("② 个人合计不一致：" + "；".join(bad[:5]))
    else:
        print("② 个人合计自洽")

    # ③ 班均 ≤ 满分
    q_avg = {}
    for no in qfull:
        vals = [float(s["scores"][no]) for s in students if no in (s.get("scores") or {})]
        q_avg[no] = sum(vals) / len(vals) if vals else 0
    if data.get("avg_mode") == "manual":
        for q in data.get("questions", []):
            if q.get("avg") is not None:
                q_avg[str(q["no"])] = float(q["avg"])
    over = [f"第{no}题 {q_avg[no]:.2f}>{qfull[no]:g}" for no in qfull if q_avg[no] > qfull[no] + 1e-9]
    if over:
        fails.append("③ 班均超过满分：" + "；".join(over))
    else:
        print("③ 班均均不超过满分")

    # ④ 每人至少有分
    no_grade = [s.get("姓名") or s.get("学号") or "?" for s in students if not (s.get("scores") or {})]
    if no_grade:
        fails.append("④ 有人一条分数都没有：" + "、".join(no_grade[:8]))
    else:
        print("④ 每人都有分数，等级可判定")

    # ⑤ 缺题（提示级）
    miss = [f'{s.get("姓名")}（缺{len([n for n in qfull if n not in (s.get("scores") or {})])}题）'
            for s in students if [n for n in qfull if n not in (s.get("scores") or {})]]
    if miss:
        warns.append("⑤ 有人题分未录全（空白按未录处理，不参与该题班均）：" + "、".join(miss[:8]))

    # ⑥ PDF 数量与页数
    rd = d / "reports"
    pdfs = sorted(p for p in rd.glob("*.pdf") if p.name != "卷面分析.pdf") if rd.exists() else []
    info_file = rd / "渲染信息.json" if rd.exists() else None
    render_info = json.loads(info_file.read_text(encoding="utf-8")) if info_file and info_file.exists() else {}
    expect_pages = render_info.get("pages") or {}
    if not pdfs:
        warns.append("⑥ 还没有生成 PDF（先跑 build_reports.py，不带 --no-pdf）")
    else:
        small, pagebad = [], []
        for p in pdfs:
            b = p.read_bytes()
            if len(b) < 8000:
                small.append(p.name)
            n = len(re.findall(rb"/Type\s*/Page[^s]", b))
            want = expect_pages.get(p.stem)
            if n and ((want and n != want) or (not want and n < 2)):
                pagebad.append(f"{p.name}({n}页{'，应为' + str(want) + '页' if want else ''})")
        if small:
            fails.append("⑥ 有 PDF 异常小（可能打印失败）：" + "、".join(small[:5]))
        if pagebad:
            fails.append("⑥ 页数与排版不符（可能溢页被裁）：" + "、".join(pagebad[:5]))
        if not small and not pagebad:
            extra = "（题量较大按明细续页）" if any(v > 2 for v in expect_pages.values()) else ""
            print(f"⑥ PDF {len(pdfs)} 份，页数与排版一致{extra}")
        if len(pdfs) != len(students):
            warns.append(f"⑥ PDF 数 {len(pdfs)} ≠ 应出数 {len(students)}（部分生成或含缺考）")
    if rd.exists() and not (rd / "卷面分析.pdf").exists():
        warns.append("⑥ 还没有 卷面分析.pdf（需要教师版时跑 build_class_panel.py）")

    # ---------- 文字层 ----------
    panels_file = d / "panels.json"
    ev_file = rd / "点评证据包.json" if rd.exists() else None
    ev = json.loads(ev_file.read_text(encoding="utf-8")) if ev_file and ev_file.exists() else {}
    if not panels_file.exists():
        warns.append("⑦ 还没有 panels.json（学生个人点评未写，报告会用兜底句）")
        print()
        for w in warns:
            print("  [提示]", w)
        if fails:
            print("\n[不通过]")
            for f in fails:
                print("  ✗", f)
            raise SystemExit(1)
        print("[数据层通过] 文字层未就绪，先写 panels.json")
        return

    pj = json.loads(panels_file.read_text(encoding="utf-8-sig"))
    texts = {}

    def walk(v, sid=None):
        if isinstance(v, str):
            if sid:
                texts.setdefault(sid, []).append(v)
        elif isinstance(v, list):
            for x in v:
                walk(x, sid)
        elif isinstance(v, dict):
            for k, x in v.items():
                walk(x, sid if not isinstance(k, str) or not re.fullmatch(r"\d{1,6}", k) else k)

    students_panel = pj.get("students") or {}
    for sid, v in students_panel.items():
        t = v if isinstance(v, str) else (v.get("comment") or v.get("personal") or "")
        if t:
            texts[str(sid)] = [t]
    class_text = (pj.get("class") or {})
    class_paper = class_text.get("paper") if isinstance(class_text, dict) else ""

    def all_text(sid):
        return "".join(texts.get(str(sid), []))

    # ⑧ 字数
    over_len, short = [], []
    for sid, parts in texts.items():
        n = clen(parts[0] if parts else "")
        if n > 200:
            over_len.append(f"{sid} {n} 字")
        elif n < 100:
            short.append(f"{sid} {n} 字")
    if over_len:
        fails.append("⑧ 个人点评超 200 字：" + "；".join(over_len[:6]))
    else:
        print("⑧ 个人点评均在 200 字内")
    if short:
        warns.append("⑧ 有偏短的点评（<100 字）：" + "、".join(short[:6]))

    # ⑨ 禁用词
    hit = []
    for sid, parts in texts.items():
        for w in BANNED:
            if w in (parts[0] if parts else ""):
                hit.append(f"{sid}「{w}」")
    if hit:
        fails.append("⑨ 出现禁用词：" + "；".join(hit[:8]))
    else:
        print("⑨ 无禁用词")

    # ⑩ 总分/排名复读
    rep = []
    for sid, parts in texts.items():
        t = parts[0] if parts else ""
        if re.search(r"总分\s*[:：]?\s*\d", t) or re.search(r"第\s*\d+\s*名", t) or "排名" in t:
            rep.append(str(sid))
    if rep:
        fails.append("⑩ 复读了总分或排名：" + "、".join(rep[:8]))
    else:
        print("⑩ 无总分/排名复读")

    # ⑪ 举例题号 ≤ 2
    many = []
    for sid, parts in texts.items():
        nos = set(re.findall(r"第\s*(\d{1,3})\s*题", parts[0] if parts else ""))
        if len(nos) > 2:
            many.append(f"{sid}（{len(nos)} 个）")
    if many:
        fails.append("⑪ 点评里点名的题号超过 2 个：" + "、".join(many[:6]))
    else:
        print("⑪ 举例题号均不超过 2 个")

    # ⑫ 相邻开篇去重
    order = [str(s["学号"]) for s in ev.get("学生", [])] or list(texts)
    dup = []
    prev = None
    for sid in order:
        t = all_text(sid)
        if not t:
            continue
        head = re.sub(r"\s+", "", t)[:6]
        if prev and head == prev[0]:
            dup.append(f"{prev[1]} / {sid}「{head}」")
        prev = (head, sid)
    if dup:
        fails.append("⑫ 相邻两份开篇相同：" + "；".join(dup[:5]))
    else:
        print("⑫ 相邻开篇不重复")

    # ⑬ 数字来源（白名单）
    if not ev:
        warns.append("⑬ 缺 点评证据包.json，数字来源无法自动核对（跑 evidence.py）")
    else:
        ev_students = {str(s["学号"]): s for s in ev.get("学生", [])}
        qnos_int = {int(float(n)) for n in qfull if re.fullmatch(r"[\d.]+", n)}
        unknown = []
        for sid, parts in texts.items():
            t = parts[0] if parts else ""
            allow = set()
            for x in (ev_students.get(sid, {}).get("数字白名单") or []):
                k = num_key(x)
                if k is not None:
                    allow.add(k)
            for n in re.findall(r"\d+(?:\.\d+)?", t):
                k = num_key(n)
                if k is None:
                    continue
                if int(k) == k and int(k) in qnos_int:      # 题号不算漂移
                    continue
                if k in (0, 1, 2, 100):                     # 序数与百分比常识值
                    continue
                if k not in allow:
                    unknown.append(f"{sid}:{n}")
        if unknown:
            fails.append("⑬ 点评里的数字在证据包里找不到：" + "、".join(unknown[:10]))
        else:
            print("⑬ 数字均可回溯到证据包")

    # ⑭ 备注隐私直抄
    if ev:
        leak = []
        for st_ in ev.get("学生", []):
            note = (st_.get("学生备注") or "").strip()
            if len(note) < 8:
                continue
            t = all_text(str(st_["学号"]))
            for i in range(len(note) - 7):
                frag = note[i:i + 8]
                if frag in t:
                    leak.append(f'{st_["姓名"]}「{frag}…」')
                    break
        if leak:
            fails.append("⑭ 学生备注被原样搬进点评（应改写成写作证据）：" + "；".join(leak[:5]))
        else:
            print("⑭ 未见备注原样直抄")

    # ⑦ 数字清单（保留给人眼扫一眼）
    nums = []
    for parts in texts.values():
        nums += re.findall(r"\d+(?:\.\d+)?", parts[0] if parts else "")
    uniq = sorted(set(nums), key=lambda x: float(x))
    print(f"⑦ 点评里的数字（{len(uniq)} 个）：{'、'.join(uniq[:32])}" + ("…" if len(uniq) > 32 else ""))

    # 覆盖度
    if ev:
        lack = [s["姓名"] for s in ev["学生"] if not all_text(str(s["学号"]))]
        if lack:
            warns.append(f"⑦ {len(lack)} 人没有点评（报告会用兜底句）：" + "、".join(lack[:8]))

    print()
    for w in warns:
        print("  [提示]", w)
    if fails:
        print("\n[不通过]")
        for f in fails:
            print("  ✗", f)
        raise SystemExit(1)
    print("[全绿] 可以交给老师")


if __name__ == "__main__":
    main()
