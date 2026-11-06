#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""收表：老师填好的《试卷分析模板.xlsx》→ paper.json + 录分结果.json + roster.csv

用法：
  python ingest_excel.py --xlsx "期末-八年级物理.xlsx" --out "考试文件夹"
  python ingest_excel.py --xlsx ... --out ... --paper "原卷.pdf"      # 可选：一并归档原卷
  python ingest_excel.py --xlsx ... --out ... --dry-run              # 只校验不落盘

做三件事：
  1. 按表头文字定位列（老师调列顺序/加列都不会错位），不按固定列号硬编码；
  2. 校验一切能校验的：题号重复、分值和≠卷面满分、得分超满分、学号重复、空白与 0 分区分，
     错误直接指到「试卷信息!C7」这种单元格级坐标；
  3. 校验通过才落盘，并把三级备注写进数据契约（exam.note / questions[].note / students[].note）。
"""
import argparse
import json
import re
import shutil
import sys
from pathlib import Path

import xlsxlite as X

HERE = Path(__file__).resolve().parent
TAGS = json.loads((HERE.parent / "assets" / "error-tags.json").read_text(encoding="utf-8"))["tags"]
LABEL2CODE = {t["label"]: t["code"] for t in TAGS}
CODE2LABEL = {t["code"]: t["label"] for t in TAGS}

# 表头 → 内部字段（按"包含关键字"匹配，容忍老师改叫法）
Q_COLS = [
    ("no", ("题号",)),
    ("type", ("题型",)),
    ("full", ("分值",)),
    ("module", ("模块",)),
    ("kp", ("知识点", "考点")),
    ("errors", ("错因", "失分点")),
    ("page", ("页码",)),
    ("item", ("复习题号", "题号范围")),
    ("advice", ("建议",)),
    ("note", ("备注",)),
]
S_COLS = [("sid", ("学号", "考号")), ("name", ("姓名",)), ("absent", ("缺考",)), ("note", ("备注",))]

META_LABELS = ["考试名称", "考试日期", "校区", "年级", "班级", "科目", "授课教师", "卷面满分", "考试备注"]

SHORT_MAP = {
    "微观与内能": "内能与微观", "电路与电流": "电路电流", "电磁现象": "电磁",
    "家庭电路": "家庭电路", "电学实验": "电学实验", "电功率计算": "电功率",
    "运动的描述": "运动描述", "运动图像": "运动图像", "匀变速直线运动": "匀变速",
}


class Report:
    def __init__(self):
        self.errors, self.warns = [], []

    def err(self, where, msg):
        self.errors.append(f"{where}：{msg}")

    def warn(self, where, msg):
        self.warns.append(f"{where}：{msg}")

    def dump(self):
        for w in self.warns:
            print("  [提示]", w)
        if self.errors:
            print("\n[不通过] 共 %d 处问题：" % len(self.errors))
            for e in self.errors:
                print("  ✗", e)
            return False
        return True


def num(v, default=None):
    try:
        return float(str(v).strip())
    except Exception:
        return default


def cell_of(col, row):
    return f"{col}{row}"


def find_header(grid, musts, start=1):
    """找到同时含 musts 全部关键字的行号"""
    for r in sorted(grid):
        text = "".join(grid[r].values())
        if all(m in text for m in musts):
            return r
    return None


def map_cols(grid, header_row, specs):
    """表头文字 → 列字母；返回 {字段: 列字母} 与未识别的表头"""
    cols, unknown = {}, []
    for col, text in sorted(grid[header_row].items(), key=lambda kv: X.col_index(kv[0])):
        t = text.strip().replace(" ", "")
        hit = None
        for field, keys in specs:
            if any(k in t for k in keys):
                hit = field
                break
        if hit and hit not in cols:
            cols[hit] = col
        elif not hit:
            unknown.append(t)
    return cols, unknown


def read_meta(grid, header_row, rep):
    """表头上方的 key/value 区：值固定取标签**紧邻**右侧一格。

    不能取「右侧第一个非空」——留空时会把下一个标签当成前一个标签的值
    （校区留空 → 考试日期读到「校区」）。右邻若是另一个标签，同样判为空。
    """
    meta = {}
    for r in range(1, header_row):
        row = grid.get(r, {})
        for col, text in sorted(row.items(), key=lambda kv: X.col_index(kv[0])):
            label = text.strip().replace(" ", "")
            if label not in META_LABELS or label in meta:
                continue
            nxt = X.col_letters(X.col_index(col) + 1)
            val = (row.get(nxt) or "").strip()
            if val.replace(" ", "") in META_LABELS:
                val = ""
            meta[label] = val
    return meta


def parse_paper(bk, rep):
    grid = bk.grid("试卷信息")
    hdr = find_header(grid, ("题号", "分值"))
    if hdr is None:
        rep.err("试卷信息", "找不到表头行（需要同时含「题号」和「分值」的新行），模板可能被改坏了")
        return None, None, None
    cols, unknown = map_cols(grid, hdr, Q_COLS)
    for need, name in (("no", "题号"), ("full", "本题分值"), ("module", "对应模块"), ("kp", "具体知识点")):
        if need not in cols:
            rep.err(f"试卷信息!{hdr}行", f"缺列「{name}」，脚本无法定位数据")
    if unknown:
        rep.warn(f"试卷信息!{hdr}行", "未识别的表头（已忽略）：" + "、".join(unknown))
    if "full" not in cols:
        return None, None, None

    meta = read_meta(grid, hdr, rep)
    qs, seen_no = [], {}
    for r in sorted(grid):
        if r <= hdr:
            continue
        row = grid[r]
        raw_no = (row.get(cols.get("no", "A")) or "").strip()
        full_raw = (row.get(cols.get("full")) or "").strip()
        module = (row.get(cols.get("module", "")) or "").strip()
        kp = (row.get(cols.get("kp", "")) or "").strip()
        if raw_no in ("合计", "填写说明", "小计"):
            break
        if not raw_no and not full_raw and not module and not kp:
            continue
        where = f"试卷信息!{cols.get('full')}{r}"
        if not raw_no:
            rep.err(where, "有分值但没写题号")
            continue
        if not full_raw:
            if module or kp:
                rep.err(where, f"第 {raw_no} 题没填「本题分值」")
            continue
        full = num(full_raw)
        if full is None or full <= 0:
            rep.err(where, f"分值「{full_raw}」不是大于 0 的数字")
            continue
        if raw_no in seen_no:
            rep.err(f"试卷信息!{cols.get('no','A')}{r}",
                    f"题号 {raw_no} 重复（上一次在第 {seen_no[raw_no]} 行）")
            continue
        seen_no[raw_no] = r

        # 错因：既认 code 也认中文标签，多个用顿号/逗号/斜杠分隔
        errors, note_extra = [], []
        raw_err = (row.get(cols.get("errors", "")) or "").strip()
        for tok in re.split(r"[、,，/；;]+", raw_err):
            tok = tok.strip()
            if not tok:
                continue
            if tok in LABEL2CODE:
                errors.append(LABEL2CODE[tok])
            elif tok in CODE2LABEL:
                errors.append(tok)
            else:
                note_extra.append(tok)
        if note_extra:
            rep.warn(f"试卷信息!{cols.get('errors','')}{r}",
                     f"「{'、'.join(note_extra)}」不在错因字典里，已并入自由错因文本（不影响出报告）")
        book = " ".join(x for x in [(row.get(cols.get("page", "")) or "").strip(),
                                    (row.get(cols.get("item", "")) or "").strip()] if x)
        advice = (row.get(cols.get("advice", "")) or "").strip()
        q = {
            "no": raw_no,
            "type": (row.get(cols.get("type", "")) or "").strip(),
            "full": int(full) if full == int(full) else full,
            "module": module or "未分组",
            "kp": kp or "未标注",
            "book": book,
            "advice": advice,
            "skill": advice,
            "note": (row.get(cols.get("note", "")) or "").strip(),
        }
        if errors:
            q["errors"] = errors
        if note_extra:
            q["errors_note"] = "、".join(note_extra)
        qs.append(q)

    if not qs:
        rep.err("试卷信息", "一道题都没读到（表头下没有有效的题目行）")
        return None, None, None

    total = sum(q["full"] for q in qs)
    declared = num(meta.get("卷面满分"))
    if declared and abs(declared - total) > 1e-6:
        extra = total - declared
        rep.err("试卷信息!卷面满分",
                f"题目分值合计 {total:g} ≠ 卷面满分 {declared:g}（差 {abs(extra):g} 分）。"
                f"先核对是否漏题/分值抄错；确认卷子本身如此，就把「卷面满分」改成 {total:g}")
    return meta, qs, total


def parse_scores(bk, qs, rep):
    grid = bk.grid("学生成绩")
    hdr = find_header(grid, ("姓名",))
    if hdr is None:
        rep.err("学生成绩", "找不到表头行（需要含「姓名」的行）")
        return []
    cols, unknown = map_cols(grid, hdr, S_COLS)
    if "name" not in cols:
        rep.err("学生成绩", "缺「姓名」列")
        return []
    # 逐题列：第k题 / Qk / 纯数字
    qcol = {}
    for col, text in grid[hdr].items():
        t = text.strip()
        m = re.fullmatch(r"第\s*(\d+)\s*题", t) or re.fullmatch(r"[Qq]\s*(\d+)", t) or re.fullmatch(r"(\d{1,3})", t)
        if m:
            qcol[m.group(1)] = col
    missing = [str(q["no"]) for q in qs if str(q["no"]) not in qcol]
    if missing:
        rep.err("学生成绩!" + str(hdr) + "行",
                f"缺少这些题号的成绩列：{'、'.join(missing[:8])}（表头写成「第N题」即可）")
    extra = [k for k in qcol if k not in {str(q["no"]) for q in qs}]
    if extra:
        rep.warn("学生成绩", f"成绩表多出未登记的题号列：{'、'.join(sorted(extra)[:8])}，已忽略")
    unknown = [t for t in unknown if not re.fullmatch(r"第\s*\d+\s*题|[Qq]\s*\d+|\d{1,3}", t)]
    if unknown:
        rep.warn(f"学生成绩!{hdr}行", "未识别的表头（已忽略）：" + "、".join(unknown))

    full_of = {str(q["no"]): float(q["full"]) for q in qs}
    students, seen_sid = [], {}
    last_row = max(grid) if grid else hdr
    for r in range(hdr + 1, last_row + 1):
        row = grid.get(r, {})
        sid = (row.get(cols.get("sid", "A")) or "").strip()
        name = (row.get(cols.get("name")) or "").strip()
        if not sid and not name:
            continue
        if not name:
            rep.err(f"学生成绩!{cols['name']}{r}", "只有学号没有姓名")
            continue
        if sid and sid in seen_sid:
            rep.err(f"学生成绩!{cols.get('sid','A')}{r}", f"学号 {sid} 重复（上一次在第 {seen_sid[sid]} 行）")
            continue
        seen_sid[sid] = r
        absent_txt = (row.get(cols.get("absent", "")) or "").strip()
        absent = absent_txt in ("是", "Y", "y", "1", "TRUE", "true", "√", "缺考")
        if absent_txt and not absent and absent_txt not in ("否", "N", "n", "0", "FALSE", "false", ""):
            rep.warn(f"学生成绩!{cols.get('absent','')}{r}", f"缺考列写了「{absent_txt}」，按「未缺考」处理")
        scores = {}
        for q in qs:
            no = str(q["no"])
            col = qcol.get(no)
            if not col:
                continue
            raw = (row.get(col) or "").strip()
            if raw == "":
                continue                       # 空白 = 未录，不等于 0 分
            if raw in ("-", "缺", "缺考", "/", "免"):
                continue
            v = num(raw)
            if v is None:
                rep.err(f"学生成绩!{col}{r}", f"第 {no} 题得分「{raw}」不是数字")
                continue
            if v < 0:
                rep.err(f"学生成绩!{col}{r}", f"第 {no} 题得分 {v:g} 小于 0")
                continue
            if v > full_of[no] + 1e-9:
                rep.err(f"学生成绩!{col}{r}", f"第 {no} 题得分 {v:g} 超过满分 {full_of[no]:g}")
                continue
            scores[no] = int(v) if v == int(v) else v
        students.append({
            "学号": sid or f"{len(students) + 1:03d}",
            "姓名": name,
            "缺考": absent,
            "scores": scores,
            "note": (row.get(cols.get("note", "")) or "").strip(),
        })
    if not students:
        rep.err("学生成绩", "一个学生都没读到")
    return students


def short_of(name):
    if name in SHORT_MAP:
        return SHORT_MAP[name]
    n = re.sub(r"的(应用|探究|计算|理解|分析)$", "", name)
    return n if len(n) <= 6 else n[:4] + "应用"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True, help="老师交回的模板文件")
    ap.add_argument("--out", required=True, help="考试文件夹（不存在会新建）")
    ap.add_argument("--paper", default="", help="可选：原卷 PDF/图片/Word，会归档到 01_试卷/")
    ap.add_argument("--logo", default="", help="可选：机构 Logo（png/jpg/svg），自动取主题色")
    ap.add_argument("--theme", default="", help="可选：主题色 #RRGGBB，或 auto（按 Logo 取色）")
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    src = Path(args.xlsx).resolve()
    if not src.exists():
        raise SystemExit(f"[错误] 找不到文件：{src}")
    if src.suffix.lower() not in (".xlsx", ".xlsm"):
        raise SystemExit(f"[错误] 只支持 .xlsx / .xlsm（.xls 请先用 Excel 另存为 xlsx）：{src.name}")

    bk = X.Book(src)
    rep = Report()
    need_sheets = ("试卷信息", "学生成绩")
    miss = [s for s in need_sheets if s not in bk.names()]
    if miss:
        raise SystemExit(f"[不通过] 缺少工作表：{'、'.join(miss)}（现有：{'、'.join(bk.names())}）\n"
                         f"         请用 skill 自带的「试卷分析模板.xlsx」填写。")

    meta, qs, total = parse_paper(bk, rep)
    students = parse_scores(bk, qs, rep) if qs else []

    print(f"文件：{src.name}")
    if qs:
        print(f"读到：{len(qs)} 题，分值和 {total:g}；{len(students)} 名学生")
    ok = rep.dump()
    if not ok:
        raise SystemExit(1)
    if args.dry_run:
        print("\n[仅校验] 未写入任何文件")
        return

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    name = meta.get("考试名称") or f"{meta.get('年级','')}{meta.get('科目','')} 试卷分析"
    paper = {
        "exam": {
            "title": f"{meta.get('年级','')}{meta.get('科目','')} 试卷分析报告".strip(),
            # 老师填的「试卷名称」，允许空：空则报告页眉不显示这一项
            "title_given": meta.get("考试名称", "").strip(),
            "name": name,
            "date": meta.get("考试日期", ""),
            "subject": meta.get("科目", ""),
            "grade": meta.get("年级", ""),
            "full": total,
            "org": "",
            "campus": meta.get("校区", ""),
            "teacher": meta.get("授课教师", ""),
            "class": meta.get("班级", ""),
            "note": meta.get("考试备注", ""),
            "paper_file": "",
        },
        "avg_mode": "auto",
        "grade_mode": "pct",
        "grades": {"A": 0.90, "B": 0.70, "C": 0.30, "D": 0.10},
        "bands": [[0, 0.3], [0.3, 0.4], [0.4, 0.5], [0.5, 0.6], [0.6, 0.7], [0.7, 0.8], [0.8, 1.0]],
        "module_shorts": {},
        "questions": qs,
    }
    mods = []
    for q in qs:
        if q["module"] and q["module"] not in [m for m, _ in mods]:
            mods.append((q["module"], short_of(q["module"])))
    paper["module_shorts"] = dict(mods)

    if args.paper:
        p = Path(args.paper).resolve()
        if not p.exists():
            rep.err("原卷", f"找不到 {p}")
        else:
            dst_dir = out / "01_试卷"
            dst_dir.mkdir(exist_ok=True)
            dst = dst_dir / p.name
            if dst.exists():                     # 上次归档的文件可能是只读的
                try:
                    dst.chmod(0o666)
                except Exception:
                    pass
                dst.unlink()
            shutil.copy2(p, dst)
            try:
                dst.chmod(0o644)
            except Exception:
                pass
            paper["exam"]["paper_file"] = f"01_试卷/{p.name}"

    (out / "paper.json").write_text(json.dumps(paper, ensure_ascii=False, indent=1), encoding="utf-8")
    sheet = {
        "考试": name,
        "导出时间": "excel",
        "来源文件": src.name,
        "avg_mode": "auto",
        "questions": [{"no": q["no"], "full": q["full"], "avg": None} for q in qs],
        "students": students,
    }
    (out / "录分结果.json").write_text(json.dumps(sheet, ensure_ascii=False, indent=1), encoding="utf-8")
    with open(out / "roster.csv", "w", encoding="utf-8-sig", newline="") as f:
        f.write("学号,姓名\n")
        for s in students:
            f.write(f'{s["学号"]},{s["姓名"]}\n')

    brand_line = ""
    if args.logo or args.theme:
        import brand
        pal, info = brand.apply(out, logo=args.logo or None, theme=args.theme or None)
        brand_line = brand.describe(info)

    print(f"\n[通过] 已写入 → {out}")
    print(f"  paper.json（{len(qs)} 题 / {total:g} 分）"
          f"　录分结果.json（{len(students)} 人）"
          + (f"　原卷 → {paper['exam']['paper_file']}" if paper["exam"]["paper_file"] else ""))
    if brand_line:
        print("  " + brand_line.replace("\n", "\n  "))
    notes = sum(1 for q in qs if q.get("note")) + sum(1 for s in students if s.get("note"))
    if paper["exam"]["note"] or notes:
        print(f"  备注：考试备注{'有' if paper['exam']['note'] else '无'}，"
              f"题目备注 {sum(1 for q in qs if q.get('note'))} 条，学生备注 {sum(1 for s in students if s.get('note'))} 条")
    blank = [s["姓名"] for s in students if len(s["scores"]) < len(qs)]
    if blank:
        print(f"  [提示] {len(blank)} 人题分未录全（空白视为未录，不等于 0 分）：{'、'.join(blank[:6])}")
    print(f"\n下一步：python scripts/build_reports.py --exam-dir \"{out}\" --no-pdf")


if __name__ == "__main__":
    main()
