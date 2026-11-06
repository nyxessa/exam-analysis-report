#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""环境自检：拿到 skill 包后先跑这个，确认这台机器能出报告。

用法：
  python scripts/selftest.py          # 查依赖、素材、Excel 读写
  python scripts/selftest.py --pdf    # 再实际打一页 PDF（会启动 Edge，约 3 秒）

退出码 0 = 全过；1 = 有硬伤（按提示修，别急着跑正式流程）。
"""
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))


def _mark(mark, fallback):
    """旧版 Windows 控制台是 GBK，✓/✗ 编不出来会直接把自检搞崩，换成 √/×"""
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        mark.encode(enc)
        return "  " + mark
    except Exception:
        return "  " + fallback


OK, BAD, WARN = _mark("✓", "√"), _mark("✗", "×"), "  !"
fails, warns = [], []


def title(text):
    print(f"\n{text}")


def report(cond, label, detail="", hard=True):
    if cond:
        print(f"{OK} {label}" + (f"　{detail}" if detail else ""))
    else:
        print(f"{(BAD if hard else WARN)} {label}" + (f"　{detail}" if detail else ""))
        (fails if hard else warns).append(label)
    return cond


def main():
    for stream in (sys.stdout, sys.stderr):     # 兜底：编不出的字符用 ? 代替，别把自检自己搞崩
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    want_pdf = "--pdf" in sys.argv
    print("试卷分析报告 · 环境自检")
    print(f"skill 目录：{ROOT}")

    title("① Python 与依赖")
    v = sys.version_info
    report(v >= (3, 9), f"Python {v.major}.{v.minor}.{v.micro}", "需要 3.9 及以上")
    try:
        import xlsxlite  # noqa: F401
        import make_template  # noqa: F401
        report(True, "标准库以外无依赖", "Excel 读写走自带 xlsxlite")
    except Exception as exc:
        report(False, "脚本导入失败", repr(exc))

    title("② 素材与脚本是否齐全")
    need = [
        "SKILL.md",
        "assets/试卷分析模板.xlsx",
        "assets/report.html",
        "assets/class_panel.html",
        "assets/error-tags.json",
        "scripts/xlsxlite.py",
        "scripts/make_template.py",
        "scripts/ingest_excel.py",
        "scripts/build_reports.py",
        "scripts/build_class_panel.py",
        "scripts/evidence.py",
        "scripts/brand.py",
        "scripts/verify.py",
    ]
    missing = [f for f in need if not (ROOT / f).exists()]
    report(not missing, f"{len(need)} 个关键文件", "缺：" + "、".join(missing) if missing else "")

    # Logo 不进包：公开版不带任何机构标识，收到 Logo 后用 scripts/brand.py 上传
    has_logo = (ROOT / "assets" / "logo.png").exists()
    report(True, "Logo（可选）",
           "assets/logo.png 已存在（内部版）" if has_logo
           else "本包不含 Logo，报告顶部留空；收到 Logo 后用 --logo 上传")
    try:
        import brand
        report(True, "主题色", f"默认蓝 {brand.palette()['accent']}；可随 Logo 自动取色，也可 --theme 指定")
    except Exception as exc:
        report(False, "主题色模块导入失败", repr(exc))

    fm = (ROOT / "SKILL.md").read_text(encoding="utf-8")
    report(fm.startswith("---") and "\nname:" in fm and "\ndescription:" in fm,
           "SKILL.md frontmatter", "name / description 齐")

    title("③ 模板能不能读")
    try:
        import xlsxlite as X
        tpl = ROOT / "assets" / "试卷分析模板.xlsx"
        with X.Book(tpl) as bk:
            names = bk.names()
            g = bk.grid("学生成绩")
        report(names == ["试卷信息", "学生成绩"], "模板两个 Sheet", "、".join(names))
        header = [str(x) for x in g.get(2, {}).values()]
        qcols = [h for h in header if h.startswith("第") and h.endswith("题")]
        report(len(qcols) >= 10, f"题目列 {len(qcols)} 个", "题更多时用 make_template.py 重生成")
    except Exception as exc:
        report(False, "模板读取失败", repr(exc))

    title("④ Excel 写读闭环")
    try:
        import make_template as T
        sheets = T.build(questions=3, students=2)
        tmp = Path(tempfile.gettempdir()) / "selftest-模板.xlsx"
        X.write_book(tmp, sheets, title="自检")
        with X.Book(tmp) as back:
            names = back.names()
            header = [str(v) for v in back.grid("学生成绩").get(2, {}).values()]
        ok = names == ["试卷信息", "学生成绩"] and "第3题" in header
        report(ok, "写一份 3 题模板再读回", f"{tmp.stat().st_size / 1024:.0f} KB，题列对得上")
        tmp.unlink(missing_ok=True)
    except Exception as exc:
        report(False, "Excel 写读失败", repr(exc))

    title("⑤ 打印 PDF 的浏览器")
    edge = None
    try:
        import build_reports as B
        edge = B.find_edge()
        report(True, "找到浏览器", edge)
    except SystemExit:
        report(False, "没找到 Edge/Chrome", "装 Edge，或用环境变量 EDGE_PATH 指定路径")

    if want_pdf and edge:
        title("⑥ 实际打一页 PDF")
        try:
            tmpdir = Path(tempfile.mkdtemp(prefix="selftest_pdf_"))
            html = tmpdir / "t.html"
            html.write_text("<html><body style='font:16px sans-serif'>试卷分析报告 · 自检</body></html>",
                            encoding="utf-8")
            pdf = tmpdir / "t.pdf"
            import build_reports as B
            B.print_pdf(edge, html, pdf, tmpdir / "profile")
            size = pdf.stat().st_size
            report(size > 2000, "打印成功", f"{size / 1024:.0f} KB")
            if size > 2000:
                raw = pdf.read_bytes()
                report(b"/Type /Page" in raw or b"/Type/Page" in raw, "PDF 结构正常")
            import shutil
            shutil.rmtree(tmpdir, ignore_errors=True)
        except SystemExit as exc:
            report(False, "打印失败", str(exc)[:200])
        except Exception as exc:
            report(False, "打印异常", repr(exc))

    print()
    if fails:
        print(f"[未通过] {len(fails)} 项需要处理：" + "；".join(fails))
        print("         按上面的提示修，修不好把这段输出发给维护人。")
        return 1
    if warns:
        print(f"[可用] 有 {len(warns)} 项提示：" + "；".join(warns))
    else:
        print("[全过] 环境没问题，可以收表出报告了。")
    if not want_pdf:
        print("       想连打印链路一起验，加 --pdf 再跑一次。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
