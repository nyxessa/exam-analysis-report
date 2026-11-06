#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 demo/ 的示例数据一键跑通全链路 —— 装完自验用。

跑通 = 收表 → 出报告 → 出教师版 → 质检全绿。三个学生都是虚构的，不碰任何真实数据。

用法：
  python scripts/demo.py                        # 输出到系统临时目录
  python scripts/demo.py --out "C:/tmp/演示"     # 指定输出目录（必须不存在或为空）
  python scripts/demo.py --no-pdf               # 只出 HTML，不打印 PDF

退出码 0 = 链路通；非 0 = 哪一步没过会直接打印那一步的输出。
"""
import argparse
import json
import os
import shutil
import struct
import subprocess
import sys
import tempfile
import time
import zlib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS = ROOT / "scripts"
DEMO = ROOT / "demo"
XLSX = DEMO / "示例-八年级物理.xlsx"
PANELS = DEMO / "panels.json"


def run(title, args, allow_fail=False):
    """跑一个子脚本；失败时把它的输出原样打出来再退出"""
    print(f"\n{title}")
    env = dict(os.environ, PYTHONPATH=str(SCRIPTS))
    r = subprocess.run([sys.executable, *args], cwd=str(SCRIPTS), env=env,
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    out = (r.stdout or "") + (r.stderr or "")
    for line in out.strip().splitlines()[-6:]:
        print("  " + line)
    if r.returncode != 0:
        if allow_fail:
            return False
        raise SystemExit(f"\n[失败] {title} —— 上面是脚本的原始输出，按 references/troubleshooting.md 查")
    return True


def demo_logo(path, w=96, h=96):
    """现场画一个示例校徽（深红圆盘 + 白心），用来演示「Logo → 主题色自动取色」"""
    rows = bytearray()
    for y in range(h):
        rows.append(0)                       # 每行 filter=0
        for x in range(w):
            dx, dy = x - w / 2, y - h / 2
            d2 = dx * dx + dy * dy
            if d2 <= (w * 0.34) ** 2:
                px = (255, 255, 255, 255) if d2 <= (w * 0.12) ** 2 else (200, 16, 46, 255)
            else:
                px = (0, 0, 0, 0)
            rows.extend(px)

    def chunk(tag, data):
        return (len(data).to_bytes(4, "big") + tag + data
                + zlib.crc32(tag + data).to_bytes(4, "big"))

    path.write_bytes(b"\x89PNG\r\n\x1a\n"
                     + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
                     + chunk(b"IDAT", zlib.compress(bytes(rows), 9))
                     + chunk(b"IEND", b""))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="", help="输出目录；不填就用系统临时目录")
    ap.add_argument("--no-pdf", action="store_true", help="只出 HTML，不打 PDF")
    args = ap.parse_args()

    for f in (XLSX, PANELS):
        if not f.exists():
            raise SystemExit(f"[错误] 缺示例文件：{f}\n"
                             "  这个包应该自带 demo/ 目录，可能是解压不完整。")

    if args.out:
        out = Path(args.out).expanduser().resolve()
        if out.exists() and any(out.iterdir()):
            raise SystemExit(f"[错误] 输出目录不是空的：{out}\n  换一个目录，或先手动清空。")
    else:
        out = Path(tempfile.gettempdir()) / f"exam-demo-{time.strftime('%H%M%S')}"
    out.mkdir(parents=True, exist_ok=True)

    logo = out / "_示例校徽.png"
    demo_logo(logo)

    print("试卷分析报告 · 示例自验")
    print(f"  skill 目录：{ROOT}")
    print(f"  输出目录：{out}")

    run("① 收表（校验到单元格 + 示例 Logo 自动取色）",
        ["ingest_excel.py", "--xlsx", str(XLSX), "--out", str(out), "--logo", str(logo)])
    logo.unlink(missing_ok=True)

    theme = ""
    try:
        theme = ((json.loads((out / "paper.json").read_text(encoding="utf-8")).get("brand") or {})
                 .get("color") or "")
    except Exception:
        pass
    if theme:
        print(f"   主题色：{theme}（从示例 Logo 自动取色，不是默认蓝）")
        import brand
        if theme.upper() == brand.DEFAULT:
            raise SystemExit("[失败] 示例 Logo 没取到颜色，主题色还是默认蓝 —— 取色链路有问题")
    else:
        raise SystemExit("[失败] 示例 Logo 没写进 paper.json 的 brand 配置")

    shutil.copyfile(PANELS, out / "panels.json")
    print(f"\n② 放入示例点评\n  用的是 demo/panels.json（3 段成品点评，供参考）")

    pdf_ok = False
    if not args.no_pdf:
        pdf_ok = run("③ 出报告（含 PDF）",
                     ["build_reports.py", "--exam-dir", str(out)], allow_fail=True)
        if not pdf_ok:
            print("  [说明] 上面是打印 PDF 失败。先按提示装 Edge，")
            print("         或加 --no-pdf 只验排版，不影响验收流程本身。")
    if args.no_pdf or not pdf_ok:
        run("③ 出报告（仅 HTML）",
            ["build_reports.py", "--exam-dir", str(out), "--no-pdf"])

    if pdf_ok:
        run("④ 教师版卷面分析",
            ["build_class_panel.py", "--exam-dir", str(out)], allow_fail=True)

    ok = run("⑤ 质检（14 项）", ["verify.py", "--exam-dir", str(out)], allow_fail=True)

    print("\n" + "=" * 56)
    if not ok:
        print("[未全绿] 质检有项没过，把上面的提示发出来一起看。")
        return 1

    reports = out / "reports"
    pdfs = sorted(reports.glob("*.pdf"))
    print(f"[全绿] 链路通了：{len(pdfs)} 份 PDF")
    for p in pdfs:
        print(f"  · {p.name}")
    if not pdfs:
        print("  （--no-pdf 模式只出 HTML，没有 PDF）")
    print(f"\n打开看看：{reports}")
    print(f"示例数据在：{DEMO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
