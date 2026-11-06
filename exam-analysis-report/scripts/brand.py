#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""机构品牌：Logo 与报告主题色。

- 默认主题色是蓝紫 #3F3B92；给了机构 Logo 就从 Logo 里取主色（自动校正到印刷可读）。
- 取色顺序：SVG 读颜色 → PNG 自带解码（零依赖）→ 系统工具 / Pillow 转 PNG（JPG 等）。
- 配置写进考试文件夹的 paper.json → "brand"，所以换 Logo / 换颜色只要重跑 build_reports.py，
  不用重新收表。

用法：
  python scripts/brand.py --show --exam-dir "考试文件夹"
  python scripts/brand.py --exam-dir "考试文件夹" --logo "校徽.png"       # 存 Logo + 自动取色
  python scripts/brand.py --exam-dir "考试文件夹" --theme "#C8102E"       # 手动指定颜色
  python scripts/brand.py --exam-dir "考试文件夹" --theme auto            # 按已有 Logo 重新取色
  python scripts/brand.py --exam-dir "考试文件夹" --clear                 # 去掉 Logo，回默认蓝
"""
import argparse
import json
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import zlib
from pathlib import Path

HERE = Path(__file__).resolve().parent
SKILL = HERE.parent

DEFAULT = "#3F3B92"          # 默认主题色（蓝紫）
CRITICAL = "#ED3C3A"         # 失分/薄弱，语义色，永远红
LOGO_SUFFIX = {".png", ".jpg", ".jpeg", ".svg", ".webp", ".bmp", ".gif"}


# ---------------------------------------------------------------- 颜色基础
def parse_color(value):
    """'#abc' / '#AABBCC' / 'aabbcc' → '#AABBCC'；不合法返回 None"""
    s = str(value or "").strip()
    if not s.startswith("#"):
        s = "#" + s
    if re.fullmatch(r"#[0-9a-fA-F]{3}", s):
        s = "#" + "".join(c * 2 for c in s[1:])
    if re.fullmatch(r"#[0-9a-fA-F]{6}", s):
        return s.upper()
    return None


def hex_to_rgb(hexv):
    h = parse_color(hexv) or DEFAULT
    return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))


def rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, round(c))) for c in rgb))


def rgb_to_hsl(r, g, b):
    r, g, b = r / 255.0, g / 255.0, b / 255.0
    mx, mn = max(r, g, b), min(r, g, b)
    l = (mx + mn) / 2
    if mx == mn:
        return 0.0, l, 0.0
    d = mx - mn
    s = d / (2 - mx - mn) if l > 0.5 else d / (mx + mn)
    if mx == r:
        h = ((g - b) / d) % 6
    elif mx == g:
        h = (b - r) / d + 2
    else:
        h = (r - g) / d + 4
    return h / 6.0, l, s


def hsl_to_rgb(h, s, l):
    if s <= 0:
        v = round(l * 255)
        return (v, v, v)
    q = l * (1 + s) if l < 0.5 else l + s - l * s
    p = 2 * l - q

    def hue(t):
        t %= 1.0
        if t < 1 / 6:
            return p + (q - p) * 6 * t
        if t < 0.5:
            return q
        if t < 2 / 3:
            return p + (q - p) * (2 / 3 - t) * 6
        return p
    return tuple(round(x * 255) for x in (hue(h + 1 / 3), hue(h), hue(h - 1 / 3)))


def hsl_to_hex(h, s, l):
    return rgb_to_hex(hsl_to_rgb(h, s, l))


def _luminance(rgb):
    def lin(c):
        c = c / 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (lin(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def palette(accent=DEFAULT):
    """由主题色派生一整套印刷用色（同色相，明度分层）"""
    hexv = parse_color(accent) or DEFAULT
    h, l, s = rgb_to_hsl(*hex_to_rgb(hexv))

    def mk(l2, s2):
        return hsl_to_hex(h, min(max(s * s2, 0.0), 1.0), l2)

    # 机构色是红的时，图表里的「低于 60%」也是红 —— 把正常柱压深一档，两者才分得开
    near_red = h > 0.94 or h < 0.06
    bar = hsl_to_hex(h, min(max(s, 0.4), 0.9), 0.30) if near_red else hexv
    return {
        "accent": hexv,
        "bar": bar,                     # 图表里的「正常/达标」柱色
        "soft": mk(0.955, 0.30),        # 表头底色一类的大面积浅底
        "line": mk(0.885, 0.26),        # 细分隔线
        "badge": mk(0.935, 0.42),       # 小标签底
        "connector": mk(0.78, 0.30),    # 图表连接线（浅一档）
        "critical": CRITICAL,
        "near_red": near_red,
    }


def printable(rgb):
    """把 Logo 里的颜色校正成「白纸上打印得清」的主题色：保留色相，压明度到可读区间"""
    h, l, s = rgb_to_hsl(*rgb)
    s = min(max(s, 0.40), 0.85)
    l = min(max(l, 0.33), 0.46)
    out = hsl_to_rgb(h, s, l)
    # 最亮的一档（黄、青）也压到对比度够用：相对亮度 ≤ 0.19 时对白底约 4.5:1
    while _luminance(out) > 0.19 and l > 0.22:
        l -= 0.02
        out = hsl_to_rgb(h, s, l)
    return rgb_to_hex(out)


# ---------------------------------------------------------------- 取色
def _dominant(pixels):
    """按 色相分桶 × 饱和度加权 找主色；黑白灰与近透明像素跳过"""
    bins = {}
    for px in pixels:
        if len(px) == 4:
            r, g, b, a = px
            if a < 96:
                continue
        else:
            r, g, b = px
        h, l, s = rgb_to_hsl(r, g, b)
        if s < 0.18 or l < 0.06 or l > 0.96:
            continue
        w = s * (1.0 - 0.55 * abs(l - 0.5))
        if w <= 0:
            continue
        d = bins.setdefault(int(h * 36) % 36, [0.0, 0.0, 0.0, 0.0])
        d[0] += w
        d[1] += w * r
        d[2] += w * g
        d[3] += w * b
    if not bins:
        return None
    w, wr, wg, wb = max(bins.values(), key=lambda v: v[0])
    if w <= 0:
        return None
    return (wr / w, wg / w, wb / w)


_SVG_COLOR = re.compile(
    r"(?:fill|stop-color|stroke)\s*[:=]\s*[\"']?\s*(#[0-9a-fA-F]{3,8}|rgb\([^)]*\))")


def _svg_colors(text):
    out = []
    for m in _SVG_COLOR.finditer(text):
        c = m.group(1)
        if c.startswith("#"):
            hexv = parse_color(c[:7])
            if hexv:
                out.append(hex_to_rgb(hexv))
        else:
            nums = re.findall(r"\d+", c)
            if len(nums) >= 3:
                out.append(tuple(min(255, int(x)) for x in nums[:3]))
    return out


def _png_pixels(data, max_pixels=4_000_000, max_rows=None):
    """最小 PNG 解码：8/16 位、非隔行，color type 0/2/3/4/6。返回 RGBA 列表"""
    if data[:8] != b"\x89PNG\r\n\x1a\n":
        raise ValueError("不是 PNG")
    pos, idat, plte, trns = 8, bytearray(), None, None
    w = h = depth = ctype = interlace = None
    while pos + 8 <= len(data):
        ln = int.from_bytes(data[pos:pos + 4], "big")
        typ = data[pos + 4:pos + 8]
        chunk = data[pos + 8:pos + 8 + ln]
        pos += 12 + ln
        if typ == b"IHDR":
            w = int.from_bytes(chunk[0:4], "big")
            h = int.from_bytes(chunk[4:8], "big")
            depth, ctype, interlace = chunk[8], chunk[9], chunk[12]
        elif typ == b"PLTE":
            plte = chunk
        elif typ == b"tRNS":
            trns = chunk
        elif typ == b"IDAT":
            idat += chunk
        elif typ == b"IEND":
            break
    if not w or not h:
        raise ValueError("PNG 头缺失")
    if interlace:
        raise ValueError("隔行 PNG 不支持")
    if depth not in (8, 16) or ctype not in (0, 2, 3, 4, 6):
        raise ValueError(f"不支持 PNG 格式（depth={depth}, type={ctype}）")
    if ctype == 3 and not plte:
        raise ValueError("调色板 PNG 缺 PLTE")
    if max_rows is None and w * h > max_pixels:
        raise ValueError("too large")

    raw = zlib.decompress(bytes(idat))
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype]
    bpp = max(1, channels * depth // 8)
    stride = w * bpp
    rows = h if max_rows is None else min(h, max_rows)
    row_step = max(1, rows // 128)
    col_step = max(1, w // 128)
    prev = bytearray(stride)
    p = 0
    out = []
    for y in range(rows):
        if p >= len(raw):
            break
        f = raw[p]
        p += 1
        line = bytearray(raw[p:p + stride])
        p += stride
        if f == 1:
            for i in range(bpp, stride):
                line[i] = (line[i] + line[i - bpp]) & 255
        elif f == 2:
            for i in range(stride):
                line[i] = (line[i] + prev[i]) & 255
        elif f == 3:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                line[i] = (line[i] + ((a + prev[i]) >> 1)) & 255
        elif f == 4:
            for i in range(stride):
                a = line[i - bpp] if i >= bpp else 0
                b = prev[i]
                c = prev[i - bpp] if i >= bpp else 0
                pp = a + b - c
                pa, pb, pc = abs(pp - a), abs(pp - b), abs(pp - c)
                pr = a if (pa <= pb and pa <= pc) else (b if pb <= pc else c)
                line[i] = (line[i] + pr) & 255
        prev = line
        if y % row_step:
            continue
        for x in range(0, w, col_step):
            out.append(_px_at(line, x, ctype, depth, plte, trns))
    return out


def _px_at(line, x, ctype, depth, plte, trns):
    step = 2 if depth == 16 else 1
    if ctype == 0:
        return (line[x * step],) * 3 + (255,)
    if ctype == 4:
        return (line[x * 2 * step],) * 3 + (line[x * 2 * step + step],)
    if ctype == 2:
        i = x * 3 * step
        return (line[i], line[i + step], line[i + 2 * step], 255)
    if ctype == 6:
        i = x * 4 * step
        return (line[i], line[i + step], line[i + 2 * step], line[i + 3 * step])
    idx = line[x * step]
    if plte is None or idx * 3 + 2 >= len(plte):
        return (255, 255, 255, 0)
    a = trns[idx] if (trns and idx < len(trns)) else 255
    return (plte[idx * 3], plte[idx * 3 + 1], plte[idx * 3 + 2], a)


def _pixels_from_png(path):
    data = Path(path).read_bytes()
    try:
        return _png_pixels(data)
    except ValueError as exc:
        if "too large" not in str(exc):
            raise
    tmp = _convert(path)
    if tmp:
        try:
            return _png_pixels(Path(tmp).read_bytes())
        except Exception:
            pass
        finally:
            try:
                Path(tmp).unlink()
            except OSError:
                pass
    return _png_pixels(data, max_pixels=0, max_rows=256)   # 超大图只采前 256 行


def _pixels_from_pillow(path):
    try:
        from PIL import Image
    except Exception:
        return None
    try:
        with Image.open(path) as im:
            im = im.convert("RGBA")
            im.thumbnail((200, 200))
            return list(im.getdata())
    except Exception:
        return None


def _ps_path(p):
    return str(Path(p)).replace("'", "''")


def _convert(src):
    """用系统自带工具把图片缩到 200px 再存 PNG；没有可用工具就返回 None"""
    src = Path(src).resolve()
    out = Path(tempfile.mkdtemp(prefix="brand_")) / "logo.png"
    try:
        if os.name == "nt":
            exe = shutil.which("pwsh") or shutil.which("powershell")
            if not exe:
                return None
            ps = (
                "Add-Type -AssemblyName System.Drawing; "
                f"$s=[System.Drawing.Image]::FromFile('{_ps_path(src)}'); "
                "$w=200; $h=[int]($s.Height*$w/$s.Width); "
                "if($h -gt 200){$h=200; $w=[int]($s.Width*$h/$s.Height)}; "
                "if($w -lt 1){$w=1}; if($h -lt 1){$h=1}; "
                "$b=New-Object System.Drawing.Bitmap $w,$h; "
                "$g=[System.Drawing.Graphics]::FromImage($b); "
                "$g.InterpolationMode='HighQualityBicubic'; "
                "$g.DrawImage($s,0,0,$w,$h); "
                f"$b.Save('{_ps_path(out)}',[System.Drawing.Imaging.ImageFormat]::Png); "
                "$g.Dispose(); $b.Dispose(); $s.Dispose()"
            )
            subprocess.run([exe, "-NoProfile", "-NonInteractive", "-Command", ps],
                           capture_output=True, timeout=40)
        elif sys.platform == "darwin":
            subprocess.run(["sips", "-s", "format", "png", "-Z", "200", str(src), "--out", str(out)],
                           capture_output=True, timeout=40)
        else:
            exe = shutil.which("convert") or shutil.which("magick")
            if not exe:
                return None
            subprocess.run([exe, str(src), "-resize", "200x200", str(out)],
                           capture_output=True, timeout=40)
    except Exception:
        return None
    return out if out.exists() and out.stat().st_size > 0 else None


def _load_pixels(path):
    path = Path(path)
    if path.suffix.lower() == ".png":
        try:
            return _pixels_from_png(path)
        except Exception:
            pass
    px = _pixels_from_pillow(path)
    if px:
        return px
    conv = _convert(path)
    if conv:
        try:
            return _pixels_from_png(conv)
        except Exception:
            return None
        finally:
            try:
                conv.unlink()
            except OSError:
                pass
    return None


def extract(path):
    """从 Logo 取一个可用的主题色 → (hex|None, 说明)"""
    p = Path(path)
    if not p.exists():
        return None, f"找不到 {p}"
    if p.suffix.lower() == ".svg":
        colors = _svg_colors(p.read_text(encoding="utf-8", errors="replace"))
        if not colors:
            return None, "SVG 里没读到彩色（只有黑白或用了外部样式）"
        dom = _dominant([c + (255,) for c in colors])
        return (printable(dom), "svg") if dom else (None, "SVG 里没有彩色")
    pixels = _load_pixels(p)
    if not pixels:
        return None, "读不了这个图片格式（JPG 可另存 PNG，或用 --theme 直接指定颜色）"
    dom = _dominant(pixels)
    if not dom:
        return None, "Logo 是黑白/灰阶，看不出品牌色"
    return printable(dom), p.suffix.lower().lstrip(".")


# ---------------------------------------------------------------- 配置读写
def _read_paper(exam_dir):
    pp = Path(exam_dir) / "paper.json"
    if not pp.exists():
        raise SystemExit(f"[错误] 找不到 {pp}（先跑 ingest_excel.py 收表）")
    return pp, json.loads(pp.read_text(encoding="utf-8-sig"))


def _write_paper(pp, paper):
    pp.write_text(json.dumps(paper, ensure_ascii=False, indent=1), encoding="utf-8")


def logo_file(paper, exam_dir):
    """考试目录里的 Logo 优先；内部版把 Logo 放在 assets/logo.png 时兜底用它"""
    cfg = (paper or {}).get("brand") or {}
    rel = cfg.get("logo")
    if rel:
        p = Path(exam_dir) / rel
        if p.exists():
            return p
    legacy = SKILL / "assets" / "logo.png"
    return legacy if legacy.exists() else None


def resolve(paper, exam_dir):
    """→ (palette, info)；info = {accent, source, logo, note}"""
    cfg = (paper or {}).get("brand") or {}
    logo = logo_file(paper, exam_dir)
    missing = bool(cfg.get("logo")) and logo is None
    accent = parse_color(cfg.get("color"))
    source = cfg.get("color_source") or ("manual" if accent else "default")
    note = ""
    if not accent and logo:
        derived, why = extract(logo)
        if derived:
            accent, source = derived, "auto"
        else:
            note = why
    accent = accent or DEFAULT
    return palette(accent), {"accent": accent, "source": source, "logo": logo,
                             "logo_missing": missing, "note": note}


def _contrast_warning(hexv):
    """手填颜色太浅时提醒一句（打印在白纸上会发虚）"""
    lum = _luminance(hex_to_rgb(hexv))
    if lum > 0.35:
        contrast = 1.05 / (lum + 0.05)
        return f"颜色偏浅（对白底约 {contrast:.1f}:1），打印可能发虚；建议同色相再深一点"
    return ""


def describe(info):
    """给命令行 / AI 转述用的一行说明"""
    src = {"auto": "从 Logo 自动取色", "manual": "手动指定", "default": "默认蓝"}.get(info["source"], info["source"])
    if info.get("logo"):
        logo = f"Logo：{Path(info['logo']).name}　"
    elif info.get("logo_missing"):
        logo = "Logo：文件丢失（看 brand/ 目录）　"
    else:
        logo = "Logo：无（报告顶部留空）　"
    line = f"{logo}主题色：{info['accent']}（{src}）"
    if palette(info["accent"]).get("near_red"):
        line += "\n  提示：机构色与「失分」红接近，图表里的达标柱已自动压深一档区分"
    warn = _contrast_warning(info["accent"])
    if warn:
        line += f"\n  提示：{warn}"
    if info.get("note"):
        line += f"\n  提示：{info['note']}，已用默认蓝；可用 --theme 指定颜色"
    return line


def apply(exam_dir, logo=None, theme=None, clear=False):
    """把 Logo / 主题色写进 paper.json，返回 (palette, info)"""
    exam_dir = Path(exam_dir)
    pp, paper = _read_paper(exam_dir)
    cfg = dict(paper.get("brand") or {})

    if clear:
        paper.pop("brand", None)
        _write_paper(pp, paper)
        d = exam_dir / "brand"
        if d.exists():
            shutil.rmtree(d, ignore_errors=True)
        return resolve(paper, exam_dir)

    if logo:
        src = Path(logo).expanduser()
        if not src.is_absolute():
            src = (Path.cwd() / src)
        if not src.exists():
            raise SystemExit(f"[错误] 找不到 Logo：{src}")
        if src.suffix.lower() not in LOGO_SUFFIX:
            raise SystemExit(f"[错误] 不支持的图片格式：{src.suffix}"
                             "（用 png / jpg / svg / webp / bmp）")
        bd = exam_dir / "brand"
        bd.mkdir(parents=True, exist_ok=True)
        for old in bd.glob("logo.*"):
            old.unlink()
        ext = ".jpg" if src.suffix.lower() == ".jpeg" else src.suffix.lower()
        dst = bd / f"logo{ext}"
        shutil.copy2(src, dst)
        cfg["logo"] = f"brand/{dst.name}"
        cfg["logo_from"] = src.name
        cfg["color_source"] = "auto"
        cfg.pop("color", None)          # 换了 Logo 必须重算颜色

    want = (theme or "").strip()
    if want:
        if want.lower() in ("auto", "自动"):
            if not cfg.get("logo"):
                raise SystemExit("[错误] 还没有 Logo，没法自动取色（先加 --logo）")
            cfg["color_source"] = "auto"
            cfg.pop("color", None)
        else:
            hexv = parse_color(want)
            if not hexv:
                raise SystemExit(f"[错误] 颜色写法不对：{want}（要 #RRGGBB，例如 #C8102E）")
            cfg["color"] = hexv
            cfg["color_source"] = "manual"

    if cfg:
        paper["brand"] = cfg
    _write_paper(pp, paper)

    # 立即算一次并落盘：这样下次渲染不用再解码 Logo，老师也能马上看到取到的颜色
    if cfg.get("color_source") == "auto" and cfg.get("logo") and not cfg.get("color"):
        derived, why = extract(exam_dir / cfg["logo"])
        if derived:
            cfg["color"] = derived
        else:
            print(f"[提示] 从 Logo 取色失败：{why}")
        paper["brand"] = cfg
        _write_paper(pp, paper)

    return resolve(paper, exam_dir)


def main():
    ap = argparse.ArgumentParser(description="机构 Logo 与报告主题色")
    ap.add_argument("--exam-dir", default="", help="考试文件夹（含 paper.json）")
    ap.add_argument("--logo", default="", help="机构 Logo 图片（png/jpg/svg），会复制到 考试文件夹/brand/")
    ap.add_argument("--theme", default="", help="主题色 #RRGGBB，或 auto（按 Logo 取色）")
    ap.add_argument("--clear", action="store_true", help="去掉 Logo 与自定义色，回默认蓝")
    ap.add_argument("--show", action="store_true", help="只看当前配置")
    args = ap.parse_args()

    if not args.exam_dir:
        raise SystemExit("[错误] 缺少 --exam-dir")
    exam_dir = Path(args.exam_dir).resolve()

    if args.show or not (args.logo or args.theme or args.clear):
        pp, paper = _read_paper(exam_dir)
        pal, info = resolve(paper, exam_dir)
        print(f"考试文件夹：{exam_dir}")
        print(describe(info))
        print(f"色阶：主色 {pal['accent']}　浅底 {pal['soft']}　细线 {pal['line']}　"
              f"标签 {pal['badge']}　图表线 {pal['connector']}")
        return

    pal, info = apply(exam_dir, logo=args.logo or None, theme=args.theme or None,
                      clear=args.clear)
    print(describe(info))
    print(f"已写入 {exam_dir / 'paper.json'}　下一步：python scripts/build_reports.py --exam-dir \"{exam_dir}\"")


if __name__ == "__main__":
    sys.exit(main())
