#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlsxlite —— 纯标准库读写 .xlsx（不依赖 openpyxl / pandas）

为什么自己写：WorkBuddy 自带 Python 只有 pillow，装不了第三方库；而 Excel 模板必须能被
老师用 Excel/WPS 打开、存回后还能被脚本读回。

读（Book）
  · 共享字符串 / inlineStr / str / b / 数字 / 公式缓存值
  · 自闭合单元格 <c r="C1" s="3"/> 与自闭合空行 <row r="4"/>（旧正则读取器在这里会错列）
  · 合并单元格（值取左上角）、隐藏行列、日期序列号（按 styles.xml 的格式判定）
  用法：bk = Book(path); bk.names(); bk.grid("Sheet1")  # {行号: {列字母: 文本}}

写（write_book）
  两张以上工作表、列宽、冻结窗格、合并、下拉数据验证、四种基础样式（标题/表头/正文/提示）。
  只写 inlineStr，不生成 sharedStrings，避免老师改完再存时索引错乱。
"""
import re
import zipfile
from datetime import datetime, timedelta
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# 样式序号（与 write_book 写出的 styles.xml 中 cellXfs 顺序一致）
ST_TITLE, ST_HEAD, ST_BODY, ST_WRAP = 1, 2, 3, 4
ST_KEY, ST_NOTE, ST_NUM = 5, 6, 7

_EPOCH = datetime(1899, 12, 30)


# ---------------------------------------------------------------- 读
class Book:
    def __init__(self, path):
        self.path = Path(path)
        self.z = zipfile.ZipFile(self.path)
        self.sst = self._shared_strings()
        self._date_styles = self._date_style_indexes()
        self.sheets = self._sheet_map()

    # -- 生命周期（Windows 上不关句柄会挡住删除/覆盖同名文件）
    def close(self):
        try:
            self.z.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()

    # -- 内部
    def _shared_strings(self):
        if "xl/sharedStrings.xml" not in self.z.namelist():
            return []
        root = ET.fromstring(self.z.read("xl/sharedStrings.xml"))
        out = []
        for si in root.findall(f"{NS}si"):
            # 富文本 <r><t> 拼起来；注音 <rPh> 不要
            parts = []
            for node in si.iter():
                if node.tag == f"{NS}t":
                    parts.append(node.text or "")
            out.append("".join(parts))
        return out

    def _date_style_indexes(self):
        """cellXfs 下标 → 是否日期格式"""
        if "xl/styles.xml" not in self.z.namelist():
            return set()
        root = ET.fromstring(self.z.read("xl/styles.xml"))
        custom = {}
        nf = root.find(f"{NS}numFmts")
        if nf is not None:
            for f in nf.findall(f"{NS}numFmt"):
                custom[int(f.get("numFmtId"))] = f.get("formatCode") or ""
        builtin_date = set(range(14, 23)) | {45, 46, 47}
        out = set()
        xfs = root.find(f"{NS}cellXfs")
        if xfs is None:
            return out
        for i, xf in enumerate(xfs.findall(f"{NS}xf")):
            fid = int(xf.get("numFmtId") or 0)
            code = custom.get(fid, "")
            if fid in builtin_date or re.search(r"[ymd]", re.sub(r'"[^"]*"', "", code), re.I):
                out.add(i)
        return out

    def _sheet_map(self):
        wb = ET.fromstring(self.z.read("xl/workbook.xml"))
        rels = ET.fromstring(self.z.read("xl/_rels/workbook.xml.rels"))
        relmap = {}
        for rel in rels:
            relmap[rel.get("Id")] = rel.get("Target")
        out = {}
        sheets = wb.find(f"{NS}sheets")
        for sh in (sheets if sheets is not None else []):
            name = sh.get("name")
            rid = sh.get(f"{REL_NS}id")
            target = relmap.get(rid, "")
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target.replace("worksheets/", "worksheets/")
            out[name] = target
        return out

    def names(self):
        return list(self.sheets)

    def grid(self, name):
        """{行号: {列字母: 文本}}，空单元格不出现"""
        if name not in self.sheets:
            raise KeyError(f"没有工作表：{name}（可用：{self.names()}）")
        root = ET.fromstring(self.z.read(self.sheets[name]))
        data = root.find(f"{NS}sheetData")
        out = {}
        if data is None:
            return out
        for row in data.findall(f"{NS}row"):
            r = int(row.get("r"))
            cells = {}
            for c in row.findall(f"{NS}c"):
                ref = c.get("r") or ""
                col = re.match(r"[A-Z]+", ref)
                if not col:
                    continue
                val = self._cell_value(c)
                if val is not None and val != "":
                    cells[col.group(0)] = val
            if cells:
                out[r] = cells
        return out

    def _cell_value(self, c):
        t = (c.get("t") or "n").lower()
        style = int(c.get("s") or 0)
        if t == "inlinestr":
            return "".join(n.text or "" for n in c.iter(f"{NS}t")).strip()
        if t == "s":
            v = c.find(f"{NS}v")
            if v is None or not (v.text or "").strip():
                return ""
            idx = int(float(v.text))
            return self.sst[idx].strip() if 0 <= idx < len(self.sst) else ""
        if t == "str":
            v = c.find(f"{NS}v")
            return (v.text or "").strip() if v is not None else ""
        if t == "b":
            v = c.find(f"{NS}v")
            return ("是" if (v is not None and v.text == "1") else "否")
        v = c.find(f"{NS}v")
        if v is None or v.text is None:
            return ""
        raw = v.text.strip()
        if style in self._date_styles and re.fullmatch(r"-?\d+(\.\d+)?", raw):
            try:
                dt = _EPOCH + timedelta(days=float(raw))
                if dt.year >= 1900:
                    return dt.strftime("%Y-%m-%d")
            except Exception:
                pass
        # 3.0 → 3，保留真实小数
        if re.fullmatch(r"-?\d+\.0+", raw):
            raw = raw.split(".")[0]
        return raw

    def merged(self, name):
        root = ET.fromstring(self.z.read(self.sheets[name]))
        mc = root.find(f"{NS}mergeCells")
        return [m.get("ref") for m in (mc if mc is not None else [])]


# ---------------------------------------------------------------- 写
def col_letters(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


STYLES_XML = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy/m/d"/></numFmts>
<fonts count="4">
 <font><sz val="11"/><color theme="1"/><name val="微软雅黑"/></font>
 <font><b/><sz val="15"/><color rgb="FF1F2B6B"/><name val="微软雅黑"/></font>
 <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="微软雅黑"/></font>
 <font><sz val="10.5"/><color rgb="FF68707C"/><name val="微软雅黑"/></font>
</fonts>
<fills count="5">
 <fill><patternFill patternType="none"/></fill>
 <fill><patternFill patternType="gray125"/></fill>
 <fill><patternFill patternType="solid"><fgColor rgb="FF3F3B92"/><bgColor indexed="64"/></patternFill></fill>
 <fill><patternFill patternType="solid"><fgColor rgb="FFF0F2F4"/><bgColor indexed="64"/></patternFill></fill>
 <fill><patternFill patternType="solid"><fgColor rgb="FFFFFBF0"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
 <border/>
 <border><left style="thin"><color rgb="FFDDE1E6"/></left><right style="thin"><color rgb="FFDDE1E6"/></right>
  <top style="thin"><color rgb="FFDDE1E6"/></top><bottom style="thin"><color rgb="FFDDE1E6"/></bottom></border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="8">
 <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
 <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment vertical="center"/></xf>
 <xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
 <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="center"/></xf>
 <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>
 <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
 <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf>
 <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>
</cellXfs>
</styleSheet>'''


def _cell_xml(ref, value, style):
    if value is None or value == "":
        if style:
            return f'<c r="{ref}" s="{style}"/>'
        return ""
    if isinstance(value, bool):
        value = "是" if value else "否"
    if isinstance(value, (int, float)):
        num = f"{value:.10g}" if isinstance(value, float) else str(value)
        return f'<c r="{ref}" s="{style}"><v>{num}</v></c>'
    return (f'<c r="{ref}" s="{style}" t="inlineStr">'
            f'<is><t xml:space="preserve">{escape(str(value))}</t></is></c>')


def _sheet_xml(sheet):
    rows = sheet.get("rows") or []
    widths = sheet.get("widths") or []
    freeze = sheet.get("freeze") or 0
    merges = sheet.get("merges") or []
    validations = sheet.get("validations") or []
    heights = sheet.get("heights") or {}

    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    out.append('<sheetViews><sheetView workbookViewId="0"'
               + (' tabSelected="1"' if sheet.get("active") else "") + ">")
    if freeze:
        out.append(f'<pane ySplit="{freeze}" topLeftCell="A{freeze + 1}" '
                   f'activePane="bottomLeft" state="frozen"/>')
    out.append("</sheetView></sheetViews>")
    out.append('<sheetFormatPr defaultRowHeight="18"/>')
    if widths:
        out.append("<cols>")
        for i, w in enumerate(widths, 1):
            out.append(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>')
        out.append("</cols>")
    out.append("<sheetData>")
    for ri, row in enumerate(rows, 1):
        cells = "".join(_cell_xml(f"{col_letters(ci)}{ri}", v, s)
                        for ci, (v, s) in enumerate(row, 1))
        ht = heights.get(ri)
        attrs = f' r="{ri}"' + (f' ht="{ht}" customHeight="1"' if ht else "")
        out.append(f"<row{attrs}>{cells}</row>" if cells else f"<row{attrs}/>")
    out.append("</sheetData>")
    if merges:
        out.append(f'<mergeCells count="{len(merges)}">')
        out += [f'<mergeCell ref="{m}"/>' for m in merges]
        out.append("</mergeCells>")
    if validations:
        out.append(f'<dataValidations count="{len(validations)}">')
        for sqref, items in validations:
            out.append(f'<dataValidation type="list" allowBlank="1" showDropDown="0" '
                       f'sqref="{sqref}"><formula1>"{items}"</formula1></dataValidation>')
        out.append("</dataValidations>")
    out.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/>')
    out.append("</worksheet>")
    return "".join(out)


def write_book(path, sheets, title="试卷分析", active=0):
    """sheets: [{"name", "rows":[[(值, 样式)]], "widths":[...], "freeze":n,
                 "merges":[...], "validations":[(sqref, "是,否")], "heights":{行:高}}]"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    n = len(sheets)
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>']
    for i in range(1, n + 1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                  f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append("</Types>")

    rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    core = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>{escape(title)}</dc:title><dc:creator>exam-analysis-report</dc:creator>
<dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
</cp:coreProperties>'''
    app = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
           '<Application>exam-analysis-report</Application></Properties>')

    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    for i, sh in enumerate(sheets, 1):
        wb.append(f'<sheet name="{escape(sh["name"])}" sheetId="{i}" r:id="rId{i}"/>')
    wb.append("</sheets></workbook>")

    wbr = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(1, n + 1):
        wbr.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/'
                   f'officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
    wbr.append(f'<Relationship Id="rId{n + 1}" Type="http://schemas.openxmlformats.org/'
               f'officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wbr.append("</Relationships>")

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", rels)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(wbr))
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, sh in enumerate(sheets, 1):
            sh = dict(sh)
            sh["active"] = (i - 1 == active)
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(sh))
    return path


def read_rows(path, sheet_name=None):
    """便捷读：返回 (表头, 数据行) 由调用方自己切；这里只给原始 grid"""
    bk = Book(path)
    name = sheet_name or bk.names()[0]
    return bk.grid(name)
