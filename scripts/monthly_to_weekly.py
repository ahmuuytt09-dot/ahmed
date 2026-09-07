#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تحويل "KAZ Project Montly EHS Report.xlsx" من تقرير شهري (12 عمودًا = 12 شهرًا)
إلى تقرير أسبوعي (7 أعمدة = أيام الأسبوع Mon..Sun + عمود مجموع الأسبوع).

يعمل بالتعديل المباشر على XML داخل حزمة الـ xlsx، وليس عبر openpyxl،
للحفاظ على: 8 رسوم بيانية، 6 جداول Excel، التنسيق الشرطي، القوائم المنسدلة،
إعدادات الطباعة، الأنماط، و customXml.

الاستخدام:
    python3 scripts/monthly_to_weekly.py <input.xlsx> <output.xlsx>
"""
import re
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

DAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
TOTAL_HDR = "Week Total"

# التاريخ المرجعي: الإثنين 1 يونيو 2026 = بداية الأسبوع رقم 1
YEAR_START = "DATE(2026,6,1)"

for p, u in (("", NS), ("r", RNS)):
    ET.register_namespace(p, u)


# ----------------------------------------------------------------- utilities
def col_to_num(c):
    n = 0
    for ch in c:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


def q(tag):
    return f"{{{NS}}}{tag}"


# ------------------------------------------------------------- sheet editing
class Sheet:
    def __init__(self, xml_bytes):
        self.tree = ET.fromstring(xml_bytes)
        self.data = self.tree.find(q("sheetData"))

    # -- rows / cells -------------------------------------------------------
    def row(self, r, create=True):
        for row in self.data.findall(q("row")):
            if int(row.get("r")) == r:
                return row
        if not create:
            return None
        row = ET.Element(q("row"), {"r": str(r)})
        rows = self.data.findall(q("row"))
        idx = len(rows)
        for i, ex in enumerate(rows):
            if int(ex.get("r")) > r:
                idx = i
                break
        self.data.insert(idx, row)
        return row

    def cell(self, ref, create=False, style=None):
        col, r = split_ref(ref)
        row = self.row(r, create=create)
        if row is None:
            return None
        for c in row.findall(q("c")):
            if c.get("r") == ref:
                return c
        if not create:
            return None
        c = ET.Element(q("c"), {"r": ref})
        if style is not None:
            c.set("s", str(style))
        target = col_to_num(col)
        idx = len(row)
        for i, ex in enumerate(row.findall(q("c"))):
            if col_to_num(split_ref(ex.get("r"))[0]) > target:
                idx = i
                break
        row.insert(idx, c)
        return c

    def style_of(self, ref):
        c = self.cell(ref)
        return c.get("s") if c is not None else None

    @staticmethod
    def _clear(c):
        for child in list(c):
            c.remove(child)
        c.attrib.pop("t", None)

    def put_text(self, ref, text, style=None):
        c = self.cell(ref, create=True, style=style)
        if style is not None:
            c.set("s", str(style))
        self._clear(c)
        c.set("t", "inlineStr")
        is_el = ET.SubElement(c, q("is"))
        t = ET.SubElement(is_el, q("t"))
        t.text = text
        if text != text.strip():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    def put_number(self, ref, value, style=None):
        c = self.cell(ref, create=True, style=style)
        if style is not None:
            c.set("s", str(style))
        self._clear(c)
        v = ET.SubElement(c, q("v"))
        v.text = str(value)

    def put_formula(self, ref, formula, style=None):
        c = self.cell(ref, create=True, style=style)
        if style is not None:
            c.set("s", str(style))
        self._clear(c)
        f = ET.SubElement(c, q("f"))
        f.text = formula

    def delete_cols(self, rows, first_col, last_col):
        lo, hi = col_to_num(first_col), col_to_num(last_col)
        for r in rows:
            row = self.row(r, create=False)
            if row is None:
                continue
            for c in list(row.findall(q("c"))):
                n = col_to_num(split_ref(c.get("r"))[0])
                if lo <= n <= hi:
                    row.remove(c)

    # -- structural ---------------------------------------------------------
    def fix_spans(self):
        for row in self.data.findall(q("row")):
            cells = row.findall(q("c"))
            if not cells:
                row.attrib.pop("spans", None)
                continue
            nums = [col_to_num(split_ref(c.get("r"))[0]) for c in cells]
            row.set("spans", f"{min(nums)}:{max(nums)}")

    def fix_dimension(self):
        max_c, max_r = 1, 1
        for row in self.data.findall(q("row")):
            r = int(row.get("r"))
            for c in row.findall(q("c")):
                n = col_to_num(split_ref(c.get("r"))[0])
                max_c, max_r = max(max_c, n), max(max_r, r)
        dim = self.tree.find(q("dimension"))
        if dim is not None:
            dim.set("ref", f"A1:{num_to_col(max_c)}{max_r}")

    def merges(self):
        return self.tree.find(q("mergeCells"))

    def set_merges(self, refs):
        mc = self.merges()
        if mc is None:
            if not refs:
                return
            mc = ET.Element(q("mergeCells"))
            anchor = self.tree.find(q("sheetData"))
            self.tree.insert(list(self.tree).index(anchor) + 1, mc)
        for ch in list(mc):
            mc.remove(ch)
        for ref in refs:
            ET.SubElement(mc, q("mergeCell"), {"ref": ref})
        mc.set("count", str(len(refs)))
        if not refs:
            self.tree.remove(mc)

    def set_col_widths(self, spec):
        """spec: list of (min, max, width)"""
        cols = self.tree.find(q("cols"))
        if cols is None:
            cols = ET.Element(q("cols"))
            sd = self.tree.find(q("sheetData"))
            self.tree.insert(list(self.tree).index(sd), cols)
        existing = {(int(c.get("min")), int(c.get("max"))): c for c in cols}
        for mn, mx, w in spec:
            key = (mn, mx)
            if key in existing:
                existing[key].set("width", str(w))
                existing[key].set("customWidth", "1")
            else:
                ET.SubElement(cols, q("col"), {
                    "min": str(mn), "max": str(mx),
                    "width": str(w), "customWidth": "1"})

    def tobytes(self):
        self.fix_spans()
        self.fix_dimension()
        return b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + \
            ET.tostring(self.tree, encoding="utf-8")


# ------------------------------------------------------- block transformation
def convert_block(sh, hdr_row, data_rows, first_data_col, n_old, total_label=True):
    """
    يحوّل كتلة من 12 عمودًا شهريًا إلى 7 أعمدة أيام + عمود مجموع.
    يعيد (آخر عمود يوم، عمود المجموع).
    """
    start = col_to_num(first_data_col)
    day_cols = [num_to_col(start + i) for i in range(7)]
    total_col = num_to_col(start + 7)
    old_last = num_to_col(start + n_old - 1)

    hdr_style = sh.style_of(f"{first_data_col}{hdr_row}")

    # عناوين الأيام
    for i, dc in enumerate(day_cols):
        sh.put_text(f"{dc}{hdr_row}", DAYS[i], style=hdr_style)
    if total_label:
        sh.put_text(f"{total_col}{hdr_row}", TOTAL_HDR, style=hdr_style)

    # صفوف البيانات: الحفاظ على القيم الموجودة + دالة المجموع
    for r in data_rows:
        style = sh.style_of(f"{first_data_col}{r}")
        had_values = False
        for dc in day_cols:
            c = sh.cell(f"{dc}{r}")
            if c is not None and (c.find(q("v")) is not None or c.find(q("is")) is not None):
                had_values = True
        for dc in day_cols:
            c = sh.cell(f"{dc}{r}")
            if c is None:
                sh.cell(f"{dc}{r}", create=True, style=style)
            elif had_values and c.find(q("v")) is None and c.find(q("is")) is None:
                sh.put_number(f"{dc}{r}", 0, style=c.get("s") or style)
        sh.put_formula(f"{total_col}{r}",
                       f"SUM({day_cols[0]}{r}:{day_cols[-1]}{r})",
                       style=sh.style_of(f"{old_last}{r}") or style)

    # حذف بقية الأعمدة الشهرية القديمة
    if n_old > 8:
        sh.delete_cols(list(range(hdr_row, max(data_rows) + 1)),
                       num_to_col(start + 8), old_last)
    return day_cols, total_col


def week_banner():
    return ('"Week "&TEXT(\'EHS KPI\'!$G$7,"00")&"  |  "'
            '&TEXT(\'EHS KPI\'!$G$8,"dd mmm yyyy")&"  -  "'
            '&TEXT(\'EHS KPI\'!$G$9,"dd mmm yyyy")')


# ------------------------------------------------------------------ tables
def rewrite_table(xml_bytes, new_ref, headers):
    """يعيد ضبط نطاق الجدول وأعمدته بحيث تطابق العناوين الجديدة."""
    txt = xml_bytes.decode("utf-8")
    txt = re.sub(r'ref="[A-Z]+\d+:[A-Z]+\d+"', f'ref="{new_ref}"', txt, count=1)

    m = re.search(r"<tableColumns count=\"\d+\">(.*?)</tableColumns>", txt, re.S)
    cols = re.findall(r"<tableColumn\b.*?(?:/>|</tableColumn>)", m.group(1), re.S)

    out = []
    for i, name in enumerate(headers):
        src = cols[i] if i < len(cols) else cols[-1]
        src = re.sub(r'\sid="\d+"', f' id="{i + 1}"', src, count=1)
        src = re.sub(r'\sname="(?:[^"\\]|\\.)*"', f' name="{name}"', src, count=1)
        src = re.sub(r'\sxr3:uid="\{[^}]*\}"', "", src, count=1)
        out.append(src)

    new_block = (f'<tableColumns count="{len(headers)}">'
                 + "".join(out) + "</tableColumns>")
    txt = txt[:m.start()] + new_block + txt[m.end():]
    return txt.encode("utf-8")


# ------------------------------------------------------------------- charts
def rewrite_chart(xml_bytes, replacements):
    txt = xml_bytes.decode("utf-8")
    for old, new in replacements:
        txt = txt.replace(old, new)
    return fix_chart_caches(txt).encode("utf-8")


MONTH_RE = re.compile(
    r"^\s*(Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|Jan|Feb|Febr|Mar|Apr|May)\s*[.']?", re.I)
REF_BLOCK = re.compile(r"<c:(strRef|numRef)>(.*?)</c:\1>", re.S)
PT_RE = re.compile(r'<c:pt idx="(\d+)"([^>]*)>(.*?)</c:pt>', re.S)


def fix_chart_caches(txt):
    """يقصّ الذاكرة المؤقتة للرسوم من 12 نقطة إلى 7، ويستبدل أسماء الشهور بأيام الأسبوع."""
    def repl(m):
        kind, body = m.group(1), m.group(2)
        fm = re.search(r"<c:f>([^<]*)</c:f>", body)
        if not fm:
            return m.group(0)
        rm = re.search(r"\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+)$", fm.group(1))
        if not rm:
            return m.group(0)
        n = col_to_num(rm.group(3)) - col_to_num(rm.group(1)) + 1
        if not 1 <= n <= 8:
            return m.group(0)

        pts = PT_RE.findall(body)
        if not pts:
            body = re.sub(r'<c:ptCount val="\d+"/>', f'<c:ptCount val="{n}"/>',
                          body, count=1)
            return f"<c:{kind}>{body}</c:{kind}>"
        vals = [re.sub(r"<[^>]*>", "", p[2]) for p in pts]
        is_month_axis = kind == "strRef" and any(MONTH_RE.match(v) for v in vals)

        new_pts = []
        for i in range(n):
            if is_month_axis:
                v = DAYS[i] if i < len(DAYS) else TOTAL_HDR
            elif i < len(vals):
                v = vals[i]
            else:
                v = "0" if kind == "numRef" else ""
            attrs = pts[i][1] if i < len(pts) else (pts[0][1] if pts else "")
            new_pts.append(f'<c:pt idx="{i}"{attrs}><c:v>{v}</c:v></c:pt>')

        body = re.sub(r'<c:ptCount val="\d+"/>', f'<c:ptCount val="{n}"/>', body, count=1)
        start = body.index(PT_RE.search(body).group(0))
        end = body.rindex("</c:pt>") + len("</c:pt>")
        body = body[:start] + "".join(new_pts) + body[end:]
        return f"<c:{kind}>{body}</c:{kind}>"

    return REF_BLOCK.sub(repl, txt)


def shrink_ref(sheet, old_first, old_last, new_last, rows):
    """يبني قائمة استبدالات لنطاقات الرسوم البيانية."""
    return [(f"{sheet}!${old_first}${r}:${old_last}${r}",
             f"{sheet}!${old_first}${r}:${new_last}${r}") for r in rows]


# ========================================================================= main
def main(src, dst):
    zin = zipfile.ZipFile(src)
    items = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    # ---- خريطة الأوراق -> ملفات
    wb = ET.fromstring(items["xl/workbook.xml"])
    rels = ET.fromstring(items["xl/_rels/workbook.xml.rels"])
    rid2t = {r.get("Id"): r.get("Target") for r in rels}
    sheets = {}
    for s in wb.find(q("sheets")):
        tgt = rid2t[s.get(f"{{{RNS}}}id")].lstrip("/")
        if not tgt.startswith("xl/"):
            tgt = "xl/" + tgt
        sheets[s.get("name")] = tgt

    S = {name: Sheet(items[path]) for name, path in sheets.items()}

    # جدول النصوص المشتركة (لقراءة نص الخلايا من نوع t="s")
    shared = []
    if "xl/sharedStrings.xml" in items:
        sst = ET.fromstring(items["xl/sharedStrings.xml"])
        for si in sst.findall(q("si")):
            shared.append("".join(t.text or "" for t in si.iter(q("t"))))

    def cell_text(c):
        if c is None:
            return ""
        v = c.find(q("v"))
        if c.get("t") == "s" and v is not None:
            return shared[int(v.text)]
        if c.get("t") == "inlineStr":
            return "".join(t.text or "" for t in c.iter(q("t")))
        return v.text if v is not None else ""

    COMM = "EHS Communication (2)"
    INC = "EHS Incıdent"
    KPI = "EHS KPI"
    ENV = "Environmental"

    day_hdrs = DAYS + [TOTAL_HDR]

    # ------------------------------------------------ 1) EHS Communication (2)
    sh = S[COMM]
    convert_block(sh, 2, list(range(3, 11)), "B", 12)
    sh.set_merges([])                       # إزالة الدمج المتبقي J1:K1 / L1:M1
    sh.delete_cols([1], "J", "M")
    sh.put_formula("K1", week_banner(), style=sh.style_of("A1"))
    sh.set_col_widths([(2, 8, 9.0), (9, 9, 12.0), (11, 11, 42.0)])

    # ------------------------------------------------------- 2) EHS Incıdent
    sh = S[INC]
    convert_block(sh, 2, list(range(3, 11)), "B", 12)          # Re-Active Events
    convert_block(sh, 2, [3, 4], "P", 12)                      # LTC - LTD
    convert_block(sh, 32, [33, 34], "B", 12)                   # Pro-Active Events
    convert_block(sh, 32, [33, 34, 35], "P", 12)               # CAPA List
    sh.put_formula("K1", week_banner(), style=sh.style_of("A1"))
    sh.set_col_widths([(2, 8, 9.0), (9, 9, 12.0),
                       (11, 11, 42.0),
                       (16, 22, 9.0), (23, 23, 12.0)])

    # ------------------------------------------------------------ 3) EHS KPI
    sh = S[KPI]
    for hdr, rows in ((11, [12, 13, 14]), (18, list(range(19, 27))), (29, [30, 31, 32])):
        convert_block(sh, hdr, rows, "C", 12, total_label=False)

    lbl = sh.style_of("A3") or sh.style_of("B3")
    val = sh.style_of("G5") or sh.style_of("H3")
    sh.put_text("B7", "Week No. (1-52)", style=lbl)
    sh.put_number("G7", 1, style=val)
    sh.put_text("B8", "Week Start (Monday)", style=lbl)
    sh.put_formula("G8", f"{YEAR_START}+(G7-1)*7", style=val)
    sh.put_text("B9", "Week End (Sunday)", style=lbl)
    sh.put_formula("G9", "G8+6", style=val)

    tot = "J"                                    # عمود مجموع الأسبوع (كان O = YTD)
    for hdr in (11, 18, 29):
        sh.put_text(f"{tot}{hdr}", TOTAL_HDR, style=sh.style_of(f"C{hdr}"))

    sh.put_formula(f"{tot}12", f"SUM(C12:I12)", style=sh.style_of("C12"))
    sh.put_formula(f"{tot}13", f"SUM(C13:I13)", style=sh.style_of("C13"))
    for c in "CDEFGHI":
        sh.put_formula(f"{c}14", f"SUM({c}12,{c}13)", style=sh.style_of("C14"))
    sh.put_formula(f"{tot}14", f"SUM({tot}12,{tot}13)", style=sh.style_of("C14"))

    for r in range(19, 27):
        sh.put_formula(f"{tot}{r}", f"SUM(C{r}:I{r})", style=sh.style_of(f"C{r}"))

    for c in "CDEFGHI":
        sh.put_formula(
            f"{c}30",
            f"IF(SUM($C14:{c}14)<>0,(SUM($C19:{c}20))*200000/(SUM($C14:{c}14)),0)",
            style=sh.style_of("C30"))
        sh.put_formula(
            f"{c}31",
            f"IF((SUM($C14:{c}14))<>0,((SUM($C20:{c}22))*200000)/(SUM($C14:{c}14)),0)",
            style=sh.style_of("C31"))
        sh.put_formula(f"{c}32", f"IF({c}14<>0,{c}26*1000/{c}14,0)",
                       style=sh.style_of("C32"))
    sh.put_formula(f"{tot}30",
                   f"IF({tot}14<>0,({tot}19+{tot}20)*200000/{tot}14,0)",
                   style=sh.style_of("C30"))
    sh.put_formula(f"{tot}31",
                   f"IF({tot}14<>0,(SUM({tot}20:{tot}22))*200000/{tot}14,0)",
                   style=sh.style_of("C31"))
    sh.put_formula(f"{tot}32", f"IF({tot}14<>0,{tot}26*1000/{tot}14,0)",
                   style=sh.style_of("C32"))

    sh.delete_cols(list(range(10, 33)), "L", "AB")   # بقايا الأعمدة الشهرية

    keep, new_m = [], []
    for mc in (sh.merges() or []):
        ref = mc.get("ref")
        a, b = ref.split(":")
        ca, ra = split_ref(a)
        cb, _ = split_ref(b)
        if ca == "O" and cb == "P":                  # YTD -> Week Total
            new_m.append(f"J{ra}:K{ra}")
        elif ca == "A" and cb == "P":                # أشرطة العناوين
            new_m.append(f"A{ra}:K{ra}")
        else:
            keep.append(ref)
    sh.set_merges(keep + new_m + ["B7:F7", "B8:F8", "B9:F9"])
    sh.set_col_widths([(3, 9, 9.0), (10, 11, 11.0)])

    # ------------------------------------------------------- 4) Environmental
    sh = S[ENV]
    convert_block(sh, 2, list(range(3, 11)), "B", 12)
    for r in range(3, 11):                           # نقل عمود الوحدات N -> J
        src = sh.cell(f"N{r}")
        if src is not None:
            txt = cell_text(src).strip()
            sh.put_text(f"J{r}", f" {txt}" if txt else " ", style=src.get("s"))
    sh.delete_cols(list(range(1, 11)), "K", "N")
    sh.set_merges([])
    sh.put_formula("L1", week_banner(), style=sh.style_of("A1"))
    sh.set_col_widths([(2, 8, 9.0), (9, 9, 12.0), (10, 10, 8.0), (12, 12, 42.0)])

    for name, path in sheets.items():
        items[path] = S[name].tobytes()

    # --------------------------------------------------------------- الجداول
    table_map = {
        "xl/tables/table1.xml": ("A2:I10", day_hdrs),     # Table19  Communication
        "xl/tables/table2.xml": ("A2:I10", day_hdrs),     # Table14  Incident
        "xl/tables/table3.xml": ("A32:I34", day_hdrs),    # Table16  Incident
        "xl/tables/table4.xml": ("O32:W35", day_hdrs),    # Table17  Incident
        "xl/tables/table5.xml": ("O2:W4", day_hdrs),      # Table15  Incident
        "xl/tables/table6.xml": ("A2:I10", day_hdrs),     # Table1   Environmental
    }
    for path, (ref, hdrs) in table_map.items():
        blk = re.search(r"<tableColumns count=\"\d+\">(.*?)</tableColumns>",
                        items[path].decode("utf-8"), re.S).group(1)
        first = re.search(r'<tableColumn\b[^>]*?\sname="((?:[^"\\]|\\.)*)"', blk).group(1)
        items[path] = rewrite_table(items[path], ref, [first] + hdrs)

    # ---------------------------------------------------------------- الرسوم
    chart_repl = {
        "xl/charts/chart1.xml": shrink_ref(f"'{COMM}'", "B", "M", "H", range(2, 11)),
        "xl/charts/chart2.xml": shrink_ref(f"'{INC}'", "B", "M", "H", range(2, 11)),
        "xl/charts/chart3.xml": shrink_ref(f"'{INC}'", "P", "AA", "V", range(2, 5)),
        "xl/charts/chart4.xml": shrink_ref(f"'{INC}'", "B", "M", "H", range(32, 35)),
        "xl/charts/chart5.xml": shrink_ref(f"'{INC}'", "P", "AA", "V", range(32, 36)),
        "xl/charts/chart7.xml": (
            shrink_ref(f"'{KPI}'", "C", "N", "I", [11, 12, 13, 30, 31])),
        "xl/charts/chart8.xml": shrink_ref(f"{ENV}", "B", "M", "H", range(2, 11)),
    }
    for path, repl in chart_repl.items():
        if path in items:
            items[path] = rewrite_chart(items[path], repl)

    # ------------------------------------------- منطقة الطباعة + إعادة الحساب
    wbtxt = items["xl/workbook.xml"].decode("utf-8")
    wbtxt = wbtxt.replace("'EHS KPI'!$A$1:$S$54", "'EHS KPI'!$A$1:$N$54")
    wbtxt = re.sub(r"<calcPr[^/]*/>", '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wbtxt)
    if "<calcPr" not in wbtxt:
        wbtxt = wbtxt.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    items["xl/workbook.xml"] = wbtxt.encode("utf-8")

    # حذف calcChain (تعيد Excel بناءه) لتفادي عدم التطابق
    items.pop("xl/calcChain.xml", None)
    ct = items["[Content_Types].xml"].decode("utf-8")
    ct = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", ct)
    items["[Content_Types].xml"] = ct.encode("utf-8")
    rl = items["xl/_rels/workbook.xml.rels"].decode("utf-8")
    rl = re.sub(r'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', "", rl)
    items["xl/_rels/workbook.xml.rels"] = rl.encode("utf-8")

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in items.items():
            z.writestr(name, data)
    print(f"OK -> {dst}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
