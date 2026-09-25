#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Render the generated Method Statement .docx to a faithful HTML preview."""
import html
import os
import re
import sys
import zipfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX = os.path.join(REPO, "MS-KAZ-EHS-MS-2026-004_Night Concrete Pouring (Rev.0).docx")
OUT = os.path.join(REPO, "preview", "index.html")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
import xml.etree.ElementTree as ET


def twips_px(v, dpi=96):
    return round(int(v) / 20.0 * dpi / 72.0, 1)


def run_html(r):
    rpr = r.find(f"{W}rPr")
    style = []
    txt = "".join(t.text or "" for t in r.findall(f"{W}t"))
    if not txt:
        # fields
        if r.find(f"{W}fldChar") is not None or r.find(f"{W}instrText") is not None:
            instr = r.find(f"{W}instrText")
            if instr is not None:
                return f'<span class="fld">[{(instr.text or "").strip()}]</span>'
            return ""
        return ""
    if rpr is not None:
        rf = rpr.find(f"{W}rFonts")
        sz = rpr.find(f"{W}sz")
        col = rpr.find(f"{W}color")
        if sz is not None:
            style.append(f"font-size:{int(sz.get(f'{W}val'))/2.0:.1f}pt")
        if col is not None:
            style.append(f"color:#{col.get(f'{W}val')}")
        if rpr.find(f"{W}b") is not None:
            style.append("font-weight:700")
        if rpr.find(f"{W}i") is not None:
            style.append("font-style:italic")
        if rpr.find(f"{W}u") is not None:
            style.append("text-decoration:underline")
    s = f' style="{";".join(style)}"' if style else ""
    return f"<span{s}>{html.escape(txt)}</span>"


def para_html(p):
    ppr = p.find(f"{W}pPr")
    cls, style = [], []
    if ppr is not None:
        jc = ppr.find(f"{W}jc")
        if jc is not None:
            cls.append({"center": "ctr", "right": "rgt", "both": "just"}.get(
                jc.get(f"{W}val"), ""))
        ind = ppr.find(f"{W}ind")
        if ind is not None:
            l = ind.get(f"{W}left")
            h = ind.get(f"{W}hanging") or ind.get(f"{W}firstLine")
            if l:
                style.append(f"margin-left:{twips_px(l)}px")
            if h and ind.get(f"{W}hanging"):
                style.append(f"text-indent:-{twips_px(h)}px")
            elif h:
                style.append(f"text-indent:{twips_px(h)}px")
        pbdr = ppr.find(f"{W}pBdr")
        if pbdr is not None:
            for edge in pbdr:
                tag = edge.tag.split("}")[1]
                col = edge.get(f"{W}color")
                style.append(f"border-{tag}:2px solid #{col}")
                style.append(f"padding-bottom:3px" if tag == "bottom" else "")
        shd = ppr.find(f"{W}shd")
        if shd is not None:
            style.append(f"background:#{shd.get(f'{W}fill')}")
    style = [s for s in style if s]
    inner = "".join(run_html(r) for r in p.findall(f"{W}r"))
    if not inner.strip():
        return '<div class="sp">&nbsp;</div>'
    a = f' class="{" ".join(c for c in cls if c)}"' if cls else ""
    b = f' style="{";".join(style)}"' if style else ""
    return f"<div{a}{b}>{inner}</div>"


def cell_html(tc):
    tcpr = tc.find(f"{W}tcPr")
    style = []
    if tcpr is not None:
        w = tcpr.find(f"{W}tcW")
        if w is not None:
            style.append(f"width:{twips_px(w.get(f'{W}w'))}px")
        shd = tcpr.find(f"{W}shd")
        if shd is not None and shd.get(f"{W}fill") not in (None, "auto"):
            style.append(f"background:#{shd.get(f'{W}fill')}")
    body = "".join(para_html(x) for x in tc.findall(f"{W}p"))
    s = f' style="{";".join(style)}"' if style else ""
    return f"<td{s}>{body}</td>"


def table_html(tbl):
    tblpr = tbl.find(f"{W}tblPr")
    tstyle = []
    jc = None
    if tblpr is not None:
        b = tblpr.find(f"{W}tblBorders")
        if b is not None:
            top = b.find(f"{W}top")
            if top is not None:
                tstyle.append(f"border:1px solid #{top.get(f'{W}color')}")
        j = tblpr.find(f"{W}jc")
        if j is not None:
            jc = j.get(f"{W}val")
    grid = tbl.find(f"{W}tblGrid")
    cols = ""
    if grid is not None:
        for gc in grid.findall(f"{W}gridCol"):
            cols += f'<col style="width:{twips_px(gc.get(f"{W}w"))}px">'
    rows = ""
    for tr in tbl.findall(f"{W}tr"):
        cells = "".join(cell_html(tc) for tc in tr.findall(f"{W}tc"))
        rows += f"<tr>{cells}</tr>"
    wrap = ' style="margin:0 auto"' if jc == "center" else ""
    return (f'<table class="t"{" ".join("")} style="{";".join(tstyle)}" '
            f'data-c="{jc or ""}"><colgroup>{cols}</colgroup>{rows}</table>'
            f'<div class="tsp"></div>')


def section_break(sectpr):
    pgsz = sectpr.find(f"{W}pgSz")
    orient = "landscape" if (pgsz is not None and pgsz.get(f"{W}orient") == "landscape") else "portrait"
    return f'<div class="pagebreak" data-o="{orient}"></div>'


def main():
    z = zipfile.ZipFile(DOCX)
    root = ET.fromstring(z.read("word/document.xml"))
    body = root.find(f"{W}body")
    out = []
    for child in body:
        tag = child.tag.split("}")[1]
        if tag == "p":
            # a section break hides inside pPr/sectPr
            ppr = child.find(f"{W}pPr")
            if ppr is not None and ppr.find(f"{W}sectPr") is not None:
                out.append(section_break(ppr.find(f"{W}sectPr")))
                continue
            out.append(para_html(child))
        elif tag == "tbl":
            out.append(table_html(child))
        elif tag == "sectPr":
            out.append(section_break(child))
    hdr = z.read("word/header1.xml").decode("utf-8") if "word/header1.xml" in z.namelist() else ""
    ftr = z.read("word/footer1.xml").decode("utf-8") if "word/footer1.xml" in z.namelist() else ""

    def strip(x):
        t = re.sub(r"<[^>]+>", " ", x)
        return re.sub(r"\s+", " ", t).strip()

    header_txt, footer_txt = strip(hdr), strip(ftr)

    doc = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>KAZ-EHS-MS-2026-004 — Method Statement: Night Concrete Pouring</title>
<style>
:root{{--teal:#003B4A;--mid:#007E8C;--border:#B7CCCF;}}
*{{box-sizing:border-box}}
body{{margin:0;background:#e8edee;font-family:Calibri,"Segoe UI",Arial,sans-serif;
 color:#222A2E;line-height:1.35;-webkit-font-smoothing:antialiased}}
.sheet{{max-width:1180px;margin:0 auto;padding:18px 14px 60px}}
.paper{{background:#fff;margin:0 auto 22px;padding:34px 40px;
 box-shadow:0 2px 14px rgba(0,40,50,.16);border-radius:2px}}
.paper.land{{max-width:1180px}}
.paper.port{{max-width:830px}}
.hdr{{display:flex;justify-content:space-between;gap:16px;font-size:8pt;
 border-bottom:1px solid var(--border);padding-bottom:4px;margin-bottom:16px;color:#445}}
.hdr b{{color:var(--mid)}}
.ftr{{border-top:1px solid var(--border);margin-top:22px;padding-top:6px;
 font-size:8pt;text-align:center;color:var(--teal)}}
div{{margin:0}}
.sp{{height:5px}}
.tsp{{height:9px}}
.ctr{{text-align:center}}.rgt{{text-align:right}}.just{{text-align:justify}}
table.t{{border-collapse:collapse;table-layout:fixed;width:100%;margin:0}}
table.t[data-c="center"]{{margin:0 auto;width:auto}}
table.t td{{border:1px solid var(--border);padding:4px 6px;vertical-align:top;
 font-size:9.5pt;word-wrap:break-word;overflow-wrap:break-word}}
.pagebreak{{height:26px;background:repeating-linear-gradient(45deg,#dfe6e7,#dfe6e7 8px,#eef2f3 8px,#eef2f3 16px);
 border-radius:3px;margin:26px 0;position:relative}}
.pagebreak::after{{content:"— new section (page break) —";position:absolute;left:50%;top:50%;
 transform:translate(-50%,-50%);font-size:8pt;color:#6b7f83;background:#eef2f3;padding:1px 10px;border-radius:9px}}
.fld{{color:#999;font-size:8pt}}
.bar{{position:sticky;top:0;z-index:9;background:var(--teal);color:#fff;padding:10px 18px;
 font-size:12px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px}}
.bar b{{color:#00B5E2}}
.bar .m{{opacity:.8;font-size:11px}}
.note{{background:#FFF8E1;border-left:4px solid #BF8F00;padding:10px 14px;margin:0 auto 18px;
 font-size:12px;max-width:1140px;border-radius:2px}}
</style></head><body>
<div class="bar"><div><b>KAZ-EHS-MS-2026-004</b> &nbsp;Method Statement — Night Concrete Pouring Operations</div>
<div class="m">HTML preview of the generated .docx · {len(out)} body elements</div></div>
<div class="note">This is a browser preview of <b>MS-KAZ-EHS-MS-2026-004_Night Concrete Pouring (Rev.0).docx</b>.
The Word file is the deliverable — it carries the exact house styling, page numbering fields,
landscape JSA section and print margins. Preview colours/widths are approximate.</div>
<div class="sheet">
"""
    # paginate: split into "paper" divs at page breaks / heuristically
    chunk, is_land = [], False
    n = 0
    for el in out:
        if el.startswith('<div class="pagebreak"'):
            if chunk:
                n += 1
                doc += (f'<div class="paper {"land" if is_land else "port"}">'
                        f'<div class="hdr"><span><b>{header_txt.split(chr(9))[0]}</b></span>'
                        f'<span>{header_txt.split(chr(9))[-1]}</span></div>'
                        + "".join(chunk)
                        + f'<div class="ftr">{footer_txt.replace("PAGE","1").replace("NUMPAGES","~30")}'
                          f' &nbsp;[sheet {n}]</div></div>')
                chunk = []
            is_land = 'data-o="landscape"' in el
            continue
        chunk.append(el)
    if chunk:
        n += 1
        doc += (f'<div class="paper {"land" if is_land else "port"}">'
                f'<div class="hdr"><span><b>{header_txt.split(chr(9))[0]}</b></span>'
                f'<span>{header_txt.split(chr(9))[-1]}</span></div>'
                + "".join(chunk)
                + f'<div class="ftr">{footer_txt} &nbsp;[sheet {n}]</div></div>')
    doc += "</div></body></html>"

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(doc)
    print("WROTE", OUT, len(doc), "bytes,", n, "sheets")


if __name__ == "__main__":
    main()
