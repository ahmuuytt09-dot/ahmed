#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Builder for: KAZ-EHS-MS-2026-004
METHOD STATEMENT - NIGHT CONCRETE POURING OPERATIONS
KAZ Power Plant Upgrade Project - Khur Al-Zubair, Basra, Iraq

Reuses the shell (styles / theme / numbering / settings) of the existing house
document "بيرمت سستم (Final).docx" (KAZ-EHS-SOP-2026-003 PTW System) so that the
new Method Statement matches the project's document style exactly.

Usage:  python3 tools/build_ms_night_concrete_pour.py
Output: MS-KAZ-EHS-MS-2026-004_Night Concrete Pouring (Rev.0).docx
"""

import os
import re
import shutil
import zipfile

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Twips

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
SHELL = os.path.join(REPO, "بيرمت سستم (Final).docx")
OUT = os.path.join(
    REPO, "MS-KAZ-EHS-MS-2026-004_Night Concrete Pouring (Rev.0).docx"
)
TMP = os.path.join(HERE, "_ms_tmp.docx")

# ---------------------------------------------------------------- house style
TEAL_DARK = "003B4A"      # banner background / section heading text
TEAL_MID = "007E8C"       # heading underline, table header fill
TEAL_CYAN = "00B5E2"      # banner subtitle
TEAL_PALE = "CFE7EB"      # banner meta line
LBL_FILL = "DCEFF1"       # label cell fill
ZEBRA = "F7FBFB"          # alternate row fill
BAND = "EAF4F5"           # sub-band / note fill
BORDER = "B7CCCF"         # table border colour
BODY = "222A2E"           # body text colour
WHITE = "FFFFFF"

DOC_ID = "KAZ-EHS-MS-2026-004"
REV = "0"
FOOTER_TXT = (
    f"{DOC_ID}  \u2022  Rev. {REV}  \u2022  Siemens Energy \u2013 KAZ Power Plant "
    "Upgrade  \u2022  Page "
)
HEADER_LEFT = "KAZ Power Plant Upgrade Project \u2013 Method Statement"
HEADER_RIGHT = f"{DOC_ID}  \u2022  Rev. {REV}  \u2022  Night Concrete Pouring"

USABLE_P = 9864           # A4 portrait usable width (twips), as per house doc
USABLE_L = 14398          # A4 landscape usable width (twips)

RATING_COLOUR = {"H": "C00000", "M": "BF8F00", "L": "2E7D32"}
RATING_WORD = {"H": "HIGH", "M": "MEDIUM", "L": "LOW"}


# ------------------------------------------------------------- low level xml
def _el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k.replace("_", ":")), str(v))
    return e


def shade(obj, fill):
    """Apply cell / paragraph shading."""
    pr = obj._tc.get_or_add_tcPr() if hasattr(obj, "_tc") else obj._p.get_or_add_pPr()
    for old in pr.findall(qn("w:shd")):
        pr.remove(old)
    pr.append(_el("w:shd", w_val="clear", w_color="auto", w_fill=fill))


def set_cell_margins(cell, top=70, bottom=70, left=110, right=110):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = _el("w:tcMar")
    for name, val in (("top", top), ("start", left), ("bottom", bottom), ("end", right)):
        mar.append(_el(f"w:{name}", w_w=val, w_type="dxa"))
    tcPr.append(mar)


def table_borders(table, colour=BORDER, sz=4):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    b = _el("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b.append(_el(f"w:{edge}", w_val="single", w_sz=sz, w_space=0, w_color=colour))
    tblPr.append(b)


def fixed_layout(table):
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblLayout")):
        tblPr.remove(old)
    tblPr.append(_el("w:tblLayout", w_type="fixed"))


def row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el("w:cantSplit"))


def header_row_repeat(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(_el("w:tblHeader"))


def set_grid_exact(table, widths, align=None):
    """Force tblGrid / tblW to exactly the given twip widths and align the table."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    for old in tbl.findall(qn("w:tblGrid")):
        tbl.remove(old)
    grid_el = _el("w:tblGrid")
    for w in widths:
        grid_el.append(_el("w:gridCol", w_w=int(w)))
    tblPr.addnext(grid_el)
    for old in tblPr.findall(qn("w:tblW")):
        tblPr.remove(old)
    tblPr.append(_el("w:tblW", w_w=int(sum(widths)), w_type="dxa"))
    if align:
        for old in tblPr.findall(qn("w:jc")):
            tblPr.remove(old)
        tblPr.append(_el("w:jc", w_val=align))


def para_border(p, colour=TEAL_MID, sz=12, edge="bottom"):
    pPr = p._p.get_or_add_pPr()
    pbdr = _el("w:pBdr")
    pbdr.append(_el(f"w:{edge}", w_val="single", w_sz=sz, w_space=3, w_color=colour))
    pPr.append(pbdr)


def add_field(paragraph, instr):
    r = paragraph.add_run()
    fc = _el("w:fldChar", w_fldCharType="begin")
    r._r.append(fc)
    r2 = paragraph.add_run()
    it = _el("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r2._r.append(it)
    r3 = paragraph.add_run()
    r3._r.append(_el("w:fldChar", w_fldCharType="separate"))
    r4 = paragraph.add_run("1")
    r5 = paragraph.add_run()
    r5._r.append(_el("w:fldChar", w_fldCharType="end"))
    return [r, r2, r3, r4, r5]


# ------------------------------------------------------------- text helpers
def run(p, text, size=10.5, bold=False, italic=False, colour=BODY, font="Calibri",
        underline=False, caps=False):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.bold = bold
    r.italic = italic
    r.underline = underline
    if caps:
        r.font.all_caps = True
    rpr = r._r.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _el("w:rFonts")
        rpr.insert(0, rf)
    for attr in ("ascii", "hAnsi", "cs"):
        rf.set(qn("w:" + attr), font)
    if colour:
        c = _el("w:color", w_val=colour)
        rpr.append(c)
    return r


def para(doc, space_after=90, space_before=0, align=None, keep_next=False,
         left=0, first_line=0, line=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_after = Pt(space_after / 20.0)
    pf.space_before = Pt(space_before / 20.0)
    if line:
        pf.line_spacing = line
    if left:
        pf.left_indent = Twips(left)
    if first_line:
        pf.first_line_indent = Twips(first_line)
    if align == "c":
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "r":
        pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "j":
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if keep_next:
        pf.keep_with_next = True
    return p


BULLET_CHARS = ("\u2022", "\u25aa", "\u25e6", "\u2013")   # • ▪ ◦ –
HANG = 200                                              # hanging indent (twips)


def cell_text(cell, text, size=10, bold=False, colour=BODY, align=None,
              space_after=20, font="Calibri", valign=None):
    """Write (possibly multi-line) text into a cell, replacing its paragraph.

    Consistency rules applied to every cell in the document:
      * bullet lines (• ▪) get an identical hanging indent so wrapped text lines up;
      * multi-line cells get slightly more leading than single-line cells;
      * vertical alignment defaults to TOP, or CENTER for header / short
        centred columns so baselines line up across a row.
    """
    lines = str(text).split("\n")
    multi = len(lines) > 1
    gap = max(space_after, 34) if multi else space_after
    cell.text = ""
    first = True
    for line in lines:
        p = cell.paragraphs[0] if first else cell.add_paragraph()
        first = False
        pf = p.paragraph_format
        pf.space_after = Pt(gap / 20.0)
        pf.space_before = Pt(0)
        if align == "c":
            pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "r":
            pf.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align == "j":
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if line[:1] in BULLET_CHARS:
            pf.left_indent = Twips(HANG)
            pf.first_line_indent = Twips(-HANG)
        run(p, line, size=size, bold=bold, colour=colour, font=font)
    cell.vertical_alignment = (WD_ALIGN_VERTICAL.CENTER if valign == "c"
                              else WD_ALIGN_VERTICAL.TOP)


def heading(doc, text, size=12.5, colour=TEAL_DARK, space_before=260,
            space_after=110, rule=True, rule_colour=TEAL_MID):
    p = para(doc, space_before=space_before, space_after=space_after, keep_next=True)
    run(p, text, size=size, bold=True, colour=colour)
    if rule:
        para_border(p, colour=rule_colour)
    return p


def subheading(doc, text, colour=TEAL_MID, size=11, space_before=150, fill=None):
    p = para(doc, space_before=space_before, space_after=70, keep_next=True)
    if fill:
        shade(p, fill)
    run(p, text, size=size, bold=True, colour=colour)
    return p


def bullets(doc, items, size=10.5, left=240, colour=BODY, space_after=50,
            marker="\u25aa", bold_lead=True):
    """items: list of str, or (lead, rest) tuples -> 'lead: rest'."""
    out = []
    for it in items:
        p = para(doc, space_after=space_after, left=left, first_line=-240, align="j")
        run(p, marker + "  ", size=size, colour=TEAL_MID, bold=True)
        if isinstance(it, tuple):
            lead, rest = it
            run(p, lead, size=size, bold=bold_lead, colour=TEAL_DARK)
            if rest:
                run(p, (" \u2014 " if not rest.startswith((":", ";", ",")) else "") + rest,
                    size=size, colour=colour)
        else:
            run(p, it, size=size, colour=colour)
        out.append(p)
    return out


def body(doc, text, size=10.5, align="j", space_after=100, bold=False, colour=BODY):
    p = para(doc, space_after=space_after, align=align)
    run(p, text, size=size, bold=bold, colour=colour)
    return p


def note_box(doc, title, text, fill=BAND, title_colour=TEAL_DARK):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    fixed_layout(t)
    table_borders(t, colour=TEAL_MID, sz=4)
    set_grid_exact(t, [USABLE_P])
    c = t.cell(0, 0)
    c.width = Twips(USABLE_P)
    shade(c, fill)
    set_cell_margins(c, top=110, bottom=110, left=150, right=150)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    run(p, title, size=10.5, bold=True, colour=title_colour)
    p2 = c.add_paragraph()
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p2, text, size=10, colour=BODY)
    spacer(doc, 60)
    return t


def spacer(doc, after=90):
    p = para(doc, space_after=after)
    for r in p.runs:
        r.font.size = Pt(2)
    return p


def grid(doc, widths, rows, header=True, header_fill=TEAL_MID, zebra=True,
         size=10, header_size=10, aligns=None, header_colour=WHITE,
         total=None, cell_fills=None, bolds=None, align=None):
    """Generic table builder.
    widths  : list of twips
    rows    : list of list of str (first row = header when header=True)
    aligns  : list per column of None/'c'/'r'
    cell_fills: dict {(r,c): fill}
    bolds   : set/list of column indexes rendered bold
    align   : 'center' to centre the whole table on the page
    """
    total = total or sum(widths)
    # scale widths so they sum exactly to `total`
    scale = total / float(sum(widths))
    widths = [int(round(w * scale)) for w in widths]
    widths[-1] += total - sum(widths)
    t = doc.add_table(rows=len(rows), cols=len(widths))
    t.autofit = False
    fixed_layout(t)
    table_borders(t)
    set_grid_exact(t, widths, align=align)
    bolds = set(bolds or [])
    for ri, rowvals in enumerate(rows):
        row = t.rows[ri]
        row_cant_split(row)
        if header and ri == 0:
            header_row_repeat(row)
        for ci, val in enumerate(rowvals):
            cell = t.cell(ri, ci)
            cell.width = Twips(widths[ci])
            set_cell_margins(cell)
            is_head = header and ri == 0
            fill = None
            if is_head:
                fill = header_fill
            elif cell_fills and (ri, ci) in cell_fills:
                fill = cell_fills[(ri, ci)]
            elif zebra and (ri % 2 == 0):
                fill = ZEBRA
            if fill:
                shade(cell, fill)
            al = "c" if is_head else (aligns[ci] if aligns else None)
            # vertical centring: always for header rows, and for any body column
            # that is horizontally centred (numbers, ticks, initials, ratings)
            va = "c" if (is_head or al == "c") else None
            cell_text(
                cell, val,
                size=header_size if is_head else size,
                bold=is_head or (ci in bolds),
                colour=header_colour if is_head else BODY,
                align=al,
                valign=va,
            )
    spacer(doc, 80)
    return t


def label_table(doc, pairs, lw=2834, vw=7030, size=10):
    rows = [[k, v] for k, v in pairs]
    t = grid(doc, [lw, vw], rows, header=False, zebra=False, size=size,
             total=USABLE_P, bolds=[0])
    for ri in range(len(rows)):
        shade(t.cell(ri, 0), LBL_FILL)
        if ri % 2 == 1:
            shade(t.cell(ri, 1), ZEBRA)
    return t


def numbered(doc, items, size=10.5, left=360, space_after=60, bold_lead=True):
    """items: list of (title, text) or plain str -> manual decimal numbering."""
    for i, it in enumerate(items, start=1):
        p = para(doc, space_after=space_after, left=left, first_line=-360, align="j")
        run(p, f"{i}.\t", size=size, bold=True, colour=TEAL_MID)
        if isinstance(it, tuple):
            lead, rest = it
            if lead:
                run(p, lead + (" " if rest else ""), size=size, bold=bold_lead,
                    colour=TEAL_DARK)
            if rest:
                run(p, rest, size=size, colour=BODY)
        else:
            run(p, it, size=size, colour=BODY)


def rating_cell(table, r, c, rating, score, size=10):
    cell = table.cell(r, c)
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    run(p, RATING_WORD[rating], size=size, bold=True, colour=RATING_COLOUR[rating])
    if score:
        p2 = cell.add_paragraph()
        p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(0)
        run(p2, score, size=size - 1.5, colour=BODY)


# --------------------------------------------------------------- page setup
def setup_section(sec, landscape=False):
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Twips(16838), Twips(11906)
        sec.left_margin = sec.right_margin = Twips(1220)
        sec.top_margin = sec.bottom_margin = Twips(900)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Twips(11906), Twips(16838)
        sec.left_margin = sec.right_margin = Twips(1020)
        sec.top_margin = sec.bottom_margin = Twips(964)
    sec.header_distance = Twips(500)
    sec.footer_distance = Twips(500)
    return sec


def build_header(sec):
    sec.header.is_linked_to_previous = False
    hp = sec.header.paragraphs[0]
    hp.text = ""
    pf = hp.paragraph_format
    pf.space_after = Pt(3)
    pf.tab_stops.add_tab_stop(Twips(USABLE_P if not sec.orientation == WD_ORIENT.LANDSCAPE
                                    else USABLE_L))
    from docx.enum.text import WD_TAB_ALIGNMENT
    pf.tab_stops.clear_all()
    pf.tab_stops.add_tab_stop(
        Twips(USABLE_L if sec.orientation == WD_ORIENT.LANDSCAPE else USABLE_P),
        WD_TAB_ALIGNMENT.RIGHT,
    )
    run(hp, HEADER_LEFT, size=8, colour=TEAL_MID, bold=True)
    run(hp, "\t" + HEADER_RIGHT, size=8, colour=BODY)
    para_border(hp, colour=BORDER, sz=4)


def build_footer(sec):
    sec.footer.is_linked_to_previous = False
    fp = sec.footer.paragraphs[0]
    fp.text = ""
    fp.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(3)
    r = run(fp, FOOTER_TXT, size=8, colour=TEAL_DARK)
    for rr in add_field(fp, "PAGE"):
        rr.font.size = Pt(8)
        rr.font.name = "Calibri"
        rr.font.bold = True
        rpr = rr._r.get_or_add_rPr()
        rpr.append(_el("w:color", w_val=TEAL_MID))
    run(fp, " of ", size=8, colour=TEAL_DARK)
    for rr in add_field(fp, "NUMPAGES"):
        rr.font.size = Pt(8)
        rr.font.name = "Calibri"
        rr.font.bold = True
        rpr = rr._r.get_or_add_rPr()
        rpr.append(_el("w:color", w_val=TEAL_MID))


# ===================================================================== CONTENT
def banner(doc):
    t = doc.add_table(rows=1, cols=1)
    t.autofit = False
    fixed_layout(t)
    set_grid_exact(t, [USABLE_P])
    c = t.cell(0, 0)
    c.width = Twips(USABLE_P)
    shade(c, TEAL_DARK)
    set_cell_margins(c, top=260, bottom=260, left=280, right=280)
    c.text = ""
    p = c.paragraphs[0]
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(2)
    run(p, "METHOD STATEMENT", size=19, bold=True, colour=WHITE)
    p2 = c.add_paragraph()
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(8)
    run(p2, "NIGHT CONCRETE POURING OPERATIONS", size=13, bold=True, colour=TEAL_CYAN)
    p3 = c.add_paragraph()
    p3.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(4)
    run(p3, "KAZ Power Plant Upgrade Project \u2014 Khur Al-Zubair Power Station, "
            "Basra, Iraq", size=11, colour=WHITE)
    p4 = c.add_paragraph()
    p4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_after = Pt(0)
    run(p4, f"Document ID: {DOC_ID}   \u2022   Revision: {REV}   \u2022   Status: "
            "DRAFT FOR REVIEW (Pending Final Approval)", size=9, colour=TEAL_PALE)
    spacer(doc, 60)


def section_0(doc):
    heading(doc, "SECTION 0: DOCUMENT CONTROL", space_before=120)
    label_table(doc, [
        ("Project Title & Location",
         "KAZ Power Plant Upgrade Project (SE GT SU 77342-150686) \u2014 Khur Al-Zubair "
         "Power Station, Basra, Iraq"),
        ("Client / Beneficiary",
         "Ministry of Electricity (MOE) & Basrah Gas Company (BGC)"),
        ("Principal Contractor", "Siemens Energy"),
        ("Civil / Concrete Subcontractor",
         "____________________________ (night-shift concrete crew, pump and mixer supplier)"),
        ("Document Title",
         "Method Statement \u2014 Night Concrete Pouring Operations (including mandatory "
         "night-shift EHS requirements and JSA risk matrix)"),
        ("Document No. / Revision", f"{DOC_ID}  /  Rev. {REV}"),
        ("Document Status",
         "DRAFT FOR REVIEW \u2014 to be re-issued as \u201cApproved for Construction\u201d only after "
         "sign-off in Section 14"),
        ("Date of Issue", "____ / ____ / 2026"),
        ("Document Author", "Ahmed Mansoor \u2014 EHS Manager (Siemens Energy)"),
        ("Night Shift Working Window",
         "18:00 \u2013 06:00 hrs (or as stated on the Permit to Work). All clauses of this "
         "Method Statement apply whenever concrete placement, pumping, vibration or "
         "finishing is performed during darkness or reduced natural visibility."),
        ("Activity Description",
         "Delivery of ready-mix concrete by transit mixer, placement by mobile concrete "
         "pump truck with placing boom, compaction by immersion vibrators, screeding / "
         "finishing and initial curing \u2014 executed at night within an operating power "
         "station environment adjacent to energized equipment and overhead conductors."),
        ("Companion Documents",
         "\u25aa KAZ-EHS-SOP-2026-003 \u2014 Permit to Work (PTW) System, Rev. 1\n"
         "\u25aa KAZ-EHS-JHA-2026-005 \u2014 Job Hazard Assessment / JSA (Siemens 5-Step JHA "
         "format) for night concrete pouring\n"
         "\u25aa Risk Assessment \u2014 KAZ PP Upgrade Project (Siemens)\n"
         "\u25aa Instruction \u2014 EHS General Requirements for Contractors\n"
         "\u25aa KAZ Project EHS Plan \u2014 Waste / Environmental Management Plan \u2014 Security Plan\n"
         "\u25aa 000480 Mobilization and Additional Controls Checklist \u2014 Near-Miss Report Form "
         "(000480) \u2014 Alarm Card"),
        ("Governing Standards",
         "Iraqi Labour Law No. 37 of 2015 \u2022 OSHA 29 CFR 1926 Subpart Q (Concrete & "
         "Masonry), Subpart P (Excavations), Subpart K (Electrical), Subpart M (Fall "
         "Protection), Subpart D (Occupational Health & Environmental Controls \u2014 "
         "illumination), Subpart N (Cranes, Derricks & Hoisting \u2014 1926.1408 clearance "
         "distances) \u2022 OSHA 29 CFR 1910 \u2022 NFPA 70E \u2022 IEC 60364 / IEC 61008 (30 mA RCD) "
         "\u2022 IEEE Std 80 (earthing \u2014 step & touch potential) \u2022 ANSI/ISEA 107 (Class 3 "
         "high-visibility apparel) \u2022 EN / EN ISO PPE standards \u2022 Siemens Energy EHS "
         "Policy, EHS Principles & Core Responsibilities and client (MOE/BGC) site rules"),
        ("Technical-Limit References",
         "\u25aa Illumination: minimum 100 lux at the concrete discharge / pouring point and "
         "minimum 50 lux on access pathways and transit routes (project requirement, "
         "verified by calibrated light meter; where a referenced standard requires more, "
         "the more stringent value applies).\n"
         "\u25aa Electrical protection: 30 mA residual current devices (RCD) on all temporary "
         "outlets feeding vibrators, distribution boxes and temporary lighting "
         "(OSHA 1926.404(b)(1) / IEC 61008).\n"
         "\u25aa Fatigue: 100 % crew replacement between day and night shifts with a minimum "
         "of 8 hours rest before the night shift starts.\n"
         "\u25aa Boom clearance: minimum safe approach distance from overhead lines and "
         "energized equipment per OSHA 1926.1408 Table A / NFPA 70E approach boundaries, "
         "or the greater distance stated on the PTW / EKAD \u2014 whichever is more stringent.\n"
         "\u25aa Wind: concrete pump boom operation suspended at sustained wind speeds above "
         "32 km/h (or the pump manufacturer's lower limit).\n"
         "\u25aa Where any limit in this document differs from Siemens Energy standards, "
         "Iraqi law or client requirements, the more stringent requirement applies."),
        ("Distribution",
         "Project Manager \u2022 Site / Construction Manager \u2022 Night EHS Officer \u2022 Permit "
         "Issuer \u2022 Area Authority \u2022 Civil Subcontractor Supervisor \u2022 Concrete Pump "
         "Operator \u2022 Site Medic \u2022 Quality \u2022 Client (MOE/BGC) EHS on request"),
        ("Review / Revision Trigger",
         "This Method Statement shall be reviewed and re-validated before every night "
         "pour campaign, and immediately revised upon: any change of pour location, "
         "plant, lighting layout or energized-network configuration; any incident, "
         "near miss or Stop-Work event; any change of client or Siemens Energy EHS "
         "requirement; or at intervals not exceeding 12 months."),
        ("Handling & Records",
         "Controlled document \u2014 issued from the project EHS document register. Signed "
         "copies of this Method Statement, the Night Pour Readiness Checklist (Section "
         "11), the Shift Handover Log (Annex A), the Toolbox Talk record (Annex B) and "
         "the illumination / RCD / earthing verification log (Annex C) shall be retained "
         "for the project retention period and be available on site for inspection."),
    ])


def section_1(doc):
    heading(doc, "SECTION 1: PURPOSE, SCOPE & APPLICABILITY")
    subheading(doc, "1.1  Purpose", space_before=60)
    body(doc,
         "This Method Statement defines the mandatory sequence of work, resources, "
         "controls and responsibilities for carrying out concrete pouring operations "
         "during the night shift on the KAZ Power Plant Upgrade Project. Night work "
         "materially increases risk because of reduced visibility, reduced supervision "
         "density, operator fatigue, and the presence of energized plant and overhead "
         "conductors that are difficult to judge in the dark. The purpose of this "
         "document is therefore to eliminate those night-specific failure modes by "
         "making a defined set of controls a pre-condition for starting the pour \u2014 not "
         "an optional improvement.")
    subheading(doc, "1.2  Scope of Application")
    bullets(doc, [
        ("Applies to", "all night-time concrete pouring, pumping, placing, vibrating, "
         "screeding, finishing and early curing activities within the KAZ Power Plant "
         "Upgrade Project site boundaries, including temporary lighting, temporary "
         "power distribution, concrete pump set-up and transit mixer movements "
         "associated with those activities."),
        ("Applies to", "all Siemens Energy supervision, civil / concrete "
         "subcontractors, ready-mix suppliers, pump and mixer operators, drivers, "
         "security, medical and third-party inspection personnel involved in the work."),
        ("Also applies", "to any pour that starts in daylight and continues into "
         "darkness, from the moment natural light falls below the illumination "
         "limits stated in Section 9.6."),
        ("Excluded", "daylight-only pours (governed by the base civil Method "
         "Statement \u2014 the EHS controls of Section 9 remain applicable where the "
         "hazard exists), marine / underwater concreting, mass concrete pours "
         "requiring a dedicated thermal-control plan, post-tensioning and grouting "
         "operations."),
    ])
    subheading(doc, "1.3  Mandatory Hold Points \u2014 No Pour Without")
    note_box(doc, "STOP CONDITIONS (absolute \u2014 no exceptions, no verbal waivers)",
             "The night concrete pour shall NOT commence, and shall be suspended "
             "immediately if already running, unless ALL of the following are "
             "physically in place and verified by signature on the Night Pour "
             "Readiness Checklist (Section 11):\n"
             "1) A valid Permit to Work for the night shift, endorsed for that shift;\n"
             "2) A fully rested, dedicated Night EHS Officer present on site;\n"
             "3) The signed Day \u2192 Night Shift Handover Log (Annex A);\n"
             "4) The night Toolbox Talk delivered and recorded (Annex B);\n"
             "5) Standby ambulance with dedicated driver AND certified Site Medic on "
             "site at the pour zone;\n"
             "6) Valid third-party certificates for pump / boom and pump operator "
             "competency card physically on site;\n"
             "7) 100 lux / 50 lux illumination measured and recorded, with backup "
             "emergency lighting staged;\n"
             "8) Pump truck body grounding connected to the station earth grid and "
             "verified; all temporary supplies protected by 30 mA RCD (test-button "
             "proved);\n"
             "9) Hard barricading complete, all exposed rebar ends capped, wheel "
             "stoppers installed, machinery spotter and banksman appointed;\n"
             "10) Emergency egress route to Al-Zubair General Hospital clear of all "
             "plant and obstructions.\n"
             "Any person \u2014 including the Night EHS Officer, Spotter, Medic or any "
             "worker \u2014 is authorised and required to exercise STOP WORK AUTHORITY if "
             "any item above is not satisfied.")


def section_2(doc):
    heading(doc, "SECTION 2: DEFINITIONS & ABBREVIATIONS")
    rows = [["Term", "Definition / Meaning as used in this Method Statement"]]
    terms = [
        ("Night Shift", "The working period between 18:00 and 06:00 hrs, or any period "
                        "of reduced natural visibility in which concrete is placed."),
        ("Dedicated Night EHS Officer", "A fully rested EHS officer assigned exclusively "
                                        "to the night shift; never a day-shift officer continuing on."),
        ("100 % Crew Rotation", "Complete replacement of the day-shift workforce by a "
                                "separate night-shift crew; no individual works both shifts."),
        ("TBT", "Toolbox Talk \u2014 pre-shift safety briefing delivered to the night crew."),
        ("PTW", "Permit to Work, per KAZ-EHS-SOP-2026-003."),
        ("JSA / JHA", "Job Safety Analysis / Job Hazard Assessment \u2014 the risk assessment "
                      "in Section 10 and companion document KAZ-EHS-JHA-2026-005."),
        ("TPI Certificate", "Third-Party Inspection certificate issued by an accredited "
                            "inspection body for plant and equipment."),
        ("Competency Card", "Third-party operator competency card for the concrete pump "
                            "operator, issued by an accredited body."),
        ("AA / AEP / PI / PH", "Area Authority / Authorized Electrical Person / Permit "
                               "Issuer / Permit Holder (roles defined in the PTW System)."),
        ("Banksman / Traffic Marshal", "Person directing vehicle movements; equipped with "
                                       "an illuminated baton at night."),
        ("Machinery Spotter", "Person whose sole duty is to monitor boom and mixer "
                              "movements against clearance limits."),
        ("Rebar Mushroom Cap", "Approved plastic / composite protective cap fitted over an "
                               "exposed reinforcing bar end to prevent impalement."),
        ("Hard Barricade", "Rigid physical barrier (handrail, fence, scaffold tube barrier "
                           "or proprietary panel) \u2014 not warning tape."),
        ("RCD / GFCI", "Residual Current Device / Ground Fault Circuit Interrupter, "
                       "30 mA rated, protecting temporary electrical supplies."),
        ("Step & Touch Potential", "Voltage differences a person can be exposed to between "
                                   "feet, or between hand and feet, during an earth fault "
                                   "(IEEE Std 80)."),
        ("Wheel Stopper", "Heavy-duty physical barrier preventing a vehicle from reversing "
                          "beyond a defined point."),
        ("Lux", "Unit of illuminance (lumen per square metre) measured at the task or "
                "walking surface with a calibrated light meter."),
        ("Lux Survey", "Documented measurement of illumination levels at defined points, "
                       "recorded on Annex C."),
        ("Standby Ambulance", "Site ambulance fully equipped and manned, stationed at the "
                              "pour zone for the whole night shift."),
        ("MOTAT", "Manifest / record system for hazardous waste movement on the project."),
        ("Cold Joint", "A discontinuity formed when fresh concrete is placed against "
                       "concrete that has already set \u2014 a quality defect created by "
                       "unplanned night interruptions."),
        ("EKAD", "Electrical work / energized-area clearance documentation required for "
                 "high-voltage systems (\u2265 1000 V)."),
    ]
    rows.extend([[a, b] for a, b in terms])
    grid(doc, [2600, 7264], rows, size=9.5, header_size=9.5, bolds=[0])


def section_3(doc):
    heading(doc, "SECTION 3: ROLES & RESPONSIBILITIES MATRIX")
    body(doc,
         "Night work requires named, present individuals \u2014 not role titles held by "
         "people who have gone home. The table below shall be completed with names "
         "before every night pour and attached to the Permit to Work.")
    rows = [["Role", "Night-Pour Responsibilities & Authority", "Name / Contact"]]
    roles = [
        ("Project Manager",
         "Owns overall delivery and EHS performance. Approves the requirement for night "
         "working, confirms the dedicated night-shift resources (crew, EHS officer, "
         "ambulance, medic, lighting) are funded and mobilized, and is notified of any "
         "Stop-Work event."),
        ("Site Manager",
         "Overall site authority for the project. Confirms that the night pour is "
         "resourced and that this Method Statement is implemented; signs the document "
         "authorisation (Section 14.2) and is notified of every Stop-Work event, "
         "incident and aborted pour."),
        ("Site / Construction Manager (Night Shift In-Charge)",
         "Single point of operational command for the night shift. Confirms the shift "
         "handover, authorises the pour sequence, appoints the Spotter and Banksman, "
         "controls the number of mixers on site, and holds authority to stop or abort "
         "the pour."),
        ("Night EHS Officer (dedicated, fully rested)",
         "Present on site for the entire night shift. Delivers/verifies the TBT, verifies "
         "all Section 9 controls, performs and records the lux survey, RCD test and "
         "grounding verification, audits PPE and barricading, monitors fatigue, leads the "
         "emergency response, and holds unconditional Stop-Work Authority. Shall not be a "
         "day-shift EHS officer continuing into the night."),
        ("Day-Shift EHS Officer / Supervisor",
         "Completes the written Shift Handover Log with the night team before leaving "
         "site: completed works, active live zones, open isolations, PTW status, lighting "
         "and barricade condition, residual hazards. Shall not remain as night EHS cover."),
        ("Permit Issuer (PI) / Area Authority (AA)",
         "Jointly verify site conditions at night before issuing/endorsing the permit for "
         "the shift, confirm live-zone boundaries and safe approach distances, and "
         "co-sign the readiness checklist. Technical approval authority rests with the "
         "Area Authority."),
        ("Civil Supervisor / Permit Holder (PH)",
         "Implements this Method Statement on the ground, briefs the crew, maintains "
         "housekeeping and lighting, controls the pour sequence and joint locations, "
         "verifies rebar caps and wheel stoppers, and closes the permit at shift end."),
        ("Concrete Pump Operator (third-party certified)",
         "Holds a valid third-party competency card. Sets up on approved ground with "
         "outriggers fully extended on pads, verifies body grounding is connected before "
         "operating, keeps the boom inside permitted clearances, obeys the Spotter and "
         "stops immediately on any signal or alarm."),
        ("Machinery Spotter (dedicated)",
         "Sole duty: continuously monitor boom and mixer movements against clearance "
         "limits and the exclusion zone. Equipped with high-intensity torch, whistle and "
         "radio. Has authority to stop all movement instantly. Performs no other task."),
        ("Banksman / Traffic Marshal",
         "Directs every mixer reversing movement using an illuminated baton; maintains "
         "one-way routing and the clear egress lane; no vehicle reverses without him."),
        ("Transit Mixer Drivers",
         "Hold valid heavy vehicle driving licences and third-party safety induction; "
         "pass a pre-shift fitness check; follow site speed limit and routing; remain in "
         "the cab or in the designated safe zone during discharge; never enter the "
         "barricaded pour area."),
        ("Certified Site Medic / First Aider",
         "On duty on site for the whole night shift, trained in CPR, trauma response and "
         "high-voltage electric-shock response. Maintains the illuminated first-aid kit "
         "and the insulated rescue hook at the pour site; records standby times."),
        ("Standby Ambulance Driver",
         "Remains with the ambulance at the pour zone, engine ready, emergency lights "
         "operational, egress lane kept clear; knows the route to Al-Zubair General "
         "Hospital."),
        ("Authorized Electrical Person / Lighting Technician",
         "Installs, tests and records temporary distribution, 30 mA RCD protection, cable "
         "elevation, light-tower aiming and the pump body-grounding connection to the "
         "station earth grid; carries out repairs; no unauthorised person touches "
         "electrical equipment."),
        ("Night Security / Access Control",
         "Enforces PTW-based access control at the perimeter, logs all persons and "
         "vehicles entering the pour zone, prevents unauthorised entry, and controls the "
         "gate for the emergency egress route."),
        ("Quality Engineer / Inspector",
         "Verifies pour approval, concrete delivery tickets, slump/temperature, "
         "compaction, finish and curing; confirms joint locations to avoid cold joints "
         "caused by night interruptions."),
        ("All Personnel",
         "Comply with this Method Statement, wear Class 3 high-visibility attire and task "
         "PPE, report hazards and near misses immediately, and exercise STOP WORK "
         "AUTHORITY without fear of reprisal."),
    ]
    for a, b in roles:
        rows.append([a, b, ""])
    grid(doc, [2350, 5664, 1850], rows, size=9.5, header_size=9.5, bolds=[0])


def section_4(doc):
    heading(doc, "SECTION 4: PLANT, EQUIPMENT & CERTIFICATION REQUIREMENTS")
    body(doc,
         "No item below may enter the site or be used during a night pour unless the "
         "corresponding certificate, test record or inspection is physically available "
         "on site and has been verified and signed on the Night Pour Readiness Checklist.")
    rows = [["#", "Plant / Equipment", "Certification & Verification Required Before Night Use", "Verified By"]]
    plant = [
        ("Concrete pump truck & placing boom",
         "Valid third-party inspection (TPI) certificate for the truck and boom before "
         "site entry; current load/hydraulic test records; daily pre-use inspection "
         "(outriggers, pins, hoses, boom sections, emergency stop, alarms, reflective "
         "marking, reverse alarm and lights)."),
        ("Concrete pump operator",
         "Valid third-party competency card issued by an accredited inspection body; "
         "site-specific induction; medical fitness; record retained on site."),
        ("Transit mixer trucks",
         "Valid heavy vehicle driving licence per driver; third-party safety induction "
         "certificate; pre-shift fitness check; vehicle inspection (brakes, tyres, "
         "lights, reverse alarm, mirrors, chute condition); site speed limit briefing."),
        ("Immersion / poker vibrators & power tools",
         "110 V or RCD-protected supply; portable appliance test tag in date; cable and "
         "connector integrity check; 30 mA RCD proved by test button before each shift."),
        ("Temporary distribution boards & cables",
         "IP-rated enclosures suitable for wet conditions; 30 mA RCD protection on all "
         "outlets; earthing verified; cables elevated off wet ground and fresh concrete "
         "on insulated cable hangers; no joints in wet areas."),
        ("Light towers & portable floodlights",
         "Sufficient units to achieve \u2265 100 lux at the discharge point and \u2265 50 lux on "
         "access/transit routes; aiming set to prevent glare; fuel/ cables secured; "
         "battery-operated backup floodlights staged at the pour site."),
        ("Calibrated light meter (lux meter)",
         "Valid calibration certificate; used to measure and record the lux survey "
         "(Annex C) before the pour starts and after any lighting change."),
        ("Hard barricades, fencing & warning devices",
         "Rigid handrails / fencing panels with reflective tape and warning flashers "
         "around the pour area, open edges and pump perimeter; sufficient quantity for "
         "the full perimeter \u2014 soft warning tape is not accepted."),
        ("Rebar mushroom caps",
         "Approved plastic / composite caps, correct diameter for the bar size, fitted to "
         "100 % of exposed vertical and horizontal rebar ends in the work zone and on "
         "access pathways."),
        ("Wheel stoppers / barriers",
         "Heavy-duty stoppers positioned to prevent mixers reversing toward live "
         "equipment, grounding conductors, excavations or the pump."),
        ("Earthing / grounding kit",
         "Rated earthing lead, clamps and connections to the station earth grid; "
         "continuity/resistance verified and recorded by the Authorized Electrical Person."),
        ("Insulated electrical rescue hook",
         "Voltage-rated, undamaged, certified insulated rescue hook stationed directly at "
         "the pour site with the illuminated first-aid kit."),
        ("Standby ambulance",
         "Inspected and roadworthy; oxygen cylinder (charged), spinal board, "
         "resuscitation kit, trauma dressings, operational emergency lights and siren; "
         "dedicated standby driver on shift."),
        ("Illuminated first-aid kit & stretchers",
         "Stocked, sealed, illuminated / internally lit kit at the pour site; expiry "
         "dates checked; stretcher and spinal board accessible."),
        ("Communication equipment",
         "Charged radios on the designated night channel for Supervisor, EHS Officer, "
         "Spotter, Banksman, Pump Operator and Medic; whistle and high-intensity torch "
         "per Spotter/Banksman; backup mobile numbers."),
        ("Anemometer",
         "Used to verify wind speed before and during boom operation; suspend above "
         "32 km/h or the manufacturer's lower limit."),
        ("Concrete washout pit / skip & spill kit",
         "Designated washout area away from drains and earth grid; spill kit and "
         "absorbents available for hydraulic oil or fuel leaks."),
    ]
    for i, (item, req) in enumerate(plant, start=1):
        rows.append([str(i), item, req, ""])
    grid(doc, [420, 2280, 5764, 1400], rows, size=9, header_size=9,
         aligns=["c", None, None, "c"], bolds=[1])


def section_5(doc):
    heading(doc, "SECTION 5: NIGHT-SHIFT PERSONNEL, CREW REPLACEMENT & FITNESS")
    subheading(doc, "5.1  Mandatory 100 % Crew Replacement Rule", space_before=60)
    note_box(doc, "FATIGUE CONTROL \u2014 MANDATORY",
             "Day-shift workers and technicians are strictly prohibited from continuing "
             "into the night shift. The contractor shall deploy a dedicated Night Shift "
             "Crew, every member of which has had a minimum of 8 hours of rest before "
             "the shift starts. Day-shift EHS officers shall not cover the night shift: a "
             "fully rested, dedicated Night EHS Officer shall be on site throughout the "
             "operation. The Night Crew Register (Annex D) shall record each individual's "
             "rest start/end time and fitness declaration before the shift.")
    subheading(doc, "5.2  Minimum Night-Shift Manning")
    rows = [["#", "Night-Shift Role", "Min.", "Mandatory Qualification / Condition"]]
    manning = [
        ("Night Shift In-Charge (Site/Construction Supervisor)", "1",
         "Appointed in writing; present for the whole shift"),
        ("Dedicated Night EHS Officer", "1",
         "Fully rested (min. 8 h rest); not a day-shift officer; present for the whole shift"),
        ("Civil / Concrete Supervisor (Permit Holder)", "1",
         "PTW Permit Holder training; night-work experience"),
        ("Concrete Pump Operator", "1",
         "Valid third-party competency card; induction; fitness check"),
        ("Machinery Spotter (dedicated)", "1",
         "Trained spotter; torch, whistle, radio; no other duties"),
        ("Banksman / Traffic Marshal", "1 (min.)",
         "Trained banksman; illuminated baton; high-visibility Class 3"),
        ("Transit Mixer Drivers", "per delivery plan",
         "Valid heavy vehicle driving licence; third-party induction; pre-shift fitness check"),
        ("Concrete Labourers / Vibrator Operators", "per pour size",
         "Night crew only; concrete PPE; TBT attendance"),
        ("Finishing / Screeding Team", "per pour size",
         "Night crew only; TBT attendance"),
        ("Authorized Electrical Person / Lighting Technician", "1",
         "Electrical competency; responsible for RCD, grounding and lighting"),
        ("Certified Site Medic / First Aider", "1",
         "CPR, trauma and high-voltage electric-shock response; on duty on site"),
        ("Standby Ambulance Driver", "1",
         "Dedicated driver; remains with the ambulance at the pour zone"),
        ("Night Security / Access Controller", "per site plan",
         "Enforces PTW access control at the perimeter"),
        ("Quality Engineer / Inspector", "1 (or on call)",
         "Verifies delivery tickets, slump, compaction and finish"),
    ]
    for i, m in enumerate(manning, start=1):
        rows.append([str(i), m[0], m[1], m[2]])
    grid(doc, [420, 3600, 1100, 4744], rows, size=9.5, header_size=9.5,
         aligns=["c", None, "c", None], bolds=[1])
    subheading(doc, "5.3  Pre-Shift Fitness & Alertness Checks")
    bullets(doc, [
        "Every night-shift member completes a pre-shift fitness check (alertness, "
        "injury/illness, medication that causes drowsiness, previous rest hours) "
        "recorded on the Night Crew Register \u2014 Annex D.",
        "Mixer drivers and the pump operator are checked individually before taking "
        "control of any machine; any person showing signs of fatigue, illness or "
        "impairment is stood down and replaced.",
        "No individual may work more than one shift in any 24-hour period; consecutive "
        "night shifts shall be limited in accordance with Iraqi Labour Law No. 37 of "
        "2015 and the project's fatigue-management rules.",
        "Rest breaks are scheduled and enforced (minimum 20-minute break every 4 hours "
        "or as required by the project fatigue standard), taken in a lit, sheltered area "
        "away from plant movement.",
        "Hydration, warm shelter and adequate meals are provided for the night crew; "
        "caffeine use is not a substitute for rest.",
        "The Night EHS Officer performs at least two unannounced alertness rounds during "
        "the shift, including the period of lowest alertness (02:00\u201305:00 hrs).",
    ])


def section_6(doc):
    heading(doc, "SECTION 6: CONSTRUCTION METHODOLOGY \u2014 STEP-BY-STEP WORK SEQUENCE")
    body(doc,
         "The sequence below is mandatory. Each step shall be completed and verified "
         "before the next begins; steps marked \u25b6 are hold points requiring a signature "
         "on the Night Pour Readiness Checklist (Section 11).")

    steps = [
        ("STEP 1 \u2014 PLANNING & PERMIT PREPARATION (T\u201324 h)",
         ["Pour planned, approved and scheduled by the Construction Manager with the "
          "planned volume, pour rate, joint locations and pump/mixer numbers confirmed.",
          "PTW application submitted with this Method Statement, the JSA, third-party "
          "certificates, competency cards, lighting plan and medical standby "
          "confirmation attached (KAZ-EHS-SOP-2026-003, Step 1).",
          "Night-specific resources booked and confirmed in writing: dedicated night "
          "crew, Night EHS Officer, standby ambulance and driver, Site Medic, light "
          "towers, backup floodlights, spotter and banksman.",
          "Live-zone / energized-equipment boundaries confirmed with the Area "
          "Authority; safe approach distances and any restricted red zones marked on "
          "the site plan and physically on the ground.",
          "Lighting layout drawn: tower positions, aiming directions, cable routes and "
          "the points where lux will be measured.",
          "Weather forecast reviewed (wind, rain, dust); boom operation cancelled if "
          "sustained wind is forecast above 32 km/h."]),
        ("STEP 2 \u2014 DAY \u2192 NIGHT SHIFT HANDOVER  \u25b6",
         ["Joint walk of the pour area by the outgoing day-shift supervisor/EHS officer "
          "and the incoming night-shift in-charge/EHS officer.",
          "Written Shift Handover Log (Annex A) completed and signed by both EHS "
          "Supervisors and both Engineering/Civil Supervisors, recording: works "
          "completed, works outstanding, active live zones, open isolations/LOTO, PTW "
          "status and validity, barricade and lighting condition, excavations and open "
          "edges, ground condition, access routes and any incidents or near misses.",
          "PTW status, endorsements and any restrictions transferred and re-verified; "
          "no permit is assumed valid without a physical check.",
          "Day-shift crew leaves site \u2014 no day-shift worker or day-shift EHS officer "
          "continues into the night shift."]),
        ("STEP 3 \u2014 NIGHT CREW MOBILIZATION & FITNESS VERIFICATION  \u25b6",
         ["Night Crew Register (Annex D) completed: names, roles, rest hours (\u2265 8 h), "
          "certificates sighted, fitness declaration signed.",
          "Third-party certificates verified on site: pump/boom TPI certificate, pump "
          "operator competency card, mixer drivers' heavy vehicle licences and "
          "inductions, medic certification, ambulance inspection record.",
          "PPE issued and checked: Class 3 high-visibility attire, steel-toe rubber "
          "safety boots, heavy-duty chemical/rubber gloves, sealed safety goggles, "
          "helmets with chin strap where required, plus task-specific PPE.",
          "Radios issued and tested on the night channel; torches and whistles issued "
          "to the Spotter and Banksman."]),
        ("STEP 4 \u2014 NIGHT TOOLBOX TALK (TBT)  \u25b6",
         ["Night EHS Officer delivers the pre-shift TBT to the entire night crew; "
          "attendance recorded on Annex B with signatures.",
          "Mandatory TBT content: the specific pour scope and sequence; safe approach "
          "distances and the location of live energized zones and restricted red zones; "
          "barricaded areas and the single controlled entry point; mixer routing, "
          "reversing rules and banksman signals; boom movement rules and the Spotter's "
          "stop signal; rebar caps and open-edge protection; electrical hazards, RCD "
          "protection, grounding and the rescue hook; wet-concrete chemical burn "
          "prevention and PPE; illumination limits and what to do if lights fail; "
          "emergency evacuation routes, muster point, alarm signal, ambulance location "
          "and the route to Al-Zubair General Hospital; fatigue, breaks and STOP WORK "
          "AUTHORITY.",
          "Crew questions answered and understanding confirmed before anyone enters the "
          "pour zone."]),
        ("STEP 5 \u2014 AREA PREPARATION, BARRICADING & REBAR PROTECTION  \u25b6",
         ["Pour area, open edges and the concrete pump perimeter fully enclosed with "
          "rigid hard barricades (handrails/fencing) fitted with reflective tape and "
          "warning flashers \u2014 soft warning tape is not accepted.",
          "One controlled entry/exit point established and manned; PTW access control "
          "enforced by security; \u201cCONCRETE POURING IN PROGRESS \u2013 AUTHORISED PERSONNEL "
          "ONLY\u201d signage displayed with reflective/illuminated panels.",
          "100 % of exposed vertical and horizontal rebar ends within the work zone and "
          "on access pathways fitted with approved plastic rebar mushroom caps before "
          "work commences; verified by the EHS Officer and Supervisor.",
          "Heavy-duty wheel stoppers/barriers installed at the mixer discharge positions "
          "to prevent reversing toward live equipment, grounding conductors or "
          "excavations.",
          "Formwork, reinforcement, inserts, embedments and cleanliness inspected and "
          "released by Quality; standing water, debris and laitance removed.",
          "Open excavations, pits and edges within or adjacent to the pour zone "
          "barricaded, lit and signed; access ladders/stairs provided and illuminated.",
          "Housekeeping completed: access pathways and transit routes cleared of "
          "obstructions, cables routed and elevated, materials stacked clear of the "
          "egress lane."]),
        ("STEP 6 \u2014 TEMPORARY POWER, GROUNDING & ILLUMINATION SET-UP AND TESTING  \u25b6",
         ["Temporary distribution installed by the Authorized Electrical Person: IP-rated "
          "boards, 30 mA RCD protection on all outlets feeding vibrators, distribution "
          "boxes and temporary lights, RCD test button operated and result logged.",
          "All cables and extension leads elevated off wet ground and poured concrete "
          "using insulated cable hangers; no cable lies in water, fresh concrete or a "
          "vehicle path; crossings protected by cable ramps.",
          "Concrete pump truck body grounding applied and connected directly to the "
          "station earth grid to mitigate step and touch potential; connection inspected "
          "and the earthing continuity/resistance recorded on Annex C.",
          "Light towers positioned and angled downward to give \u2265 100 lux at the "
          "concrete discharge/pouring point and \u2265 50 lux along access pathways and "
          "transit routes, without blinding mixer or pump operators and without leaving "
          "dark shadows in the working zone.",
          "Lux survey performed with the calibrated light meter at the defined points "
          "and recorded on Annex C; additional towers deployed until the values are met.",
          "Battery-operated portable floodlights staged at the pour site, charged and "
          "tested, ready for immediate use on primary power failure.",
          "Insulated electrical rescue hook and illuminated first-aid kit positioned "
          "directly at the pour site; fire extinguishers positioned at the pump and "
          "board locations."]),
        ("STEP 7 \u2014 CONCRETE PUMP POSITIONING & PRE-POUR READINESS VERIFICATION  \u25b6",
         ["Pump positioned on firm, compacted, level ground clear of excavations, "
          "trenches and underground services; outriggers 100 % extended on timber/steel "
          "pads; ground bearing capacity confirmed.",
          "Boom clearance checked against overhead lines and energized equipment: the "
          "boom shall never breach the restricted red zones and shall remain at or "
          "beyond the minimum safe approach distance (OSHA 1926.1408 Table A / NFPA 70E "
          "or the greater PTW-specified distance) throughout all slewing and extension.",
          "The maximum permitted boom envelope is physically marked (flagged line, "
          "ground marking or laser/measured reference) and briefed to the operator and "
          "Spotter.",
          "Emergency stop functions, alarms, hydraulic lines, hoses and clamps checked; "
          "delivery line secured and supported; end hose handled by two persons with "
          "correct gloves \u2014 never wrapped around the body.",
          "Spotter and Banksman posted with torches, whistles, illuminated batons and "
          "radios; hand/radio stop signals agreed and rehearsed with the operator.",
          "Standby ambulance stationed directly at the pouring area with its dedicated "
          "driver; Site Medic on duty; emergency egress route confirmed clear of all "
          "mixers, machinery and obstructions for immediate transfer to Al-Zubair "
          "General Hospital.",
          "Night Pour Readiness Checklist (Section 11) completed and signed by the "
          "Permit Holder, Night EHS Officer, Permit Issuer and Area Authority. "
          "\u25b6 THE POUR SHALL NOT START UNTIL THIS IS SIGNED."]),
        ("STEP 8 \u2014 CONCRETE DELIVERY, PLACEMENT & COMPACTION",
         ["Mixers follow the approved one-way route at site speed limit; each waits in "
          "the designated holding area \u2014 never queued inside the pour zone or across "
          "the egress lane.",
          "Every reversing movement directed by the Banksman with an illuminated baton; "
          "no driver reverses without direction; wheel stoppers engaged; reverse alarms "
          "and lights functioning.",
          "Delivery tickets checked against the approved mix design, pour location, "
          "batching time, slump and temperature; rejected loads returned \u2014 no water "
          "added on site without Quality approval.",
          "Placement by boom/pipeline in controlled layers per the approved pour "
          "sequence; discharge point kept within the Spotter's line of sight; personnel "
          "never stand under the boom, the end hose or the delivery line.",
          "Compaction with RCD-protected immersion vibrators operated by trained "
          "personnel wearing chemical-resistant gloves and sealed goggles; vibrator "
          "cables kept clear of the concrete, rebar and vehicle paths.",
          "The Spotter continuously monitors boom and mixer movements and calls an "
          "immediate stop on any approach to a clearance limit, an uncontrolled "
          "movement, or loss of visibility.",
          "Continuous pour maintained to prevent cold joints: mixer flow managed against "
          "the pour rate; if a delivery interruption exceeds the concrete's initial set "
          "window, the Construction Manager and Quality agree the construction joint "
          "location before stopping.",
          "Spillage cleaned immediately; washout only in the designated washout pit; no "
          "concrete, washout water or hydraulic oil enters drains, ground or the earth "
          "grid area."]),
        ("STEP 9 \u2014 FINISHING, CURING & PROTECTION",
         ["Screeding, floating and finishing performed under direct task lighting "
          "(\u2265 100 lux at the working face); finishing machines operated only by trained "
          "personnel with the area barricaded.",
          "Fresh concrete protected from traffic, plant, water and debris; edges and "
          "fresh faces barricaded and signed \u2014 \u201cFRESH CONCRETE \u2013 DO NOT WALK\u201d.",
          "Curing applied immediately after finishing per specification (curing "
          "compound, wet hessian/polythene or ponding); curing materials stored and "
          "handled per MSDS with gloves and goggles.",
          "Open edges, penetrations and construction joints left safe and barricaded "
          "with hard protection and lighting maintained until the following day shift."]),
        ("STEP 10 \u2014 HOUSEKEEPING, DEMOBILIZATION & PERMIT CLOSURE",
         ["Pump, boom, delivery lines and vibrators cleaned at the washout area; "
          "residual concrete and washout water contained in the washout pit/skip and "
          "disposed via the approved route (MOTAT records where applicable).",
          "Barricades, light towers, boards and cables recovered or left in a safe, lit "
          "and barricaded condition; temporary lighting left on over fresh concrete and "
          "open edges.",
          "All plant refuelled/parked in the designated area, keys controlled; the "
          "egress lane restored and kept clear.",
          "Waste segregated (concrete debris, packaging, oily absorbents, general "
          "waste) and removed to the approved storage area per the Environmental / "
          "Waste Management Plan.",
          "PTW closed jointly by the Permit Holder and Permit Issuer: site condition, "
          "isolations, remaining hazards and any incidents recorded (PTW System Step 3).",
          "Any incident, near miss, unsafe condition or Stop-Work event reported "
          "immediately and recorded on the Near-Miss Report Form (000480) and in the "
          "Daily EHS Report."]),
        ("STEP 11 \u2014 NIGHT \u2192 DAY SHIFT HANDOVER  \u25b6",
         ["Night team completes the handover log with the incoming day-shift "
          "supervisor/EHS officer: concrete placed (volume, locations, joints), curing "
          "status, live zones still active, isolations, lighting and barricade status, "
          "plant left on site, defects and outstanding works.",
          "Joint walk of the pour area at first light; fresh concrete, edges, openings "
          "and washout areas re-inspected before day-shift activities start nearby.",
          "Any hazard not eliminated during the night is escalated in writing to the "
          "Site Manager and recorded as an open item with an owner and due date."]),
    ]
    for title, items in steps:
        subheading(doc, title, space_before=160)
        bullets(doc, items, size=10, marker="\u2013", left=280)


def section_7(doc):
    heading(doc, "SECTION 7: PERMIT TO WORK (PTW) LINKAGE & PRE-CONDITIONS")
    rows = [["PTW / Interface Requirement", "Application to Night Concrete Pouring"]]
    items = [
        ("Permit category",
         "General Cold Work Permit for the civil concrete activity; additionally an "
         "Electrical Work Permit (and EKAD where the high-voltage system is involved) "
         "for temporary power, lighting installation and the pump body-grounding "
         "connection to the station earth grid; Excavation Permit where the pour is "
         "above/beside an open excavation; Lifting Permit where a crane is used to "
         "handle pump lines, skip or equipment."),
        ("Validity / endorsement",
         "The permit is issued or endorsed for the night shift only and must be "
         "re-evaluated and re-endorsed every shift. A permit issued for day work does "
         "not automatically authorise night work."),
        ("Mandatory attachments",
         "This Method Statement; the JSA / Risk Assessment (Section 10); signed Night "
         "Pour Readiness Checklist; Shift Handover Log; TBT attendance record; TPI "
         "certificates and competency cards; lux survey record; RCD test and earthing "
         "verification record; medical standby confirmation (ambulance + medic)."),
        ("Joint verification",
         "Permit Issuer, Permit Holder, Area Authority and Night EHS Officer jointly "
         "verify the site at night \u2014 physically, with the lighting on \u2014 before "
         "signing. Verification in daylight is not accepted for a night pour."),
        ("PTW / LOTO separation",
         "PTW validity and LOTO validity are managed separately; expiry of the permit "
         "does not release any isolation. Night isolations remain under Isolation "
         "Authority control and are listed in the Shift Handover Log."),
        ("Shift handover of permits",
         "At shift change the incoming Permit Issuer/Supervisor, Permit Holder and "
         "(for energized or isolated works) the Isolation Authority jointly review the "
         "permit, walk the area and sign the handover before work resumes."),
        ("Emergency suspension triggers",
         "Site alarm or emergency; loss of illumination below the required lux; RCD "
         "trip that cannot be cleared safely; failure of pump body grounding; wind above "
         "32 km/h; lightning; unexpected energization or change in live-network "
         "configuration; absence of the Night EHS Officer, Spotter, ambulance or Medic; "
         "any breach of safe approach distance \u2014 in each case work stops immediately "
         "and a Stop-Work notice is raised."),
        ("Access control",
         "Only persons named on the permit and recorded by security may enter the "
         "barricaded pour zone; visitors require escort, TBT and full PPE."),
    ]
    rows.extend(items)
    grid(doc, [2700, 7164], rows, size=9.5, header_size=9.5, bolds=[0])


EHS_SECTIONS = [
    ("9.1", "SHIFT ROTATION & FATIGUE CONTROLS", [
        ("Full Workforce Replacement (100 % Rotation)",
         "Day-shift workers and technicians are strictly prohibited from continuing into "
         "the night shift; the contractor must deploy a dedicated Night Shift Crew with a "
         "minimum of 8 hours of prior rest."),
        ("Dedicated Night EHS Supervision",
         "Day-shift EHS officers must not cover the night shift. A fully rested, dedicated "
         "Night EHS Officer must be on-site throughout the operation."),
        ("Formal Shift Handover Protocol",
         "Mandatory written handover log signed between Day and Night EHS / Engineering "
         "Supervisors to log completed works, active live zones, and PTW updates."),
        ("Night Toolbox Talk (TBT)",
         "The EHS Officer must conduct a pre-shift TBT with the night crew, highlighting "
         "safe approach distances, live energized zones, and emergency evacuation routes."),
        ("Fitness & Alertness Verification",
         "Pre-shift fitness checks and a Night Crew Register recording rest hours for "
         "every individual; alertness rounds by the Night EHS Officer during the "
         "02:00\u201305:00 hrs low-alertness window; enforced rest breaks in a lit, "
         "sheltered area."),
    ]),
    ("9.2", "PHYSICAL BARRICADING, REBAR PROTECTION & CERTIFICATION", [
        ("Hard Barricading & Perimeter Isolation",
         "The entire pouring area, open edges and concrete pump perimeter must be fully "
         "enclosed using rigid hard barricades (handrails / fencing) with reflective tape "
         "and warning flashers, instead of soft warning tape."),
        ("Rebar Impalement Protection",
         "All exposed vertical and horizontal rebar ends within the work zone and access "
         "pathways must be fitted with approved plastic rebar mushroom caps prior to work "
         "commencement."),
        ("Concrete Pump & Boom Certification",
         "The concrete pump truck and placement boom must hold a valid Third-Party "
         "Inspection Certificate before site entry."),
        ("Certified Pump Operator",
         "The concrete pump operator must possess a valid Third-Party Competency Card "
         "issued by an accredited inspection body."),
        ("Transit Mixer Drivers",
         "All concrete mixer drivers must present valid Heavy Vehicle Driving Licenses, "
         "third-party safety induction certifications, and undergo pre-shift fitness "
         "checks."),
        ("Controlled Access",
         "A single manned entry/exit point with PTW-based access control, illuminated "
         "signage and a visitor escort rule; no person enters the barricaded pour zone "
         "without TBT and full PPE."),
    ]),
    ("9.3", "MEDICAL PREPAREDNESS & STANDBY AMBULANCE", [
        ("On-Site Standby Ambulance",
         "A fully functional, inspected site ambulance with a dedicated standby driver "
         "must be stationed directly at the concrete pouring area throughout the entire "
         "night shift."),
        ("Ambulance Readiness",
         "The ambulance must be equipped with an oxygen cylinder, spinal board, "
         "resuscitation kit, and operational emergency lights."),
        ("Certified Site Medic",
         "A certified Site Medic / First Aider must be on duty on-site, trained in CPR, "
         "trauma response, and high-voltage electrical shock response."),
        ("Clear Emergency Egress",
         "Keep the emergency evacuation route clear of all mixer trucks, machinery and "
         "obstructions for immediate transfer to Al-Zubair General Hospital if required."),
        ("Rescue & Casualty Handling Readiness",
         "Stretcher and spinal board accessible at the pour site; casualty handling plan "
         "briefed at the TBT (including \u201cdo not remove an impaling object\u201d); hospital "
         "pre-notified of night works; evacuation drill or tabletop exercise completed "
         "before the first night pour of the campaign."),
    ]),
    ("9.4", "SAFE APPROACH CLEARANCES & MACHINERY SPOTTING", [
        ("Dedicated EHS Machinery Spotter",
         "Assign a dedicated spotter whose sole duty is to continuously monitor the "
         "concrete pump boom and mixer movements, equipped with a high-intensity torch "
         "and whistle."),
        ("Safe Clearance Limits",
         "Strict adherence to minimum safe clearance distances from overhead lines and "
         "energized equipment; the pump boom must never breach restricted red zones."),
        ("Physical Wheel Stoppers",
         "Heavy-duty wheel stoppers / barriers must be positioned to prevent concrete "
         "mixers from reversing toward live equipment or grounding conductors."),
        ("Marked Boom Envelope",
         "The maximum permitted boom working envelope is measured, marked and briefed "
         "before the shift; the operator and Spotter both hold the same reference, and "
         "any approach to the limit triggers an immediate stop."),
        ("Vehicle / Pedestrian Segregation",
         "Banksman with illuminated baton directs all reversing; one-way routing; "
         "segregated and lit pedestrian access; no person walks under the boom, the end "
         "hose or a loaded chute; site speed limit enforced."),
    ]),
    ("9.5", "ELECTRICAL SAFETY & EQUIPMENT GROUNDING", [
        ("Concrete Pump Body Grounding",
         "Body grounding must be applied to the concrete pump truck, connected directly "
         "to the station earth grid to mitigate step and touch potential risks."),
        ("RCD / GFCI Electrical Protection",
         "All power cables feeding concrete vibrators, distribution boxes and temporary "
         "lights must be protected by 30 mA residual current devices (RCD); the test "
         "button is operated and the result recorded before each shift."),
        ("Elevated Cable Management",
         "All electrical cables and extension leads must be elevated off wet ground and "
         "poured concrete using insulated cable hangers."),
        ("Authorized Person Only",
         "Installation, testing, repair and modification of temporary electrical systems "
         "are performed only by the Authorized Electrical Person; damaged equipment is "
         "removed from service, tagged and replaced \u2014 never taped."),
        ("Wet-Condition Integrity",
         "IP-rated enclosures and connectors suitable for wet conditions; no joints or "
         "sockets lying in water, laitance or fresh concrete; earth continuity verified "
         "and recorded; insulated rescue hook and insulated mats available at the pour "
         "site."),
    ]),
    ("9.6", "ILLUMINATION & VISIBILITY STANDARDS", [
        ("Pouring Task Lighting",
         "Minimum 100 lux focused lighting at the concrete discharge / pouring point and "
         "50 lux along access pathways and transit routes, measured with a calibrated "
         "light meter and recorded."),
        ("Glare & Shadow Control",
         "Light towers must be angled downward to prevent blinding mixer / pump operators "
         "and eliminate dark shadows in working zones."),
        ("Backup Emergency Lighting",
         "Battery-operated portable floodlights must be stationed at the pour site in "
         "case of primary power failure; personal torches issued to the Spotter, "
         "Banksman, EHS Officer and Medic."),
        ("Continuous Verification",
         "Lux re-measured after any change in tower position, after refuelling/relocation "
         "and at least once mid-shift; if illumination falls below the required value, "
         "placement stops until it is restored."),
        ("Reflective Marking & Signage",
         "Reflective tape on all barricades, plant, outriggers, delivery lines and "
         "vehicle tails; illuminated or retro-reflective signage at the entry point, "
         "open edges and the egress route."),
    ]),
    ("9.7", "SPECIALIZED CONCRETE PPE & EMERGENCY RESCUE GEAR", [
        ("High-Visibility Attire",
         "Class 3 high-visibility reflective vests / attire mandatory for all night-shift "
         "personnel (ANSI/ISEA 107 Class 3 or equivalent EN standard)."),
        ("Concrete Handling PPE",
         "Steel-toe rubber safety boots, heavy-duty chemical / rubber gloves, and sealed "
         "safety goggles to protect against chemical burns and splatter; long sleeves and "
         "waterproof trousers / apron for anyone handling wet concrete, plus face shield "
         "where splatter risk is high."),
        ("Electrical Rescue Hook",
         "An insulated electrical rescue hook and an illuminated first-aid kit must be "
         "stationed directly at the pour site."),
        ("Additional Task PPE",
         "Helmet with chin strap where fall or struck-by risk exists; hearing protection "
         "near pump, vibrator and mixer; dust mask (FFP3) for cement handling, cutting or "
         "dry sweeping; waterproof knee protection for finishers."),
        ("PPE Inspection & Replacement",
         "PPE checked at the TBT and during the shift; contaminated or damaged PPE "
         "replaced immediately; skin contact with wet concrete washed off at once \u2014 "
         "cement is corrosive (pH \u2248 13) and causes burns and dermatitis."),
    ]),
]


def section_9(doc):
    heading(doc, "SECTION 9: MANDATORY EHS REQUIREMENTS \u2014 NIGHT CONCRETE POURING")
    body(doc,
         "The requirements in this Section are mandatory minimum conditions for night "
         "concrete pouring on this project. They shall be read together with the work "
         "sequence in Section 6, verified on the Night Pour Readiness Checklist in "
         "Section 11, and evidenced by the records listed in Section 12. Where a "
         "Siemens Energy standard, Iraqi legal requirement or client (MOE/BGC) "
         "requirement is more stringent, the more stringent requirement applies.")
    for num, title, items in EHS_SECTIONS:
        subheading(doc, f"{num}  {title}", space_before=190)
        rows = [["#", "Mandatory EHS Requirement", "Detail / Verification Evidence"]]
        for i, (lead, txt) in enumerate(items, start=1):
            rows.append([str(i), lead, txt])
        grid(doc, [420, 2900, 6544], rows, size=9.5, header_size=9.5,
             aligns=["c", None, None], bolds=[1], zebra=True)


RISK_ROWS = [
    # (hazard, risk, controls, S, L, rS, rL)
    ("Exposed vertical rebar ends",
     "Impalement injuries, deep puncture wounds from trips / falls at night.",
     ["Install approved rebar mushroom caps on all exposed rebar ends.",
      "Ensure 100 lux task lighting over accessways.",
      "Verify 100 % cap coverage by joint EHS/Supervisor walk before the shift (checklist item).",
      "Barricade and light all rebar-dense zones; prohibit walking on rebar mats.",
      "Replace damaged/missing caps immediately; spare caps kept at the pour site."],
     4, 4, 4, 1),
    ("Unauthorized access to pour zone",
     "Struck by machinery, fall into open pits due to low visibility.",
     ["Enclose the work perimeter with rigid hard barricading and flashing lights.",
      "Enforce strict PTW access control with one manned entry point and security log.",
      "Illuminated/retro-reflective signage: \u201cConcrete pouring in progress \u2013 authorised "
      "personnel only\u201d.",
      "Barricade, light and sign every open pit, edge and excavation; provide lit access ladders.",
      "Night TBT covers boundaries; visitors escorted with PPE and TBT."],
     4, 3, 4, 1),
    ("Mechanical failure of concrete pump",
     "Boom collapse, hydraulic line rupture near live lines.",
     ["Mandatory valid third-party certificate for pump and boom before site entry.",
      "Mandatory certified third-party operator (valid competency card).",
      "Daily pre-use inspection: outriggers, pins, hoses, boom sections, emergency stops, alarms.",
      "Set up on firm, level, compacted ground; outriggers 100 % extended on pads.",
      "Suspend boom operation above 32 km/h wind (or manufacturer's lower limit); exclude "
      "personnel from under the boom and delivery line.",
      "Hydraulic spill kit and fire extinguisher at the pump; emergency stop drill briefed."],
     4, 3, 4, 1),
    ("Transit mixer reversing accidents",
     "Property damage, personnel injury, contact with live equipment.",
     ["Verify valid heavy vehicle driving licences for all drivers.",
      "Deploy banksman / traffic marshal with illuminated baton; no reversing without direction.",
      "Install physical wheel stoppers at discharge positions.",
      "Approved one-way routing, site speed limit, segregated and lit pedestrian access.",
      "Functioning reverse alarms, lights and mirrors verified pre-shift; drivers' fitness check.",
      "Mixer holding area outside the pour zone; egress lane never used for queuing."],
     4, 3, 4, 1),
    ("Human error & fatigue",
     "Operational mistakes and accidents due to lack of alertness.",
     ["Enforce 100 % crew rotation (minimum 8 h rest).",
      "Deploy a dedicated Night EHS Officer (fully rested; not a day-shift officer).",
      "Execute the formal Shift Handover Log signed by both shifts.",
      "Night TBT with confirmation of understanding; pre-shift fitness checks and Night Crew Register.",
      "Enforced rest breaks, hydration and sheltered lit rest area; two unannounced alertness "
      "rounds including 02:00\u201305:00 hrs.",
      "STOP WORK AUTHORITY for every person; no individual works two shifts in 24 hours."],
     3, 4, 3, 1),
    ("Delayed medical emergency response",
     "Delayed casualty stabilization during night incidents.",
     ["Station standby ambulance with dedicated driver directly at the pour zone.",
      "Certified Site Medic on duty (CPR, trauma, high-voltage electric-shock response).",
      "Keep the hospital evacuation route clear at all times.",
      "Ambulance equipped with oxygen cylinder, spinal board, resuscitation kit and operational "
      "emergency lights; inspected before the shift.",
      "Illuminated first-aid kit and insulated rescue hook at the pour site; emergency contacts "
      "and muster point briefed at the TBT.",
      "Al-Zubair General Hospital pre-notified; route map and travel time confirmed; drill or "
      "tabletop exercise before the first night pour."],
     5, 3, 5, 1),
]

EXTRA_RISK_ROWS = [
    ("Contact with overhead lines / energized equipment by boom, chute or delivery line",
     "Electrocution, arc flash, burns, fatality, plant damage, fire, network outage.",
     ["Observe minimum safe approach distances (OSHA 1926.1408 Table A / NFPA 70E) and any "
      "greater distance stated on the PTW / EKAD; boom never breaches the restricted red zone.",
      "Live zones identified with the Area Authority, marked physically and shown on the night "
      "lighting / site plan; briefed at the TBT.",
      "Dedicated Spotter continuously monitors the boom envelope with torch, whistle and radio; "
      "immediate stop signal on any approach to the limit.",
      "Pump body grounding connected to the station earth grid; earth continuity verified and "
      "recorded (step & touch potential, IEEE Std 80).",
      "Insulated rescue hook stationed at the pour site; rescuers trained not to touch the "
      "casualty or plant until isolation is confirmed.",
      "Electrical Work Permit / EKAD for any work near energized parts; work suspended on any "
      "change in network configuration."],
     5, 4, 5, 1),
    ("Temporary electrical systems, vibrators and lighting in wet conditions",
     "Electric shock, electrocution, burns, fire from damaged cables, RCD failure, water ingress.",
     ["30 mA RCD protection on all outlets feeding vibrators, distribution boxes and temporary "
      "lights; test button operated and recorded each shift.",
      "All cables and extension leads elevated off wet ground and poured concrete on insulated "
      "cable hangers; crossings protected by cable ramps; no joints in wet areas.",
      "IP-rated enclosures and connectors; insulation, PAT tags and connector condition inspected "
      "before use; damaged equipment removed from service and tagged.",
      "Only the Authorized Electrical Person installs, tests, repairs or modifies; LOTO applied "
      "for any intervention.",
      "Fire extinguisher (ABC/CO2) at each board and the pump; no overload of temporary supplies.",
      "Insulated mats and rescue hook available; electric-shock response briefed at the TBT."],
     4, 3, 4, 1),
    ("Wet concrete, cement splatter and chemical exposure",
     "Chemical burns, cement dermatitis, eye injury (alkaline pH \u2248 13), allergic sensitisation.",
     ["Steel-toe rubber safety boots, heavy-duty chemical/rubber gloves, sealed safety goggles; "
      "face shield, long sleeves and waterproof apron/trousers where splatter risk is high.",
      "Skin contact washed off immediately with clean water; eyewash and clean water available at "
      "the pour site.",
      "MSDS available for cement, admixtures and curing compounds; gloves/goggles per MSDS; "
      "barrier cream provided.",
      "No kneeling or standing in fresh concrete; kneeling on protected boards with knee pads.",
      "Contaminated PPE and clothing removed and replaced; first-aid kit illuminated and stocked "
      "for burn treatment.",
      "Curing compounds applied per MSDS with appropriate gloves, goggles and mask."],
     3, 4, 3, 1),
    ("Slips, trips and falls on wet surfaces, hoses and debris at night",
     "Sprains, fractures, head injury, fall into rebar or an open edge.",
     ["Maintain \u2265 50 lux on access pathways and transit routes and \u2265 100 lux at the working "
      "face; re-measure after any lighting change.",
      "Immediate housekeeping: spilled concrete, laitance, water, hoses and debris removed; "
      "walkways kept clear and lit.",
      "Cables and hoses routed off walkways and elevated; anti-slip footwear (EN ISO 20345 S5 "
      "rubber boots) mandatory.",
      "Open edges, pits and penetrations hard-barricaded, lit and signed; lit ladders/stairs for "
      "access; no running, no mobile phone use while walking.",
      "Work at height \u2265 1.8 m under a Work at Height Permit with 100 % tie-off and certified "
      "anchor points.",
      "Light towers angled to eliminate dark shadows in the working zone."],
     3, 4, 3, 1),
    ("Plant / pedestrian interface and congested night logistics",
     "Struck-by or crushed injuries, collision between mixers and plant, entrapment in blind spots.",
     ["Segregated, lit and signed pedestrian routes; physical separation from vehicle routes.",
      "Trained banksman with illuminated baton for all movements; radio contact with operators; "
      "agreed stop signals rehearsed before the shift.",
      "Controlled number of mixers on site; queuing in the holding area only; one-way circuit with "
      "no reversing except under direction.",
      "Reflective marking on plant, barricades and personnel (Class 3 hi-vis); reverse alarms, "
      "lights and mirrors functional.",
      "No person stands under a boom, loaded chute, end hose or suspended load; exclusion zones "
      "barricaded.",
      "Speed limit enforced; keys controlled; only competent operators drive or operate plant."],
     4, 4, 4, 1),
    ("Manual handling of pump lines, vibrator heads and heavy equipment",
     "Musculoskeletal injury, back strain, crush and pinch injuries to hands and feet.",
     ["Mechanical aids and two-person lifting for pump lines, end hoses and vibrator heads; correct "
      "lifting technique briefed at the TBT.",
      "Delivery lines supported and secured to reduce weight carried by hand; hoses never wrapped "
      "around the body or shoulder.",
      "Gloves (EN 388) and safety boots with metatarsal protection where required; task rotation "
      "and enforced rest breaks to limit fatigue-related handling errors.",
      "Materials stacked and stored at the point of use to avoid double handling in the dark.",
      "Any handling injury reported and treated immediately; first aider on site."],
     3, 3, 3, 1),
    ("Adverse night weather \u2014 wind, rain, lightning, dust storm",
     "Boom instability and collapse, loss of visibility, electrical water ingress, slippery "
     "surfaces, heat/cold stress.",
     ["Wind speed measured with an anemometer before and during boom operation; suspend above "
      "32 km/h or the manufacturer's lower limit.",
      "Lightning: all work stopped, plant lowered, personnel moved to a safe shelter on detection "
      "or first thunder.",
      "Rain: RCD protection and IP-rated equipment verified; cables elevated; pouring suspended if "
      "surface water cannot be controlled; extra lighting deployed.",
      "Dust storm: work stopped, visibility threshold enforced, plant parked and barricaded, "
      "personnel mustered.",
      "Weather forecast reviewed at planning; abort/re-schedule decision made by the Construction "
      "Manager with EHS.",
      "Shelter, warm drinks, hydration and rest area provided for the night crew."],
     4, 3, 4, 1),
    ("Concrete supply interruption / extended pour duration leading to a cold joint",
     "Structural/quality defect, unplanned night stoppage, re-mobilization in darkness, rushed "
     "work and fatigue at the end of shift.",
     ["Pour rate matched to mixer supply; delivery schedule and holding area agreed before the "
      "shift; backup supplier identified.",
      "Construction joint location pre-agreed with Quality/Engineering if interruption exceeds the "
      "initial set window.",
      "Shift duration monitored against the fatigue limit; relief crew available; no shift extension "
      "without the Construction Manager's written approval and a fresh fitness check.",
      "Lighting, barricades and curing protection maintained over any joint or unfinished work until "
      "the day shift takes over.",
      "Interruption recorded in the handover log and escalated to Quality and the Site Manager."],
     3, 3, 3, 1),
    ("Environmental release \u2014 washout water, cement slurry, hydraulic/fuel spill",
     "Soil and drain contamination, regulatory non-compliance, slip hazard, damage to the earth grid "
     "area.",
     ["Designated concrete washout pit/skip located away from drains, ground water paths and the "
      "station earth grid; washout only at that location.",
      "Spill kits and absorbents at the pump, boards and refuelling point; immediate containment of "
      "any hydraulic oil or fuel release.",
      "Waste segregated (concrete debris, packaging, oily absorbents, general waste) and removed to "
      "the approved storage area per the Environmental / Waste Management Plan; MOTAT records "
      "maintained for hazardous waste.",
      "No discharge of slurry into drains, excavations or the sea; empty mixer drums washed only at "
      "the washout point.",
      "Environmental incident reported immediately to EHS and recorded in the Daily EHS Report."],
     3, 3, 3, 1),
]


def section_10(doc):
    heading(doc, "SECTION 10: RISK ASSESSMENT MATRIX (JSA TABLE)")
    body(doc,
         "Risk scoring follows the Siemens Energy 5 \u00d7 5 matrix used across this project: "
         "Risk Score R = Severity (S) \u00d7 Likelihood (L), rated Low (1\u20136), Medium (8\u201312) "
         "or High (15\u201325). Initial risk is the score with no precautions in place; "
         "residual risk is the score after the mandatory controls in this Method "
         "Statement are fully implemented and verified. Table 10.1 (overleaf, landscape) "
         "is the JSA table for the principal night-pouring hazards; Table 10.2 covers the "
         "additional hazards identified for the activity. All hazards shall also be "
         "carried into the project Job Hazard Assessment document KAZ-EHS-JHA-2026-005 in "
         "the Siemens 5-step JHA format.")
    note_box(doc, "RESIDUAL RISK STATEMENT",
             "With every mandatory control of Section 9 implemented, verified and signed "
             "on the Night Pour Readiness Checklist, all residual risk ratings for night "
             "concrete pouring are reduced to LOW and are considered acceptable, subject "
             "to continuous monitoring by the Night EHS Officer. If any control cannot be "
             "implemented, the residual rating remains at its initial value and the pour "
             "shall not proceed \u2014 the activity must be re-planned for daylight hours or "
             "the control must be provided. Any residual rating of MEDIUM or HIGH requires "
             "the written approval of the Project Manager and EHS Manager before work "
             "starts.")


def section_10_landscape(doc):
    """Landscape section holding the two JSA tables."""
    widths = [2500, 3000, 6298, 1300, 1300]
    total = USABLE_L

    def build_table(rows, caption, heading_text):
        subheading(doc, heading_text, size=11.5, space_before=60)
        data = [["Hazard / Activity", "Identified Potential Risk",
                 "Mandatory Risk Controls", "Initial Risk\nS \u00d7 L = R",
                 "Residual Risk\nS \u00d7 L = R"]]
        for h, r, ctrls, S, L, rS, rL in rows:
            ctrl_text = "\n".join("\u2022 " + c for c in ctrls)
            data.append([h, r, ctrl_text, f"{S} \u00d7 {L} = {S*L}",
                         f"{rS} \u00d7 {rL} = {rS*rL}"])
        t = grid(doc, widths, data, size=8.5, header_size=9,
                 aligns=[None, None, None, "c", "c"], bolds=[0], total=total, zebra=True)
        # colour-code the rating cells
        for ri, (_h, _r, _c, S, L, rS, rL) in enumerate(rows, start=1):
            init = S * L
            res = rS * rL
            rating_cell(t, ri, 3, "H" if init >= 15 else ("M" if init >= 8 else "L"),
                        f"S{S} \u00d7 L{L} = {init}", size=9)
            rating_cell(t, ri, 4, "H" if res >= 15 else ("M" if res >= 8 else "L"),
                        f"S{rS} \u00d7 L{rL} = {res}", size=9)
            shade(t.cell(ri, 3), {"L": "E7F2E7", "M": "FDF3DC", "H": "FBE2E2"}[
                "H" if init >= 15 else ("M" if init >= 8 else "L")])
            shade(t.cell(ri, 4), {"L": "E7F2E7", "M": "FDF3DC", "H": "FBE2E2"}[
                "H" if res >= 15 else ("M" if res >= 8 else "L")])
        body(doc, caption, size=8.5, align="j", space_after=140)
        return t

    build_table(
        RISK_ROWS,
        "Table 10.1 \u2014 JSA / Risk Assessment Matrix: principal night concrete pouring "
        "hazards (as issued in the project EHS requirements). Controls listed here are "
        "mandatory minimums; the full clause reference is given in Section 9.",
        "TABLE 10.1 \u2014 JSA RISK ASSESSMENT MATRIX: PRINCIPAL NIGHT-POURING HAZARDS")

    build_table(
        EXTRA_RISK_ROWS,
        "Table 10.2 \u2014 JSA / Risk Assessment Matrix: additional hazards identified for "
        "night concrete pouring in an operating power-station environment (electrical "
        "proximity, temporary power in wet conditions, cement chemistry, slips/trips, "
        "plant-pedestrian interface, manual handling, night weather, supply interruption "
        "and environmental release). These rows shall be read together with Table 10.1.",
        "TABLE 10.2 \u2014 JSA RISK ASSESSMENT MATRIX: ADDITIONAL NIGHT-POURING HAZARDS")

    subheading(doc, "TABLE 10.3 \u2014 SIEMENS 5 \u00d7 5 RISK MATRIX & RATING BANDS", size=11.5,
               space_before=60)
    sev = [["Severity (S) \u2014 Consequence of harm", "Likelihood (L) \u2014 Probability of occurrence"],
           ["1 \u2014 No injury", "1 \u2014 Very unlikely"],
           ["2 \u2014 Minor injury (first aid)", "2 \u2014 Unlikely"],
           ["3 \u2014 Lost-time injury (medical treatment)", "3 \u2014 Likely"],
           ["4 \u2014 Serious injury (disability / hospitalisation)", "4 \u2014 Very likely"],
           ["5 \u2014 Death or permanent disability", "5 \u2014 Certain"]]
    grid(doc, [4200, 4200], sev, size=9.5, header_size=9.5, total=8400, align="center")

    matrix = [["R = S \u00d7 L", "L = 1", "L = 2", "L = 3", "L = 4", "L = 5"]]
    for s in range(1, 6):
        matrix.append([f"S = {s}"] + [str(s * l) for l in range(1, 6)])
    t = grid(doc, [1400, 1400, 1400, 1400, 1400, 1400], matrix, size=9.5,
             header_size=9.5, aligns=["c"] * 6, bolds=[0], total=8400, zebra=False,
             align="center")
    for ri in range(1, 6):
        for ci in range(1, 6):
            v = ri * ci
            key = "H" if v >= 15 else ("M" if v >= 8 else "L")
            shade(t.cell(ri, ci), {"L": "E7F2E7", "M": "FDF3DC", "H": "FBE2E2"}[key])
            cell_text(t.cell(ri, ci), str(v), size=9.5, bold=True,
                      colour=RATING_COLOUR[key], align="c")
    spacer(doc, 60)

    bands = [["Risk Rating", "Score (R)", "Action Required"],
             ["LOW (L)", "1 \u2013 6",
              "Acceptable \u2014 monitor and review; controls maintained and verified each shift."],
             ["MEDIUM (M)", "8 \u2013 12",
              "Additional controls required; work may proceed only with documented controls, "
              "supervision and EHS verification. Written approval of the Project Manager and "
              "EHS Manager required for night work."],
             ["HIGH (H)", "15 \u2013 25",
              "STOP \u2014 work shall not proceed. The activity must be eliminated, re-engineered "
              "or re-scheduled (e.g. moved to daylight) until the rating is reduced."]]
    t = grid(doc, [2200, 1600, 8200], bands, size=9.5, header_size=9.5,
             aligns=["c", "c", None], bolds=[0], total=12000, zebra=False,
             align="center")
    for ri, key in enumerate(["L", "M", "H"], start=1):
        shade(t.cell(ri, 0), {"L": "E7F2E7", "M": "FDF3DC", "H": "FBE2E2"}[key])
        cell_text(t.cell(ri, 0), ["LOW (L)", "MEDIUM (M)", "HIGH (H)"][ri - 1], size=9.5,
                  bold=True, colour=RATING_COLOUR[key], align="c")
    body(doc,
         "P.A.E.R. target convention (as used in the project JHA): P = Personnel, "
         "A = Asset/Property, E = Environment, R = Reputation. All hazards in Tables 10.1 "
         "and 10.2 affect P and A; electrical proximity, environmental release and any "
         "incident involving live plant also affect E and R.",
         size=8.5, align="j", space_after=60)


def section_11(doc):
    heading(doc, "SECTION 11: NIGHT POUR READINESS CHECKLIST (PRE-POUR VERIFICATION)")
    body(doc,
         "This checklist shall be completed physically on site, at night, with the "
         "lighting on and the plant positioned as it will be used during the pour. Every "
         "line requires a tick, the verification method and the verifier's initials. Any "
         "\u201cNo\u201d answer is a hold point: the pour shall not start until the item is "
         "closed or the work is re-scheduled.")
    groups = [
        ("A. DOCUMENTATION, PERMITS & CERTIFICATES", [
            "Valid PTW issued/endorsed for this night shift, with all mandatory attachments.",
            "This Method Statement and the JSA (Tables 10.1 / 10.2) approved and available on site.",
            "Concrete pump truck & placing boom \u2014 valid third-party inspection certificate sighted on site.",
            "Pump operator \u2014 valid third-party competency card sighted on site.",
            "All mixer drivers \u2014 valid heavy vehicle driving licences and third-party induction certificates sighted.",
            "Site Medic certification (CPR / trauma / high-voltage shock response) sighted.",
            "Ambulance inspection record and equipment check completed.",
            "Calibrated light meter certificate valid; anemometer available.",
        ]),
        ("B. CREW, FATIGUE & BRIEFING", [
            "100 % night crew deployed \u2014 no day-shift worker or technician continuing into the night shift.",
            "Every night-crew member has had a minimum of 8 hours rest (Night Crew Register, Annex D).",
            "Dedicated, fully rested Night EHS Officer appointed and present for the whole shift.",
            "Pre-shift fitness checks completed for all crew, drivers and operators.",
            "Day \u2192 Night Shift Handover Log (Annex A) completed and signed by EHS and Engineering supervisors.",
            "Night Toolbox Talk delivered; attendance and signatures recorded (Annex B).",
            "TBT covered safe approach distances, live energized zones and emergency evacuation routes.",
            "Spotter, Banksman, Permit Holder and Medic named and present; radios tested on the night channel.",
        ]),
        ("C. BARRICADING, ACCESS & WORK-FACE PREPARATION", [
            "Pour area, open edges and pump perimeter fully enclosed with rigid hard barricades (not warning tape).",
            "Reflective tape and warning flashers fitted to all barricades; illuminated signage at the entry point.",
            "Single controlled entry/exit point manned; PTW access control enforced; visitor log in use.",
            "100 % of exposed vertical and horizontal rebar ends fitted with approved mushroom caps.",
            "Access pathways clear, lit and free of rebar, debris and standing water.",
            "Heavy-duty wheel stoppers installed at mixer discharge positions (protecting live equipment and grounding conductors).",
            "Open excavations/pits/edges barricaded, lit and signed; lit ladders or stairs provided.",
            "Formwork, reinforcement and embedments inspected and released by Quality; pour sequence and joint locations agreed.",
        ]),
        ("D. PLANT, PUMP & TRAFFIC", [
            "Pump positioned on firm, level, compacted ground clear of excavations and underground services.",
            "Outriggers 100 % extended on pads; ground bearing confirmed.",
            "Boom maximum permitted envelope measured, marked and briefed; clearance from overhead lines / energized equipment verified against the permit limit.",
            "Emergency stops, alarms, hydraulic lines, hoses and clamps inspected; delivery line secured and supported.",
            "Mixers routed one-way; holding area defined; site speed limit briefed; egress lane free of queuing.",
            "Reverse alarms, lights, mirrors and brakes functional on every mixer; banksman's illuminated baton working.",
            "Wind speed measured \u2014 below 32 km/h (or manufacturer's lower limit); weather monitored during the shift.",
            "Washout pit/skip and spill kit in position; fire extinguishers at the pump and boards.",
        ]),
        ("E. ELECTRICAL SAFETY & GROUNDING", [
            "Concrete pump truck body grounding connected directly to the station earth grid; connection inspected.",
            "Earthing continuity / resistance measured and recorded (Annex C) by the Authorized Electrical Person.",
            "30 mA RCD protection verified on all outlets feeding vibrators, distribution boxes and temporary lights.",
            "RCD test button operated on every device; trip and reset confirmed; result logged.",
            "All cables and extension leads elevated off wet ground and poured concrete on insulated cable hangers.",
            "Enclosures/connectors IP-rated for wet conditions; no joints, sockets or boards standing in water.",
            "No damaged cables, taped repairs or unauthorised extensions; defective equipment removed and tagged.",
            "Insulated rescue hook and insulated mats stationed at the pour site.",
        ]),
        ("F. ILLUMINATION & VISIBILITY", [
            "\u2265 100 lux measured and recorded at the concrete discharge / pouring point (Annex C).",
            "\u2265 50 lux measured and recorded along access pathways and transit routes (Annex C).",
            "Light towers angled downward \u2014 no glare into mixer/pump operator positions; no dark shadows in the working zone.",
            "Battery-operated backup floodlights staged at the pour site, charged and tested.",
            "Personal high-intensity torches issued to Spotter, Banksman, EHS Officer and Medic; whistles issued.",
            "Reflective marking on plant, outriggers, delivery lines, barricades and vehicle tails.",
            "Class 3 high-visibility reflective attire worn by 100 % of night-shift personnel.",
        ]),
        ("G. PPE & RESCUE GEAR", [
            "Steel-toe rubber safety boots, heavy-duty chemical/rubber gloves and sealed safety goggles issued and worn by all concrete handlers.",
            "Face shields / waterproof aprons and long sleeves available for high-splatter tasks; helmets with chin strap where required.",
            "Hearing protection, FFP3 masks and knee protection available for the relevant tasks.",
            "PPE inspected at the TBT; damaged or contaminated PPE replaced before the shift.",
            "Illuminated first-aid kit stationed directly at the pour site; contents and expiry dates checked.",
            "Clean water / eyewash available at the pour site for cement contact; MSDS accessible.",
        ]),
        ("H. MEDICAL & EMERGENCY PREPAREDNESS", [
            "Standby ambulance stationed directly at the concrete pouring area for the entire night shift.",
            "Dedicated standby ambulance driver on duty and remaining with the vehicle.",
            "Ambulance equipped with oxygen cylinder, spinal board, resuscitation kit; emergency lights operational.",
            "Certified Site Medic on duty on site for the whole shift.",
            "Emergency evacuation route verified clear of all mixer trucks, machinery and obstructions.",
            "Route, travel time and receiving point at Al-Zubair General Hospital confirmed; hospital pre-notified of night works.",
            "Alarm signal, muster point, emergency contacts and radio channel briefed at the TBT.",
            "Emergency scenario responses briefed: electric shock, impalement, struck-by, fall, chemical burn, pump/boom failure.",
        ]),
        ("I. ENVIRONMENT & HOUSEKEEPING", [
            "Concrete washout restricted to the designated pit/skip away from drains and the earth grid.",
            "Waste segregation in place (concrete debris, packaging, oily absorbents, general waste); MOTAT records ready.",
            "Spill kit available and crew briefed on hydraulic/fuel spill response.",
            "Housekeeping completed; housekeeping maintained during the pour; walkways and egress kept clear.",
        ]),
    ]
    for gtitle, items in groups:
        rows = [["#", "Verification Item", "Yes", "N/A", "Verification Method", "Initials"]]
        for i, it in enumerate(items, start=1):
            rows.append([str(i), it, "\u2610", "\u2610", "", ""])
        subheading(doc, gtitle, size=10.5, space_before=150, fill=BAND)
        grid(doc, [400, 5400, 520, 520, 2200, 824], rows, size=9, header_size=9,
             aligns=["c", None, "c", "c", None, "c"], total=USABLE_P, zebra=True)

    subheading(doc, "READINESS AUTHORISATION \u2014 POUR MAY / MAY NOT COMMENCE", space_before=170)
    rows = [["Declaration", "Name", "Signature", "Date / Time"],
            ["All items above are verified as implemented. The night concrete pour is "
             "authorised to commence.\n(Any open \u201cNo\u201d item = NO POUR.)", "", "", ""],
            ["Permit Holder / Civil Supervisor", "", "", ""],
            ["Dedicated Night EHS Officer", "", "", ""],
            ["Permit Issuer (Siemens Energy)", "", "", ""],
            ["Area Authority (client / operations, where applicable)", "", "", ""],
            ["Night Shift In-Charge / Construction Manager", "", "", ""],
            ["Site Manager (Siemens Energy)", "", "", ""],
            ["Concrete Pouring Subcontractor \u2014 Authorized Representative", "", "", ""],
            ["Concrete Pouring Subcontractor \u2014 Site Supervisor / Permit Holder", "", "", ""]]
    grid(doc, [4200, 2200, 2200, 1264], rows, size=9.5, header_size=9.5,
         bolds=[0], total=USABLE_P)


def section_12(doc):
    heading(doc, "SECTION 12: EMERGENCY RESPONSE, MEDICAL FACILITY & COMMUNICATION")
    subheading(doc, "12.1  Emergency Contacts (to be completed and displayed at the pour site)",
               space_before=60)
    rows = [["Contact", "Name / Location", "Number"]]
    for a, b, c in [
        ("Site Emergency / Control Room", "KAZ PP Upgrade Project site control room", "____________"),
        ("Standby Ambulance (on site)", "Stationed at the pour zone \u2014 driver: ____________", "____________"),
        ("Certified Site Medic", "On duty at the pour site \u2014 name: ____________", "____________"),
        ("Night EHS Officer", "On site for the whole shift \u2014 name: ____________", "____________"),
        ("Night Shift In-Charge", "Name: ____________", "____________"),
        ("Siemens Energy Emergency Contact", "24 h emergency line (per project JHA cover sheet)", "+90 533 625 9494"),
        ("Ambulance (public \u2014 Al-Is'aaf)", "Iraqi national emergency medical service", "122"),
        ("Fire / Civil Defense Directorate", "Iraqi civil defence \u2014 fire and technical rescue", "115"),
        ("Police (Al-Najda)", "Iraqi emergency police", "104"),
        ("Traffic Police", "Iraqi traffic directorate", "130"),
        ("Site Emergency Control Room", "KAZ PP Upgrade Project \u2014 24 h manned", "____________"),
        ("Note", "112 may not operate reliably in Iraq \u2014 always dial the direct Iraqi "
                 "codes above. Verify all numbers locally and display them at the pour "
                 "site, in the ambulance and on the Alarm Card.", "\u2014"),
        ("Receiving Hospital", "Al-Zubair General Hospital, Al-Zubair, Basra", "____________"),
        ("Alternate Hospital", "____________ (as defined in the Project EHS Plan)", "____________"),
        ("Client / MOE-BGC Duty Officer", "____________", "____________"),
    ]:
        rows.append([a, b, c])
    grid(doc, [3000, 4700, 2164], rows, size=9.5, header_size=9.5, bolds=[0])

    subheading(doc, "12.2  Medical Standby & Evacuation Arrangements")
    bullets(doc, [
        "A fully functional, inspected site ambulance with a dedicated standby driver is "
        "stationed directly at the concrete pouring area throughout the entire night "
        "shift; it is not shared with another work front and is not used for transport of "
        "materials or personnel.",
        "Ambulance equipment verified before the shift: oxygen cylinder (charged), spinal "
        "board, resuscitation kit, trauma dressings, burns kit, eyewash and operational "
        "emergency lights and siren.",
        "A certified Site Medic / First Aider trained in CPR, trauma response and "
        "high-voltage electric-shock response is on duty on site for the whole shift.",
        "The emergency evacuation route from the pour zone to the site gate and on to "
        "Al-Zubair General Hospital is kept clear of all mixer trucks, machinery and "
        "obstructions at all times; the route is lit, marked and driven once before the "
        "shift to confirm travel time.",
        "The receiving hospital is pre-notified of night works and the expected crew "
        "numbers; the hospital name, route, travel time and contact number are recorded on "
        "the emergency board at the pour site and briefed at the TBT.",
        "Muster point, alarm signal (site horn / radio announcement) and the Alarm Card "
        "procedure apply; headcount is taken by the Night Shift In-Charge with the "
        "security access log as the reference.",
        "At least one emergency exercise (practical or tabletop \u2014 electric shock and "
        "impalement scenarios) is completed and recorded before the first night pour of "
        "each campaign.",
    ])

    subheading(doc, "12.3  Immediate Response Actions by Scenario")
    rows = [["Scenario", "Immediate Actions (all personnel briefed at the TBT)"]]
    scen = [
        ("Electric shock / contact with energized equipment",
         "Do NOT touch the casualty or the plant. Raise the alarm and call an immediate stop. "
         "Isolate/de-energize (Authorized Electrical Person / LOTO) \u2014 if isolation is not "
         "possible, separate the casualty using the insulated rescue hook stationed at the pour "
         "site. Start CPR/AED via the Site Medic; treat burns; keep the casualty still and warm; "
         "evacuate to Al-Zubair General Hospital; preserve the scene for investigation."),
        ("Rebar impalement / deep puncture",
         "Raise the alarm; do NOT remove or cut the impaling object; stabilise it in place with "
         "padding. Control bleeding around the wound, treat for shock, immobilise and transport "
         "on the spinal board with the object secured. Site Medic accompanies the casualty."),
        ("Struck-by / crush injury from plant or mixer",
         "Stop all plant immediately, isolate keys and secure the area. Do not move the casualty "
         "unless there is immediate danger. Site Medic stabilises; spinal precautions applied; "
         "evacuate by ambulance. Banksman and operators preserved as witnesses; scene photographs "
         "taken by EHS."),
        ("Fall into excavation / pit or from height",
         "Stop work; keep the casualty still; do not attempt an uncontrolled rescue. Site Medic "
         "assesses from a safe position; rescue performed by trained personnel using the "
         "stretcher/tripod and spinal board; evacuate immediately."),
        ("Cement / chemical burn or eye splatter",
         "Move the casualty clear; irrigate skin or eyes with clean water or eyewash for at least "
         "15\u201320 minutes; remove contaminated clothing, boots and gloves; cover with a clean "
         "dressing; Site Medic assesses and refers to hospital for any burn larger than the palm "
         "or any eye exposure."),
        ("Pump / boom failure, hydraulic rupture or collapse",
         "Emergency stop activated; all personnel withdrawn to the muster point and counted; boom "
         "and delivery line depressurised only by the certified operator; area barricaded; spill "
         "contained with the spill kit; the pump is not returned to service until inspected and "
         "re-certified."),
        ("Lighting failure / loss of illumination",
         "Stop placement immediately; deploy the battery-operated backup floodlights; do not move "
         "plant until illumination is restored to \u2265 100 lux / 50 lux and re-recorded on Annex C."),
        ("Fire at the pump, board or vehicle",
         "Raise the alarm; isolate fuel and power; attack with ABC/CO2 extinguisher only if safe "
         "and trained; evacuate to the muster point; call the site fire team and 115/112."),
        ("Adverse weather (wind > 32 km/h, lightning, dust storm)",
         "Stop boom operation and all placement; lower the boom; park and secure plant; move "
         "personnel to shelter; re-assess with the anemometer and resume only on the Construction "
         "Manager's instruction with a re-check of the readiness checklist."),
    ]
    for a, b in scen:
        rows.append([a, b])
    grid(doc, [2700, 7164], rows, size=9.5, header_size=9.5, bolds=[0])

    subheading(doc, "12.4  Reporting & Records")
    bullets(doc, [
        "All incidents, injuries, dangerous occurrences and near misses are reported "
        "immediately to the Night EHS Officer and Night Shift In-Charge, recorded on the "
        "Near-Miss Report Form (000480) / project incident form, and entered in the Daily "
        "EHS Report.",
        "Siemens Energy and client (MOE/BGC) notification timelines per the Project EHS "
        "Plan are observed; the scene is preserved until released by EHS.",
        "Records generated by this Method Statement: PTW and endorsements; readiness "
        "checklist; shift handover log; TBT record; Night Crew Register; lux survey, RCD "
        "test and earthing records; third-party certificates and competency cards; "
        "delivery tickets and quality records; washout/waste (MOTAT) records; incident and "
        "near-miss reports.",
    ])


def section_13(doc):
    heading(doc, "SECTION 13: TRAINING, COMMUNICATION & COMPETENCY")
    rows = [["Training / Communication", "Audience", "Frequency", "Record"]]
    items = [
        ("Site EHS induction and third-party safety induction", "All night-shift personnel, drivers, suppliers",
         "Before first entry / per client requirement", "Induction register"),
        ("Method Statement & JSA briefing (this document)", "Entire night crew, supervisors, operators",
         "Before the first night pour and after any revision", "Briefing register / signatures"),
        ("Night Toolbox Talk (TBT) \u2014 scope, live zones, approach distances, evacuation routes",
         "Entire night crew", "Every shift, before work starts", "Annex B \u2014 TBT record"),
        ("Banksman / Spotter / signaller training", "Appointed banksman and machinery spotter",
         "Before appointment; refresher annually", "Competency records"),
        ("Concrete pump operator competency (third party)", "Pump operator",
         "Valid card; verified every shift", "Competency card copy on site"),
        ("Heavy vehicle driving licence and defensive driving", "All mixer drivers",
         "Valid licence; induction before mobilization", "Licence copies / induction record"),
        ("First aid, CPR, trauma and high-voltage shock response", "Site Medic; nominated first aiders",
         "Valid certification; refresher per standard", "Certificates on site"),
        ("Electrical safety, RCD testing, earthing and LOTO", "Authorized Electrical Person",
         "Valid authorization; task briefing each shift", "Authorization record"),
        ("Emergency / evacuation exercise (shock + impalement scenarios)",
         "Night crew, medic, ambulance driver, security", "Before the first night pour of each campaign",
         "Exercise report"),
        ("STOP WORK AUTHORITY and fatigue awareness", "All personnel", "Induction + each TBT",
         "TBT record"),
    ]
    rows.extend(items)
    grid(doc, [4100, 2300, 2100, 1364], rows, size=9.5, header_size=9.5, bolds=[0])


def section_14(doc):
    heading(doc, "SECTION 14: SIGN-OFF & APPROVAL MATRIX")
    subheading(doc, "14.1  Document Preparation, Review & Approval", space_before=60)
    prepared = ("Name: Ahmed Mansoor\nRole: EHS Manager (Siemens Energy)\n"
                "Company: Siemens Energy\nStatus: Prepared\nDate: ____ / ____ / 2026\n\n"
                "Signature: ______________________")
    reviewed = ("Name: Dogukan Arandi\nRole: EHS / Site EHS Lead (Reviewer)\n"
                "Company: Siemens Energy\nStatus: Under Review\nDate: ____ / ____ / 2026\n\n"
                "Signature: ______________________")
    approved = ("Name: Onur Ozvatan\nRole: Project Manager\n"
                "Company: Siemens Energy\nStatus: For Approval\nDate: ____ / ____ / 2026\n\n"
                "Signature: ______________________")
    rows = [["PREPARED BY", "REVIEWED BY", "APPROVED BY"],
            [prepared, reviewed, approved]]
    grid(doc, [3288, 3288, 3288], rows, size=9.5, header_size=10,
         aligns=["c", None, None], total=USABLE_P, zebra=False)

    subheading(doc, "14.2  Site Management & Concrete Subcontractor Sign-Off",
               space_before=170)
    body(doc,
         "This Method Statement is not valid for execution until it is signed below by "
         "the Site Manager and by the authorized representative of the concrete pouring "
         "subcontractor. The subcontractor's signature confirms that it has read the "
         "document, will provide every resource and control it requires \u2014 including the "
         "dedicated night crew, certified pump operator, standby ambulance and medic \u2014 "
         "and will not start any night pour without the signed Night Pour Readiness "
         "Checklist (Section 11).", size=10, space_after=90)
    rows = [["Role / Entity", "Name", "Company", "Signature", "Date", "Stamp / Seal"]]
    signatories = [
        "Site Manager (Siemens Energy)",
        "Construction Manager / Night Shift In-Charge",
        "Concrete Pouring Subcontractor \u2014 Authorized Representative",
        "Concrete Pouring Subcontractor \u2014 Site Supervisor / Permit Holder",
        "Concrete Pouring Subcontractor \u2014 EHS Officer",
        "Concrete Pump Supplier \u2014 Authorized Representative",
        "Ready-Mix Concrete Supplier \u2014 Authorized Representative",
        "Night EHS Officer (dedicated night shift)",
        "Client / Area Authority (MOE-BGC) \u2014 Acknowledgement (where required)",
    ]
    for s in signatories:
        rows.append([s, "", "", "", "", ""])
    grid(doc, [2600, 1600, 1500, 1900, 1100, 1164], rows, size=9.5, header_size=9,
         aligns=[None, None, None, None, "c", "c"], bolds=[0], total=USABLE_P)
    spacer(doc, 60)

    subheading(doc, "14.3  Night-Pour Authorisation per Shift (in addition to the PTW)")
    rows = [["Shift / Date", "Night Shift In-Charge", "Night EHS Officer", "Permit Issuer / AA",
             "Pour Authorised? (Yes/No)"]]
    for _ in range(4):
        rows.append(["", "", "", "", ""])
    grid(doc, [1900, 2100, 2000, 2200, 1664], rows, size=9.5, header_size=9,
         aligns=[None] * 5, total=USABLE_P)

    subheading(doc, "14.4  Revision History")
    rows = [["Rev.", "Date", "Description of Change", "Prepared", "Reviewed", "Approved"]]
    rows.append(["0", "____ / ____ / 2026",
                 "First issue for review \u2014 Method Statement for Night Concrete Pouring "
                 "Operations, incorporating the comprehensive night-shift EHS requirements "
                 "(Sections 9.1\u20139.7) and the JSA risk assessment matrix (Section 10). "
                 "Sign-off includes the Site Manager and the concrete pouring "
                 "subcontractor (Section 14.2).",
                 "AM", "", ""])
    for _ in range(3):
        rows.append(["", "", "", "", "", ""])
    grid(doc, [700, 1500, 4200, 1000, 1200, 1264], rows, size=9, header_size=9,
         aligns=["c", "c", None, "c", "c", "c"], total=USABLE_P)


def section_15(doc):
    heading(doc, "SECTION 15: REFERENCES & TECHNICAL LIMIT SOURCES")
    bullets(doc, [
        "Iraqi Labour Law No. 37 of 2015 \u2014 occupational safety and health requirements, "
        "including working hours, rest and night-work provisions applicable in Iraq.",
        "OSHA 29 CFR 1926 Subpart Q \u2014 Concrete, Masonry, Bricklaying and Blocklaying "
        "construction (concrete pumps, vibrators, formwork, washout, personal protection).",
        "OSHA 29 CFR 1926 Subpart K \u2014 Electrical (temporary wiring, GFCI/RCD protection "
        "at 1926.404(b)(1), grounding, wet conditions).",
        "OSHA 29 CFR 1926 Subpart D \u2014 Occupational Health and Environmental Controls "
        "(1926.56 illumination of work areas).",
        "OSHA 29 CFR 1926 Subpart N \u2014 Cranes, Derricks, Hoists, Elevators and Conveyors "
        "(1926.1408 Table A power-line clearance distances; 1926.1400 wind and operational limits).",
        "OSHA 29 CFR 1926 Subparts M and P \u2014 Fall protection (\u2265 1.8 m) and Excavations "
        "(protective systems > 1.2 m, edge protection, egress).",
        "OSHA 29 CFR 1926 Subpart E \u2014 Personal Protective and Life-Saving Equipment; "
        "Subpart F \u2014 Fire Protection.",
        "NFPA 70E \u2014 Standard for Electrical Safety in the Workplace (approach boundaries, "
        "energized work controls).",
        "IEC 60364 / IEC 61008 \u2014 Low-voltage electrical installations and residual current "
        "operated circuit breakers (30 mA RCD).",
        "IEEE Std 80 \u2014 Guide for Safety in AC Substation Grounding (step and touch "
        "potential; equipment/body grounding practice).",
        "ANSI/ISEA 107 \u2014 High-Visibility Safety Apparel (Class 3) and EN ISO 20471 equivalent.",
        "EN 397 (helmets), EN 166 (eye protection), EN 388 / EN ISO 374 (gloves \u2014 mechanical "
        "and chemical), EN ISO 20345 S5 (rubber safety boots), EN 361 (harnesses).",
        "Siemens Energy \u2014 Annex 1 EHS Policy; Standard EHS Principles and Core "
        "Responsibilities; Instruction \u2014 EHS General Requirements for Contractors; SU EQS "
        "Minimum Requirements Instruction.",
        "KAZ-EHS-SOP-2026-003 \u2014 Permit to Work (PTW) System, Rev. 1 (permit categories, "
        "validity, shift handover, suspension triggers).",
        "KAZ Project EHS Plan; Environmental / Waste Management Plan; Security Plan; Daily EHS "
        "Report; Near-Miss Report Form (000480); Mobilization and Additional Controls Checklist.",
        "Client (MOE/BGC) site EHS rules, operating-plant restrictions and Area Authority "
        "instructions for the Khur Al-Zubair Power Station.",
        "Concrete pump, light tower and vibrator manufacturer operating manuals \u2014 where a "
        "manufacturer limit is stricter than any value in this document, the manufacturer "
        "limit applies.",
        "Where numerical limits in this Method Statement differ between standards or client "
        "requirements, the more stringent requirement applies.",
    ])


def annexes(doc):
    heading(doc, "ANNEX A \u2014 DAY \u2194 NIGHT SHIFT HANDOVER LOG")
    label_table(doc, [
        ("Project / Document", f"KAZ Power Plant Upgrade Project \u2014 {DOC_ID} Rev. {REV}"),
        ("Pour location / area", ""),
        ("Date of handover", ""),
        ("Handover type", "\u2610 Day \u2192 Night    \u2610 Night \u2192 Day"),
        ("PTW number(s) & validity", ""),
    ])
    rows = [["#", "Handover Item", "Outgoing Shift Entry", "Incoming Shift Verification (initials)"]]
    items = [
        "Works completed during the shift (locations, volumes, joints)",
        "Works outstanding / to be continued",
        "Active live zones and energized equipment; restricted red zones",
        "Open isolations / LOTO applied and remaining under Isolation Authority control",
        "PTW status, endorsements, restrictions and expiry",
        "Barricading condition and any temporary removals",
        "Lighting condition, lux values recorded, tower positions, backup floodlights",
        "Earthing/grounding connections and RCD test status",
        "Open excavations, pits, edges and rebar cap condition",
        "Plant left on site, keys, defects and out-of-service equipment",
        "Egress route condition; ambulance and medic standby status",
        "Incidents, near misses, unsafe acts/conditions and corrective actions raised",
        "Environmental issues (washout, spills, waste) and MOTAT records",
        "Instructions, restrictions or changes from the Area Authority / client",
        "Other hazards or information the incoming shift must know",
    ]
    for i, it in enumerate(items, start=1):
        rows.append([str(i), it, "", ""])
    grid(doc, [400, 4300, 3300, 1864], rows, size=9, header_size=9,
         aligns=["c", None, None, "c"], total=USABLE_P)
    rows = [["Role", "Name", "Signature", "Time"],
            ["Outgoing EHS Supervisor / Officer", "", "", ""],
            ["Incoming Night EHS Officer", "", "", ""],
            ["Outgoing Engineering / Civil Supervisor", "", "", ""],
            ["Incoming Night Shift In-Charge", "", "", ""],
            ["Permit Issuer (noted)", "", "", ""]]
    grid(doc, [3400, 2600, 2400, 1464], rows, size=9.5, header_size=9.5, bolds=[0],
         total=USABLE_P)

    heading(doc, "ANNEX B \u2014 NIGHT TOOLBOX TALK (TBT) RECORD")
    label_table(doc, [
        ("Date / Time of TBT", ""),
        ("Pour location / activity", ""),
        ("Delivered by (Night EHS Officer)", ""),
        ("PTW number", ""),
        ("Language(s) used", "\u2610 English   \u2610 Arabic   \u2610 Other: ____________"),
    ])
    subheading(doc, "Mandatory TBT Topics \u2014 tick when covered", size=10, space_before=90)
    topics = [
        "Scope, sequence and planned volume of tonight's pour; joint locations",
        "Safe approach distances; location of live energized zones and restricted red zones",
        "Marked boom envelope and the Spotter's stop signal; never stand under the boom or hose",
        "Barricaded perimeter, single controlled entry point and PTW access control",
        "Mixer routing, one-way circuit, reversing rules, banksman signals and wheel stoppers",
        "Rebar mushroom caps, open edges, excavations and fresh-concrete exclusion",
        "Electrical hazards: 30 mA RCDs, elevated cables, pump body grounding, rescue hook, no unauthorised repairs",
        "Wet concrete chemical burns: gloves, goggles, boots, immediate washing, eyewash location",
        "Illumination limits (100 lux / 50 lux), glare control and what to do if lights fail",
        "Emergency evacuation routes, muster point, alarm signal, ambulance location and route to Al-Zubair General Hospital",
        "Fatigue, rest breaks, hydration and STOP WORK AUTHORITY",
        "Environmental controls: washout pit, spill kit, waste segregation",
        "Weather limits: wind > 32 km/h, lightning, dust storm \u2014 stop rules",
        "Lessons learned from previous incidents / near misses on the project",
    ]
    rows = [["#", "Topic", "Covered (\u2713)"]]
    for i, tp in enumerate(topics, start=1):
        rows.append([str(i), tp, ""])
    grid(doc, [400, 7800, 1664], rows, size=9, header_size=9, aligns=["c", None, "c"],
         total=USABLE_P)
    subheading(doc, "Attendance & Acknowledgement", size=10, space_before=90)
    rows = [["#", "Name / Surname", "Company", "Role", "Signature"]]
    for i in range(1, 15):
        rows.append([str(i), "", "", "", ""])
    grid(doc, [400, 2800, 2200, 2200, 2264], rows, size=9, header_size=9,
         aligns=["c", None, None, None, None], total=USABLE_P)
    label_table(doc, [("Questions raised / actions", ""), ("TBT closed at (time)", ""),
                      ("EHS Officer signature", "")])

    heading(doc, "ANNEX C \u2014 ILLUMINATION, RCD & EARTHING VERIFICATION LOG")
    label_table(doc, [
        ("Date / Shift", ""),
        ("Pour location", ""),
        ("Verified by (Authorized Electrical Person / Night EHS Officer)", ""),
        ("Light meter \u2014 make / serial / calibration due", ""),
    ])
    subheading(doc, "C.1  Illumination (Lux) Survey", size=10, space_before=90)
    rows = [["#", "Measurement Point", "Required", "Measured (lux)", "Pass / Fail", "Action if Fail", "Time"]]
    pts = [
        ("Concrete discharge / pouring point (boom end / chute)", "\u2265 100 lux"),
        ("Working face \u2014 screeding / finishing area", "\u2265 100 lux"),
        ("Concrete pump set-up area and operator position", "\u2265 100 lux"),
        ("Rebar / formwork access pathways", "\u2265 50 lux"),
        ("Transit mixer route, turning and holding area", "\u2265 50 lux"),
        ("Mixer discharge / reversing position (banksman's view)", "\u2265 50 lux"),
        ("Emergency egress route to site gate", "\u2265 50 lux"),
        ("Muster point, ambulance station and first-aid point", "\u2265 50 lux"),
        ("Material storage, washout pit and waste area", "\u2265 50 lux"),
        ("Open edges, excavations and barricaded zones", "\u2265 50 lux"),
    ]
    for i, (p, req) in enumerate(pts, start=1):
        rows.append([str(i), p, req, "", "", "", ""])
    grid(doc, [380, 3200, 900, 1250, 950, 1900, 1284], rows, size=9, header_size=8.5,
         aligns=["c", None, "c", "c", "c", None, "c"], total=USABLE_P)
    subheading(doc, "C.2  Electrical Verification", size=10, space_before=90)
    rows = [["#", "Verification Item", "Result / Reading", "Pass / Fail", "Initials", "Time"]]
    for i, it in enumerate([
        "Pump truck body grounding connected to the station earth grid (connection point identified)",
        "Earthing continuity / resistance measured (record value and limit applied)",
        "RCD 1 \u2014 board/ID: ________  test button operated, tripped and reset",
        "RCD 2 \u2014 board/ID: ________  test button operated, tripped and reset",
        "RCD 3 \u2014 board/ID: ________  test button operated, tripped and reset",
        "All vibrator and tool cables elevated on insulated hangers, clear of water and concrete",
        "Enclosures IP-rated, dry, undamaged; no exposed conductors or taped repairs",
        "Light towers aimed downward \u2014 no glare at operator positions; no dark shadows in the work zone",
        "Battery-operated backup floodlights charged, tested and staged at the pour site",
        "Insulated rescue hook present, undamaged and within certification date",
        "Illuminated first-aid kit present at the pour site, stocked and sealed",
        "Fire extinguishers present at pump and boards, in date and accessible",
    ], start=1):
        rows.append([str(i), it, "", "", "", ""])
    grid(doc, [380, 4600, 1700, 1000, 1000, 1184], rows, size=9, header_size=8.5,
         aligns=["c", None, None, "c", "c", "c"], total=USABLE_P)
    rows = [["Re-check during shift (time)", "Re-check result", "EHS Officer signature"],
            ["", "", ""], ["", "", ""], ["", "", ""]]
    grid(doc, [3200, 3800, 2864], rows, size=9.5, header_size=9.5, total=USABLE_P)

    heading(doc, "ANNEX D \u2014 NIGHT SHIFT CREW REGISTER, REST & FITNESS DECLARATION")
    rows = [["#", "Name / Surname", "Company", "Role on Night Shift",
             "Rest Start / End (min. 8 h)", "Fitness Check (\u2713)", "Certificates Verified (\u2713)",
             "Signature"]]
    for i in range(1, 17):
        rows.append([str(i), "", "", "", "", "", "", ""])
    grid(doc, [340, 1880, 1240, 1760, 1520, 880, 1180, 1064], rows, size=8.5,
         header_size=8, aligns=["c", None, None, None, "c", "c", "c", None], total=USABLE_P)
    label_table(doc, [
        ("Confirmation", "I confirm that every person listed above is a dedicated "
                         "night-shift member, has had a minimum of 8 hours of rest before this shift, "
                         "has passed the pre-shift fitness check, holds the certificates required for "
                         "their role, and has attended the night Toolbox Talk."),
        ("Night EHS Officer \u2014 name / signature", ""),
        ("Night Shift In-Charge \u2014 name / signature", ""),
        ("Date / Time", ""),
    ])

    heading(doc, "ANNEX E \u2014 MANDATORY MINIMUM STANDARDS SUMMARY (SITE QUICK CARD)")
    rows = [["Parameter", "Mandatory Minimum"]]
    quick = [
        ("Illumination \u2014 discharge / pouring point", "100 lux (measured, recorded)"),
        ("Illumination \u2014 access pathways & transit routes", "50 lux (measured, recorded)"),
        ("Electrical protection", "30 mA RCD on all temporary outlets; test proved each shift"),
        ("Cable management", "Elevated on insulated hangers \u2014 never on wet ground or fresh concrete"),
        ("Pump grounding", "Body grounded directly to the station earth grid; verified and recorded"),
        ("Crew rotation", "100 % replacement; minimum 8 hours rest; no day\u2192night continuation"),
        ("Night EHS cover", "Dedicated, fully rested Night EHS Officer on site for the whole shift"),
        ("Handover", "Written, signed Day\u2192Night handover log every shift"),
        ("Briefing", "Night TBT every shift \u2014 approach distances, live zones, evacuation routes"),
        ("Barricading", "Rigid hard barricades + reflective tape + flashers (no soft tape)"),
        ("Rebar protection", "Approved mushroom caps on 100 % of exposed bar ends"),
        ("Access", "Single controlled entry point; PTW access control; illuminated signage"),
        ("Pump & boom", "Valid third-party inspection certificate before site entry"),
        ("Pump operator", "Valid third-party competency card"),
        ("Mixer drivers", "Valid heavy vehicle driving licence + third-party induction + fitness check"),
        ("Reversing", "Banksman with illuminated baton + physical wheel stoppers"),
        ("Spotter", "Dedicated spotter, torch and whistle, no other duties"),
        ("Boom clearance", "Never inside the marked envelope / restricted red zone; permit distance or OSHA/NFPA minimum, whichever is greater"),
        ("Wind limit", "Suspend boom operation above 32 km/h (or manufacturer's lower limit)"),
        ("Medical standby", "Ambulance + dedicated driver + certified Site Medic at the pour zone all shift"),
        ("Ambulance equipment", "Oxygen cylinder, spinal board, resuscitation kit, working emergency lights"),
        ("Egress", "Route to Al-Zubair General Hospital clear at all times"),
        ("Hi-vis", "Class 3 high-visibility attire \u2014 100 % of night personnel"),
        ("Concrete PPE", "Steel-toe rubber boots + heavy-duty chemical/rubber gloves + sealed goggles"),
        ("Rescue gear", "Insulated rescue hook + illuminated first-aid kit at the pour site"),
        ("Backup lighting", "Battery-operated floodlights staged, charged and tested"),
        ("Environmental", "Washout only in the designated pit; spill kit available; waste segregated (MOTAT)"),
        ("Authorisation", "Signed Night Pour Readiness Checklist + endorsed night PTW before any placement"),
    ]
    for a, b in quick:
        rows.append([a, b])
    grid(doc, [4000, 5864], rows, size=9.5, header_size=9.5, bolds=[0], total=USABLE_P)

    spacer(doc, 120)
    note_box(doc, "END OF DOCUMENT",
             f"{DOC_ID} Rev. {REV} \u2014 Method Statement: Night Concrete Pouring Operations, "
             "KAZ Power Plant Upgrade Project (Khur Al-Zubair, Basra, Iraq). Status: DRAFT "
             "FOR REVIEW. This document is a controlled record; verify you are using the "
             "current revision from the project EHS document register before every night "
             "pour.", fill=LBL_FILL)


# ======================================================================= BUILD
def make_shell(tmp_path):
    """Copy the house document, empty its body and replace the footer text."""
    shutil.copy(SHELL, tmp_path)
    doc = Document(tmp_path)
    body_el = doc.element.body
    for child in list(body_el):
        body_el.remove(child)
    sectPr = OxmlElement("w:sectPr")
    body_el.append(sectPr)
    doc.save(tmp_path)

    # rebuild footer part with the new document id text + PAGE / NUMPAGES fields
    zin = zipfile.ZipFile(tmp_path, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    zin.close()

    if "word/footer1.xml" in parts:
        f = parts["word/footer1.xml"].decode("utf-8")
        # strip existing paragraph content, keep root element and properties
        m = re.match(r"^(.*?<w:ftr\b[^>]*>)(.*)(</w:ftr>\s*)$", f, re.S)
        head, _old, tail = m.group(1), m.group(2), m.group(3)

        def fld(instr):
            return (
                '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
                '<w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
                f'<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
                '<w:fldChar w:fldCharType="separate"/></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
                '<w:t>1</w:t></w:r>'
                '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                '<w:b/><w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
                '<w:fldChar w:fldCharType="end"/></w:r>'
            )

        def txt(s):
            return ('<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
                    f'<w:color w:val="003B4A"/><w:sz w:val="16"/></w:rPr>'
                    f'<w:t xml:space="preserve">{s}</w:t></w:r>')

        para_xml = (
            '<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="3" '
            'w:color="B7CCCF"/></w:pBdr><w:spacing w:before="60" w:after="0"/>'
            '<w:jc w:val="center"/></w:pPr>'
            + txt(FOOTER_TXT) + fld("PAGE") + txt(" of ") + fld("NUMPAGES")
            + '</w:p>'
        )
        parts["word/footer1.xml"] = (head + para_xml + tail).encode("utf-8")

    # header part (new)
    header_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="3" '
        'w:color="B7CCCF"/></w:pBdr><w:tabs><w:tab w:val="right" w:pos="9864"/></w:tabs>'
        '<w:spacing w:before="0" w:after="40"/></w:pPr>'
        '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:b/>'
        '<w:color w:val="007E8C"/><w:sz w:val="16"/></w:rPr>'
        f'<w:t xml:space="preserve">{HEADER_LEFT}</w:t></w:r>'
        '<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>'
        '<w:color w:val="222A2E"/><w:sz w:val="16"/></w:rPr>'
        f'<w:t xml:space="preserve">\t{HEADER_RIGHT}</w:t></w:r>'
        '</w:p></w:hdr>'
    )
    parts["word/header1.xml"] = header_xml.encode("utf-8")

    # relationships: add header, keep footer
    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")
    if 'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"' not in rels:
        rels = rels.replace(
            "</Relationships>",
            '<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/header" Target="header1.xml"/>'
            "</Relationships>")
    parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")

    ct = parts["[Content_Types].xml"].decode("utf-8")
    if 'PartName="/word/header1.xml"' not in ct:
        ct = ct.replace(
            "</Types>",
            '<Override PartName="/word/header1.xml" ContentType="application/vnd.'
            'openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>')
    parts["[Content_Types].xml"] = ct.encode("utf-8")

    # core properties
    if "docProps/core.xml" in parts:
        core = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/'
            'metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/'
            'XMLSchema-instance">'
            '<dc:title>Method Statement - Night Concrete Pouring Operations</dc:title>'
            '<dc:subject>KAZ Power Plant Upgrade Project - Basra, Iraq</dc:subject>'
            '<dc:creator>Ahmed Mansoor (EHS Manager) - Siemens Energy</dc:creator>'
            '<cp:keywords>EHS; Method Statement; Night Work; Concrete Pouring; JSA; '
            'KAZ-EHS-MS-2026-004</cp:keywords>'
            '<cp:lastModifiedBy>Ahmed Mansoor</cp:lastModifiedBy>'
            '<cp:category>EHS Document</cp:category>'
            '</cp:coreProperties>'
        )
        parts["docProps/core.xml"] = core.encode("utf-8")

    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, parts[n])
        if "word/header1.xml" not in names:
            zout.writestr("word/header1.xml", parts["word/header1.xml"])
    return tmp_path


def link_header_footer(sec, first=False):
    """Attach header1.xml / footer1.xml to a section via raw sectPr manipulation."""
    sectPr = sec._sectPr
    # remove existing refs
    for tag in ("w:headerReference", "w:footerReference"):
        for el in sectPr.findall(qn(tag)):
            sectPr.remove(el)
    for tag, rid in (("w:headerReference", "rId100"), ("w:footerReference", "rId9")):
        el = OxmlElement(tag)
        el.set(qn("w:type"), "default")
        el.set(qn("r:id"), rid)
        sectPr.insert(0, el)


def main():
    make_shell(TMP)
    doc = Document(TMP)
    sec0 = doc.sections[0]
    setup_section(sec0, landscape=False)

    banner(doc)
    section_0(doc)
    section_1(doc)
    section_2(doc)
    section_3(doc)
    section_4(doc)
    section_5(doc)
    section_6(doc)
    section_7(doc)

    heading(doc, "SECTION 8: KEY EHS PERFORMANCE INDICATORS FOR NIGHT POURS")
    rows = [["#", "Indicator", "Target", "Monitored By / Frequency"]]
    kpis = [
        ("Readiness checklist completed and signed before every pour", "100 %", "Night EHS Officer / each shift"),
        ("Night TBT delivered with signed attendance", "100 %", "Night EHS Officer / each shift"),
        ("Signed shift handover logs", "100 %", "EHS Supervisors / each shift"),
        ("Crew members with \u2265 8 h recorded rest", "100 %", "Night EHS Officer / each shift"),
        ("Lux readings at or above 100 / 50 lux", "100 % of measurement points", "Night EHS Officer / each shift + mid-shift re-check"),
        ("RCD test-button proved and earthing verified", "100 %", "Authorized Electrical Person / each shift"),
        ("Exposed rebar ends capped", "100 %", "Supervisor + EHS / each shift"),
        ("Valid third-party certificates on site (pump, boom, operator, drivers)", "100 %", "EHS / before mobilization and each shift"),
        ("Standby ambulance + medic present at the pour zone for the whole shift", "100 %", "Night EHS Officer / continuous"),
        ("Egress route kept clear", "100 %", "Banksman + Security / continuous"),
        ("Class 3 hi-vis compliance", "100 %", "EHS / each shift"),
        ("Near misses and unsafe conditions reported", "\u2265 1 per 10 workers per month", "All personnel / daily report"),
        ("Recordable incidents during night pours", "Zero", "EHS Manager / monthly"),
        ("Stop-Work events raised and closed", "100 % closed with corrective action", "EHS / each event"),
    ]
    for i, k in enumerate(kpis, start=1):
        rows.append([str(i), k[0], k[1], k[2]])
    grid(doc, [400, 4900, 1800, 2764], rows, size=9.5, header_size=9.5,
         aligns=["c", None, "c", None], total=USABLE_P)

    section_9(doc)
    section_10(doc)

    # --- landscape section for the JSA tables -------------------------------
    sec1 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec1, landscape=True)
    section_10_landscape(doc)

    # --- back to portrait ---------------------------------------------------
    sec2 = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_section(sec2, landscape=False)
    section_11(doc)
    section_12(doc)
    section_13(doc)
    section_14(doc)
    section_15(doc)
    annexes(doc)

    for sec in doc.sections:
        link_header_footer(sec)

    # ensure rId9 (footer) exists; if the shell used a different id, fix it
    doc.save(TMP)
    fix_rel_ids(TMP)
    shutil.move(TMP, OUT)
    print("WROTE:", OUT)


def fix_rel_ids(path):
    """Guarantee header/footer rel ids, and strip the template's stale thumbnail
    and app.xml leftovers inherited from the shell document."""
    zin = zipfile.ZipFile(path, "r")
    names = zin.namelist()
    parts = {n: zin.read(n) for n in names}
    zin.close()

    # --- drop the inherited thumbnail preview (belongs to the PTW template) ---
    drop = {n for n in parts if "thumbnail" in n.lower()}
    for n in drop:
        parts.pop(n, None)
    root_rels = parts["_rels/.rels"].decode("utf-8")
    root_rels = re.sub(r'<Relationship[^>]*thumbnail[^>]*/>', "", root_rels, flags=re.I)
    parts["_rels/.rels"] = root_rels.encode("utf-8")

    # --- neutralise app.xml (template metadata) ---
    if "docProps/app.xml" in parts:
        parts["docProps/app.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/'
            'extended-properties" xmlns:vt="http://schemas.openxmlformats.org/'
            'officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Office Word</Application>'
            '<Company>Siemens Energy</Company>'
            '<Manager>Ahmed Mansoor (EHS Manager)</Manager>'
            '</Properties>'
        ).encode("utf-8")

    rels = parts["word/_rels/document.xml.rels"].decode("utf-8")

    def ensure(rels, rid, target, rel_type):
        if f'Id="{rid}"' in rels:
            rels = re.sub(rf'<Relationship Id="{rid}"[^>]*/>',
                          f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/>',
                          rels)
        else:
            rels = rels.replace(
                "</Relationships>",
                f'<Relationship Id="{rid}" Type="{rel_type}" Target="{target}"/></Relationships>')
        return rels

    rels = ensure(rels, "rId9", "footer1.xml",
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
    rels = ensure(rels, "rId100", "header1.xml",
                  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header")
    parts["word/_rels/document.xml.rels"] = rels.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in parts:
            zout.writestr(n, parts[n])


if __name__ == "__main__":
    main()
