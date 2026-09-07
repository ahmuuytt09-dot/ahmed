#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بناء "KAZ Project Weekly EHS Report.xlsx" بتصميم احترافي.

- 7 أعمدة = أيام الأسبوع، عناوينها تواريخ حقيقية بالإنجليزية ([$-409] لضمان
  اللغة الإنجليزية بغضّ النظر عن إعدادات ويندوز المحلية)
- محدّد أسبوع واحد (1..52) في ورقة Dashboard يقود كل الأوراق بالدوال
- نظام ألوان وخطوط موحّد، خلايا الإدخال صفراء والخلايا المحسوبة رمادية
- رسوم بيانية، تنسيق شرطي، تحقق من صحة الإدخال، وإعداد طباعة

الاستخدام:  python3 scripts/build_weekly_report.py <output.xlsx>
"""
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

# ------------------------------------------------------------ نظام التصميم
NAVY = "1F3864"
BLUE = "2E75B6"
STEEL = "8EA9DB"
LIGHT = "D9E2F3"
BAND = "F4F7FC"
GREY = "595959"
LINE = "BFBFBF"
INPUT = "FFF2CC"
INPUT_BD = "BF9000"
GREEN = "548235"
ORANGE = "BF8F00"
RED = "C00000"
WHITE = "FFFFFF"

F_TITLE = Font(name="Calibri", size=18, bold=True, color=WHITE)
F_SUB = Font(name="Calibri", size=10, bold=False, color="D6DCE5", italic=True)
F_SECT = Font(name="Calibri", size=11, bold=True, color=WHITE)
F_HDR = Font(name="Calibri", size=10, bold=True, color=WHITE)
F_ITEM = Font(name="Calibri", size=10, bold=False, color="262626")
F_ITEMB = Font(name="Calibri", size=10, bold=True, color=NAVY)
F_DATA = Font(name="Calibri", size=10, color="262626")
F_NOTE = Font(name="Calibri", size=8, italic=True, color=GREY)
F_CARDL = Font(name="Calibri", size=9, bold=True, color=WHITE)
F_CARDV = Font(name="Calibri", size=20, bold=True, color=NAVY)
F_LBL = Font(name="Calibri", size=10, bold=True, color=NAVY)

FILL_TITLE = PatternFill("solid", fgColor=NAVY)
FILL_SECT = PatternFill("solid", fgColor=BLUE)
FILL_HDR = PatternFill("solid", fgColor=NAVY)
FILL_SUBHDR = PatternFill("solid", fgColor=STEEL)
FILL_BAND = PatternFill("solid", fgColor=BAND)
FILL_LIGHT = PatternFill("solid", fgColor=LIGHT)
FILL_INPUT = PatternFill("solid", fgColor=INPUT)
FILL_WHITE = PatternFill("solid", fgColor=WHITE)

thin = Side(style="thin", color=LINE)
med = Side(style="medium", color=NAVY)
B_ALL = Border(left=thin, right=thin, top=thin, bottom=thin)
B_IN = Border(left=Side(style="thin", color=INPUT_BD),
              right=Side(style="thin", color=INPUT_BD),
              top=Side(style="thin", color=INPUT_BD),
              bottom=Side(style="thin", color=INPUT_BD))

C_CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
C_LEFT = Alignment(horizontal="left", vertical="center", indent=1)
C_RIGHT = Alignment(horizontal="right", vertical="center", indent=1)

# صيغة التاريخ الإنجليزية مفروضة عبر معرّف اللغة 409 (en-US)
FMT_DAY = '[$-409]ddd\\ dd\\-mmm'
FMT_LONG = '[$-409]dddd", "dd\\ mmmm\\ yyyy'
FMT_INT = '#,##0'
FMT_RATE = '0.00'

DASH = "Dashboard"
WK_NO, WK_START, WK_END = "$D$5", "$D$6", "$D$7"

DAYS = 7
FIRST = 2                      # العمود B = الاثنين
LAST = FIRST + DAYS - 1        # العمود H = الأحد
C_WEEK = LAST + 1              # I = Week Total
C_YTD = LAST + 2               # J = YTD


def cl(i):
    return get_column_letter(i)


# ------------------------------------------------------------------ أدوات
def sheet_setup(ws, tab, last_col, widths=None):
    ws.sheet_properties.tabColor = tab
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    for c, w in (widths or {}).items():
        ws.column_dimensions[c].width = w


def title_bar(ws, last_col, title, subtitle_formula=None):
    ws.merge_cells(f"A1:{cl(last_col)}1")
    c = ws["A1"]
    c.value = title
    c.font = F_TITLE
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    for i in range(1, last_col + 1):
        ws.cell(1, i).fill = FILL_TITLE

    ws.merge_cells(f"A2:{cl(last_col)}2")
    c = ws["A2"]
    c.value = subtitle_formula or ""
    c.font = F_SUB
    c.fill = FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18
    for i in range(1, last_col + 1):
        ws.cell(2, i).fill = FILL_TITLE
    ws.row_dimensions[3].height = 7


def banner():
    return (f'="WEEK "&TEXT({DASH}!{WK_NO},"00")&"   \u2022   "'
            f'&TEXT({DASH}!{WK_START},"[$-409]dd mmm yyyy")&"  \u2013  "'
            f'&TEXT({DASH}!{WK_END},"[$-409]dd mmm yyyy")')


def section(ws, row, last_col, text):
    ws.merge_cells(f"A{row}:{cl(last_col)}{row}")
    c = ws.cell(row, 1)
    c.value = text
    c.font = F_SECT
    c.fill = FILL_SECT
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[row].height = 20
    for i in range(1, last_col + 1):
        ws.cell(row, i).fill = FILL_SECT
        ws.cell(row, i).border = Border(bottom=med)


def day_header(ws, row, item_label="ITEM", totals=True):
    """صف عناوين: التسمية + 7 تواريخ إنجليزية + المجاميع."""
    c = ws.cell(row, 1, item_label)
    c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, C_CTR, B_ALL
    for i in range(DAYS):
        cell = ws.cell(row, FIRST + i)
        cell.value = f"={DASH}!{WK_START}+{i}"
        cell.number_format = FMT_DAY
        cell.font, cell.fill, cell.alignment, cell.border = F_HDR, FILL_HDR, C_CTR, B_ALL
    if totals:
        for col, lab in ((C_WEEK, "WEEK\nTOTAL"), (C_YTD, "YTD")):
            cell = ws.cell(row, col, lab)
            cell.font, cell.fill = F_HDR, PatternFill("solid", fgColor=BLUE)
            cell.alignment, cell.border = C_CTR, B_ALL
    ws.row_dimensions[row].height = 30


def data_row(ws, row, label, fmt=FMT_INT, band=False, bold=False,
             totals=True, ytd_input=True, zero=False):
    c = ws.cell(row, 1, label)
    c.font = F_ITEMB if bold else F_ITEM
    c.alignment, c.border = C_LEFT, B_ALL
    c.fill = FILL_LIGHT if bold else (FILL_BAND if band else FILL_WHITE)
    for i in range(DAYS):
        cell = ws.cell(row, FIRST + i)
        if zero:
            cell.value = 0
        cell.number_format = fmt
        cell.font, cell.alignment, cell.border = F_DATA, C_CTR, B_IN
        cell.fill = FILL_INPUT
    if totals:
        t = ws.cell(row, C_WEEK)
        t.value = f"=SUM({cl(FIRST)}{row}:{cl(LAST)}{row})"
        t.number_format = fmt
        t.font, t.alignment, t.border = F_ITEMB, C_CTR, B_ALL
        t.fill = FILL_LIGHT
        y = ws.cell(row, C_YTD)
        if ytd_input:
            y.fill, y.border = FILL_INPUT, B_IN
        else:
            y.fill, y.border = FILL_LIGHT, B_ALL
        y.number_format = fmt
        y.font, y.alignment = F_DATA, C_CTR
    ws.row_dimensions[row].height = 17


def formula_row(ws, row, label, per_day, week_f, ytd_f, fmt=FMT_RATE, bold=True):
    c = ws.cell(row, 1, label)
    c.font = F_ITEMB if bold else F_ITEM
    c.alignment, c.border, c.fill = C_LEFT, B_ALL, FILL_LIGHT
    for i in range(DAYS):
        col = cl(FIRST + i)
        cell = ws.cell(row, FIRST + i, per_day(col))
        cell.number_format = fmt
        cell.font, cell.alignment, cell.border = F_DATA, C_CTR, B_ALL
        cell.fill = FILL_BAND
    for col, f in ((C_WEEK, week_f), (C_YTD, ytd_f)):
        cell = ws.cell(row, col, f)
        cell.number_format = fmt
        cell.font, cell.alignment, cell.border = F_ITEMB, C_CTR, B_ALL
        cell.fill = FILL_LIGHT
    ws.row_dimensions[row].height = 17


def legend(ws, row, last_col, col0=1):
    ws.cell(row, col0, "Legend:").font = F_NOTE
    b = ws.cell(row, col0 + 1)
    b.fill, b.border = FILL_INPUT, B_IN
    ws.cell(row, col0 + 2, "manual input").font = F_NOTE
    d = ws.cell(row, col0 + 4)
    d.fill, d.border = FILL_LIGHT, B_ALL
    ws.cell(row, col0 + 5, "calculated \u2014 do not type").font = F_NOTE
    ws.row_dimensions[row].height = 14


WIDTHS = {"A": 42, **{cl(i): 12.5 for i in range(FIRST, LAST + 1)},
          cl(C_WEEK): 13, cl(C_YTD): 11}


# ------------------------------------------------------------- Dashboard
def build_dashboard(wb):
    ws = wb.create_sheet(DASH)
    LASTC = 13                                   # A فاصل + B..M محتوى
    sheet_setup(ws, NAVY, LASTC,
                {"A": 2.5, **{cl(i): 12.9 for i in range(2, LASTC + 1)}})

    ws.merge_cells(f"A1:{cl(LASTC)}1")
    c = ws["A1"]
    c.value = "KAZ PROJECT  \u2014  WEEKLY EHS REPORT"
    c.font, c.fill = F_TITLE, FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[1].height = 42
    ws.merge_cells(f"A2:{cl(LASTC)}2")
    c = ws["A2"]
    c.value = ('="Environment, Health & Safety   \u2022   Reporting year "'
               '&TEXT(DATE(2026,6,1),"[$-409]mmm yyyy")&" \u2013 "'
               '&TEXT(DATE(2027,5,31),"[$-409]mmm yyyy")')
    c.font, c.fill = F_SUB, FILL_TITLE
    c.alignment = Alignment(horizontal="left", vertical="center", indent=2)
    ws.row_dimensions[2].height = 20
    for r in (1, 2):
        for i in range(1, LASTC + 1):
            ws.cell(r, i).fill = FILL_TITLE
    ws.row_dimensions[3].height = 7

    section(ws, 4, LASTC, "REPORTING PERIOD")
    for r, lab, val, fmt, inp in (
            (5, "Report Week No.", 1, "0", True),
            (6, "Week Start", "=DATE(2026,6,1)+(D5-1)*7", FMT_LONG, False),
            (7, "Week End", "=D6+6", FMT_LONG, False)):
        ws.merge_cells(f"B{r}:C{r}")
        lc = ws.cell(r, 2, lab)
        lc.font, lc.alignment, lc.border, lc.fill = F_LBL, C_RIGHT, B_ALL, FILL_LIGHT
        ws.cell(r, 3).border = B_ALL
        ws.cell(r, 3).fill = FILL_LIGHT
        ws.merge_cells(f"D{r}:F{r}")
        vc = ws.cell(r, 4, val)
        vc.number_format = fmt
        vc.font = Font(name="Calibri", size=11, bold=True, color=NAVY)
        vc.alignment, vc.border = C_CTR, (B_IN if inp else B_ALL)
        vc.fill = FILL_INPUT if inp else FILL_BAND
        for k in (5, 6):
            ws.cell(r, k).border = B_IN if inp else B_ALL
            ws.cell(r, k).fill = FILL_INPUT if inp else FILL_BAND
        ws.row_dimensions[r].height = 21

    dv = DataValidation(type="whole", operator="between", formula1=1, formula2=52,
                        allow_blank=False, showErrorMessage=True,
                        errorTitle="Invalid week",
                        error="Enter a week number between 1 and 52.")
    ws.add_data_validation(dv)
    dv.add(ws["D5"])
    ws.merge_cells("H5:M5")
    n = ws.cell(5, 8, "\u25b6  Type the week number in the yellow cell \u2014 "
                      "every sheet updates its dates automatically.")
    n.font = Font(name="Calibri", size=9, italic=True, color=GREY)
    n.alignment = C_LEFT
    ws.row_dimensions[8].height = 7

    section(ws, 9, LASTC, "KEY FIGURES \u2014 THIS WEEK")
    cards = [("TOTAL MANHOURS", "='EHS KPI'!I8", FMT_INT, BLUE),
             ("RECORDABLE INJURIES", "=SUM('EHS KPI'!I13:I15)", "0", RED),
             ("FIRST AID CASES", "='EHS KPI'!I16", "0", ORANGE),
             ("NEAR MISSES", "='EHS KPI'!I23", "0", GREEN),
             ("AFR", "='EHS KPI'!I29", FMT_RATE, NAVY),
             ("TRIR", "='EHS KPI'!I30", FMT_RATE, NAVY)]
    for k, (lab, f, fmt, colr) in enumerate(cards):
        col = 2 + k * 2
        a, b = cl(col), cl(col + 1)
        ws.merge_cells(f"{a}10:{b}10")
        lc = ws.cell(10, col, lab)
        lc.font, lc.alignment = F_CARDL, C_CTR
        for i in (col, col + 1):
            ws.cell(10, i).fill = PatternFill("solid", fgColor=colr)
        ws.merge_cells(f"{a}11:{b}11")
        vc = ws.cell(11, col, f)
        vc.number_format, vc.font, vc.alignment = fmt, F_CARDV, C_CTR
        for i in (col, col + 1):
            ws.cell(11, i).fill, ws.cell(11, i).border = FILL_BAND, B_ALL
    ws.row_dimensions[10].height = 19
    ws.row_dimensions[11].height = 36
    ws.row_dimensions[12].height = 7

    section(ws, 13, LASTC, "TRENDS")
    kpi = wb["EHS KPI"]
    ch = BarChart()
    ch.type, ch.style, ch.title = "col", 10, "Manhours by day"
    ch.y_axis.title, ch.height, ch.width = "Hours", 8.2, 18.5
    ch.add_data(Reference(kpi, min_col=1, min_row=6, max_col=LAST, max_row=7),
                from_rows=True, titles_from_data=True)
    ch.set_categories(Reference(kpi, min_col=FIRST, max_col=LAST, min_row=5))
    ws.add_chart(ch, "B15")

    ch2 = BarChart()
    ch2.type, ch2.style, ch2.title = "col", 10, "EHS events \u2014 week total"
    ch2.y_axis.title, ch2.height, ch2.width, ch2.legend = "Count", 8.2, 18.5, None
    ch2.add_data(Reference(kpi, min_col=C_WEEK, min_row=12, max_row=19))
    ch2.set_categories(Reference(kpi, min_col=1, min_row=12, max_row=19))
    ws.add_chart(ch2, "H15")

    for col, lab in ((2, "Prepared by:"), (6, "Reviewed by:"), (10, "Date:")):
        ws.cell(33, col, lab).font = F_LBL
        for i in (col + 1, col + 2):
            ws.cell(33, i).border = Border(bottom=Side(style="thin", color=GREY))
    ws.row_dimensions[33].height = 22
    legend(ws, 35, LASTC, col0=2)
    ws.freeze_panes = "A4"
    return ws


# ------------------------------------------------------------- EHS KPI
def build_kpi(wb):
    ws = wb.create_sheet("EHS KPI")
    sheet_setup(ws, BLUE, C_YTD, WIDTHS)
    title_bar(ws, C_YTD, "EHS KPI", banner())

    section(ws, 4, C_YTD, "MANHOURS")
    day_header(ws, 5, "CATEGORY")
    data_row(ws, 6, "Siemens Energy")
    data_row(ws, 7, "Subcontractors", band=True)
    formula_row(ws, 8, "Total Manhours",
                lambda c: f"=SUM({c}6:{c}7)",
                f"=SUM({cl(C_WEEK)}6:{cl(C_WEEK)}7)",
                f"=SUM({cl(C_YTD)}6:{cl(C_YTD)}7)", fmt=FMT_INT)
    ws.row_dimensions[9].height = 8

    section(ws, 10, C_YTD, "RE-ACTIVE EVENTS")
    day_header(ws, 11, "EVENT TYPE")
    for i, e in enumerate(["Fatality", "Lost Time Case", "Restricted Work Case",
                           "Medical Treatment Case", "First Aid Case", "Damage",
                           "Travel Incident", "Environmental Event"]):
        data_row(ws, 12 + i, e, fmt="0", band=(i % 2 == 1))
    ws.row_dimensions[20].height = 8

    section(ws, 21, C_YTD, "PRO-ACTIVE EVENTS")
    day_header(ws, 22, "EVENT TYPE")
    for i, e in enumerate(["Near Miss", "Unsafe Condition / Act",
                           "Safety Observation"]):
        data_row(ws, 23 + i, e, fmt="0", band=(i % 2 == 1))
    ws.row_dimensions[26].height = 8

    section(ws, 27, C_YTD, "SAFETY PERFORMANCE INDICATORS")
    day_header(ws, 28, "INDICATOR")
    W, Y = cl(C_WEEK), cl(C_YTD)
    formula_row(ws, 29, "Accident Frequency Rate (AFR)",
                lambda c: f"=IF(SUM($B8:{c}8)=0,0,SUM($B12:{c}13)*200000/SUM($B8:{c}8))",
                f"=IF({W}8=0,0,SUM({W}12:{W}13)*200000/{W}8)",
                f"=IF({Y}8=0,0,SUM({Y}12:{Y}13)*200000/{Y}8)")
    formula_row(ws, 30, "Total Recordable Injury Rate (TRIR)",
                lambda c: f"=IF(SUM($B8:{c}8)=0,0,SUM($B13:{c}15)*200000/SUM($B8:{c}8))",
                f"=IF({W}8=0,0,SUM({W}13:{W}15)*200000/{W}8)",
                f"=IF({Y}8=0,0,SUM({Y}13:{Y}15)*200000/{Y}8)")
    formula_row(ws, 31, "Environmental Frequency Rate (EFR)",
                lambda c: f"=IF(SUM($B8:{c}8)=0,0,SUM($B19:{c}19)*1000/SUM($B8:{c}8))",
                f"=IF({W}8=0,0,{W}19*1000/{W}8)",
                f"=IF({Y}8=0,0,{Y}19*1000/{Y}8)")

    ws.conditional_formatting.add(f"{cl(FIRST)}29:{Y}31", ColorScaleRule(
        start_type="min", start_color="C6EFCE",
        mid_type="percentile", mid_value=50, mid_color="FFEB9C",
        end_type="max", end_color="FFC7CE"))

    legend(ws, 33, C_YTD)
    ws.freeze_panes = "B6"
    return ws


# --------------------------------------------------- الأوراق الجدولية العامة
def build_grid(wb, name, tab, title, section_title, items, header="ITEM",
               fmt=FMT_INT, chart_title=None, units=None):
    ws = wb.create_sheet(name)
    widths = dict(WIDTHS)
    if units:
        widths[cl(C_YTD + 1)] = 9
    sheet_setup(ws, tab, C_YTD + (1 if units else 0), widths)
    title_bar(ws, C_YTD + (1 if units else 0), title, banner())
    section(ws, 4, C_YTD + (1 if units else 0), section_title)
    day_header(ws, 5, header)
    if units:
        u = ws.cell(5, C_YTD + 1, "UNIT")
        u.font, u.fill, u.alignment, u.border = F_HDR, FILL_SUBHDR, C_CTR, B_ALL
    for i, it in enumerate(items):
        r = 6 + i
        data_row(ws, r, it if not units else it[0], fmt=fmt, band=(i % 2 == 1))
        if units:
            uc = ws.cell(r, C_YTD + 1, it[1])
            uc.font, uc.alignment, uc.border = F_NOTE, C_CTR, B_ALL
            uc.fill = FILL_BAND
    last = 5 + len(items)
    legend(ws, last + 2, C_YTD)

    if chart_title:
        ch = BarChart()
        ch.type, ch.style, ch.title = "col", 10, chart_title
        ch.height, ch.width, ch.legend = 8, 24, None
        ch.add_data(Reference(ws, min_col=C_WEEK, min_row=6, max_row=last))
        ch.set_categories(Reference(ws, min_col=1, min_row=6, max_row=last))
        ws.add_chart(ch, f"A{last + 4}")
    ws.freeze_panes = "B6"
    return ws


# ------------------------------------------------------------- CAPA List
def build_capa(wb):
    ws = wb.create_sheet("CAPA List")
    cols = [("CAPA No.", 12), ("Date of Detection", 17), ("Location", 18),
            ("Unsafe Situation / Behaviour", 40), ("Action to Take", 40),
            ("Responsible", 18), ("Deadline", 14), ("Status", 14)]
    sheet_setup(ws, ORANGE, len(cols),
                {cl(i + 1): w for i, (_, w) in enumerate(cols)})
    title_bar(ws, len(cols), "CAPA REGISTER", banner())
    section(ws, 4, len(cols), "CORRECTIVE & PREVENTIVE ACTIONS")
    for i, (h, _) in enumerate(cols):
        c = ws.cell(5, i + 1, h)
        c.font, c.fill, c.alignment, c.border = F_HDR, FILL_HDR, C_CTR, B_ALL
    ws.row_dimensions[5].height = 28

    n = 30
    for r in range(6, 6 + n):
        for i in range(len(cols)):
            c = ws.cell(r, i + 1)
            c.font, c.border, c.alignment = F_DATA, B_IN, C_LEFT
            c.fill = FILL_INPUT
        ws.cell(r, 1).value = f"DOF{r - 5:04d}"
        ws.cell(r, 1).alignment = C_CTR
        ws.cell(r, 1).fill = FILL_LIGHT
        ws.cell(r, 1).font = F_ITEMB
        ws.cell(r, 1).border = B_ALL
        for i in (1, 6):
            ws.cell(r, i + 1).number_format = '[$-409]dd\\-mmm\\-yyyy'
            ws.cell(r, i + 1).alignment = C_CTR
        ws.cell(r, 8).alignment = C_CTR
        ws.row_dimensions[r].height = 18

    dv = DataValidation(type="list", formula1='"OPEN,IN PROGRESS,CLOSED"',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Invalid status",
                        error="Choose OPEN, IN PROGRESS or CLOSED.")
    ws.add_data_validation(dv)
    dv.add(f"H6:H{5 + n}")

    rng = f"H6:H{5 + n}"
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"OPEN"'],
        fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"IN PROGRESS"'],
        fill=PatternFill("solid", fgColor="FFEB9C"),
        font=Font(color="9C6500", bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(
        operator="equal", formula=['"CLOSED"'],
        fill=PatternFill("solid", fgColor="C6EFCE"),
        font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(f"G6:G{5 + n}", CellIsRule(
        operator="lessThan", formula=["TODAY()"],
        fill=PatternFill("solid", fgColor="FFC7CE"),
        font=Font(color="9C0006", bold=True)))

    ws.auto_filter.ref = f"A5:{cl(len(cols))}{5 + n}"
    ws.freeze_panes = "A6"
    return ws


# ==================================================================== main
def main(dst):
    wb = Workbook()
    wb.remove(wb.active)

    build_kpi(wb)
    build_grid(wb, "EHS Communication", GREEN,
               "EHS COMMUNICATION & TRAINING", "ACTIVITIES",
               ["Toolbox Talk", "EHS Meeting", "Site Induction Training",
                "Site Inspection", "EHS Award", "SWAT (DIV/BU)",
                "Work Permit Issued", "Onsite EHS Induction"],
               header="ACTIVITY", chart_title="Activities \u2014 week total")
    build_grid(wb, "EHS Incidents", RED,
               "EHS INCIDENTS", "RE-ACTIVE & PRO-ACTIVE EVENTS",
               ["Fatality", "Lost Time Case", "Restricted Work Case",
                "Medical Treatment Case", "First Aid Case", "Damage",
                "Travel Incident", "Environmental Event", "Near Miss",
                "Unsafe Condition / Act"],
               header="EVENT TYPE", fmt="0",
               chart_title="Events \u2014 week total")
    build_capa(wb)
    build_grid(wb, "Environmental", "1F6F3F",
               "ENVIRONMENTAL PERFORMANCE", "CONSUMPTION & WASTE",
               [("Water Consumption", "m\u00b3"), ("Electricity Consumption", "kWh"),
                ("Domestic Waste", "kg"), ("Plastic Waste", "kg"),
                ("Metal Waste", "kg"), ("Paper Waste", "kg"),
                ("Glass Waste", "kg"), ("Hazardous Waste", "kg")],
               header="PARAMETER", chart_title="Consumption & waste \u2014 week total",
               units=True)

    build_dashboard(wb)
    wb.move_sheet(DASH, offset=-len(wb.sheetnames) + 1)
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(dst)
    print("OK ->", dst)


if __name__ == "__main__":
    main(sys.argv[1])
