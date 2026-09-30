#!/usr/bin/env python3
"""Build the EHS Night Shift Concrete Pouring Audit & Findings presentation.

Output: EHS_Night_Concrete_Pouring_Audit_Findings.pptx

The deck is intentionally built with native PowerPoint shapes and editable text
so that site names, dates, owners and photographs can be updated in PowerPoint.
"""
from __future__ import annotations

import os
from typing import Iterable, List, Optional, Sequence, Tuple

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT = os.path.join(REPO, "EHS_Night_Concrete_Pouring_Audit_Findings.pptx")
LOGO = os.path.join(HERE, "assets", "siemens_energy_logo.png")

# ---------------------------------------------------------------------------
# Palette: night charcoal, Siemens Energy cyan/teal and safety accents.
BG = "0D171B"
BG_ALT = "122328"
PANEL = "162A30"
PANEL_2 = "1B353B"
GRID = "28464C"
WHITE = "F5FAFB"
MUTED = "A9BEC3"
MUTED_2 = "6E8B91"
CYAN = "00B5E2"
TEAL = "007E8C"
TEAL_LIGHT = "33D0D2"
RED = "F04A4A"
RED_DARK = "6E2026"
AMBER = "F3C45D"
AMBER_DARK = "6C4C13"
GREEN = "4EC7A2"
PURPLE = "7D3FB2"

SLIDE_W = 13.333
SLIDE_H = 7.5
FONT = "Aptos"
FONT_DISPLAY = "Aptos Display"


def rgb(hex_value: str) -> RGBColor:
    return RGBColor.from_string(hex_value)


def set_bg(slide, color: str = BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def set_line(line, color: str = GRID, width: float = 1.0, dash=None, transparency: int = 0):
    line.color.rgb = rgb(color)
    line.width = Pt(width)
    if dash:
        line.dash_style = dash
    try:
        line.transparency = transparency
    except Exception:
        pass


def add_shape(slide, kind, x, y, w, h, fill: Optional[str] = None,
              line: Optional[str] = None, line_width: float = 1.0,
              radius_color: Optional[str] = None):
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(line_width)
    return shape


def add_rect(slide, x, y, w, h, fill: Optional[str] = None,
             line: Optional[str] = None, line_width: float = 1.0):
    return add_shape(slide, MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, w, h, fill, line, line_width)


def add_round_rect(slide, x, y, w, h, fill: Optional[str] = None,
                   line: Optional[str] = None, line_width: float = 1.0):
    return add_shape(slide, MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, line_width)


def add_text(slide, text: str, x: float, y: float, w: float, h: float,
             size: float = 12, color: str = WHITE, bold: bool = False,
             font: str = FONT, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
             margin: float = 0.0, italic: bool = False,
             char_spacing: Optional[int] = None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(margin)
    tf.margin_right = Inches(margin)
    tf.margin_top = Inches(margin)
    tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(0)
    p.space_before = Pt(0)
    run = p.add_run()
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = rgb(color)
    if char_spacing is not None:
        # python-pptx does not expose tracking, but the XML property is simple.
        from pptx.oxml.xmlchemy import OxmlElement
        rpr = run._r.get_or_add_rPr()
        spacing = OxmlElement("a:spc")
        spacing.set("val", str(char_spacing))
        rpr.append(spacing)
    return box


def add_rich_text(slide, runs: Sequence[Tuple[str, str, float, bool]],
                  x, y, w, h, align=PP_ALIGN.LEFT, valign=MSO_ANCHOR.TOP,
                  margin=0.0, line_spacing=None):
    """Add runs as [(text, color, size, bold), ...]."""
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Inches(margin)
    tf.vertical_anchor = valign
    p = tf.paragraphs[0]
    p.alignment = align
    if line_spacing:
        p.line_spacing = line_spacing
    for txt, color, size, bold in runs:
        run = p.add_run()
        run.text = txt
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = rgb(color)
    return box


def add_line(slide, x1, y1, x2, y2, color=GRID, width=1.0, dash=None):
    line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    set_line(line.line, color, width, dash)
    return line


def add_circle(slide, x, y, d, fill=None, line=None, width=1.0):
    return add_shape(slide, MSO_AUTO_SHAPE_TYPE.OVAL, x, y, d, d, fill, line, width)


def add_logo(slide, x=11.57, y=0.19, w=1.45, h=0.33):
    # Existing repo asset is used rather than recreating the mark. A small
    # cyan rule anchors it into the night-work header.
    if os.path.exists(LOGO):
        slide.shapes.add_picture(LOGO, Inches(x), Inches(y), width=Inches(w), height=Inches(h))
    else:
        add_round_rect(slide, x, y, w, h, None, CYAN, 1)
        add_text(slide, "SIEMENS ENERGY LOGO", x, y + 0.07, w, 0.18, 6.3, CYAN, True, align=PP_ALIGN.CENTER)


def add_header(slide, section: str, page: int, label: str = "EHS / NIGHT SHIFT"):
    # persistent top rail
    add_rect(slide, 0.0, 0.0, SLIDE_W, 0.055, CYAN)
    add_text(slide, label, 0.48, 0.18, 2.4, 0.20, 7.2, CYAN, True, char_spacing=90)
    add_logo(slide)
    add_line(slide, 0.48, 0.67, 12.86, 0.67, GRID, 0.8)
    add_text(slide, section, 0.48, 7.12, 3.5, 0.16, 6.8, MUTED_2, True, char_spacing=80)
    add_text(slide, f"SIEMENS ENERGY  /  {page:02d}", 10.53, 7.12, 2.35, 0.16, 6.8, MUTED_2, True, align=PP_ALIGN.RIGHT, char_spacing=60)


def add_kicker(slide, text, x=0.48, y=0.90, color=CYAN):
    add_text(slide, text.upper(), x, y, 5.9, 0.20, 8.0, color, True, char_spacing=100)


def add_title(slide, title, subtitle=None, y=1.18, w=11.3, size=25):
    add_text(slide, title, 0.48, y, w, 0.58, size, WHITE, True, FONT_DISPLAY)
    if subtitle:
        add_text(slide, subtitle, 0.50, y + 0.66, w - 0.3, 0.32, 10.5, MUTED, False)


def add_pill(slide, text, x, y, w, color=RED, text_color=WHITE, height=0.27):
    add_round_rect(slide, x, y, w, height, color, None)
    add_text(slide, text.upper(), x + 0.05, y + 0.065, w - 0.1, height - 0.08, 7.0, text_color, True, align=PP_ALIGN.CENTER, char_spacing=70)


def add_section_marker(slide, number, label, x=0.50, y=1.02):
    add_text(slide, f"{number:02d}", x, y, 0.43, 0.30, 14, CYAN, True, FONT_DISPLAY)
    add_text(slide, label.upper(), x + 0.55, y + 0.06, 3.0, 0.18, 7.5, MUTED, True, char_spacing=90)


def add_photo_placeholder(slide, x, y, w, h, label, tag="FIELD EVIDENCE"):
    # A deliberately obvious, editable placeholder for a site photograph.
    add_round_rect(slide, x, y, w, h, "162429", TEAL, 1.3)
    # diagonal crosshair frame
    add_line(slide, x + 0.18, y + 0.18, x + w - 0.18, y + h - 0.18, GRID, 0.7)
    add_line(slide, x + w - 0.18, y + 0.18, x + 0.18, y + h - 0.18, GRID, 0.7)
    add_line(slide, x + w / 2, y + 0.16, x + w / 2, y + h - 0.16, GRID, 0.6)
    add_line(slide, x + 0.16, y + h / 2, x + w - 0.16, y + h / 2, GRID, 0.6)
    add_circle(slide, x + w / 2 - 0.22, y + h / 2 - 0.22, 0.44, None, TEAL, 1.0)
    add_circle(slide, x + w / 2 - 0.055, y + h / 2 - 0.055, 0.11, CYAN, None)
    add_rect(slide, x + 0.16, y + 0.16, 1.02, 0.22, BG, None)
    add_text(slide, tag, x + 0.22, y + 0.205, 0.9, 0.11, 6.4, CYAN, True, char_spacing=55)
    add_text(slide, "[ INSERT PHOTO HERE ]", x + 0.22, y + h / 2 - 0.13, w - 0.44, 0.28, 13, WHITE, True, FONT_DISPLAY, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
    add_text(slide, label, x + 0.20, y + h - 0.34, w - 0.40, 0.18, 7.3, MUTED, False, align=PP_ALIGN.CENTER)


def add_photo_pair(slide, x, y, w, h, labels: Tuple[str, str]):
    gap = 0.18
    each = (w - gap) / 2
    add_photo_placeholder(slide, x, y, each, h, labels[0], "FIELD EVIDENCE A")
    add_photo_placeholder(slide, x + each + gap, y, each, h, labels[1], "FIELD EVIDENCE B")


def add_callout(slide, kicker, text, x, y, w, h, accent=AMBER, fill=PANEL):
    add_round_rect(slide, x, y, w, h, fill, None)
    add_rect(slide, x, y, 0.055, h, accent)
    add_text(slide, kicker.upper(), x + 0.22, y + 0.18, w - 0.40, 0.16, 7.2, accent, True, char_spacing=80)
    add_text(slide, text, x + 0.22, y + 0.40, w - 0.42, h - 0.52, 11.0, WHITE, True, FONT_DISPLAY)


def add_action_box(slide, text, x=0.50, y=5.83, w=6.65, h=0.82, accent=RED):
    add_round_rect(slide, x, y, w, h, RED_DARK if accent == RED else AMBER_DARK, None)
    add_text(slide, "ACTION REQUIRED", x + 0.22, y + 0.16, 1.40, 0.15, 7.1, accent, True, char_spacing=75)
    add_text(slide, text, x + 0.22, y + 0.36, w - 0.40, 0.30, 10.6, WHITE, True, FONT_DISPLAY)


def add_owner_strip(slide, owner="[Responsible Party]", deadline="[Closure Deadline]", x=0.50, y=6.82, w=6.65):
    add_text(slide, "OWNER", x, y, 0.45, 0.13, 6.6, MUTED_2, True, char_spacing=70)
    add_text(slide, owner, x + 0.52, y - 0.02, 2.35, 0.17, 7.7, WHITE, True)
    add_text(slide, "CLOSE", x + 3.28, y, 0.45, 0.13, 6.6, MUTED_2, True, char_spacing=70)
    add_text(slide, deadline, x + 3.80, y - 0.02, 2.85, 0.17, 7.7, WHITE, True)


def add_observation_slide(prs, num: int, title: str, finding: str, prompt_label: str,
                          prompt: str, action: str, photo_label: str,
                          risk: str = "HIGH", accent: str = RED,
                          owner: str = "[Responsible Party]", deadline: str = "[Closure Deadline]",
                          pair: Optional[Tuple[str, str]] = None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    add_header(slide, f"OBSERVATION {num:02d}  /  CORRECTIVE ACTION", num + 2)
    add_section_marker(slide, num, "FIELD OBSERVATION", 0.50, 0.91)
    add_pill(slide, risk, 10.75, 0.94, 1.18 if risk != "CRITICAL" else 1.34, accent)
    add_text(slide, title, 0.50, 1.43, 7.15, 0.54, 24, WHITE, True, FONT_DISPLAY)
    add_text(slide, "CONTROL GAP  •  NIGHT CONCRETE POURING OPERATION", 0.52, 2.08, 6.8, 0.17, 7.1, CYAN, True, char_spacing=55)

    # Observation block
    add_round_rect(slide, 0.50, 2.45, 6.65, 1.27, PANEL, None)
    add_text(slide, "FINDING", 0.73, 2.68, 1.0, 0.16, 7.1, CYAN, True, char_spacing=70)
    add_text(slide, finding, 0.73, 2.95, 6.06, 0.54, 13.1, WHITE, True, FONT_DISPLAY)

    # Question / risk block
    add_callout(slide, prompt_label, prompt, 0.50, 3.96, 6.65, 1.43, AMBER if accent == RED else accent, PANEL_2)
    # Photo(s)
    if pair:
        add_photo_pair(slide, 7.52, 1.43, 5.30, 3.96, pair)
    else:
        add_photo_placeholder(slide, 7.52, 1.43, 5.30, 3.96, photo_label)
    add_text(slide, "REPLACE WITH SITE PHOTO BEFORE ISSUE", 7.55, 5.54, 5.20, 0.15, 6.3, MUTED_2, True, align=PP_ALIGN.RIGHT, char_spacing=45)
    add_action_box(slide, action, 0.50, 5.83, 6.65, 0.82, accent)
    add_owner_strip(slide, owner, deadline)
    return slide


def add_cover(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    # Night grid and schematic concrete-pour motif.
    for x in [0.4, 2.0, 3.6, 5.2, 6.8, 8.4, 10.0, 11.6, 13.2]:
        add_line(slide, x, 0.0, x - 1.0, 7.5, GRID, 0.4)
    for y in [1.05, 2.2, 3.35, 4.5, 5.65, 6.8]:
        add_line(slide, 0, y, 13.33, y, GRID, 0.4)
    add_rect(slide, 0.0, 0.0, 0.08, 7.5, CYAN)
    add_logo(slide, 10.98, 0.45, 1.85, 0.43)
    add_text(slide, "SIEMENS ENERGY  /  EHS", 0.62, 0.55, 3.6, 0.20, 8.0, CYAN, True, char_spacing=100)
    add_text(slide, "AUDIT & FINDINGS", 0.62, 1.53, 4.0, 0.28, 12.0, AMBER, True, char_spacing=120)
    add_text(slide, "EHS Site Inspection Report –", 0.62, 1.96, 8.0, 0.52, 28, WHITE, True, FONT_DISPLAY)
    add_text(slide, "Night Concrete Pouring Operation", 0.62, 2.56, 8.7, 0.58, 28, WHITE, True, FONT_DISPLAY)
    add_text(slide, "Key Observations, Compliance Gaps & Corrective Actions", 0.65, 3.38, 7.0, 0.28, 12.5, MUTED, False)

    # Right-side visual card
    add_round_rect(slide, 8.83, 1.45, 3.37, 4.40, "14282F", TEAL, 1.0)
    add_circle(slide, 9.35, 2.04, 2.34, None, TEAL, 1.0)
    add_circle(slide, 9.78, 2.47, 1.48, None, CYAN, 1.2)
    add_circle(slide, 10.30, 2.99, 0.44, CYAN, None)
    add_line(slide, 10.52, 2.15, 10.52, 4.36, CYAN, 1.2)
    add_line(slide, 9.44, 3.25, 11.60, 3.25, CYAN, 1.2)
    add_line(slide, 9.55, 4.58, 11.48, 1.92, AMBER, 1.7)
    add_text(slide, "NIGHT / POUR", 9.22, 4.89, 2.58, 0.22, 11.0, WHITE, True, align=PP_ALIGN.CENTER, char_spacing=85)
    add_text(slide, "CONTROL VISIBILITY", 9.22, 5.22, 2.58, 0.16, 6.7, CYAN, True, align=PP_ALIGN.CENTER, char_spacing=65)
    # metadata band
    add_rect(slide, 0.62, 5.18, 7.28, 0.03, CYAN)
    add_text(slide, "DATE", 0.64, 5.45, 0.65, 0.15, 7.1, MUTED_2, True, char_spacing=75)
    add_text(slide, "[Date]", 0.64, 5.68, 1.45, 0.23, 11.0, WHITE, True)
    add_text(slide, "LOCATION", 2.23, 5.45, 0.95, 0.15, 7.1, MUTED_2, True, char_spacing=75)
    add_text(slide, "[Site Name]", 2.23, 5.68, 1.70, 0.23, 11.0, WHITE, True)
    add_text(slide, "PREPARED BY", 4.32, 5.45, 1.08, 0.15, 7.1, MUTED_2, True, char_spacing=75)
    add_text(slide, "Ahmed Al-Mansoury", 4.32, 5.68, 2.35, 0.23, 11.0, WHITE, True)
    add_text(slide, "EHS Manager", 4.32, 5.96, 2.35, 0.18, 8.0, MUTED, False)
    add_text(slide, "FOR MANAGEMENT REVIEW  •  INSERT FIELD PHOTOS BEFORE ISSUE", 0.64, 6.87, 7.8, 0.18, 7.0, MUTED_2, True, char_spacing=55)
    add_text(slide, "01 / 10", 11.13, 6.86, 1.07, 0.18, 7.0, MUTED_2, True, align=PP_ALIGN.RIGHT, char_spacing=80)
    return slide


def add_summary(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    add_header(slide, "SUMMARY / IMMEDIATE ACTION PLAN", 10)
    add_kicker(slide, "CLOSE THE GAP BEFORE THE NEXT POUR", 0.50, 0.95, AMBER)
    add_title(slide, "Summary & immediate action plan", "One accountable owner. One verified closure. No repeat exposure.", 1.28, 10.7, 25)

    # matrix
    x0, y0 = 0.50, 2.35
    widths = [0.42, 3.62, 1.20, 2.24, 1.88]
    headers = ["#", "OBSERVATION", "RISK", "RESPONSIBLE PARTY", "CLOSURE DEADLINE"]
    total = sum(widths)
    add_rect(slide, x0, y0, total, 0.43, TEAL, None)
    xx = x0
    for w, h in zip(widths, headers):
        add_text(slide, h, xx + 0.08, y0 + 0.13, w - 0.16, 0.13, 6.7, WHITE, True, align=PP_ALIGN.CENTER if h == "#" else PP_ALIGN.LEFT, char_spacing=50)
        xx += w
    rows = [
        ("01", "PPE compliance", "HIGH", "Site Management", "Before next entry"),
        ("02", "Safe access / movement", "HIGH", "Main Contractor", "Immediate"),
        ("03", "Concrete vibrator handling", "HIGH", "Concrete Subcontractor", "Before next pour"),
        ("04", "Pump outrigger clearance", "CRITICAL", "Pump Supervisor", "STOP / re-set now"),
        ("05", "Driver licensing & competency", "CRITICAL", "Transport Contractor", "Before site entry"),
        ("06", "Concrete waste disposal", "HIGH", "Environmental Lead", "Same shift"),
        ("07", "Lighting cable joints", "HIGH", "Electrical Contractor", "Before night work"),
    ]
    row_h = 0.47
    for i, row in enumerate(rows):
        y = y0 + 0.43 + i * row_h
        fill = PANEL if i % 2 == 0 else BG_ALT
        add_rect(slide, x0, y, total, row_h, fill, GRID, 0.5)
        xx = x0
        for j, (w, val) in enumerate(zip(widths, row)):
            color = WHITE
            bold = j in (0, 1, 2)
            if j == 2:
                color = RED if val == "CRITICAL" else AMBER
            add_text(slide, val, xx + 0.08, y + 0.16, w - 0.16, 0.17, 7.8 if j != 1 else 8.0, color, bold, align=PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT)
            xx += w
    # bottom management instruction
    add_round_rect(slide, 0.50, 6.36, 12.20, 0.54, RED_DARK, None)
    add_text(slide, "MANAGEMENT DECISION", 0.73, 6.54, 1.55, 0.13, 7.0, RED, True, char_spacing=70)
    add_text(slide, "Release the next night pour only after critical controls are verified and evidence is attached.", 2.45, 6.48, 9.85, 0.22, 11.0, WHITE, True, FONT_DISPLAY)
    return slide


def build():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    # Keep the cover first, then executive context, seven findings, and action plan.
    add_cover(prs)

    # Slide 2: Executive summary & context
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    add_header(slide, "EXECUTIVE SUMMARY / CONTEXT", 2)
    add_kicker(slide, "WHY THIS AUDIT MATTERS", 0.50, 0.95)
    add_title(slide, "Night work demands visible control", "Concrete pouring compresses people, plant and temporary works into a low-light operating window.", 1.28, 11.2, 25)

    # left scope/purpose blocks
    add_round_rect(slide, 0.50, 2.42, 4.05, 1.47, PANEL, None)
    add_text(slide, "SCOPE", 0.76, 2.70, 1.0, 0.16, 7.2, CYAN, True, char_spacing=85)
    add_text(slide, "Night shift concrete pouring activities inspection.", 0.76, 3.02, 3.46, 0.50, 16.0, WHITE, True, FONT_DISPLAY)
    add_round_rect(slide, 0.50, 4.08, 4.05, 1.47, PANEL, None)
    add_text(slide, "PURPOSE", 0.76, 4.36, 1.0, 0.16, 7.2, CYAN, True, char_spacing=85)
    add_text(slide, "Confirm EHS compliance, structural safety and personnel protection during night operations.", 0.76, 4.67, 3.46, 0.60, 13.0, WHITE, True, FONT_DISPLAY)

    # right: three control lenses
    add_text(slide, "CONTROL LENSES", 5.10, 2.42, 2.1, 0.16, 7.2, AMBER, True, char_spacing=85)
    lenses = [
        ("01", "PEOPLE", "PPE, competence & safe movement", CYAN),
        ("02", "PLANT", "Pump stability & tool handling", AMBER),
        ("03", "PLACE", "Lighting, waste & housekeeping", RED),
    ]
    yy = 2.77
    for n, head, body, col in lenses:
        add_circle(slide, 5.10, yy, 0.42, col, None)
        add_text(slide, n, 5.10, yy + 0.12, 0.42, 0.13, 7.5, BG, True, align=PP_ALIGN.CENTER)
        add_text(slide, head, 5.75, yy - 0.01, 1.35, 0.22, 12.5, WHITE, True, FONT_DISPLAY)
        add_text(slide, body, 7.00, yy + 0.02, 4.8, 0.18, 9.6, MUTED, False)
        add_line(slide, 5.10, yy + 0.60, 12.55, yy + 0.60, GRID, 0.7)
        yy += 0.88
    add_callout(slide, "CONTROL MESSAGE", "High-risk controls must be verified in the field — not assumed from the method statement.", 5.10, 5.67, 7.45, 0.92, AMBER, PANEL_2)
    return slide

    # unreachable, keeps type checkers happy


def add_all_findings(prs):
    add_observation_slide(prs, 1, "Personal Protective Equipment (PPE) Compliance",
        "Personnel were observed inside the active site without full mandatory PPE.",
        "KEY QUESTION", "Why is full PPE not being worn during hazardous night operations?",
        "100% mandatory PPE compliance for all personnel on site at all times.",
        "Personnel without PPE", "HIGH", RED, "Site Management", "Before next entry")
    add_observation_slide(prs, 2, "Access Routes Within Concrete Pouring Area",
        "No safe, unobstructed access route was provided for workers moving between work areas.",
        "KEY QUESTION", "How can workers navigate safely across the pouring area without a defined route?",
        "Establish clear, safe and illuminated walkways across the pouring zone immediately.",
        "Poor / missing access area", "HIGH", RED, "Main Contractor", "Immediate")
    add_observation_slide(prs, 3, "Manual Operation of Concrete Vibrator",
        "A hazard was observed during manual movement and handling of the vibrator while actively operating.",
        "KEY QUESTION", "Is manual movement of an operating vibrator without proper risk controls acceptable?",
        "Enforce safe handling procedures and ergonomics for vibrator operators.",
        "Concrete vibrator operation", "HIGH", RED, "Concrete Subcontractor", "Before next pour")
    add_observation_slide(prs, 4, "Concrete Pump Setup & Outrigger Clearance",
        "Required outrigger clearance was not achieved per the approved setup plan; setback from the pit edge was insufficient.",
        "RISK", "Potential equipment instability, overturning and structural collapse.",
        "Immediately relocate / re-set the pump with fully extended outriggers and safe setback from the excavation edge.",
        "Pump outrigger setup", "CRITICAL", RED, "Pump Supervisor", "STOP / re-set now")
    add_observation_slide(prs, 5, "Concrete Mixer Drivers’ Qualification",
        "None of the concrete mixer operators possessed a valid driving license or third-party competency certificate.",
        "CONTROL GAP", "Unverified drivers create an immediate authorization and road-vehicle interface exposure.",
        "Stop unauthorized drivers. Validate and collect third-party certificates and valid licenses before site entry.",
        "Mixer truck / verification check", "CRITICAL", RED, "Transport Contractor", "Before site entry")
    add_observation_slide(prs, 6, "Improper Concrete Waste Disposal",
        "Leftover concrete was repeatedly discharged directly onto the ground instead of designated sheets or disposal pits.",
        "ENVIRONMENTAL RISK", "Uncontrolled discharge contaminates the work area and defeats the approved waste route.",
        "Stop illegal dumping immediately. Use polythene sheets and designated wash-out areas for every load.",
        "Concrete dumped on ground", "HIGH", AMBER, "Environmental Lead", "Same shift",
        ("Dumped concrete", "Designated wash-out area"))
    add_observation_slide(prs, 7, "Temporary Lighting Cable Joint Safety",
        "Lighting cables contained exposed / temporary joints, creating severe electrical and trip hazards during night shift.",
        "RISK", "Low-light conditions amplify both electric shock and trip potential at the workface.",
        "Repair and properly insulate all joints using industrial-grade weather-proof connectors.",
        "Exposed cable joint", "HIGH", RED, "Electrical Contractor", "Before night work")


def build_deck():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    add_cover(prs)
    # Reuse the executive summary function body without needing a second PRS.
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide)
    add_header(slide, "EXECUTIVE SUMMARY / CONTEXT", 2)
    add_kicker(slide, "WHY THIS AUDIT MATTERS", 0.50, 0.95)
    add_title(slide, "Night work demands visible control", "Concrete pouring compresses people, plant and temporary works into a low-light operating window.", 1.28, 11.2, 25)
    add_round_rect(slide, 0.50, 2.42, 4.05, 1.47, PANEL, None)
    add_text(slide, "SCOPE", 0.76, 2.70, 1.0, 0.16, 7.2, CYAN, True, char_spacing=85)
    add_text(slide, "Night shift concrete pouring activities inspection.", 0.76, 3.02, 3.46, 0.50, 16.0, WHITE, True, FONT_DISPLAY)
    add_round_rect(slide, 0.50, 4.08, 4.05, 1.47, PANEL, None)
    add_text(slide, "PURPOSE", 0.76, 4.36, 1.0, 0.16, 7.2, CYAN, True, char_spacing=85)
    add_text(slide, "Confirm EHS compliance, structural safety and personnel protection during night operations.", 0.76, 4.67, 3.46, 0.60, 13.0, WHITE, True, FONT_DISPLAY)
    add_text(slide, "CONTROL LENSES", 5.10, 2.42, 2.1, 0.16, 7.2, AMBER, True, char_spacing=85)
    lenses = [("01", "PEOPLE", "PPE, competence & safe movement", CYAN), ("02", "PLANT", "Pump stability & tool handling", AMBER), ("03", "PLACE", "Lighting, waste & housekeeping", RED)]
    yy = 2.77
    for n, head, body, col in lenses:
        add_circle(slide, 5.10, yy, 0.42, col, None)
        add_text(slide, n, 5.10, yy + 0.12, 0.42, 0.13, 7.5, BG, True, align=PP_ALIGN.CENTER)
        add_text(slide, head, 5.75, yy - 0.01, 1.35, 0.22, 12.5, WHITE, True, FONT_DISPLAY)
        add_text(slide, body, 7.00, yy + 0.02, 4.8, 0.18, 9.6, MUTED, False)
        add_line(slide, 5.10, yy + 0.60, 12.55, yy + 0.60, GRID, 0.7)
        yy += 0.88
    add_callout(slide, "CONTROL MESSAGE", "High-risk controls must be verified in the field — not assumed from the method statement.", 5.10, 5.67, 7.45, 0.92, AMBER, PANEL_2)
    add_all_findings(prs)
    add_summary(prs)
    # Document properties help recipients identify the artifact in Office.
    prs.core_properties.title = "EHS Site Inspection Report – Night Concrete Pouring Operation"
    prs.core_properties.subject = "Key observations, compliance gaps and corrective actions"
    prs.core_properties.author = "Ahmed Al-Mansoury, EHS Manager"
    prs.core_properties.keywords = "EHS, night shift, concrete pouring, audit, findings, Siemens Energy"
    prs.save(OUT)
    print(f"Wrote {OUT} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    build_deck()
