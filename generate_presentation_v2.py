#!/usr/bin/env python3
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Colors Siemens Energy
DEEP_PETROL = RGBColor(0x0E, 0x2F, 0x3E)
DARK_PETROL = RGBColor(0x08, 0x24, 0x30)
TEAL = RGBColor(0x00, 0x99, 0x99)
TEAL_LIGHT = RGBColor(0x00, 0xBF, 0xBF)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xE6, 0xE6, 0xE6)
GREY2 = RGBColor(0xCC, 0xCC, 0xCC)
AMBER = RGBColor(0xFF, 0xB9, 0x00)
RED = RGBColor(0xE3, 0x06, 0x13)
CARD_BG = RGBColor(0x14, 0x3D, 0x4E)
CARD_BG2 = RGBColor(0x1A, 0x4A, 0x5C)

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)

def set_bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def shape(slide, l, t, w, h, fill=None, border=None, border_w=Pt(1)):
    s = slide.shapes.add_shape(1, l, t, w, h)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if border:
        s.line.color.rgb = border
        s.line.width = border_w
    else:
        s.line.fill.background()
    return s

def text_box(slide, l, t, w, h, txt, size=11, bold=False, color=WHITE, align=PP_ALIGN.LEFT, name="Calibri"):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = name
    p.alignment = align
    return tb

def add_para(tf, txt, size=10, bold=False, color=WHITE, space=Pt(4)):
    p = tf.add_paragraph()
    p.text = txt
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Calibri"
    p.space_after = space
    return p

slides_info = [
    {
        "id": "01",
        "title": "01 - UNAUTHORIZED WORK - WITHOUT PERMIT TO WORK (PTW)",
        "arabic": "أعمال بدون تصريح عمل",
        "obs": [
            "• Location: [Add Area] - Date: [Add Date]",
            "• Activity: Excavation / Mechanical / Hot Work without approved PTW",
            "• Violation: No TRA review, no isolation, no SIMOPS check",
            "• PTW logbook not maintained, permit not displayed at work front",
            "• Reference: SE EHS General Requirements Sec 5.3, KAZ EHS Plan Sec 7.2"
        ],
        "risk": "CRITICAL",
        "risk_text": "Fatality, fire/explosion, loss of control. Breach of ISO 45001 Cl 8.1.2. Authority may issue prohibition notice.",
        "action": [
            "1. STOP WORK immediately - secure area",
            "2. Re-training PTW for all supervisors in 24h",
            "3. Issue NCR Level 1 + warning letter",
            "4. Daily PTW audit for 14 days",
            "5. 3-strikes removal policy"
        ],
        "photos": ["PHOTO 1: Work activity without PTW display", "PHOTO 2: PTW Logbook / Missing permit at site"]
    },
    {
        "id": "02",
        "title": "02 - SITE ACCESS & INDUCTION - WITHOUT EHS INDUCTION",
        "arabic": "دخول موقع بدون تدريب السلامة",
        "obs": [
            "• No. of persons: [8-12] found working without induction card",
            "• Gate register vs actual headcount mismatch",
            "• No induction sticker / QR verification",
            "• Violation: KAZ EHS Plan Sec 6.1, Local labor law",
            "• System failure: Security-EHS interface weak"
        ],
        "risk": "CRITICAL",
        "risk_text": "Workers unaware of hazards, emergency assembly, life-saving rules. Risk of unsafe acts + regulatory stoppage. ISO 45001 Cl 7.2/7.3",
        "action": [
            "1. Remove all non-inducted personnel now",
            "2. 100% gate check - Security + EHS",
            "3. Submit manpower list vs induction in 12h",
            "4. QR/Biometric tracking system",
            "5. Back-charge re-induction cost"
        ],
        "photos": ["PHOTO 1: Workers without induction cards", "PHOTO 2: Gate register / Induction cards check"]
    },
    {
        "id": "03",
        "title": "03 - PPE NON-COMPLIANCE - FAILURE TO WEAR MANDATORY PPE",
        "arabic": "عدم ارتداء معدات الحماية الشخصية",
        "obs": [
            "• PPE missing: Helmet chin strap, safety glasses, gloves, hi-vis vest",
            "• Location: [Add location] - Work: [Add task]",
            "• Supervision present but not enforcing",
            "• PPE available in store - behavioral failure",
            "• Violation: SE Life Saving Rule #3, EN 397/166"
        ],
        "risk": "HIGH",
        "risk_text": "Head injury, eye injury, hand cuts, struck-by. Weak safety culture. Increases LTIFR/TRIR.",
        "action": [
            "1. On-spot correction + TBT with photos",
            "2. Issue Observation Card + violation notice",
            "3. Daily PPE checks - start of shift + random",
            "4. Link PPE to KPI & payment milestone",
            "5. Supervisor accountability - 3 strikes"
        ],
        "photos": ["PHOTO 1: PPE violation - wide shot", "PHOTO 2: PPE violation - close-up + supervisor present"]
    },
    {
        "id": "04",
        "title": "04 - EXCAVATION SAFETY - INADEQUATE HARD BARRICADES",
        "arabic": "أعمال حفر بدون حواجز صلبة",
        "obs": [
            "• Depth: [1.2m - 2.5m] - No hard barricade, only loose tape (prohibited)",
            "• No edge protection, no ladder for access/egress",
            "• Spoil pile within 1m of edge - surcharge risk",
            "• No reflective markers / lighting at night",
            "• Violation: OSHA 1926 Subpart P, SE-GEN-EHS-007"
        ],
        "risk": "CRITICAL",
        "risk_text": "Fatal fall, cave-in engulfment, vehicle fall. High-risk activity - MAH. Authority prohibition risk.",
        "action": [
            "1. IMMEDIATE hard barricade - scaffold tubes / concrete",
            "2. Ladders every 7.5m + spoil 1.5m away",
            "3. Excavation permit re-validation + geotech check",
            "4. Reflective delineators + solar lights",
            "5. Competent person checklist daily"
        ],
        "photos": ["PHOTO 1: Excavation without hard barricade - full view", "PHOTO 2: Spoil pile too close + no ladder - close view"]
    },
    {
        "id": "05",
        "title": "05 - DIRECTIVES NON-COMPLIANCE - FAILURE TO FOLLOW EHS INSTRUCTIONS",
        "arabic": "عدم الالتزام بتعليمات السلامة",
        "obs": [
            "• Date of instruction: [Add date] - Instruction: [Add details]",
            "• TBT not conducted / no attendance sheet",
            "• Previous NCRs closed without evidence - falsification",
            "• Verbal + written instructions ignored - documented in daily report",
            "• Violation: ISO 45001 Cl 7.3 & 8.1 - Willful disregard"
        ],
        "risk": "HIGH",
        "risk_text": "Erosion of EHS authority, culture failure. If directives optional, all critical controls (LOTO, WAH) will fail.",
        "action": [
            "1. Formal letter to Subcontractor MD - written commitment",
            "2. Daily EHS Directives Log - sign-off PM + EHS",
            "3. Leadership stand-down with all foremen",
            "4. Link to contractual penalty",
            "5. Weekly scorecard in PM meeting"
        ],
        "photos": ["PHOTO 1: Evidence of ignored instruction (e.g., open excavation)", "PHOTO 2: TBT register missing / NCR closure without evidence"]
    },
    {
        "id": "06",
        "title": "06 - TRAFFIC & SIGNAGE - MISSING / DAMAGED SAFETY SIGNS",
        "arabic": "إشارات مرورية وتحذيرية مفقودة",
        "obs": [
            "• Missing: Speed limit (15 km/h), parking, pedestrian walkway signs",
            "• No flagmen at intersections / blind spots",
            "• Speeding observed: 30 km/h in 15 km/h zone - radar check",
            "• Pedestrian routes not marked / segregated",
            "• Violation: Project TMP, ISO 39001"
        ],
        "risk": "HIGH",
        "risk_text": "Vehicle-pedestrian collision, equipment damage, fatality - Top 3 fatality cause. Emergency access blocked.",
        "action": [
            "1. Reinstate all signs - retro-reflective grade",
            "2. Pedestrian routes with Jersey barriers + green paint",
            "3. Deploy flagmen + radar checks 2x daily",
            "4. Traffic safety campaign + driver re-induction",
            "5. EHS + Logistics joint daily inspection"
        ],
        "photos": ["PHOTO 1: Missing / damaged speed limit sign", "PHOTO 2: No pedestrian segregation + speeding area"]
    },
    {
        "id": "07",
        "title": "07 - WASTE MANAGEMENT - NO SEGREGATION & LABELING",
        "arabic": "عدم فرز النفايات ووضع ملصقات",
        "obs": [
            "• Waste mixed: Hazardous + recyclable + general in same skip",
            "• No color-coding, no labeling, no MSDS for haz waste",
            "• Oil rags / filters in general waste - fire + env violation",
            "• Location: [Add waste area location]",
            "• Violation: ISO 14001 Cl 8.1, Project Waste Plan, EPA"
        ],
        "risk": "HIGH",
        "risk_text": "Soil/groundwater pollution, fire from incompatible mixing, regulatory fine, reputational damage, sustainability KPI failure.",
        "action": [
            "1. Provide color-coded skips: Black-General, Green-Recyclable, Red-Haz",
            "2. Clean-up & correct segregation under EHS supervision",
            "3. Training on hazardous waste + manifest",
            "4. Appoint Waste Focal Point",
            "5. Weekly waste audit + disposal manifests"
        ],
        "photos": ["PHOTO 1: Mixed waste in one container - overview", "PHOTO 2: Oil-contaminated rags in general waste - close-up"]
    },
    {
        "id": "08",
        "title": "08 - EQUIPMENT INTEGRITY - OIL & FUEL LEAKS FROM MACHINERY",
        "arabic": "تسريب زيوت ووقود من المعدات الثقيلة",
        "obs": [
            "• Equipment: Excavator CAT 320 / Loader [Add ID] - active leak",
            "• Leak type: Hydraulic oil / fuel - dripping observed",
            "• No drip trays, soil contamination ~5 sqm under parking",
            "• No maintenance records, checklist falsified - same signature 7 days",
            "• Violation: PUWER, SE Equipment Standard, ISO 14001"
        ],
        "risk": "CRITICAL",
        "risk_text": "Fire/explosion (fuel on hot surface), slip, soil pollution, equipment failure loss of control. Maintenance system breakdown.",
        "action": [
            "1. Ground equipment - Tag DO NOT USE, remove keys",
            "2. Excavate contaminated soil as haz waste + absorbents",
            "3. Submit maintenance history + 3rd party cert in 48h",
            "4. Daily checklist verified by SE + photo of no leak",
            "5. Drip trays + spill kits at all parking zones"
        ],
        "photos": ["PHOTO 1: Active oil leak under equipment - with drip", "PHOTO 2: Soil contamination + missing drip tray - parking area"]
    },
    {
        "id": "09",
        "title": "09 - HAZARDOUS MATERIAL - DIESEL STORAGE WITHOUT BUNDING",
        "arabic": "خزن الديزل بدون حوض احتواء ثانوي",
        "obs": [
            "• Diesel storage: 2000L (2x IBC) - no secondary containment / bunding",
            "• No 110% capacity bund, no impervious base, bare ground",
            "• No fire extinguisher, no HAZCHEM sign, no spill kit",
            "• No earthing/bonding for fuel transfer - static risk",
            "• Ignition source 10m away (welding) - Violation HSG 176, SE-ENV-003"
        ],
        "risk": "CRITICAL",
        "risk_text": "Major fire/explosion, catastrophic env release, soil pollution, prosecution, MAH pool fire. Immediate danger to life.",
        "action": [
            "1. STOP fueling - 15m exclusion + fire watch",
            "2. Compliant bund: 110%, concrete floor, roofed, ventilated, 15m from drains",
            "3. Install 2x 9kg DCP + foam, HAZCHEM, 150L spill kit, earthing",
            "4. Fuel Handling SOP + emergency drill",
            "5. Location approved by SE EHS + risk assessment"
        ],
        "photos": ["PHOTO 1: Diesel tanks without bunding - overview", "PHOTO 2: Bare ground + no fire extinguisher / signage - close-up"]
    },
    {
        "id": "10",
        "title": "10 - RECURRING VIOLATIONS - PERSISTENT PPE VIOLATIONS AFTER WARNINGS",
        "arabic": "مخالفات متكررة رغم التحذيرات",
        "obs": [
            "• Same workgroup - 3 warnings + 5 Obs cards + 2 NCRs in 14 days",
            "• Trend: 60% increase (5 to 12 per week) - Daily EHS Reports",
            "• Intentional non-compliance, not lack of awareness",
            "• Subcontractor not enforcing disciplinary procedure per contract",
            "• Indicates normalization of deviance - cultural failure"
        ],
        "risk": "CRITICAL",
        "risk_text": "Precursor to fatality (Heinrich triangle). Total failure of safety leadership. If PPE (last defense) ignored, higher controls will fail. Client confidence at risk.",
        "action": [
            "1. Contractual penalty + stop work repeat crew + PM presents plan",
            "2. Remove/replace supervision - CVs + competency interview by SE",
            "3. BBS program - daily observations + feedback + recognition",
            "4. Daily management safety walks - SE PM + Sub PM",
            "5. Prepare partial termination / back-charge if no improvement in 7 days"
        ],
        "photos": ["PHOTO 1: Repeat PPE violation - same crew", "PHOTO 2: Trend chart / NCR log showing increase"]
    },
]

def create_boxed_slide(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DEEP_PETROL)

    # Header
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.9), fill=DARK_PETROL)
    shape(slide, Inches(0), Inches(0.9), Inches(13.33), Pt(4), fill=TEAL)
    # Slide ID badge
    shape(slide, Inches(0.2), Inches(0.15), Inches(0.7), Inches(0.6), fill=TEAL)
    text_box(slide, Inches(0.2), Inches(0.15), Inches(0.7), Inches(0.6), f"{data['id']}", size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Title
    text_box(slide, Inches(1.0), Inches(0.1), Inches(8.5), Inches(0.35), data['title'], size=16, bold=True, color=WHITE)
    text_box(slide, Inches(1.0), Inches(0.45), Inches(8.5), Inches(0.3), data['arabic'], size=11, bold=False, color=TEAL_LIGHT)

    # Risk badge
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_text_color = WHITE if data['risk'] == "CRITICAL" else DARK_PETROL
    shape(slide, Inches(10.5), Inches(0.15), Inches(2.6), Inches(0.6), fill=risk_color)
    text_box(slide, Inches(10.5), Inches(0.15), Inches(2.6), Inches(0.6), f"RISK: {data['risk']} | خطر: {data['risk']}", size=13, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    # === BOX 1: NON-COMPLIANCE (Left top, 5.5w x 2.6h) ===
    shape(slide, Inches(0.2), Inches(1.2), Inches(5.5), Inches(2.8), fill=CARD_BG, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(0.3), Inches(1.2), Inches(5.3), Inches(0.3), "📋 BOX 1: NON-COMPLIANCE DETAILS / خانة تفاصيل المخالفة", size=10, bold=True, color=TEAL_LIGHT)
    tf_box = slide.shapes.add_textbox(Inches(0.3), Inches(1.5), Inches(5.3), Inches(2.4)).text_frame
    tf_box.word_wrap = True
    for i, line in enumerate(data['obs']):
        if i==0:
            p = tf_box.paragraphs[0]
        else:
            p = tf_box.add_paragraph()
        p.text = line
        p.font.size = Pt(9.5)
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"
        p.space_after = Pt(3)

    # === BOX 2: PHOTO EVIDENCE 1 (Right top) ===
    shape(slide, Inches(5.9), Inches(1.2), Inches(3.6), Inches(2.8), fill=RGBColor(0x0A,0x22,0x2B), border=GREY2, border_w=Pt(1))
    text_box(slide, Inches(6.0), Inches(1.2), Inches(3.4), Inches(0.3), "📷 BOX 2: PHOTO EVIDENCE 1 / صورة 1", size=10, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(6.0), Inches(1.6), Inches(3.4), Inches(1.8), f"{data['photos'][0]}\n\n[ INSERT PHOTO HERE ]\n\nFig {data['id']}.1\n\nAdd: Date, Time, Location, GPS if possible", size=9, bold=False, color=GREY2, align=PP_ALIGN.CENTER)

    # === BOX 3: PHOTO EVIDENCE 2 (Far right top) ===
    shape(slide, Inches(9.7), Inches(1.2), Inches(3.4), Inches(2.8), fill=RGBColor(0x0A,0x22,0x2B), border=GREY2, border_w=Pt(1))
    text_box(slide, Inches(9.8), Inches(1.2), Inches(3.2), Inches(0.3), "📷 BOX 3: PHOTO EVIDENCE 2 / صورة 2", size=10, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(9.8), Inches(1.6), Inches(3.2), Inches(1.8), f"{data['photos'][1]}\n\n[ INSERT PHOTO HERE ]\n\nFig {data['id']}.2\n\nAdd: Close-up / Different angle", size=9, bold=False, color=GREY2, align=PP_ALIGN.CENTER)

    # === BOX 4: RISK (Left bottom) ===
    risk_bg = RGBColor(0x2A,0x0A,0x0A) if data['risk']=="CRITICAL" else RGBColor(0x2B,0x20,0x00)
    shape(slide, Inches(0.2), Inches(4.2), Inches(5.5), Inches(1.3), fill=risk_bg, border=risk_color, border_w=Pt(2))
    text_box(slide, Inches(0.3), Inches(4.2), Inches(5.3), Inches(0.3), f"⚠️ BOX 4: RISK LEVEL - {data['risk']} / خانة مستوى الخطورة", size=10, bold=True, color=risk_color)
    text_box(slide, Inches(0.3), Inches(4.5), Inches(5.3), Inches(0.9), data['risk_text'], size=9.5, bold=False, color=LIGHT_GREY)

    # === BOX 5: CORRECTIVE ACTION (Right bottom, large) ===
    shape(slide, Inches(5.9), Inches(4.2), Inches(7.2), Inches(2.8), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(6.0), Inches(4.2), Inches(7.0), Inches(0.3), "✅ BOX 5: REQUIRED CORRECTIVE ACTION / خانة الإجراءات التصحيحية + STANDARD REFERENCE", size=10, bold=True, color=DARK_PETROL)
    tf2 = slide.shapes.add_textbox(Inches(6.0), Inches(4.5), Inches(4.5), Inches(2.4)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(9.5)
        p.font.color.rgb = DARK_PETROL
        p.font.name = "Calibri"
        p.space_after = Pt(3)

    # Standard ref sub-box inside corrective
    shape(slide, Inches(10.7), Inches(4.6), Inches(2.3), Inches(2.2), fill=RGBColor(0xF0,0xF0,0xF0), border=GREY2)
    text_box(slide, Inches(10.8), Inches(4.6), Inches(2.1), Inches(0.3), "📚 STANDARD REF", size=8, bold=True, color=DARK_PETROL)
    std_text = "• ISO 45001:2018\n• ISO 14001:2015\n• OSHA 29 CFR 1926\n• SE EHS Contractor Req\n• Project EHS Plan\n• Local EPA / HSG"
    text_box(slide, Inches(10.8), Inches(4.9), Inches(2.1), Inches(1.8), std_text, size=8, color=DARK_PETROL)

    # Footer
    shape(slide, Inches(0), Inches(7.15), Inches(13.33), Pt(1), fill=GREY2)
    text_box(slide, Inches(0.2), Inches(7.2), Inches(6), Inches(0.2), "Siemens Energy | Zero Harm | Confidential | Boxed Format with Photo Placeholders", size=7, color=GREY2)
    text_box(slide, Inches(11), Inches(7.2), Inches(2), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY2, align=PP_ALIGN.RIGHT)

# === TITLE SLIDE ===
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DEEP_PETROL)
shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(6), fill=TEAL)
text_box(slide, Inches(0.5), Inches(0.3), Inches(4), Inches(0.4), "SIEMENS ENERGY", size=16, bold=True, color=TEAL_LIGHT)
text_box(slide, Inches(9.5), Inches(0.3), Inches(3.5), Inches(0.4), "CONFIDENTIAL", size=11, bold=True, color=AMBER, align=PP_ALIGN.RIGHT)
shape(slide, Inches(0.5), Inches(1.0), Inches(3), Pt(3), fill=AMBER)
text_box(slide, Inches(0.5), Inches(1.4), Inches(12), Inches(1.2), "SUBCONTRACTOR EHS NON-COMPLIANCE & SAFETY VIOLATION REPORT", size=28, bold=True, color=WHITE)
text_box(slide, Inches(0.5), Inches(2.6), Inches(12), Inches(0.5), "BOXED FORMAT WITH PHOTO PLACEHOLDERS / نسخة مقسمة بخانات وأماكن صور", size=14, bold=False, color=TEAL_LIGHT)
text_box(slide, Inches(0.5), Inches(3.3), Inches(6), Inches(1.5), "Project: KAZ Power Plant Upgrade\nRef: SE-EHS-NC-2026-001 BOXED V2\nDate: September 2026\nPrepared by: EHS Manager - Siemens Energy\n\nFORMAT:\n• Each slide = 5 Boxes\n• Box 1: Non-Compliance Details\n• Box 2 & 3: Photo Evidence Placeholders (2 photos per violation)\n• Box 4: Risk Level\n• Box 5: Corrective Action + Standard Reference", size=11, color=LIGHT_GREY)

# Visual guide box
shape(slide, Inches(7.5), Inches(3.3), Inches(5.3), Inches(3.2), fill=CARD_BG, border=TEAL)
text_box(slide, Inches(7.6), Inches(3.3), Inches(5.1), Inches(0.3), "HOW TO USE THIS TEMPLATE / كيفية الاستخدام", size=11, bold=True, color=AMBER)
guide = """1. لكل مخالفة، يوجد 5 خانات واضحة
2. خانة 1: اكتب تفاصيل المخالفة + الموقع + التاريخ
3. خانة 2 و 3: اسحب صورتين من الموقع وألصقهما هنا
   - صورة 1: نظرة عامة
   - صورة 2: تفاصيل قريبة / زاوية أخرى
4. خانة 4: مستوى الخطورة تلقائياً ملون
5. خانة 5: الإجراءات التصحيحية + المرجع القانوني

DESIGN:
• Deep Teal #0E2F3E background
• White boxes for actions
• Red for CRITICAL, Amber for HIGH
• Photo boxes have [INSERT PHOTO HERE] placeholder
• Just drag & drop your site photos"""
text_box(slide, Inches(7.6), Inches(3.6), Inches(5.1), Inches(2.8), guide, size=9, color=WHITE)

text_box(slide, Inches(0.5), Inches(6.8), Inches(12), Inches(0.3), "Zero Harm is Non-Negotiable | 5 Boxes Per Slide | 2 Photos Per Violation | Ready for Management Decision", size=9, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

# Create all boxed slides
for info in slides_info:
    create_boxed_slide(info)

# === FINAL ACTION PLAN SLIDE - BOXED ===
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DEEP_PETROL)
shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.8), fill=DARK_PETROL)
shape(slide, Inches(0), Inches(0.8), Inches(13.33), Pt(4), fill=TEAL)
text_box(slide, Inches(0.3), Inches(0.1), Inches(12), Inches(0.6), "ACTION PLAN & MANAGEMENT DECISION MATRIX - BOXED FORMAT", size=18, bold=True, color=WHITE)

# 3 columns boxed
# Immediate
shape(slide, Inches(0.2), Inches(1.1), Inches(4.2), Inches(0.4), fill=RED)
text_box(slide, Inches(0.2), Inches(1.1), Inches(4.2), Inches(0.4), "🔴 BOX A: IMMEDIATE (0-24h) / فوري", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(0.2), Inches(1.5), Inches(4.2), Inches(4.0), fill=WHITE)
text_box(slide, Inches(0.3), Inches(1.5), Inches(4.0), Inches(4.0), "• STOP WORK all critical violations\n• Secure areas - hard barricade + fire watch\n• Remove non-inducted personnel\n• Ground leaking equipment - Tag Out\n• Issue NCR Level 1 & 2 + warning letter\n• Emergency meeting - Subcontractor MD in 24h\n\nPHOTO BOX: [Insert photo of stop work / secured area]\n\nDECISION NEEDED: Approve Stop-Work Authority to EHS", size=9, color=DARK_PETROL)

# Short
shape(slide, Inches(4.6), Inches(1.1), Inches(4.2), Inches(0.4), fill=AMBER)
text_box(slide, Inches(4.6), Inches(1.1), Inches(4.2), Inches(0.4), "🟡 BOX B: SHORT-TERM (1-7 Days) / قصير", size=11, bold=True, color=DARK_PETROL, align=PP_ALIGN.CENTER)
shape(slide, Inches(4.6), Inches(1.5), Inches(4.2), Inches(4.0), fill=WHITE)
text_box(slide, Inches(4.7), Inches(1.5), Inches(4.0), Inches(4.0), "• PTW Audit + Re-Training all supervisors\n• Diesel storage - 110% bund + fire extinguishers\n• Traffic signs - retro-reflective + flagmen\n• Waste - color-coded skips + training\n• Excavation - hard barricades + ladders\n• PPE Campaign + BBS Launch\n• EHS Directives Log - daily sign-off\n\nPHOTO BOX: [Insert after-correction photos]\n\nDECISION: Approve $15k cost - back-charge to sub", size=9, color=DARK_PETROL)

# Strategic
shape(slide, Inches(9.0), Inches(1.1), Inches(4.1), Inches(0.4), fill=TEAL)
text_box(slide, Inches(9.0), Inches(1.1), Inches(4.1), Inches(0.4), "🟢 BOX C: STRATEGIC (7-30 Days) / استراتيجي", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(9.0), Inches(1.5), Inches(4.1), Inches(4.0), fill=WHITE)
text_box(slide, Inches(9.1), Inches(1.5), Inches(3.9), Inches(4.0), "• Penalty matrix linked to KPIs (TRIR, NCR closure)\n• Replace non-performing supervisors - competency check\n• QR/Biometric induction + digital PTW app\n• Weekly dashboard Red/Amber/Green to Top Mgmt\n• 3rd party EHS audit ISO 45001/14001\n• Re-sign Life Saving Rules - Townhall\n• Demobilization / replacement contingency plan\n\nPHOTO BOX: [Insert dashboard / training photos]\n\nDECISION: Authorize partial termination if recur", size=9, color=DARK_PETROL)

# Decision matrix box
shape(slide, Inches(0.2), Inches(5.7), Inches(13.0), Inches(1.5), fill=CARD_BG, border=AMBER, border_w=Pt(2))
text_box(slide, Inches(0.3), Inches(5.7), Inches(12.8), Inches(0.3), "✅ BOX D: MANAGEMENT DECISION MATRIX - REQUIRED TODAY / مصفوفة قرارات الإدارة المطلوبة اليوم", size=11, bold=True, color=AMBER)
text_box(slide, Inches(0.3), Inches(6.0), Inches(12.8), Inches(1.0), "1. Approve IMMEDIATE Stop-Work Authority to EHS for Critical Violations [ ] Yes [ ] No  |  2. Authorize Formal Contractual Warning & Penalty Enforcement [ ] Yes [ ] No  |  3. Mandate Subcontractor Top Management On-Site within 48h [ ] Yes [ ] No  |  4. Approve Budget for Additional EHS Controls - Back-charge to Subcontractor [ ] Yes [ ] No  |  5. Endorse Zero Tolerance Policy: 3 Strikes = Removal - Communicate in Townhall [ ] Yes [ ] No\n\nSignature: _________________  Project Director  Date: _______  |  Signature: _________________  Construction Manager  Date: _______", size=9, color=WHITE)

out = "/home/user/ahmed/Subcontractor_EHS_Boxed_V2_with_Photo_Placeholders.pptx"
prs.save(out)
print(f"Saved to {out}")
