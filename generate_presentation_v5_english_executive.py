#!/usr/bin/env python3
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY_BG = RGBColor(0xF5, 0xF5, 0xF7)
LIGHT_GREY_CARD = RGBColor(0xF9, 0xF9, 0xFA)
GREY_BORDER = RGBColor(0xE0, 0xE0, 0xE0)
DARK_TEXT = RGBColor(0x21, 0x21, 0x21)
GREY_TEXT = RGBColor(0x61, 0x61, 0x61)
RED = RGBColor(0xE3, 0x06, 0x13)
RED_LIGHT = RGBColor(0xFF, 0xEB, 0xEB)
AMBER = RGBColor(0xFF, 0xB9, 0x00)
AMBER_LIGHT = RGBColor(0xFF, 0xF8, 0xE1)
TEAL = RGBColor(0x00, 0x99, 0x99)
TEAL_DARK = RGBColor(0x0E, 0x2F, 0x3E)
BLACK = RGBColor(0x00, 0x00, 0x00)

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

def text_box(slide, l, t, w, h, txt, size=11, bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Calibri"
    p.alignment = align
    return tb

slides_data = [
    {
        "id": "01",
        "title": "PPE Violations - Failure to Wear Personal Protective Equipment",
        "desc": [
            "• Subcontractor personnel observed without basic PPE: Safety Helmet, Safety Glasses, Safety Boots.",
            "• Failure to use Full Body Harness when working at heights above 1.8m - fall protection missing.",
            "• Partial or incorrect use of PPE - helmet without chin strap, high-visibility vest not fastened, gloves not worn.",
            "• Supervision present on site but failing to enforce PPE compliance - clear supervisory failure.",
            "• Violation of Siemens Energy Life Saving Rule #3 and PPE Standard EN 397 / EN 166 / ANSI Z89.1"
        ],
        "risk": "HIGH",
        "risk_text": "Immediate risk of head injury, eye injury, hand lacerations, fall from height, slip/trip incidents. Direct violation of Life Saving Rule #3. Indicates weak safety culture and significantly increases LTIFR/TRIR. Leading indicator for more severe violations.",
        "action": [
            "1. Immediate on-site correction + documented Toolbox Talk with photos and attendance sheet",
            "2. Issue Safety Observation Cards and formal EHS violation notice to individual and supervisor",
            "3. Launch daily PPE compliance campaign - checks at start of shift + random inspections by HSE Officers",
            "4. Link PPE compliance to subcontractor KPI and payment milestone - include in weekly scorecard",
            "5. Enforce supervisor accountability - 3-strikes removal policy - formal briefing and signed commitment"
        ],
        "photo_label": "PPE Violation - Helmet / Glasses / Harness missing"
    },
    {
        "id": "02",
        "title": "Working Without PTW - High-Risk Activities Without Permit to Work",
        "desc": [
            "• Commencement of high-risk activities (Hot Work, Confined Space Entry, Electrical Work) without issuance or activation of Permit to Work (PTW).",
            "• No Task Risk Assessment (TRA) review, no energy isolation verification, no SIMOPS coordination check.",
            "• PTW logbook not maintained, permits not displayed at work front as required by procedure.",
            "• Direct violation of Siemens Energy EHS General Requirements for Contractors Sec 5.3 and KAZ Project EHS Plan Sec 7.2",
            "• Represents total loss of primary control barrier - critical procedural failure"
        ],
        "risk": "CRITICAL",
        "risk_text": "Potential fatality, fire/explosion, uncontrolled release of hazardous energy, SIMOPS conflict. Total loss of primary control barrier - leading indicator of major accident potential (MAH). Breach of ISO 45001 Cl 8.1.2. Regulatory authority may issue prohibition notice and project stoppage.",
        "action": [
            "1. IMMEDIATE STOP WORK order for all activities without PTW - enforce until full compliance verified by SE EHS",
            "2. Mandatory re-training of all subcontractor supervision (foremen, engineers) on Siemens PTW System within 24 hours - competency test 80% pass mark",
            "3. Issue formal Non-Conformance Report (NCR) Level 1 and contractual warning letter",
            "4. EHS to conduct full PTW system audit and daily spot-checks for next 14 days - findings in daily report",
            "5. Supervisor accountability: Removal of non-compliant supervisors if recurrence observed - 3-strikes policy"
        ],
        "photo_label": "Work without PTW displayed / PTW logbook missing"
    },
    {
        "id": "03",
        "title": "Hard Barricades for Excavations - Inadequate Protection Around Excavation Zones",
        "desc": [
            "• Excavation zones >1.2m depth (up to 2.5m observed) left without hard barricades - only loose warning tape used, which is prohibited for deep excavations.",
            "• No edge protection, no safe access/egress (ladders missing), spoil pile stored within 1m of excavation edge - surcharge loading risk.",
            "• No reflective markers, delineators, or night lighting to prevent fall incidents during low visibility.",
            "• Violation of OSHA 29 CFR 1926 Subpart P, ISO 45001 Working at Height/Excavation, and Siemens Standard SE-GEN-EHS-007",
            "• Night work hazard: Excavation invisible - high risk of personnel and vehicle fall"
        ],
        "risk": "CRITICAL",
        "risk_text": "High potential for fatal fall into excavation, cave-in engulfment, personnel/vehicle fall. Excavation is classified as High-Risk activity with significant fatality history in construction - Major Accident Hazard (MAH). Regulatory violation may trigger authority intervention and prohibition notice.",
        "action": [
            "1. IMMEDIATE hard barricading with scaffold tubes / concrete barriers / hard fencing around all open excavations - compliant with SE-GEN-EHS-007",
            "2. Provide compliant access ladders every 7.5m and keep spoil minimum 1.5m away from edge - toe boards where required",
            "3. Excavation permit re-validation with geotechnical check and competent person inspection",
            "4. Install reflective delineators and solar blinker lights for low-light hours",
            "5. Competent person inspection checklist to be submitted daily before start of work - signed by engineer and verified by SE EHS"
        ],
        "photo_label": "Excavation without hard barricade - tape only / Spoil pile too close"
    },
    {
        "id": "04",
        "title": "Parking & Speed Limits - Non-Compliance with Site Traffic Signage",
        "desc": [
            "• Non-compliance with directional signage - speed limit signs missing/damaged, no designated parking signage, pedestrian walkways not marked.",
            "• Speed limit violation inside plant (15 km/h limit) - observed 30 km/h confirmed by radar gun spot check.",
            "• Parking in non-designated areas - obstructing emergency routes and pedestrian walkways.",
            "• No flagmen at high-risk intersections and blind spots - vehicle-pedestrian interface unmanaged.",
            "• Violation of Project Traffic Management Plan (TMP) and ISO 39001 Road Traffic Safety Management"
        ],
        "risk": "HIGH",
        "risk_text": "High risk of vehicle-pedestrian collision, equipment damage, and potential fatality. Vehicle incidents are top 3 fatality causes in construction industry. Poor signage increases liability in case of incident. Emergency vehicle access may be impeded - risk to emergency response.",
        "action": [
            "1. Immediate reinstatement of all traffic signs per approved TMP - retro-reflective grade, proper height and visibility",
            "2. Establish dedicated pedestrian routes with hard segregation (Jersey barriers) - marked with green walkway paint",
            "3. Deploy trained flagmen at high-risk intersections and implement speed monitoring - radar checks twice daily",
            "4. Conduct traffic safety campaign and driver re-induction - all drivers including subcontractor personnel",
            "5. EHS + Logistics joint daily inspection of traffic controls - checklist to be signed and filed"
        ],
        "photo_label": "Missing speed limit sign / Random parking / No pedestrian segregation"
    },
    {
        "id": "05",
        "title": "Waste Management & Segregation - Poor Housekeeping and Waste Handling",
        "desc": [
            "• Accumulation of site waste and lack of segregation - hazardous, wood, plastic, debris mixed in same container.",
            "• No clear labeling, color-coding, or MSDS reference for hazardous waste bins - violates visual management standard.",
            "• Oil-contaminated rags and filters disposed in general waste - environmental violation and fire hazard.",
            "• Location: Waste area not defined, containers overflowing - poor housekeeping throughout site.",
            "• Non-compliance with ISO 14001:2015 Cl 8.1, Project Waste Management Plan, and local EPA regulations"
        ],
        "risk": "HIGH",
        "risk_text": "Environmental pollution risk (soil/groundwater contamination), fire hazard from incompatible waste mixing, regulatory fine, and reputational damage to Siemens Energy. Cross-contamination makes recycling impossible - sustainability KPI failure. Potential for legal non-compliance and environmental prosecution.",
        "action": [
            "1. Immediate provision of color-coded, labeled waste skips per SE Environmental Standard: Black-General, Green-Recyclable, Red-Hazardous with HAZCHEM labels",
            "2. Clean-up and correct segregation of existing mixed waste by competent team under EHS supervision",
            "3. Training on Waste Management - focus on hazardous waste handling, spill prevention, and manifest system",
            "4. Appoint Waste Management Focal Point from subcontractor - responsible for daily housekeeping and segregation",
            "5. Weekly waste audit - track waste volumes and disposal manifests - report in monthly environmental report"
        ],
        "photo_label": "Mixed waste in one skip - hazardous + general / Oil rags in general waste"
    },
    {
        "id": "06",
        "title": "Unsafe Workshop Equipment - Damaged and Uninspected Tools",
        "desc": [
            "• Use of damaged or uninspected workshop equipment - exposed electrical cables, equipment without protective guards (e.g., angle grinder without guard).",
            "• No preventive maintenance records available on site - daily inspection checklists falsified or not completed - same signature for 7 days.",
            "• No drip trays provided - soil contamination observed under equipment parking area.",
            "• Equipment without third-party inspection certificates - no evidence of competency for operators.",
            "• Violation of Equipment Safety Standard, PUWER Regulations, and Siemens Energy Equipment Standard"
        ],
        "risk": "CRITICAL",
        "risk_text": "Electric shock, fire/explosion hazard (fuel on hot surfaces), slip hazard, soil and groundwater contamination, equipment failure leading to loss of control/incident. Represents breakdown of maintenance management system and lack of competent operators. Breach of ISO 14001 (pollution prevention) and ISO 45001 (equipment safety).",
        "action": [
            "1. IMMEDIATE grounding of damaged/leaking equipment - Tag 'DO NOT USE - EHS HOLD' and remove keys - secure area",
            "2. Contaminated soil to be excavated and disposed as hazardous waste - spill response team to clean area with absorbents",
            "3. Subcontractor to submit full equipment maintenance history and third-party inspection certificates within 48 hours - No certificate, no operation",
            "4. Implement mandatory daily equipment checklist - verified by SE EHS/Mechanical - with photo evidence of no leaks and guard in place",
            "5. Introduction of drip trays and spill kits at all equipment parking zones - 100% coverage - and replacement of damaged cables/guards"
        ],
        "photo_label": "Exposed cable / Equipment without guard / Active oil leak"
    },
    {
        "id": "07",
        "title": "Uncoordinated Work & Lack of Induction - Working Without EHS Coordination",
        "desc": [
            "• Introduction of manpower to site without completing EHS Induction (Safety Induction) - 8-12 persons found working without induction card/sticker.",
            "• Commencement of work without prior coordination with EHS Department - no gate entry log, induction register incomplete and not cross-checked at gate.",
            "• Toolbox Talks (TBT) not conducted or without documentation, attendance sheets, or relevance to task.",
            "• Previous NCRs and observation reports closed without evidence or verification - falsification of closure - willful disregard for EHS Management System.",
            "• Indicates systemic failure in access control, security-EHS interface, and subcontractor self-verification process"
        ],
        "risk": "CRITICAL",
        "risk_text": "Workers unaware of site-specific hazards, emergency assembly points, muster procedures, life-saving rules, and incident reporting channels. High probability of unsafe acts, delayed emergency response, and regulatory non-compliance leading to potential project stoppage by authorities. Breach of ISO 45001 Cl 7.2 Competence and 7.3 Awareness.",
        "action": [
            "1. Immediate removal of all non-inducted personnel from site - escort off-site by Security - 100% gate check from next shift",
            "2. Implement joint Security + EHS verification at main gate - No induction card, no entry - effective immediately",
            "3. Subcontractor to submit full manpower list vs induction records within 12 hours - with gap closure plan and accountability statement",
            "4. Introduce QR/Biometric induction tracking system linked to access control + Daily EHS Directives Log sign-off by Subcontractor PM",
            "5. Formal warning and back-charge for re-induction + Hold EHS Leadership stand-down meeting with all foremen - documented commitment and action plan signed"
        ],
        "photo_label": "Workers without induction cards / Gate register mismatch / Missing TBT register"
    },
]

def create_slide(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.75), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), fill=TEAL_DARK)
    text_box(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), data['id'], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(1.0), Inches(0.15), Inches(8.5), Inches(0.5), data['title'], size=14, bold=True, color=DARK_TEXT)
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.2), Inches(0.15), Inches(2.8), Inches(0.5), fill=risk_bg, border=risk_color, border_w=Pt(1.5))
    text_box(slide, Inches(10.2), Inches(0.15), Inches(2.8), Inches(0.5), f"RISK: {data['risk']}", size=11, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.3), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.3), "1. TECHNICAL DESCRIPTION OF NON-COMPLIANCE", size=10, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.4), Inches(7.2), Inches(2.0), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    tf = slide.shapes.add_textbox(Inches(0.4), Inches(1.45), Inches(7.0), Inches(1.9)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc']):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(8.8)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    shape(slide, Inches(0.3), Inches(3.6), Inches(7.2), Inches(0.3), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.6), Inches(7.0), Inches(0.3), f"2. RISK ASSESSMENT & IMPACT - {data['risk']}", size=10, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.9), Inches(7.2), Inches(1.1), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(3.9), Inches(7.0), Inches(1.0), data['risk_text'], size=8.8, bold=False, color=DARK_TEXT)

    shape(slide, Inches(0.3), Inches(5.2), Inches(7.2), Inches(0.3), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(5.2), Inches(7.0), Inches(0.3), "3. REQUIRED CORRECTIVE ACTION (Contractor)", size=10, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.5), Inches(7.2), Inches(1.6), fill=WHITE, border=GREY_BORDER)
    tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(5.55), Inches(7.0), Inches(1.5)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(8.2)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    # Photo area
    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.3), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.3), "📷 PHOTO EVIDENCE AREA - Dedicated Space for Site Photos", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(1.4), Inches(5.2), Inches(4.0), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1.5))
    text_box(slide, Inches(7.9), Inches(1.5), Inches(5.0), Inches(0.4), "[Insert Violation Photo Here]", size=14, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.0), Inches(5.0), Inches(0.5), f"{data['photo_label']}\nFig {data['id']}.1 - {data['title'][:40]}", size=9, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.8), Inches(5.0), Inches(1.2), "⬆️ DRAG & DROP YOUR SITE PHOTO HERE\n\nPhoto will auto-fit inside this box\n\nAdd: Date, Time, Location, GPS Coordinates\nTaken by: ___________\n\nThis box is reserved for approved site photo evidence", size=10, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(5.5), Inches(5.2), Inches(0.6), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(5.5), Inches(5.0), Inches(0.6), "Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________  Taken by: ___________  Camera: ___________\nStandard Ref: ISO 45001:2018, OSHA 29 CFR 1926, SE EHS Requirements, Project EHS Plan", size=7.5, color=GREY_TEXT)
    shape(slide, Inches(7.8), Inches(6.2), Inches(2.5), Inches(0.9), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(6.2), Inches(2.3), Inches(0.9), "[Additional Photo 2]\nFig X.2\nDifferent Angle / Close-up\n[Insert Second Photo Here]", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(10.5), Inches(6.2), Inches(2.5), Inches(0.9), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL)
    text_box(slide, Inches(10.6), Inches(6.2), Inches(2.3), Inches(0.9), "[After Correction Photo]\nFig X.3\nCompliant State\n[Insert After Photo Here]", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0), Inches(7.2), Inches(13.33), Pt(1), fill=GREY_BORDER)
    text_box(slide, Inches(0.3), Inches(7.25), Inches(6), Inches(0.2), "Siemens Energy | KAZ Project | EHS Department | Confidential - Management Review | Executive White Design", size=7, color=GREY_TEXT)
    text_box(slide, Inches(10.5), Inches(7.25), Inches(2.5), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

def create_title(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), fill=TEAL)
    shape(slide, Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(0.35), Inches(3), Inches(0.3), "SIEMENS ENERGY", size=14, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.6), Inches(0.65), Inches(5), Inches(0.3), "KAZ Power Plant Upgrade Project | Zero Harm Culture", size=9, color=GREY_TEXT)
    text_box(slide, Inches(10), Inches(0.35), Inches(2.8), Inches(0.6), "CONFIDENTIAL\nManagement Eyes Only", size=10, bold=True, color=RED, align=PP_ALIGN.RIGHT)
    shape(slide, Inches(0.5), Inches(1.5), Inches(0.08), Inches(1.2), fill=RED)
    text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.7), "Subcontractor EHS Non-Compliance Report", size=28, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(0.8), Inches(2.2), Inches(11.5), Inches(0.5), "Field Operations Safety Violations - Subcontractor", size=16, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.8), Inches(2.7), Inches(11.5), Inches(0.4), "Operational Compliance & Risk Assessment | Executive Management Review", size=12, bold=False, color=GREY_TEXT)
    shape(slide, Inches(0.5), Inches(3.3), Inches(5.5), Inches(2.2), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(3.3), Inches(5.3), Inches(0.3), "Report Information", size=11, bold=True, color=TEAL_DARK)
    info = """Reference: SE-EHS-NC-2026-EXEC-EN
Date: September 2026
Prepared by: EHS Manager - Siemens Energy
Classification: Confidential - Management Review & Decision

Standards: ISO 45001:2018, ISO 14001:2015, OSHA 29 CFR 1926
Siemens Energy EHS Principles & Life Saving Rules
Project EHS Plan & Traffic Management Plan

Distribution: Project Director, Construction Manager, Commercial Manager, QA/QC Manager, Subcontractor Top Management

Design: Executive & Modern - Pure White #FFFFFF background with Red #E30613 / Amber #FFB900 warning highlights"""
    text_box(slide, Inches(0.6), Inches(3.6), Inches(5.3), Inches(1.8), info, size=8.5, color=DARK_TEXT)
    shape(slide, Inches(6.5), Inches(3.3), Inches(6.3), Inches(2.2), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(6.6), Inches(3.3), Inches(6.1), Inches(0.3), "Design & Usage Instructions - Executive White Format", size=11, bold=True, color=TEAL_DARK)
    guide = """Executive Design:
• Background: Pure White #FFFFFF / Light Grey #F5F5F7
• Text: Dark #212121 / Grey #616161
• Warning: Red #E30613 (CRITICAL) / Amber #FFB900 (HIGH)
• Accent: Teal #009999 / Dark Petrol #0E2F3E

Each Slide Structure (4 Boxes):
1. Violation Title - Top header with ID and Risk badge
2. Technical Description - 5 points with location/date fields
3. Risk Assessment - Impact and regulatory breach
4. Corrective Action - 5 numbered actions - owner & deadline
5. Photo Evidence Area - Large dedicated box [Insert Violation Photo Here]

How to Add Photos:
• Drag & drop your approved site photo into large right-side box
• Box will auto-fit photo - no resizing needed
• Add 2 photos: Overview + Close-up + After Correction photo
• Fill Photo Details: Date, Time, Location, GPS, Taken by

All slides include:
• [Insert Violation Photo Here] placeholder - clear and visible
• Fig number, photo label, details field
• Standard reference footer"""
    text_box(slide, Inches(6.6), Inches(3.6), Inches(6.1), Inches(1.8), guide, size=8, color=DARK_TEXT)
    shape(slide, Inches(0.5), Inches(5.8), Inches(12.33), Inches(0.9), fill=RED_LIGHT, border=RED, border_w=Pt(1.5))
    text_box(slide, Inches(0.6), Inches(5.8), Inches(12.1), Inches(0.2), "Executive Statement:", size=10, bold=True, color=RED)
    text_box(slide, Inches(0.6), Inches(6.05), Inches(12.1), Inches(0.6), "This report documents critical, systemic EHS non-compliances observed during site inspections of subcontractor field operations. Findings indicate fundamental breakdown of subcontractor EHS management system and direct threat to our Zero Harm commitment, project schedule, and legal compliance. Immediate management intervention is required. The cost of inaction far exceeds the cost of intervention.", size=9, color=DARK_TEXT)

def create_closing(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.5), Inches(0.1), Inches(12), Inches(0.6), "Final Recommendations & Corrective Action Plan - Management Decision Required", size=18, bold=True, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), fill=RED)
    text_box(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), "🔴 IMMEDIATE (0-24 HOURS)", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0.3), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(1.4), Inches(3.9), Inches(3.8), "• STOP WORK all critical violations (PTW, Excavation, Diesel, Leaking Equipment)\n• Secure hazardous areas - hard barricade + fire watch + 15m exclusion zone\n• Remove all non-inducted personnel - 100% gate verification from next shift\n• Ground all leaking equipment - Tag DO NOT USE - remove keys\n• Issue formal NCR Level 1 & 2 + contractual warning letter\n• Convene Emergency EHS Meeting - Subcontractor MD/PM within 24h\n\nPHOTO BOX: [Insert photo of stop work / secured area]\n\nDECISION REQUIRED: Approve Stop-Work Authority to EHS\n\nStandard: ISO 45001 Cl 8.1.2, SE EHS General Requirements", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), fill=AMBER)
    text_box(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), "🟡 SHORT-TERM (1-7 DAYS)", size=11, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(4.6), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(4.7), Inches(1.4), Inches(3.9), Inches(3.8), "• Full PTW System Audit + Re-Training - all supervisors certified - 80% pass mark\n• Compliant diesel storage - 110% bund + impervious floor + fire extinguishers + signage\n• Reinstate traffic signage - retro-reflective + pedestrian segregation + flagmen + radar\n• Waste management system - color-coded skips + labeling + focal point + training\n• Excavation compliance - hard barricades + ladders every 7.5m + competent person checks\n• PPE Campaign + BBS Launch - daily checks + scorecard + incentives\n• EHS Directives Log - daily sign-off by Subcontractor PM + SE EHS Manager\n\nPHOTO BOX: [Insert after-correction photos - compliant state]\n\nDECISION: Approve $15k cost - back-charge to subcontractor", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), fill=TEAL)
    text_box(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), "🟢 STRATEGIC (7-30 DAYS) - CULTURE FIX", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(8.9), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(9.0), Inches(1.4), Inches(3.9), Inches(3.8), "• Contractual Enforcement: Penalty matrix linked to EHS KPIs (TRIR, NCR closure, PTW compliance) - deduction from invoice\n• Leadership Accountability: Replace non-performing supervisors - CV + EHS competency + interview by SE\n• System Upgrade: QR/Biometric induction tracking + digital PTW system + equipment checklist app\n• KPI & Scorecard: Weekly EHS compliance dashboard Red/Amber/Green to Top Management\n• Independent Audit: 3rd party EHS audit of subcontractor system - ISO 45001/14001 gap analysis\n• Zero Harm Commitment: Re-sign Life Saving Rules - Subcontractor Top Management visible commitment in townhall\n• Contingency: Prepare demobilization / replacement plan if no sustained improvement\n\nPHOTO BOX: [Insert dashboard / training photos]\n\nDECISION: Authorize partial termination if critical violations recur", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(5.5), Inches(12.7), Inches(1.7), fill=LIGHT_GREY_CARD, border=TEAL_DARK, border_w=Pt(2))
    text_box(slide, Inches(0.4), Inches(5.5), Inches(12.5), Inches(0.3), "✅ MANAGEMENT DECISION MATRIX - REQUIRED TODAY", size=12, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(1.2), "1. Approve IMMEDIATE Stop-Work Authority to EHS for Critical Violations [ ] Yes [ ] No\n2. Authorize Formal Contractual Warning & Penalty Enforcement [ ] Yes [ ] No\n3. Mandate Subcontractor Top Management Presence On-Site within 48h [ ] Yes [ ] No\n4. Approve Budget for Additional EHS Controls (Bunded Storage, Signage, Barricades) - Back-charge to Subcontractor [ ] Yes [ ] No\n5. Endorse Zero Tolerance Policy: 3 Strikes = Removal - Communicate in Townhall [ ] Yes [ ] No\n\nSignature: _________________  Project Director  Date: _______  |  Signature: _________________  Construction Manager  Date: _______  |  Signature: _________________  EHS Manager  Date: _______\n\nClosing Statement: Zero Harm is Non-Negotiable. Cost of inaction far exceeds cost of intervention. Request Top Management endorsement.", size=8.5, color=DARK_TEXT)

final_prs = Presentation()
final_prs.slide_width = Inches(13.33)
final_prs.slide_height = Inches(7.5)
create_title(final_prs)
for d in slides_data:
    create_slide(d)
create_closing(final_prs)

out = "/home/user/ahmed/Subcontractor_EHS_Executive_White_ENGLISH_ONLY_7Slides.pptx"
final_prs.save(out)
print(f"Saved to {out}")
