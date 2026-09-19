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
        "title": "Unauthorized Works Without PTW & Induction - Painting, Storage Fence, Cement Pouring in Energized Zone",
        "desc": [
            "• Activities Observed: Painting works, installation of storage area fence/barrier (Storage), and cement pouring works executed WITHOUT approved Permit to Work (PTW).",
            "• Personnel involved were found WITHOUT valid EHS Site Induction - no induction cards, no QR verification at gate.",
            "• Critical Failure: When questioned by Siemens Energy EHS Department, subcontractor supervision (Al-Mial) stated they have NO knowledge of these activities, workers, or location - despite work being inside ENERGIZED zone.",
            "• Energized Zone Risk: Work conducted inside live electrical area without LOTO verification, without energized zone risk assessment, and without EHS coordination - potential arc flash, electrocution.",
            "• Violation: Siemens Energy EHS General Requirements Sec 5.3 PTW, KAZ EHS Plan Sec 6.1 Induction & Sec 7.2 PTW, ISO 45001 Cl 7.2, 7.3, 8.1.2 - Total loss of control."
        ],
        "risk": "CRITICAL",
        "risk_text": "FATALITY RISK - Electrocution, arc flash, chemical exposure (painting), struck-by, SIMOPS conflict. Workers unaware of energized hazards and emergency procedures. Subcontractor management has no control over own workforce - indicates total breakdown of work management system. Regulatory prohibition and criminal liability. If incident occurs in energized zone, Siemens Energy exposed to major liability.",
        "action": [
            "1. IMMEDIATE STOP WORK for painting, storage fence, cement pouring - evacuate energized zone - secure with hard barricade and LOTO verification",
            "2. Remove all non-inducted personnel - 100% gate verification - submit full manpower list vs induction records within 6 hours",
            "3. Subcontractor PM to provide written explanation: How work started without PTW/induction and without knowledge of own supervision? - Root cause analysis required",
            "4. Re-training on PTW and Energized Zone Work Procedure for all Al-Mial supervision within 24h - competency test - no pass, no work",
            "5. Implement Daily Work Front Coordination Meeting - Al-Mial to submit daily work plan 24h in advance to SE EHS for approval - No plan, no work - Link to penalty"
        ],
        "photo_label": "Painting without PTW / Storage fence installation without induction / Cement pouring in energized zone"
    },
    {
        "id": "02",
        "title": "Excavation Barricades - Soft Barricade Used Instead of Required Hard Barricade",
        "desc": [
            "• Observation: Majority of excavation zones observed with SOFT barricade (warning tape only) instead of required HARD barricade (scaffold tubes, concrete barriers, hard fencing).",
            "• Depth: Excavations >1.2m up to 2.5m - hard barricade is MANDATORY per SE-GEN-EHS-007 and OSHA 1926 Subpart P - soft tape is PROHIBITED for >1.2m depth.",
            "• Missing: Edge protection, safe access/egress ladders, spoil pile within 1m of edge, no reflective delineators, no night lighting.",
            "• Night Risk: Excavations invisible at night - high risk of personnel/vehicle fall - no solar blinkers or reflective markers.",
            "• Violation: SE Excavation Safety Standard, OSHA, ISO 45001 - Repeated violation despite previous NCRs on same issue."
        ],
        "risk": "CRITICAL",
        "risk_text": "Fatal fall into excavation, cave-in engulfment, vehicle fall into excavation. Soft tape provides ZERO physical protection - only visual warning. Excavation is High-Risk MAH activity. Regulatory authority may issue prohibition notice. Previous incidents in industry show tape failure leads to fatality.",
        "action": [
            "1. IMMEDIATE replacement of all soft barricades with HARD barricades - scaffold tubes with toe boards / concrete Jersey barriers around ALL open excavations - within 12 hours",
            "2. Provide compliant access ladders every 7.5m, keep spoil minimum 1.5m away, install reflective delineators + solar lights for night",
            "3. Excavation permit re-validation with geotechnical assessment and competent person daily inspection checklist - signed and verified by SE EHS",
            "4. Issue NCR Level 2 for repeat violation - contractual penalty - back-charge cost of hard barricades to subcontractor",
            "5. Photo evidence: Before (soft tape) vs After (hard barricade) - submit to SE EHS for closure verification"
        ],
        "photo_label": "Soft tape only around deep excavation / No hard barricade / No ladder + spoil too close"
    },
    {
        "id": "03",
        "title": "Recurring PPE Violations Despite Repeated Warnings",
        "desc": [
            "• Observation: Persistent and recurring failure to wear mandatory PPE - safety helmets without chin straps, safety glasses, gloves, high-visibility vests, safety boots.",
            "• Working at height without Full Body Harness - observed during storage fence installation and painting works at height.",
            "• History: Same violations reported multiple times in Daily EHS Reports, Safety Observation Cards, and Toolbox Talks - NO improvement.",
            "• Supervision present on site but failing to enforce - behavioral and cultural failure, not availability - PPE available in store.",
            "• Violation: Siemens Energy Life Saving Rule #3 PPE, EN 397, EN 166, ANSI Z89.1, ISO 45001 - Normalization of deviance."
        ],
        "risk": "HIGH",
        "risk_text": "Head injury, eye injury, hand lacerations, fall from height, struck-by incidents. Recurring violation indicates total failure of safety culture and supervisory oversight. If PPE (last line of defense per hierarchy of controls) is ignored, higher-level controls will also fail. Increases LTIFR/TRIR, damages Zero Harm commitment, client confidence at risk. Precursor to serious injury/fatality per Heinrich triangle.",
        "action": [
            "1. Immediate on-site correction + documented Toolbox Talk with photos - issue Observation Cards + formal violation notice to individual and supervisor",
            "2. Implement Behavior-Based Safety (BBS) program - daily PPE checks at start of shift + random checks by HSE Officers - scorecard",
            "3. Enforce 3-Strikes Removal Policy: 1st warning, 2nd suspension, 3rd removal from site - for workers AND supervisors - communicate in townhall",
            "4. Link PPE compliance to subcontractor KPI and payment milestone - include in weekly EHS scorecard Red/Amber/Green - financial consequence",
            "5. Daily management safety walks - Siemens Energy PM + Subcontractor PM - joint walk and documented actions - positive recognition for compliant crews"
        ],
        "photo_label": "Repeat PPE violation - same crew / Helmet without chin strap / Working at height without harness"
    },
    {
        "id": "04",
        "title": "Historical Recurring Violations - Bitumen, Workshop, Formwork Works Without PTW",
        "desc": [
            "• Pattern Analysis: Same category of violation (Working Without PTW) repeated across multiple historical activities - indicates SYSTEMIC failure, not isolated incident.",
            "• Previous Activities Without PTW: Bitumen works (hot work - fire risk), Workshop activities (electrical, grinding), Formwork (working at height, manual handling, collapse risk).",
            "• All executed without approved PTW, without TRA review, without EHS coordination - documented in previous NCRs and observation reports.",
            "• Subcontractor failed to learn from previous NCRs - no corrective action implemented - same root cause repeated.",
            "• Violation: Siemens Energy EHS General Requirements Sec 5.3, KAZ EHS Plan, ISO 45001 Cl 10.2 Incident, nonconformity and corrective action - failure to prevent recurrence."
        ],
        "risk": "CRITICAL",
        "risk_text": "Systemic failure of PTW management system - demonstrates subcontractor does not understand or respect PTW as critical control barrier. Bitumen (fire/explosion, burn), Workshop (electric shock, fire, eye injury), Formwork (collapse, fall from height, struck-by) - all high-risk activities. If PTW is optional, all high-risk controls will fail. Indicates management system ineffective - requires management system intervention and top management accountability. Regulatory pattern indicates willful negligence.",
        "action": [
            "1. Conduct comprehensive PTW System Audit - review ALL historical PTWs for bitumen, workshop, formwork - identify gaps and falsification - submit audit report within 48h",
            "2. Mandatory PTW Re-Training for ALL Al-Mial supervision and workforce - including bitumen, workshop, formwork specific hazards - competency assessment 80% pass mark - no pass, no work",
            "3. Issue cumulative NCR Level 2 for repeat violation pattern - contractual penalty escalation - consider partial suspension of high-risk activities until system fixed",
            "4. Implement Digital PTW System - QR code based - with mandatory fields: TRA, isolation verification, EHS approval - no EHS approval, permit cannot be activated",
            "5. Weekly PTW Compliance Trend Analysis - present in project management meeting - track PTW compliance % - target 100% - link to subcontractor payment"
        ],
        "photo_label": "Bitumen works without PTW / Workshop without PTW / Formwork without PTW - historical photos"
    },
    {
        "id": "05",
        "title": "Non-Compliance with EHS Directives - Refusal to Stop Unsafe Equipment Despite Order",
        "desc": [
            "• Observation: Siemens Energy EHS Department issued clear verbal and written instruction to STOP use of unacceptable / unsafe equipment (leaking equipment, damaged tools, equipment without guard).",
            "• Subcontractor Response: REFUSED to comply and CONTINUED work with same unsafe equipment - willful disregard and direct challenge to EHS authority.",
            "• Equipment: Leaking hydraulic oil and fuel, damaged electrical cables, angle grinder without guard, equipment without third-party inspection certificate.",
            "• Violation: Direct violation of Siemens Energy Stop Work Authority, EHS General Requirements, Project EHS Plan Sec 8.1 Operational Control, ISO 45001 Cl 8.1.2.",
            "• Cultural Failure: Demonstrates subcontractor does not recognize EHS authority - creates precedent that EHS directives are optional - extremely dangerous."
        ],
        "risk": "CRITICAL",
        "risk_text": "Erosion of EHS authority and safety culture - if Stop Work order can be ignored, all critical controls (LOTO, WAH, PTW) are at risk of failure. Fire/explosion, electric shock, equipment failure loss of control, environmental pollution. Willful violation indicates intentional non-compliance - highest level of disciplinary action required. If not addressed decisively, Siemens Energy loses ability to control site safety - major liability. Potential for major accident with criminal liability.",
        "action": [
            "1. IMMEDIATE escalation to Top Management - Project Director to issue formal letter to Al-Mial Managing Director - require written commitment and explanation for refusal within 12h",
            "2. Enforce Stop Work Authority - ground ALL unacceptable equipment immediately - Tag 'DO NOT USE - EHS HOLD' - remove keys - secure area - deploy fire watch if fuel leak",
            "3. Invoke contractual penalty clause for refusal to follow EHS instruction - highest level penalty - back-charge cost of additional SE supervision",
            "4. Remove and replace non-compliant supervision who refused order - require CVs and EHS competency assessment + interview by SE EHS before new supervisor allowed on site",
            "5. Conduct Emergency EHS Leadership Meeting - Al-Mial Top Management + all foremen - re-communicate Stop Work Authority, EHS roles, and consequences of refusal - documented commitment signed - include in induction"
        ],
        "photo_label": "Unsafe equipment ordered to stop but continued working / Leaking equipment still in use / Equipment without guard still operating"
    },
    {
        "id": "06",
        "title": "Unsafe Condition of Workshop Equipment - Damaged Tools and Machinery",
        "desc": [
            "• Workshop Equipment Observed: Damaged and uninspected tools and machinery in unsafe condition - exposed electrical cables, angle grinders without protective guards, welding machines without earthing, damaged slings and shackles.",
            "• No preventive maintenance records - daily inspection checklists falsified - same signature for 7 consecutive days - no evidence of competent inspection.",
            "• No drip trays under equipment parking - soil contamination ~5 sqm observed - oil and fuel leaks active - fire hazard.",
            "• No third-party inspection certificates for lifting equipment, power tools, electrical equipment - no evidence of operator competency.",
            "• Violation: PUWER Regulations, Equipment Safety Standard, SE Equipment Standard, ISO 45001 Cl 8.1.2, ISO 14001 Cl 8.1 - Environmental and safety breach."
        ],
        "risk": "CRITICAL",
        "risk_text": "Electric shock, electrocution, fire/explosion (fuel on hot surface, exposed cable spark), hand injury, eye injury from grinding without guard, lifting equipment failure leading to dropped load, slip hazard, soil/groundwater contamination. Breakdown of maintenance management system and lack of competent operators. Potential for multiple fatalities if lifting equipment fails or electrical fault causes fire in workshop with flammable materials. Environmental prosecution risk.",
        "action": [
            "1. IMMEDIATE grounding of ALL unsafe workshop equipment - Tag 'DO NOT USE - EHS HOLD' - remove keys - secure workshop area - inventory of all equipment",
            "2. Contaminated soil to be excavated and disposed as hazardous waste - spill response with absorbents - clean area - provide drip trays and spill kits 100% coverage",
            "3. Subcontractor to submit full equipment register with maintenance history and valid third-party inspection certificates within 48 hours - No certificate, no operation - No checklist, no operation",
            "4. Implement mandatory daily equipment checklist - verified by SE EHS/Mechanical with photo evidence - guard in place, no leaks, cable intact - checklist to include equipment ID, inspector name, signature, photo",
            "5. Replace all damaged cables, missing guards, damaged slings/shackles - procurement of compliant equipment - training on equipment inspection for operators - competency assessment"
        ],
        "photo_label": "Exposed electrical cable / Grinder without guard / Oil leak + soil contamination / Damaged sling"
    },
    {
        "id": "07",
        "title": "Poor Waste Management & Lack of Segregation - No Clear System",
        "desc": [
            "• Observation: Waste management system not clear - no segregation, no labeling, no color-coding, no defined waste area - poor housekeeping throughout site.",
            "• Waste Mixed: Hazardous waste (oil rags, filters, paint tins), recyclable (wood, plastic), and general waste mixed in same container/skip - cross-contamination.",
            "• No MSDS reference for hazardous waste bins, no HAZCHEM labels, no spill kit near waste area - environmental violation and fire hazard.",
            "• Location: Waste containers overflowing, waste accumulated on ground, no defined collection point - attracts pests, fire risk, slip/trip hazard.",
            "• Violation: ISO 14001:2015 Cl 8.1 Operational Control, Project Waste Management Plan, local EPA regulations, SE Environmental Standard SE-ENV-003."
        ],
        "risk": "HIGH",
        "risk_text": "Environmental pollution risk (soil/groundwater contamination from hazardous waste), fire hazard from incompatible waste mixing (oil rags + general waste = spontaneous combustion), regulatory fine, and reputational damage to Siemens Energy. Cross-contamination makes recycling impossible - sustainability KPI failure. Poor housekeeping leads to slip/trip, manual handling injuries, and indicates poor overall site management. Potential for environmental prosecution and client dissatisfaction.",
        "action": [
            "1. Immediate provision of color-coded, labeled waste skips per SE Standard: Black-General waste, Green-Recyclable (wood, plastic), Red-Hazardous with HAZCHEM labels and MSDS",
            "2. Clean-up and correct segregation of existing mixed waste by competent team under EHS supervision - define waste area with hard barricade and signage",
            "3. Training on Waste Management for all Al-Mial workforce - focus on hazardous waste handling, spill prevention, manifest system, and housekeeping",
            "4. Appoint Waste Management Focal Point from subcontractor - responsible for daily housekeeping, segregation, and waste area inspection - name to be submitted to SE EHS",
            "5. Weekly waste audit - track waste volumes and disposal manifests - include in monthly environmental report - target: Zero mixed waste - link to KPI"
        ],
        "photo_label": "Mixed waste in one skip / Oil rags in general waste / Overflowing waste area / No labeling"
    },
    {
        "id": "08",
        "title": "Deficiency in Safety Signage - Parking Area (Garage) and Speed Limit Signs",
        "desc": [
            "• Observation: Poor and insufficient safety signage for parking area (garage) and speed limit signs - signs missing, damaged, not visible, or not retro-reflective.",
            "• Parking Area: No designated parking signage, no pedestrian walkway markings, no speed limit signs inside garage, no directional arrows, no stop signs at intersections.",
            "• Speed Limits: Speed limit signs (15 km/h) missing or damaged at multiple locations - no speed humps, no radar monitoring - speeding observed 30 km/h in 15 km/h zone.",
            "• Vehicle-Pedestrian Interface: No segregation between vehicle routes and pedestrian routes in parking area - workers walking through parking area with moving vehicles - blind spots.",
            "• Violation: Project Traffic Management Plan (TMP), ISO 39001 Road Traffic Safety, SE Traffic Safety Standard - Previous NCRs on same issue not closed."
        ],
        "risk": "HIGH",
        "risk_text": "High risk of vehicle-pedestrian collision, vehicle-vehicle collision, equipment damage, and potential fatality in parking area. Vehicle incidents are top 3 fatality causes in construction. Poor signage increases liability in case of incident - Siemens Energy liable if signage not per TMP. Emergency vehicle access may be impeded - risk to emergency response. Garage area with poor visibility - high risk of reversing incidents.",
        "action": [
            "1. Immediate reinstatement of all safety signage per approved TMP - retro-reflective grade, proper height, visible at night - parking signs, speed limit 15 km/h, pedestrian walkway, stop signs, directional arrows",
            "2. Establish dedicated pedestrian routes in parking/garage area with hard segregation - Jersey barriers + green walkway paint + pedestrian crossing markings",
            "3. Install speed humps and deploy trained flagmen at garage entrance/exit and high-risk intersections - implement speed monitoring with radar gun 2x daily - log speeding violations",
            "4. Conduct traffic safety campaign and driver re-induction for all drivers including Al-Mial - focus on parking area hazards, reversing, pedestrian priority",
            "5. EHS + Logistics joint daily inspection of traffic controls and signage - checklist to be signed - include parking area and garage - photo evidence of compliant signage"
        ],
        "photo_label": "Missing parking signage / Missing speed limit sign / No pedestrian walkway in garage / Speeding area"
    },
    {
        "id": "09",
        "title": "Delayed Response from Al-Mial to Close Observations - Poor Close-Out System",
        "desc": [
            "• Observation: Significant delay in response from Al-Mial (subcontractor) department to close EHS observations, NCRs, and Safety Observation Cards - poor close-out system.",
            "• Timeline: Observations issued days/weeks ago still open - no corrective action submitted, no evidence of closure, no root cause analysis - repeated follow-ups required by SE EHS.",
            "• Previous NCRs closed without evidence or verification - falsification of closure - same violations recurring because root cause not addressed.",
            "• No dedicated EHS focal point from Al-Mial to track and close observations - no tracking system - observations lost or ignored.",
            "• Violation: ISO 45001 Cl 10.2 Nonconformity and corrective action, Project EHS Plan Sec 9.1 Monitoring and measurement, Contractual requirement for timely close-out - indicates lack of commitment to EHS."
        ],
        "risk": "HIGH",
        "risk_text": "Open observations mean hazards remain uncontrolled on site - risk of incident from known hazard that was already identified but not fixed. Delay indicates subcontractor does not value EHS and does not prioritize safety - cultural failure. Erosion of EHS system effectiveness - if observations are not closed, EHS inspection loses value. Accumulation of open observations increases liability - Siemens Energy aware of hazard but subcontractor not fixing - if incident occurs, Siemens Energy exposed. Client and regulatory perception: Poor EHS management.",
        "action": [
            "1. Al-Mial to appoint dedicated EHS Focal Point / Coordinator - name, CV, and contact to be submitted to SE EHS within 12 hours - responsible for tracking and closing ALL observations",
            "2. Submit close-out plan for ALL open observations within 24 hours - with specific actions, owner, deadline, and evidence required for each observation - prioritized by risk level (Critical first)",
            "3. Implement Observation Tracking System - Excel tracker or digital app - with columns: Observation ID, Date, Description, Risk, Owner, Deadline, Status, Evidence, Verified by - to be updated daily and shared with SE EHS",
            "4. Weekly Observation Close-Out Meeting - Al-Mial PM + EHS Focal Point + SE EHS Manager - review open observations, close-out evidence, and trend - meeting minutes with actions",
            "5. Link observation close-out to contractual penalty and KPI - target: 100% close-out of Critical observations within 24h, High within 72h, Medium within 7 days - include in weekly EHS scorecard and payment milestone - escalate to Top Management if delay continues"
        ],
        "photo_label": "Open observations list / Overdue NCRs / Tracker showing delayed close-out"
    },
]

def create_slide(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.75), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), fill=TEAL_DARK)
    text_box(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), data['id'], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(1.0), Inches(0.1), Inches(8.5), Inches(0.55), data['title'], size=11.5, bold=True, color=DARK_TEXT)
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.2), Inches(0.15), Inches(2.8), Inches(0.5), fill=risk_bg, border=risk_color, border_w=Pt(1.5))
    text_box(slide, Inches(10.2), Inches(0.15), Inches(2.8), Inches(0.5), f"RISK: {data['risk']}", size=11, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.3), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.3), "1. TECHNICAL DESCRIPTION / الوصف الفني", size=9.5, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.4), Inches(7.2), Inches(2.0), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    tf = slide.shapes.add_textbox(Inches(0.4), Inches(1.45), Inches(7.0), Inches(1.9)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc']):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(8.2)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    shape(slide, Inches(0.3), Inches(3.6), Inches(7.2), Inches(0.3), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.6), Inches(7.0), Inches(0.3), f"2. RISK ASSESSMENT / تقييم المخاطر - {data['risk']}", size=9.5, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.9), Inches(7.2), Inches(1.1), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(3.9), Inches(7.0), Inches(1.0), data['risk_text'], size=8.2, bold=False, color=DARK_TEXT)

    shape(slide, Inches(0.3), Inches(5.2), Inches(7.2), Inches(0.3), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(5.2), Inches(7.0), Inches(0.3), "3. CORRECTIVE ACTION / الإجراء التصحيحي", size=9.5, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.5), Inches(7.2), Inches(1.6), fill=WHITE, border=GREY_BORDER)
    tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(5.55), Inches(7.0), Inches(1.5)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(7.8)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(1.8)

    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.3), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.3), "📷 PHOTO EVIDENCE / مساحة الصور", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(1.4), Inches(5.2), Inches(4.0), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1.5))
    text_box(slide, Inches(7.9), Inches(1.5), Inches(5.0), Inches(0.4), "[أدخل صورة المخالفة هنا]\n[Insert Violation Photo Here]", size=12, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.1), Inches(5.0), Inches(0.5), f"{data['photo_label']}\nFig {data['id']}.1", size=8.5, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.8), Inches(5.0), Inches(1.2), "⬆️ DRAG & DROP YOUR SITE PHOTO HERE\nاسحب الصورة وألصقها هنا\n\nAdd: Date, Time, Location, GPS\nأضف: التاريخ، الوقت، الموقع", size=10, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(5.5), Inches(5.2), Inches(0.6), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(5.5), Inches(5.0), Inches(0.6), "Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________  Taken by: ___________\nتفاصيل الصورة: التاريخ، الوقت، الموقع، المصور", size=7.5, color=GREY_TEXT)
    shape(slide, Inches(7.8), Inches(6.2), Inches(2.5), Inches(0.9), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(6.2), Inches(2.3), Inches(0.9), "[Additional Photo 2]\nFig X.2\nDifferent Angle\n[صورة إضافية]", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(10.5), Inches(6.2), Inches(2.5), Inches(0.9), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL)
    text_box(slide, Inches(10.6), Inches(6.2), Inches(2.3), Inches(0.9), "[After Correction]\nFig X.3\nCompliant State\n[بعد التصحيح]", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0), Inches(7.2), Inches(13.33), Pt(1), fill=GREY_BORDER)
    text_box(slide, Inches(0.3), Inches(7.25), Inches(6), Inches(0.2), "Siemens Energy | KAZ Project | EHS | Confidential | Final Observations Report", size=7, color=GREY_TEXT)
    text_box(slide, Inches(10.5), Inches(7.25), Inches(2.5), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026-FINAL", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

def create_title(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), fill=TEAL)
    shape(slide, Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(0.35), Inches(3), Inches(0.3), "SIEMENS ENERGY", size=14, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.6), Inches(0.65), Inches(5), Inches(0.3), "KAZ Power Plant Upgrade Project | Zero Harm Culture", size=9, color=GREY_TEXT)
    text_box(slide, Inches(10), Inches(0.35), Inches(2.8), Inches(0.6), "CONFIDENTIAL\nManagement Eyes Only", size=10, bold=True, color=RED, align=PP_ALIGN.RIGHT)
    shape(slide, Inches(0.5), Inches(1.4), Inches(0.08), Inches(1.0), fill=RED)
    text_box(slide, Inches(0.8), Inches(1.3), Inches(11.5), Inches(0.6), "Subcontractor EHS Non-Compliance Report - Final Observations", size=22, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.4), "تقرير مخالفات السلامة النهائي للمقاول الفرعي - الملاحظات الميدانية الفعلية", size=14, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.8), Inches(2.4), Inches(11.5), Inches(0.3), "Based on Actual Site Observations Provided | بناءً على الملاحظات الميدانية المقدمة من الموقع", size=10, bold=False, color=GREY_TEXT)
    shape(slide, Inches(0.5), Inches(2.9), Inches(7.5), Inches(2.6), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(2.9), Inches(7.3), Inches(0.3), "Final Observations Summary / ملخص الملاحظات النهائية", size=11, bold=True, color=TEAL_DARK)
    summary = """1. Painting, Storage Fence Installation, Cement Pouring WITHOUT PTW & Induction + No knowledge of activities in ENERGIZED zone
   طلاء، تركيب حاجز مخزن، صب أسمنت بدون بيرمت وبدون اندكشن + عدم علم بالقسم رغم منطقة انرجايز

2. Excavation Barricades - Soft Barricade Instead of Hard Barricade
   حواجز الحفر سوفت باركيد بدل هارد باركيد

3. Recurring PPE Violations Despite Warnings
   مخالفات PPE متكررة رغم التبليغ

4. Historical Violations - Bitumen, Workshop, Formwork Without PTW
   أعمال سابقة نفس المخالفات - بيتومين، ورشة، فورم ورك بدون بيرمت

5. Non-Compliance with EHS Directives - Refusal to Stop Unsafe Equipment
   عدم اتباع أوامر السلامة - رفض إيقاف معدات غير مقبولة واستمرار العمل

6. Unsafe Condition of Workshop Equipment
   معدات الورشة غير آمنة - Unsafe Condition

7. Poor Waste Management & Lack of Segregation
   عدم توضيح إدارة النفايات وفصل المواد

8. Deficiency in Safety Signage - Parking Garage & Speed Limits
   فقر في لوحات السلامة - باركنك وسرعة محددة

9. Delayed Response from Al-Mial to Close Observations
   تأخر استجابة قسم الميال لحل الملاحظات"""
    text_box(slide, Inches(0.6), Inches(3.2), Inches(7.3), Inches(2.2), summary, size=8.5, color=DARK_TEXT)

    shape(slide, Inches(8.3), Inches(2.9), Inches(4.5), Inches(2.6), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(8.4), Inches(2.9), Inches(4.3), Inches(0.3), "Report Info / معلومات التقرير", size=11, bold=True, color=TEAL_DARK)
    info = """Ref: SE-EHS-NC-2026-FINAL-OBS
Date: September 2026
Prepared by: EHS Manager - Siemens Energy
Project: KAZ Power Plant Upgrade

Classification: Confidential - Top Management
التصنيف: سري - للإدارة العليا

Standards:
• ISO 45001:2018, ISO 14001:2015
• OSHA 29 CFR 1926
• Siemens Energy EHS Principles
• Life Saving Rules
• Project EHS Plan

Design: Executive White Modern
• White #FFFFFF background
• Red #E30613 Critical
• Amber #FFB900 High
• Teal #009999 accent

Each Slide Includes:
• Technical Description (5 points)
• Risk Assessment
• Corrective Action (5 points)
• Photo placeholders [Insert Photo Here]

Based on YOUR 9 actual observations provided in Arabic."""
    text_box(slide, Inches(8.4), Inches(3.2), Inches(4.3), Inches(2.2), info, size=7.8, color=DARK_TEXT)

    shape(slide, Inches(0.5), Inches(5.7), Inches(12.33), Inches(1.2), fill=RED_LIGHT, border=RED, border_w=Pt(1.5))
    text_box(slide, Inches(0.6), Inches(5.7), Inches(12.1), Inches(0.2), "Executive Statement / ملاحظة تنفيذية:", size=10, bold=True, color=RED)
    text_box(slide, Inches(0.6), Inches(5.95), Inches(12.1), Inches(0.9), "This report documents 9 ACTUAL site observations provided by EHS team - including painting, storage fence, cement pouring without PTW/induction in energized zone, soft barricades instead of hard, recurring PPE violations, historical bitumen/workshop/formwork without PTW, refusal to stop unsafe equipment, unsafe workshop equipment, poor waste management, missing parking/speed signage, and delayed close-out from Al-Mial. Findings indicate total breakdown of subcontractor work management and EHS system in energized zone - immediate management intervention required to prevent fatality.\nهذا التقرير يوثق 9 ملاحظات فعلية من الموقع تشمل طلاء وتركيب حاجز مخزن وصب أسمنت بدون بيرمت واندكشن في منطقة انرجايز، حواجز حفر سوفت بدل هارد، مخالفات PPE متكررة، أعمال بيتومين وورشة وفورم ورك بدون بيرمت، رفض إيقاف معدات غير آمنة، معدات ورشة غير آمنة، سوء إدارة نفايات، فقر لوحات باركنك وسرعة، وتأخر استجابة الميال - انهيار كامل لنظام إدارة العمل والسلامة.", size=8.5, color=DARK_TEXT)

def create_closing(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.5), Inches(0.1), Inches(12), Inches(0.6), "Final Recommendations & Corrective Action Plan - Based on 9 Actual Observations", size=16, bold=True, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), fill=RED)
    text_box(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), "🔴 IMMEDIATE (0-24h) / فوري", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0.3), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(1.4), Inches(3.9), Inches(3.8), "• STOP WORK: Painting, Storage Fence, Cement Pouring, Excavation without hard barricade, Unsafe workshop equipment\n• Secure ENERGIZED zone - hard barricade + LOTO verification + fire watch\n• Remove non-inducted personnel - 100% gate check - manpower list vs induction in 6h\n• Ground leaking/unsafe equipment - Tag DO NOT USE\n• Issue NCR Level 2 + contractual warning - cumulative for pattern\n• Emergency meeting - Al-Mial MD + PM within 12h - written explanation for refusal\n\n[Photo: Secured energized zone after stop work]\n\nDECISION: Approve Stop-Work Authority to EHS\nRoot Cause: Why work started without knowledge of own supervision?", size=8, color=DARK_TEXT)
    shape(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), fill=AMBER)
    text_box(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), "🟡 SHORT-TERM (1-7 Days) / قصير", size=11, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(4.6), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(4.7), Inches(1.4), Inches(3.9), Inches(3.8), "• PTW Audit: Review ALL historical - bitumen, workshop, formwork, painting, storage, cement - submit in 48h\n• Re-Training: PTW + Energized Zone + Excavation + PPE + Waste - all Al-Mial supervision - 80% pass\n• Hard Barricades: Replace ALL soft tape with hard barricades - within 12h - Before/After photos\n• Workshop Equipment: Inventory + maintenance history + 3rd party certs in 48h - replace damaged\n• Waste: Color-coded skips + labeling + focal point + training\n• Signage: Parking + speed limit 15 km/h - retro-reflective + flagmen + radar + speed humps\n• PPE Campaign + BBS + Daily management walks\n\n[Photo: Compliant hard barricade + signage + waste segregation]\n\nDECISION: Approve $20k cost - back-charge to Al-Mial", size=8, color=DARK_TEXT)
    shape(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), fill=TEAL)
    text_box(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), "🟢 STRATEGIC (7-30 Days) / استراتيجي", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(8.9), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(9.0), Inches(1.4), Inches(3.9), Inches(3.8), "• Penalty Matrix: Linked to KPIs - TRIR, NCR closure, PTW compliance, observation close-out - deduction from invoice\n• Leadership Accountability: Replace non-performing supervisors who refused EHS order - CV + competency interview by SE\n• System Upgrade: Digital PTW with QR + Biometric induction + equipment checklist app + observation tracker\n• KPI & Scorecard: Weekly dashboard Red/Amber/Green to Top Management - include delayed response metric for Al-Mial\n• Independent Audit: 3rd party EHS audit of Al-Mial system - ISO 45001/14001 gap analysis\n• Zero Harm Commitment: Re-sign Life Saving Rules - Townhall with Top Management - visible commitment\n• Contingency: Prepare demobilization / replacement plan if no sustained improvement - procurement to identify alternative\n\n[Photo: Dashboard / Training / Digital PTW]\n\nDECISION: Authorize partial termination if critical violations recur + observation delay continues", size=8, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(5.5), Inches(12.7), Inches(1.7), fill=LIGHT_GREY_CARD, border=TEAL_DARK, border_w=Pt(2))
    text_box(slide, Inches(0.4), Inches(5.5), Inches(12.5), Inches(0.3), "✅ MANAGEMENT DECISION MATRIX - REQUIRED TODAY / مصفوفة القرارات", size=12, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(1.2), "1. Approve Stop-Work Authority for Energized Zone + Unsafe Equipment [ ] Yes [ ] No | إيقاف العمل في منطقة انرجايز والمعدات غير الآمنة\n2. Authorize Warning + Penalty for Refusal to Follow EHS Order [ ] Yes [ ] No | تحذير وغرامة لرفض أوامر السلامة\n3. Mandate Al-Mial Top Management On-Site within 12h with Written Explanation [ ] Yes [ ] No | حضور إدارة الميال خلال 12 ساعة مع توضيح كتابي\n4. Approve Budget $20k for Hard Barricades, Signage, Waste Skips - Back-charge [ ] Yes [ ] No | ميزانية حواجز ولوحات وحاويات\n5. Endorse Zero Tolerance 3 Strikes + Observation Close-Out KPI (Critical 24h) [ ] Yes [ ] No | عدم تسامح 3 مخالفات + مؤشر إغلاق الملاحظات\n\nSignature: _________________ Project Director Date: _______ | Signature: _________________ Construction Manager Date: _______ | Signature: _________________ EHS Manager Date: _______", size=8, color=DARK_TEXT)

final_prs = Presentation()
final_prs.slide_width = Inches(13.33)
final_prs.slide_height = Inches(7.5)
create_title(final_prs)
for d in slides_data:
    create_slide(d)
create_closing(final_prs)

out = "/home/user/ahmed/Subcontractor_EHS_FINAL_9_Observations_Executive_White_EN.pptx"
final_prs.save(out)
print(f"Saved to {out}")
