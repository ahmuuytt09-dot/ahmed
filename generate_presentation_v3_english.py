#!/usr/bin/env python3
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

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

def text_box(slide, l, t, w, h, txt, size=11, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
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

slides_info = [
    {
        "id": "01",
        "title": "01 - UNAUTHORIZED WORK - WITHOUT PERMIT TO WORK (PTW)",
        "obs": [
            "Location: [Insert Area] | Date: [Insert Date] | Time: [Insert Time]",
            "Activity: Excavation / Mechanical / Hot Work executed without approved PTW",
            "Violation: No TRA review, no energy isolation verification, no SIMOPS coordination",
            "PTW logbook not maintained, permit not displayed at work front as required",
            "Reference: Siemens Energy EHS General Requirements Sec 5.3, KAZ EHS Plan Sec 7.2, ISO 45001 Cl 8.1.2"
        ],
        "risk": "CRITICAL",
        "risk_text": "Potential fatality, uncontrolled hazardous energy release, fire/explosion, SIMOPS conflict. Total loss of primary control barrier. Authority may issue prohibition notice.",
        "action": [
            "1. IMMEDIATE STOP WORK - secure area and remove workforce",
            "2. Mandatory PTW re-training for all supervisors within 24h - competency test",
            "3. Issue formal NCR Level 1 + contractual warning letter",
            "4. Daily PTW audit and spot-checks for 14 days - EHS verification",
            "5. Enforce 3-strikes removal policy for non-compliant supervisors"
        ],
        "photos": ["PHOTO 1: Work activity without PTW displayed at work front", "PHOTO 2: PTW logbook - missing entry / permit not available"]
    },
    {
        "id": "02",
        "title": "02 - SITE ACCESS & INDUCTION - PERSONNEL WITHOUT EHS INDUCTION",
        "obs": [
            "No. of persons: [8-12] found working without induction card / sticker",
            "Gate register vs actual headcount mismatch - onboarding register incomplete",
            "No induction verification at main gate - Security-EHS interface failure",
            "Violation: KAZ Project EHS Plan Sec 6.1 Competency & Training, local labor law",
            "Systemic failure in access control and subcontractor self-verification"
        ],
        "risk": "CRITICAL",
        "risk_text": "Workers unaware of site hazards, emergency assembly points, muster procedures, life-saving rules. High probability of unsafe acts, delayed emergency response, regulatory stoppage.",
        "action": [
            "1. Immediate removal of all non-inducted personnel from site - escort by Security",
            "2. Implement 100% gate check - Joint Security + EHS verification - No card, no entry",
            "3. Subcontractor to submit full manpower list vs induction records within 12h + gap closure plan",
            "4. Introduce QR / Biometric induction tracking linked to access control",
            "5. Formal warning + back-charge for re-induction sessions"
        ],
        "photos": ["PHOTO 1: Workers without induction cards - site overview", "PHOTO 2: Gate register vs induction cards verification"]
    },
    {
        "id": "03",
        "title": "03 - PPE NON-COMPLIANCE - FAILURE TO WEAR MANDATORY PPE",
        "obs": [
            "PPE missing: Helmet chin strap not secured, safety glasses, gloves, hi-vis vest not worn",
            "Location: [Insert Location] | Task: [Insert Task] | Date: [Insert Date]",
            "Subcontractor supervision present on site but failing to enforce compliance",
            "PPE available in store - confirms behavioral and supervisory failure, not availability",
            "Violation: Siemens Energy Life Saving Rule #3, PPE Standard EN 397, EN 166, ANSI Z89.1"
        ],
        "risk": "HIGH",
        "risk_text": "Immediate risk of head injury, eye injury, hand lacerations, struck-by incidents. Demonstrates weak safety culture. Significantly increases LTIFR/TRIR.",
        "action": [
            "1. On-the-spot correction + toolbox talk - documented with photos and attendance",
            "2. Issue Safety Observation Cards + formal EHS violation notice to individual and supervisor",
            "3. Mandatory PPE compliance campaign - daily checks at start of shift + random checks",
            "4. Link PPE compliance to subcontractor KPI and payment milestone - weekly scorecard",
            "5. Supervisor accountability briefing - 3-strikes removal policy signed"
        ],
        "photos": ["PHOTO 1: PPE violation - wide shot showing multiple workers", "PHOTO 2: PPE violation - close-up + supervisor present not enforcing"]
    },
    {
        "id": "04",
        "title": "04 - EXCAVATION SAFETY - INADEQUATE HARD BARRICADES & PROTECTION",
        "obs": [
            "Depth: [1.2m - 2.5m observed] - No hard barricade, only loose warning tape (prohibited for >1.2m)",
            "No edge protection, no safe access/egress (ladders missing), spoil pile within 1m of edge",
            "No reflective markers, delineators, or lighting for night-time visibility",
            "Violation: OSHA 29 CFR 1926 Subpart P, ISO 45001 Excavation Standard, Siemens SE-GEN-EHS-007",
            "Night work risk: Excavation invisible - vehicle fall hazard"
        ],
        "risk": "CRITICAL",
        "risk_text": "High potential for fatal fall into excavation, cave-in engulfment, personnel/vehicle fall. Excavation is High-Risk activity with fatality history - classified as Major Accident Hazard.",
        "action": [
            "1. IMMEDIATE hard barricading - scaffold tubes / concrete barriers / hard fencing around all excavations",
            "2. Provide compliant access ladders every 7.5m + keep spoil minimum 1.5m away + toe boards",
            "3. Excavation permit re-validation with geotechnical check and competent person inspection",
            "4. Install reflective delineators + solar blinker lights for low-light hours",
            "5. Competent person inspection checklist daily before start - signed by engineer + verified by SE EHS"
        ],
        "photos": ["PHOTO 1: Excavation without hard barricade - full view with tape only", "PHOTO 2: Spoil pile too close to edge + no ladder - close view"]
    },
    {
        "id": "05",
        "title": "05 - DIRECTIVES NON-COMPLIANCE - FAILURE TO FOLLOW EHS INSTRUCTIONS",
        "obs": [
            "Date of instruction: [Insert Date] | Instruction: [Insert EHS instruction details]",
            "Repeated failure to comply with verbal and written EHS instructions - documented in daily reports + email trail",
            "Toolbox Talks (TBT) not conducted or without documentation / attendance sheets / task relevance",
            "Previous NCRs closed without evidence or verification - falsification of closure",
            "Violation: ISO 45001 Cl 7.3 Awareness & 8.1 Operational Control - Willful disregard"
        ],
        "risk": "HIGH",
        "risk_text": "Erosion of EHS authority and safety culture. Creates precedent for further violations, normalizes deviance. If directives optional, all critical controls (LOTO, WAH) at risk of failure. Systemic cultural failure.",
        "action": [
            "1. Escalate to Project Management - formal letter to Subcontractor Top Management - require written commitment",
            "2. Implement Daily EHS Directives Log - all instructions logged with date/time, owner, sign-off by PM + SE EHS",
            "3. Hold EHS Leadership stand-down meeting with all foremen - documented commitment signed",
            "4. Link non-compliance to contractual penalty clause - financial consequence",
            "5. Weekly EHS compliance scorecard Red/Amber/Green reviewed in project management meeting"
        ],
        "photos": ["PHOTO 1: Evidence of ignored instruction (e.g., open excavation still not barricaded)", "PHOTO 2: TBT register missing / NCR closure without evidence"]
    },
    {
        "id": "06",
        "title": "06 - TRAFFIC & SIGNAGE - POOR / DAMAGED / MISSING SAFETY SIGNAGE",
        "obs": [
            "Missing / damaged: Speed limit signs (15 km/h), parking signage, pedestrian walkway markings",
            "Vehicle-pedestrian interface unmanaged - no flagmen at high-risk intersections and blind spots",
            "Speeding observed: 30 km/h in construction zone (limit 15 km/h) - confirmed by radar gun",
            "Pedestrian routes not marked or segregated - workers walking on vehicle routes",
            "Violation: Project Traffic Management Plan (TMP), ISO 39001 Road Traffic Safety"
        ],
        "risk": "HIGH",
        "risk_text": "High risk of vehicle-pedestrian collision, equipment damage, fatality - vehicle incidents are top 3 fatality causes. Poor signage increases liability. Emergency vehicle access may be impeded.",
        "action": [
            "1. Immediate reinstatement of all traffic signs per approved TMP - retro-reflective grade, proper height",
            "2. Establish dedicated pedestrian routes with hard segregation - Jersey barriers + green walkway paint",
            "3. Deploy trained flagmen at intersections + implement speed monitoring - radar checks 2x daily",
            "4. Conduct traffic safety campaign and driver re-induction - all drivers including subcontractor",
            "5. EHS + Logistics joint daily inspection of traffic controls - checklist signed"
        ],
        "photos": ["PHOTO 1: Missing / damaged speed limit sign - location view", "PHOTO 2: No pedestrian segregation + vehicle-pedestrian interface risk"]
    },
    {
        "id": "07",
        "title": "07 - WASTE MANAGEMENT - LACK OF SEGREGATION & LABELING",
        "obs": [
            "Waste containers without segregation - hazardous, recyclable, general waste mixed in same skip",
            "No clear labeling, color-coding, or MSDS reference for hazardous waste bins",
            "Oil-contaminated rags and filters disposed in general waste - environmental violation + fire hazard",
            "Location: [Insert waste area location] | Date observed: [Insert Date]",
            "Violation: ISO 14001:2015 Cl 8.1, Project Waste Management Plan, local EPA regulations"
        ],
        "risk": "HIGH",
        "risk_text": "Environmental pollution risk (soil/groundwater contamination), fire hazard from incompatible mixing, regulatory fine, reputational damage to Siemens Energy. Cross-contamination makes recycling impossible - sustainability KPI failure.",
        "action": [
            "1. Immediate provision of color-coded, labeled waste skips per SE Standard: Black-General, Green-Recyclable, Red-Hazardous",
            "2. Clean-up and correct segregation of existing mixed waste by competent team under EHS supervision",
            "3. Training on Waste Management - hazardous waste handling, spill prevention, manifest system",
            "4. Appoint Waste Management Focal Point from subcontractor - responsible for daily housekeeping",
            "5. Weekly waste audit - track volumes and disposal manifests - report in monthly environmental report"
        ],
        "photos": ["PHOTO 1: Mixed waste in one container - overview showing hazardous + general", "PHOTO 2: Oil-contaminated rags in general waste - close-up + fire hazard"]
    },
    {
        "id": "08",
        "title": "08 - EQUIPMENT INTEGRITY - ACTIVE OIL & FUEL LEAKS FROM MACHINERY",
        "obs": [
            "Equipment: Excavator CAT 320 [ID: ___] / Loader [ID: ___] - active hydraulic oil and fuel leaks observed dripping",
            "Drip trays not provided - soil contamination ~5 sqm observed under equipment parking area",
            "No preventive maintenance records on site - daily inspection checklists falsified - same signature for 7 days",
            "Violation: Equipment Safety Standard, PUWER, Environmental Protection, SE Equipment Standard",
            "Operator competency: No evidence of daily checks by competent operator"
        ],
        "risk": "CRITICAL",
        "risk_text": "Fire and explosion hazard (fuel on hot surfaces), slip hazard, soil/groundwater contamination, equipment failure leading to loss of control. Breakdown of maintenance management system. Breach of ISO 14001 and PUWER.",
        "action": [
            "1. IMMEDIATE grounding of leaking equipment - Tag 'DO NOT USE - EHS HOLD' + keys removed",
            "2. Contaminated soil to be excavated and disposed as hazardous waste - spill response with absorbents",
            "3. Subcontractor to submit full maintenance history + third-party inspection certificates within 48h - No cert, no operation",
            "4. Implement mandatory daily equipment checklist - verified by SE EHS/Mechanical + photo evidence of no leaks",
            "5. Introduction of drip trays and spill kits at all equipment parking zones - 100% coverage"
        ],
        "photos": ["PHOTO 1: Active oil/fuel leak under equipment - dripping with circle highlight", "PHOTO 2: Soil contamination under parking area + missing drip tray"]
    },
    {
        "id": "09",
        "title": "09 - HAZARDOUS MATERIAL STORAGE - DIESEL STORAGE WITHOUT SECONDARY CONTAINMENT",
        "obs": [
            "Diesel storage: Approx. 2000L (2x IBC tanks) - no secondary containment / bunding - 110% capacity required",
            "No isolation distance from ignition sources (welding 10m away) - no fire extinguisher, no HAZCHEM signage, no spill kit",
            "Storage on bare ground - no impervious base, no earthing/bonding for fuel transfer - static electricity risk",
            "Direct violation: HSG 176 Storage of Flammable Liquids, EPA SPCC, SE Hazardous Materials Standard SE-ENV-003",
            "Location: [Insert location] - Near drains / buildings - environmental receptor risk"
        ],
        "risk": "CRITICAL",
        "risk_text": "Major fire/explosion risk, catastrophic environmental release, soil/groundwater contamination leading to prosecution. Potential Major Accident Hazard (MAH) - pool fire affecting adjacent works. Immediate danger to life and environment.",
        "action": [
            "1. STOP all fueling operations - secure area with fire watch and 15m exclusion zone until compliant",
            "2. Provide compliant bunded storage: 110% secondary containment, impervious concrete floor, roofed, ventilated, 15m from drains/buildings",
            "3. Install fire extinguishers (2x 9kg DCP + 1x foam), HAZCHEM signage, spill kit 150L, earthing/bonding system tested",
            "4. Develop and communicate Fuel Storage & Handling SOP + conduct emergency response drill (spill + fire)",
            "5. Location to be approved by SE EHS - minimum 15m from excavations/buildings/drains - risk assessment required"
        ],
        "photos": ["PHOTO 1: Diesel tanks without bunding - overview showing 2x IBC", "PHOTO 2: Bare ground + no fire extinguisher / no signage / no spill kit - close-up"]
    },
    {
        "id": "10",
        "title": "10 - RECURRING VIOLATIONS - PERSISTENT PPE VIOLATIONS DESPITE WARNINGS",
        "obs": [
            "Continued and repeated PPE violations after 3 formal warnings, 5 Safety Observation Cards, 2 NCRs in last 14 days",
            "Same workgroup and supervision involved - indicates intentional non-compliance, not lack of awareness",
            "Trend analysis from Daily EHS Reports: 60% increase in PPE violations in last 2 weeks (5 to 12 per week)",
            "Subcontractor management failing to enforce disciplinary procedure per contract Appendix EHS - no action taken",
            "Demonstrates normalization of deviance - systemic failure, not individual mistake"
        ],
        "risk": "CRITICAL",
        "risk_text": "Normalization of deviance - precursor to serious injury/fatality (Heinrich/Bird triangle). Total failure of subcontractor safety leadership and culture. If PPE (last line of defense per hierarchy) ignored, higher-level controls will also fail. Project reputation and Siemens Energy Zero Harm commitment at direct risk.",
        "action": [
            "1. Invoke contractual penalty - stop work for repeat offender crew + mandatory leadership engagement - Subcontractor PM presents corrective plan to SE Management",
            "2. Remove and replace non-compliant supervision - require CVs + EHS competency assessment + interview by SE EHS",
            "3. Implement Behavior-Based Safety (BBS) program with positive/negative reinforcement - daily observations + feedback + recognition",
            "4. Daily management safety walks - Siemens Energy PM + Subcontractor PM - joint walk and documented actions",
            "5. Consider partial termination / back-charge for additional EHS supervision if no improvement in 7 days - prepare contingency replacement crew"
        ],
        "photos": ["PHOTO 1: Repeat PPE violation - same crew - evidence of intentional non-compliance", "PHOTO 2: Trend chart / NCR log showing 60% increase - screenshot from Daily Reports"]
    },
]

def create_boxed_slide(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DEEP_PETROL)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.9), fill=DARK_PETROL)
    shape(slide, Inches(0), Inches(0.9), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0.2), Inches(0.15), Inches(0.7), Inches(0.6), fill=TEAL)
    text_box(slide, Inches(0.2), Inches(0.15), Inches(0.7), Inches(0.6), f"{data['id']}", size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(1.0), Inches(0.1), Inches(9), Inches(0.7), data['title'], size=15, bold=True, color=WHITE)
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_text_color = WHITE if data['risk'] == "CRITICAL" else DARK_PETROL
    shape(slide, Inches(10.5), Inches(0.15), Inches(2.6), Inches(0.6), fill=risk_color)
    text_box(slide, Inches(10.5), Inches(0.15), Inches(2.6), Inches(0.6), f"RISK: {data['risk']}", size=13, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    # BOX 1
    shape(slide, Inches(0.2), Inches(1.2), Inches(5.5), Inches(2.8), fill=CARD_BG, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(0.3), Inches(1.2), Inches(5.3), Inches(0.3), "BOX 1: NON-COMPLIANCE DETAILS - Insert Location/Date/Time", size=9, bold=True, color=TEAL_LIGHT)
    tf_box = slide.shapes.add_textbox(Inches(0.3), Inches(1.5), Inches(5.3), Inches(2.4)).text_frame
    tf_box.word_wrap = True
    for i, line in enumerate(data['obs']):
        if i==0:
            p = tf_box.paragraphs[0]
        else:
            p = tf_box.add_paragraph()
        p.text = line
        p.font.size = Pt(9)
        p.font.color.rgb = WHITE
        p.font.name = "Calibri"
        p.space_after = Pt(2.5)

    # BOX 2 Photo 1
    shape(slide, Inches(5.9), Inches(1.2), Inches(3.6), Inches(2.8), fill=RGBColor(0x0A,0x22,0x2B), border=GREY2, border_w=Pt(1))
    text_box(slide, Inches(6.0), Inches(1.2), Inches(3.4), Inches(0.3), "BOX 2: PHOTO EVIDENCE 1", size=9, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(6.0), Inches(1.55), Inches(3.4), Inches(2.2), f"{data['photos'][0]}\n\n[ INSERT PHOTO HERE ]\nDRAG & DROP YOUR SITE PHOTO\n\nFig {data['id']}.1\nAdd: Date, Time, Location, GPS", size=8.5, bold=False, color=GREY2, align=PP_ALIGN.CENTER)

    # BOX 3 Photo 2
    shape(slide, Inches(9.7), Inches(1.2), Inches(3.4), Inches(2.8), fill=RGBColor(0x0A,0x22,0x2B), border=GREY2, border_w=Pt(1))
    text_box(slide, Inches(9.8), Inches(1.2), Inches(3.2), Inches(0.3), "BOX 3: PHOTO EVIDENCE 2", size=9, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(9.8), Inches(1.55), Inches(3.2), Inches(2.2), f"{data['photos'][1]}\n\n[ INSERT PHOTO HERE ]\nDRAG & DROP YOUR SITE PHOTO\n\nFig {data['id']}.2\nAdd: Close-up / Different Angle", size=8.5, bold=False, color=GREY2, align=PP_ALIGN.CENTER)

    # BOX 4 Risk
    risk_bg = RGBColor(0x2A,0x0A,0x0A) if data['risk']=="CRITICAL" else RGBColor(0x2B,0x20,0x00)
    shape(slide, Inches(0.2), Inches(4.2), Inches(5.5), Inches(1.3), fill=risk_bg, border=risk_color, border_w=Pt(2))
    text_box(slide, Inches(0.3), Inches(4.2), Inches(5.3), Inches(0.3), f"BOX 4: RISK LEVEL & IMPACT - {data['risk']}", size=9, bold=True, color=risk_color)
    text_box(slide, Inches(0.3), Inches(4.5), Inches(5.3), Inches(0.9), data['risk_text'], size=9, bold=False, color=LIGHT_GREY)

    # BOX 5 Corrective
    shape(slide, Inches(5.9), Inches(4.2), Inches(7.2), Inches(2.8), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(6.0), Inches(4.2), Inches(5.0), Inches(0.3), "BOX 5: REQUIRED CORRECTIVE ACTION", size=9, bold=True, color=DARK_PETROL)
    tf2 = slide.shapes.add_textbox(Inches(6.0), Inches(4.5), Inches(4.5), Inches(2.4)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(9)
        p.font.color.rgb = DARK_PETROL
        p.font.name = "Calibri"
        p.space_after = Pt(2.5)

    shape(slide, Inches(10.7), Inches(4.6), Inches(2.3), Inches(2.2), fill=RGBColor(0xF0,0xF0,0xF0), border=GREY2)
    text_box(slide, Inches(10.8), Inches(4.6), Inches(2.1), Inches(0.3), "STANDARD REF", size=8, bold=True, color=DARK_PETROL)
    std_text = "• ISO 45001:2018\n• ISO 14001:2015\n• OSHA 29 CFR 1926\n• SE EHS Contractor Req\n• Project EHS Plan\n• Local EPA / HSG 176\n• SE-GEN-EHS-007"
    text_box(slide, Inches(10.8), Inches(4.9), Inches(2.1), Inches(1.8), std_text, size=7.5, color=DARK_PETROL)

    shape(slide, Inches(0), Inches(7.15), Inches(13.33), Pt(1), fill=GREY2)
    text_box(slide, Inches(0.2), Inches(7.2), Inches(6), Inches(0.2), "Siemens Energy | Zero Harm | Confidential | Boxed Format - 5 Boxes + 2 Photo Placeholders Per Slide", size=7, color=GREY2)
    text_box(slide, Inches(11), Inches(7.2), Inches(2), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY2, align=PP_ALIGN.RIGHT)

# TITLE SLIDE
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DEEP_PETROL)
shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(6), fill=TEAL)
text_box(slide, Inches(0.5), Inches(0.3), Inches(4), Inches(0.4), "SIEMENS ENERGY", size=16, bold=True, color=TEAL_LIGHT)
text_box(slide, Inches(9.5), Inches(0.3), Inches(3.5), Inches(0.4), "CONFIDENTIAL", size=11, bold=True, color=AMBER, align=PP_ALIGN.RIGHT)
shape(slide, Inches(0.5), Inches(1.0), Inches(3), Pt(3), fill=AMBER)
text_box(slide, Inches(0.5), Inches(1.4), Inches(12), Inches(1.0), "SUBCONTRACTOR EHS NON-COMPLIANCE & SAFETY VIOLATION REPORT", size=26, bold=True, color=WHITE)
text_box(slide, Inches(0.5), Inches(2.4), Inches(12), Inches(0.4), "BOXED FORMAT WITH PHOTO PLACEHOLDERS - ENGLISH VERSION", size=14, bold=False, color=TEAL_LIGHT)
text_box(slide, Inches(0.5), Inches(2.9), Inches(6), Inches(2.0), "Project: KAZ Power Plant Upgrade Project\nReference: SE-EHS-NC-2026-001 BOXED V3 ENGLISH\nDate: September 2026\nPrepared by: EHS Manager - Siemens Energy\nClassification: Confidential - Management Eyes Only\n\nFORMAT OVERVIEW:\n• 12 Slides Total\n• Each Observation Slide = 5 Distinct Boxes\n• Box 1: Non-Compliance Details (with location/date fields)\n• Box 2 & 3: Photo Evidence Placeholders - 2 photos per violation\n• Box 4: Risk Level & Impact (Red=Critical, Amber=High)\n• Box 5: Required Corrective Action + Standard Reference\n\nSTANDARDS: ISO 45001:2018, ISO 14001:2015, OSHA, SE EHS Requirements", size=10, color=LIGHT_GREY)

shape(slide, Inches(7.0), Inches(2.9), Inches(5.8), Inches(3.5), fill=CARD_BG, border=TEAL)
text_box(slide, Inches(7.1), Inches(2.9), Inches(5.6), Inches(0.3), "HOW TO USE THIS TEMPLATE - INSTRUCTIONS", size=11, bold=True, color=AMBER)
guide = """1. For each violation slide, fill Box 1 with specific details:
   - Location, Date, Time, Equipment ID
   - Replace [Insert...] placeholders

2. Box 2 & 3 - Photo Placeholders:
   - Drag & drop your site photos directly into the boxes
   - Photo 1: Overview / wide shot
   - Photo 2: Close-up / different angle
   - Add Fig number, date, GPS coordinates

3. Box 4 - Risk is pre-filled - color-coded:
   - RED = CRITICAL (fatality potential)
   - AMBER = HIGH (serious injury / env)

4. Box 5 - Corrective actions are actionable:
   - Assign owner and deadline
   - Track closure in NCR system

5. Final Slide: Action Plan with 4 Boxes:
   - Immediate (0-24h), Short (1-7d), Strategic (7-30d)
   - Management Decision Matrix with Yes/No checkboxes

VISUAL DESIGN:
• Background: Deep Petrol #0E2F3E
• Boxes: Card #143D4E + White for actions
• Photo boxes: Dark #0A222B with dashed border effect
• All boxes have titles for clarity"""
text_box(slide, Inches(7.1), Inches(3.2), Inches(5.6), Inches(3.1), guide, size=8.5, color=WHITE)

text_box(slide, Inches(0.5), Inches(6.8), Inches(12), Inches(0.3), "Zero Harm is Non-Negotiable | 5 Boxes Per Slide | 2 Photo Placeholders Per Violation | English Version | Ready for Management Review", size=9, bold=True, color=AMBER, align=PP_ALIGN.CENTER)

for info in slides_info:
    create_boxed_slide(info)

# ACTION PLAN
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, DEEP_PETROL)
shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.8), fill=DARK_PETROL)
shape(slide, Inches(0), Inches(0.8), Inches(13.33), Pt(4), fill=TEAL)
text_box(slide, Inches(0.3), Inches(0.1), Inches(12), Inches(0.6), "ACTION PLAN & MANAGEMENT DECISION MATRIX - BOXED FORMAT", size=17, bold=True, color=WHITE)

shape(slide, Inches(0.2), Inches(1.1), Inches(4.2), Inches(0.4), fill=RED)
text_box(slide, Inches(0.2), Inches(1.1), Inches(4.2), Inches(0.4), "BOX A: IMMEDIATE (0-24 HOURS)", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(0.2), Inches(1.5), Inches(4.2), Inches(4.0), fill=WHITE)
text_box(slide, Inches(0.3), Inches(1.5), Inches(4.0), Inches(4.0), "• STOP WORK all critical violations (PTW, Excavation, Diesel, Leaking Equipment)\n• Secure hazardous areas - hard barricade + fire watch + 15m exclusion\n• Remove all non-inducted personnel - 100% gate verification\n• Ground all leaking equipment - Tag DO NOT USE\n• Issue formal NCR Level 1 & 2 + contractual warning letter\n• Convene Emergency EHS Meeting - Subcontractor MD/PM within 24h\n\nPHOTO BOX: [Insert photo of stop work order / secured area]\n\nDECISION REQUIRED:\nApprove Stop-Work Authority to EHS", size=8.5, color=DARK_PETROL)

shape(slide, Inches(4.6), Inches(1.1), Inches(4.2), Inches(0.4), fill=AMBER)
text_box(slide, Inches(4.6), Inches(1.1), Inches(4.2), Inches(0.4), "BOX B: SHORT-TERM (1-7 DAYS)", size=10, bold=True, color=DARK_PETROL, align=PP_ALIGN.CENTER)
shape(slide, Inches(4.6), Inches(1.5), Inches(4.2), Inches(4.0), fill=WHITE)
text_box(slide, Inches(4.7), Inches(1.5), Inches(4.0), Inches(4.0), "• Full PTW System Audit + Re-Training - all supervisors certified - 80% pass mark\n• Compliant diesel storage - 110% bund + impervious floor + extinguishers\n• Reinstate traffic signage - retro-reflective + pedestrian segregation + flagmen\n• Waste management system - color-coded skips + labeling + focal point\n• Excavation compliance - hard barricades + ladders every 7.5m + competent person checks\n• PPE Campaign + BBS Launch - daily checks + scorecard\n• EHS Directives Log - daily sign-off by PM + EHS\n\nPHOTO BOX: [Insert after-correction photos - compliant state]\n\nDECISION: Approve $15k cost - back-charge to subcontractor", size=8.5, color=DARK_PETROL)

shape(slide, Inches(9.0), Inches(1.1), Inches(4.1), Inches(0.4), fill=TEAL)
text_box(slide, Inches(9.0), Inches(1.1), Inches(4.1), Inches(0.4), "BOX C: STRATEGIC (7-30 DAYS)", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(9.0), Inches(1.5), Inches(4.1), Inches(4.0), fill=WHITE)
text_box(slide, Inches(9.1), Inches(1.5), Inches(3.9), Inches(4.0), "• Contractual Enforcement: Penalty matrix linked to KPIs (TRIR, NCR closure, PTW compliance)\n• Leadership Accountability: Replace non-performing supervisors - CV + competency interview\n• System Upgrade: QR/Biometric induction + digital PTW + equipment checklist app\n• KPI & Scorecard: Weekly dashboard Red/Amber/Green to Top Management\n• Independent Audit: 3rd party EHS audit ISO 45001/14001 gap analysis\n• Zero Harm Commitment: Re-sign Life Saving Rules - Townhall with Top Management\n• Contingency: Prepare demobilization / replacement plan if no improvement\n\nPHOTO BOX: [Insert dashboard / training photos]\n\nDECISION: Authorize partial termination if critical violations recur", size=8.5, color=DARK_PETROL)

shape(slide, Inches(0.2), Inches(5.7), Inches(13.0), Inches(1.5), fill=CARD_BG, border=AMBER, border_w=Pt(2))
text_box(slide, Inches(0.3), Inches(5.7), Inches(12.8), Inches(0.3), "BOX D: MANAGEMENT DECISION MATRIX - REQUIRED TODAY", size=11, bold=True, color=AMBER)
text_box(slide, Inches(0.3), Inches(6.0), Inches(12.8), Inches(1.0), "1. Approve IMMEDIATE Stop-Work Authority to EHS for Critical Violations [ ] Yes [ ] No  |  2. Authorize Formal Contractual Warning & Penalty Enforcement [ ] Yes [ ] No  |  3. Mandate Subcontractor Top Management Presence On-Site within 48h [ ] Yes [ ] No  |  4. Approve Budget for Additional EHS Controls (Bunded Storage, Signage, Barricades) - Back-charge to Subcontractor [ ] Yes [ ] No  |  5. Endorse Zero Tolerance Policy: 3 Strikes = Removal - Communicate in Townhall [ ] Yes [ ] No\n\nSignature: _________________  Project Director  Date: _______  |  Signature: _________________  Construction Manager  Date: _______  |  Signature: _________________  EHS Manager  Date: _______", size=8.5, color=WHITE)

out = "/home/user/ahmed/Subcontractor_EHS_Boxed_V3_ENGLISH_ONLY.pptx"
prs.save(out)
print(f"Saved to {out}")
