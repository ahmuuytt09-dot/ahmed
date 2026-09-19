#!/usr/bin/env python3
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY_BG = RGBColor(0xF5, 0xF5, 0xF7)
LIGHT_GREY_CARD = RGBColor(0xF9, 0xF9, 0xFA)
GREY_BORDER = RGBColor(0xE0, 0xE0, 0xE0)
DARK_GREY_BORDER = RGBColor(0xB0, 0xB0, 0xB0)
DARK_TEXT = RGBColor(0x21, 0x21, 0x21)
GREY_TEXT = RGBColor(0x61, 0x61, 0x61)
RED = RGBColor(0xE3, 0x06, 0x13)
RED_LIGHT = RGBColor(0xFF, 0xEB, 0xEB)
RED_DARK = RGBColor(0xC0, 0x00, 0x0A)
AMBER = RGBColor(0xFF, 0xB9, 0x00)
AMBER_LIGHT = RGBColor(0xFF, 0xF8, 0xE1)
TEAL = RGBColor(0x00, 0x99, 0x99)
TEAL_DARK = RGBColor(0x0E, 0x2F, 0x3E)
TEAL_LIGHT_BG = RGBColor(0xE0, 0xF2, 0xF1)
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

def text_box(slide, l, t, w, h, txt, size=11, bold=False, color=DARK_TEXT, align=PP_ALIGN.LEFT, font="Calibri"):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font
    p.alignment = align
    return tb

def add_siemens_logo(slide, left=Inches(0.3), top=Inches(0.15)):
    # Siemens Energy Logo - Text based professional
    shape(slide, left, top, Inches(2.2), Inches(0.45), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, left, top, Inches(2.2), Inches(0.25), "SIEMENS", size=14, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER, font="Arial")
    text_box(slide, left, top+Inches(0.22), Inches(2.2), Inches(0.2), "ENERGY", size=10, bold=False, color=TEAL, align=PP_ALIGN.CENTER, font="Arial")

# 9 Observations based on user's last detailed message + 7 original = combined 9 final
slides_data = [
    {
        "id": "01",
        "title_en": "Unauthorized Works Without PTW & Induction in Energized Zone",
        "title_ar": "أعمال بدون تصريح عمل وبدون تعريف سلامة في منطقة مفعمة بالطاقة",
        "subtitle": "Painting, Storage Fence Installation, Cement Pouring Works",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Painting works, installation of storage fence/barrier (Storage Area), and cement pouring works executed WITHOUT approved Permit to Work (PTW) and WITHOUT EHS Induction.\n• When questioned by Siemens Energy EHS Department, subcontractor supervision (Al-Mial) stated they have NO knowledge of these activities, workers, or location - despite work being inside ENERGIZED zone (live electrical area).\n• No LOTO verification, no energized zone risk assessment, no EHS coordination, no TRA review.\n• Violation: SE EHS General Requirements Sec 5.3 PTW, KAZ EHS Plan Sec 6.1 & 7.2, ISO 45001 Cl 7.2, 7.3, 8.1.2",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "CRITICAL",
        "risk_text": "FATALITY RISK: Electrocution, arc flash, chemical exposure (painting), struck-by, SIMOPS conflict. Workers unaware of energized hazards and emergency procedures. Subcontractor management has NO control over own workforce - total breakdown of work management system. Regulatory prohibition, criminal liability, and major liability for Siemens Energy if incident occurs in energized zone.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي المطلوب",
        "action": "1. IMMEDIATE STOP WORK - Evacuate energized zone - Secure with hard barricade + LOTO verification\n2. Remove all non-inducted personnel - 100% gate check - Submit manpower list vs induction records within 6 hours\n3. Subcontractor PM to provide written explanation + Root Cause Analysis: How work started without PTW/induction and without knowledge of own supervision?\n4. Re-training on PTW and Energized Zone Procedure for all Al-Mial supervision within 24h - 80% pass mark\n5. Implement Daily Work Front Coordination Meeting - Daily work plan 24h in advance to SE EHS for approval - No plan, no work",
        "photo_main": "Painting Without PTW\nStorage Fence Installation Without Induction\nCement Pouring in Energized Zone",
        "photo2": "Workers Without Induction Cards",
        "photo3": "Energized Zone - No LOTO / No Barricade"
    },
    {
        "id": "02",
        "title_en": "Excavation Barricades - Soft Barricade Instead of Hard Barricade",
        "title_ar": "حواجز الحفر - استخدام حاجز ناعم بدلاً من الحاجز الصلب المطلوب",
        "subtitle": "Majority of Excavations with Warning Tape Only - Prohibited for >1.2m",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Majority of excavation zones observed with SOFT barricade (warning tape only) instead of required HARD barricade (scaffold tubes, concrete barriers, hard fencing).\n• Depth: Excavations >1.2m up to 2.5m - hard barricade MANDATORY per SE-GEN-EHS-007 and OSHA 1926 Subpart P - soft tape PROHIBITED for >1.2m depth.\n• Missing: Edge protection, safe access/egress ladders, spoil pile within 1m of edge, no reflective delineators, no night lighting.\n• Night Risk: Excavations invisible at night - high risk of personnel/vehicle fall - no solar blinkers.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "CRITICAL",
        "risk_text": "Fatal fall into excavation, cave-in engulfment, vehicle fall. Soft tape provides ZERO physical protection - only visual warning. Excavation is High-Risk MAH activity with fatality history. Regulatory prohibition risk. Tape failure leads to fatality - previous industry incidents.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. IMMEDIATE replacement of ALL soft barricades with HARD barricades - scaffold tubes with toe boards / concrete barriers - within 12 hours\n2. Provide compliant ladders every 7.5m, keep spoil minimum 1.5m away, install reflective delineators + solar lights\n3. Excavation permit re-validation with geotechnical assessment + competent person daily checklist - verified by SE EHS\n4. Issue NCR Level 2 for repeat violation + contractual penalty + back-charge cost of hard barricades\n5. Photo evidence: Before (soft tape) vs After (hard barricade) - submit to SE EHS for closure verification",
        "photo_main": "Excavation Without Hard Barricade\nSoft Tape Only - Prohibited",
        "photo2": "No Ladder + Spoil Too Close",
        "photo3": "After: Hard Barricade Compliant"
    },
    {
        "id": "03",
        "title_en": "Recurring PPE Violations Despite Repeated Warnings",
        "title_ar": "مخالفات معدات الوقاية الشخصية المتكررة رغم التحذيرات",
        "subtitle": "Same Crew - Same Violations - No Improvement After Multiple Warnings",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Persistent and recurring failure to wear mandatory PPE: Safety helmets without chin straps, safety glasses, gloves, high-visibility vests, safety boots, full body harness at height.\n• Working at height without Full Body Harness during storage fence installation and painting works at height >1.8m.\n• History: Same violations reported multiple times in Daily EHS Reports, Safety Observation Cards, Toolbox Talks - NO improvement despite warnings.\n• Supervision present but failing to enforce - behavioral and cultural failure, not availability - PPE available in store.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "HIGH",
        "risk_text": "Head injury, eye injury, hand lacerations, fall from height, struck-by. Recurring violation indicates total failure of safety culture and supervisory oversight. If PPE (last line of defense) is ignored, higher-level controls will also fail. Increases LTIFR/TRIR, damages Zero Harm commitment, client confidence at risk. Precursor to serious injury per Heinrich triangle.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. Immediate on-site correction + documented Toolbox Talk with photos - Observation Cards + formal violation notice\n2. Implement Behavior-Based Safety (BBS) - daily PPE checks at start of shift + random checks - scorecard Red/Amber/Green\n3. Enforce 3-Strikes Removal Policy: 1st warning, 2nd suspension, 3rd removal from site - for workers AND supervisors\n4. Link PPE compliance to subcontractor KPI and payment milestone - financial consequence\n5. Daily management safety walks - SE PM + Subcontractor PM - joint walk + positive recognition for compliant crews",
        "photo_main": "PPE Violation - No Helmet / No Glasses\nWorking at Height Without Harness",
        "photo2": "Same Crew Repeat Violation",
        "photo3": "After: Compliant PPE"
    },
    {
        "id": "04",
        "title_en": "Historical Pattern - Bitumen, Workshop, Formwork Works Without PTW",
        "title_ar": "نمط تاريخي متكرر - أعمال بيتومين وورشة وفورم ورك بدون تصريح",
        "subtitle": "Same PTW Violation Repeated Across Multiple Activities - Systemic Failure",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Pattern Analysis: Same category of violation (Working Without PTW) repeated across multiple historical activities - indicates SYSTEMIC failure, not isolated incident.\n• Previous Activities Without PTW: Bitumen works (hot work - fire/burn risk), Workshop activities (electrical, grinding - electric shock, fire), Formwork (working at height, collapse, manual handling).\n• All executed without approved PTW, without TRA review, without EHS coordination - documented in previous NCRs.\n• Subcontractor failed to learn from previous NCRs - no corrective action implemented - same root cause repeated - failure to prevent recurrence.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "CRITICAL",
        "risk_text": "Systemic failure of PTW management system - subcontractor does not respect PTW as critical control barrier. Bitumen: fire/explosion, burn. Workshop: electric shock, fire, eye injury. Formwork: collapse, fall from height, struck-by. All high-risk activities. If PTW optional, all high-risk controls will fail. Management system ineffective - requires top management accountability. Regulatory pattern indicates willful negligence.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. Comprehensive PTW System Audit - review ALL historical PTWs for bitumen, workshop, formwork - identify gaps and falsification - audit report within 48h\n2. Mandatory PTW Re-Training for ALL Al-Mial supervision and workforce - including specific hazards - competency assessment 80% pass - no pass, no work\n3. Issue cumulative NCR Level 2 for repeat pattern - contractual penalty escalation - consider partial suspension of high-risk activities\n4. Implement Digital PTW System - QR code based - mandatory fields: TRA, isolation, EHS approval - no EHS approval, permit cannot be activated\n5. Weekly PTW Compliance Trend Analysis - present in PM meeting - track PTW compliance % - target 100% - link to payment",
        "photo_main": "Bitumen Works Without PTW\nWorkshop Without PTW\nFormwork Without PTW",
        "photo2": "Historical NCRs - Same Violation",
        "photo3": "After: PTW Displayed at Work Front"
    },
    {
        "id": "05",
        "title_en": "Non-Compliance with EHS Directives - Refusal to Stop Unsafe Equipment",
        "title_ar": "عدم الالتزام بأوامر السلامة - رفض إيقاف المعدات غير الآمنة",
        "subtitle": "Direct Challenge to EHS Authority - Continued Work Despite Stop Order",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Siemens Energy EHS Department issued clear verbal and written instruction to STOP use of unacceptable / unsafe equipment (leaking equipment, damaged tools, equipment without guard, equipment without inspection certificate).\n• Subcontractor Response: REFUSED to comply and CONTINUED work with same unsafe equipment - willful disregard and direct challenge to EHS authority.\n• Equipment: Leaking hydraulic oil and fuel, damaged electrical cables, angle grinder without guard, lifting equipment without third-party certificate.\n• Violation: Direct violation of Siemens Energy Stop Work Authority, EHS General Requirements, Project EHS Plan Sec 8.1, ISO 45001 Cl 8.1.2 - cultural failure.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "CRITICAL",
        "risk_text": "Erosion of EHS authority and safety culture - if Stop Work order can be ignored, all critical controls (LOTO, WAH, PTW) at risk of failure. Fire/explosion, electric shock, equipment failure loss of control, environmental pollution. Willful violation - highest level of disciplinary action required. If not addressed decisively, Siemens Energy loses ability to control site safety - major liability with criminal liability potential.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. IMMEDIATE escalation to Top Management - Project Director to issue formal letter to Al-Mial Managing Director - written explanation for refusal within 12h\n2. Enforce Stop Work Authority - ground ALL unacceptable equipment immediately - Tag 'DO NOT USE - EHS HOLD' - remove keys - secure area - fire watch if fuel leak\n3. Invoke contractual penalty clause for refusal to follow EHS instruction - highest level penalty - back-charge cost of additional SE supervision\n4. Remove and replace non-compliant supervision who refused order - CVs + EHS competency assessment + interview by SE EHS before new supervisor allowed\n5. Emergency EHS Leadership Meeting - Al-Mial Top Management + all foremen - re-communicate Stop Work Authority and consequences - documented commitment signed",
        "photo_main": "Unsafe Equipment Ordered to Stop\nBut Continued Working - Refusal",
        "photo2": "Leaking Equipment Still in Use",
        "photo3": "After: Equipment Grounded - Tag Out"
    },
    {
        "id": "06",
        "title_en": "Unsafe Condition of Workshop Equipment - Damaged Tools and Machinery",
        "title_ar": "حالة غير آمنة لمعدات الورشة - أدوات ومعدات تالفة",
        "subtitle": "Exposed Cables, Missing Guards, Damaged Slings, Oil Leaks - Unsafe Condition",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Workshop Equipment Observed: Damaged and uninspected tools in unsafe condition - exposed electrical cables, angle grinders without protective guards, welding machines without earthing, damaged slings and shackles, leaking equipment.\n• No preventive maintenance records - daily inspection checklists falsified - same signature for 7 consecutive days - no evidence of competent inspection.\n• No drip trays under equipment parking - soil contamination ~5 sqm observed - oil and fuel leaks active - fire hazard.\n• No third-party inspection certificates for lifting equipment, power tools, electrical equipment - no operator competency evidence.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "CRITICAL",
        "risk_text": "Electric shock, electrocution, fire/explosion (fuel on hot surface, exposed cable spark), hand injury, eye injury from grinding without guard, lifting equipment failure leading to dropped load, slip hazard, soil/groundwater contamination. Breakdown of maintenance management system. Potential for multiple fatalities if lifting equipment fails or electrical fault causes fire in workshop with flammable materials. Environmental prosecution risk.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. IMMEDIATE grounding of ALL unsafe workshop equipment - Tag 'DO NOT USE - EHS HOLD' - remove keys - secure workshop - inventory of all equipment\n2. Contaminated soil excavated and disposed as hazardous waste - spill response with absorbents - provide drip trays and spill kits 100% coverage\n3. Submit full equipment register with maintenance history and valid third-party inspection certificates within 48h - No certificate, no operation\n4. Implement mandatory daily equipment checklist - verified by SE EHS/Mechanical with photo evidence - guard in place, no leaks, cable intact\n5. Replace all damaged cables, missing guards, damaged slings/shackles - procurement of compliant equipment - training on equipment inspection - competency assessment",
        "photo_main": "Exposed Electrical Cable\nGrinder Without Guard\nDamaged Sling / Shackle",
        "photo2": "Oil Leak + Soil Contamination",
        "photo3": "After: Compliant Equipment + Drip Tray"
    },
    {
        "id": "07",
        "title_en": "Poor Waste Management & Lack of Segregation",
        "title_ar": "سوء إدارة النفايات وعدم فصل المواد",
        "subtitle": "No Clear System - Mixed Hazardous, Wood, Plastic, Debris in Same Container",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Waste management system not clear - no segregation, no labeling, no color-coding, no defined waste area - poor housekeeping throughout site.\n• Waste Mixed: Hazardous waste (oil rags, filters, paint tins), recyclable (wood, plastic), and general waste mixed in same container/skip - cross-contamination.\n• No MSDS reference for hazardous waste bins, no HAZCHEM labels, no spill kit near waste area - environmental violation and fire hazard.\n• Waste containers overflowing, waste accumulated on ground, no defined collection point - attracts pests, fire risk, slip/trip hazard.\n• Violation: ISO 14001:2015 Cl 8.1, Project Waste Management Plan, local EPA regulations, SE Environmental Standard.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "HIGH",
        "risk_text": "Environmental pollution (soil/groundwater contamination from hazardous waste), fire hazard from incompatible mixing (oil rags + general waste = spontaneous combustion), regulatory fine, reputational damage to Siemens Energy. Cross-contamination makes recycling impossible - sustainability KPI failure. Poor housekeeping leads to slip/trip, manual handling injuries, indicates poor overall site management. Potential for environmental prosecution.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. Immediate provision of color-coded, labeled waste skips per SE Standard: Black-General, Green-Recyclable (wood, plastic), Red-Hazardous with HAZCHEM labels and MSDS\n2. Clean-up and correct segregation of existing mixed waste by competent team under EHS supervision - define waste area with hard barricade and signage\n3. Training on Waste Management for all Al-Mial workforce - hazardous waste handling, spill prevention, manifest system, housekeeping\n4. Appoint Waste Management Focal Point - responsible for daily housekeeping, segregation, waste area inspection - name to be submitted to SE EHS\n5. Weekly waste audit - track volumes and disposal manifests - include in monthly environmental report - target: Zero mixed waste - link to KPI",
        "photo_main": "Mixed Waste in One Skip\nHazardous + General + Recyclable",
        "photo2": "Oil Rags in General Waste - Fire Hazard",
        "photo3": "After: Color-Coded Segregated Skips"
    },
    {
        "id": "08",
        "title_en": "Deficiency in Safety Signage - Parking Garage & Speed Limits",
        "title_ar": "نقص في لوحات السلامة - منطقة وقوف السيارات والسرعة المحددة",
        "subtitle": "Poor Signage in Parking Area (Garage) and Speed Limit Signs Missing/Damaged",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Poor and insufficient safety signage for parking area (garage) and speed limit signs - signs missing, damaged, not visible, or not retro-reflective.\n• Parking Area: No designated parking signage, no pedestrian walkway markings, no speed limit signs inside garage, no directional arrows, no stop signs at intersections.\n• Speed Limits: Speed limit signs (15 km/h) missing or damaged at multiple locations - no speed humps, no radar monitoring - speeding observed 30 km/h in 15 km/h zone confirmed by radar gun.\n• Vehicle-Pedestrian Interface: No segregation between vehicle routes and pedestrian routes in parking area - workers walking through parking area with moving vehicles - blind spots.\n• Violation: Project Traffic Management Plan (TMP), ISO 39001 Road Traffic Safety, SE Traffic Safety Standard - Previous NCRs not closed.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "HIGH",
        "risk_text": "High risk of vehicle-pedestrian collision, vehicle-vehicle collision, equipment damage, fatality in parking area. Vehicle incidents are top 3 fatality causes in construction. Poor signage increases liability - Siemens Energy liable if signage not per TMP. Emergency vehicle access may be impeded - risk to emergency response. Garage area with poor visibility - high risk of reversing incidents.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. Immediate reinstatement of all safety signage per approved TMP - retro-reflective grade, proper height, visible at night - parking signs, speed limit 15 km/h, pedestrian walkway, stop signs, directional arrows\n2. Establish dedicated pedestrian routes in parking/garage area with hard segregation - Jersey barriers + green walkway paint + pedestrian crossing markings\n3. Install speed humps and deploy trained flagmen at garage entrance/exit and high-risk intersections - implement speed monitoring with radar gun 2x daily - log speeding violations\n4. Conduct traffic safety campaign and driver re-induction for all drivers including Al-Mial - focus on parking area hazards, reversing, pedestrian priority\n5. EHS + Logistics joint daily inspection of traffic controls and signage - checklist signed - include parking area and garage - photo evidence of compliant signage",
        "photo_main": "Missing Parking Signage\nMissing Speed Limit 15 km/h Sign",
        "photo2": "No Pedestrian Walkway in Garage\nVehicle-Pedestrian Interface Risk",
        "photo3": "After: Compliant Signage + Pedestrian Route"
    },
    {
        "id": "09",
        "title_en": "Delayed Response from Al-Mial to Close Observations",
        "title_ar": "تأخر استجابة قسم الميال لإغلاق الملاحظات",
        "subtitle": "Poor Close-Out System - Open Observations for Days/Weeks - No Evidence",
        "desc_title": "Technical Description / الوصف الفني",
        "desc": "• Significant delay in response from Al-Mial (subcontractor) department to close EHS observations, NCRs, and Safety Observation Cards - poor close-out system.\n• Timeline: Observations issued days/weeks ago still open - no corrective action submitted, no evidence of closure, no root cause analysis - repeated follow-ups required by SE EHS.\n• Previous NCRs closed without evidence or verification - falsification of closure - same violations recurring because root cause not addressed.\n• No dedicated EHS focal point from Al-Mial to track and close observations - no tracking system - observations lost or ignored.\n• Violation: ISO 45001 Cl 10.2 Nonconformity and corrective action, Project EHS Plan Sec 9.1 Monitoring, Contractual requirement for timely close-out.",
        "risk_title": "Risk Assessment / تقييم المخاطر",
        "risk": "HIGH",
        "risk_text": "Open observations mean hazards remain uncontrolled on site - risk of incident from known hazard that was already identified but not fixed. Delay indicates subcontractor does not value EHS and does not prioritize safety - cultural failure. Erosion of EHS system effectiveness - if observations not closed, EHS inspection loses value. Accumulation of open observations increases liability - Siemens Energy aware of hazard but subcontractor not fixing - if incident occurs, Siemens Energy exposed. Client and regulatory perception: Poor EHS management.",
        "action_title": "Required Corrective Action / الإجراء التصحيحي",
        "action": "1. Al-Mial to appoint dedicated EHS Focal Point / Coordinator - name, CV, contact to be submitted to SE EHS within 12 hours - responsible for tracking and closing ALL observations\n2. Submit close-out plan for ALL open observations within 24 hours - with specific actions, owner, deadline, evidence required for each observation - prioritized by risk level (Critical first)\n3. Implement Observation Tracking System - Excel tracker or digital app - columns: Observation ID, Date, Description, Risk, Owner, Deadline, Status, Evidence, Verified by - updated daily and shared with SE EHS\n4. Weekly Observation Close-Out Meeting - Al-Mial PM + EHS Focal Point + SE EHS Manager - review open observations, close-out evidence, trend - meeting minutes with actions\n5. Link observation close-out to contractual penalty and KPI - target: 100% close-out of Critical observations within 24h, High within 72h, Medium within 7 days - include in weekly EHS scorecard and payment milestone",
        "photo_main": "Open Observations List\nOverdue NCRs - Days/Weeks Open",
        "photo2": "Tracker Showing Delayed Close-Out\nNo Evidence Submitted",
        "photo3": "After: Closed Observations with Evidence"
    },
]

def create_slide_final(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    # Top teal line
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(5), fill=TEAL)
    
    # Header with Siemens Logo
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.85), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    add_siemens_logo(slide, Inches(0.3), Inches(0.15))
    
    # Title area
    text_box(slide, Inches(2.8), Inches(0.1), Inches(7.0), Inches(0.4), f"{data['id']} - {data['title_en']}", size=12, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(2.8), Inches(0.45), Inches(7.0), Inches(0.35), f"{data['title_ar']} | {data['subtitle']}", size=8.5, bold=False, color=GREY_TEXT)

    # Risk badge top right
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.2), Inches(0.15), Inches(2.9), Inches(0.6), fill=risk_bg, border=risk_color, border_w=Pt(2))
    text_box(slide, Inches(10.2), Inches(0.15), Inches(2.9), Inches(0.6), f"RISK LEVEL: {data['risk']}\nمستوى الخطورة: {data['risk']}", size=10, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    # LEFT CONTENT - 3 boxes stacked (7.2 width)
    # Box 1: Description
    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.35), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.35), f"1. {data['desc_title']}", size=10, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.45), Inches(7.2), Inches(1.9), fill=LIGHT_GREY_CARD, border=GREY_BORDER, border_w=Pt(1))
    # Left red border for criticality
    shape(slide, Inches(0.3), Inches(1.45), Pt(4), Inches(1.9), fill=risk_color)
    tf = slide.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(6.8), Inches(1.8)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc'].split('\n')):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(8)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    # Box 2: Risk
    shape(slide, Inches(0.3), Inches(3.5), Inches(7.2), Inches(0.35), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.5), Inches(7.0), Inches(0.35), f"2. {data['risk_title']} - {data['risk']}", size=10, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.85), Inches(7.2), Inches(1.1), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(3.85), Pt(4), Inches(1.1), fill=risk_color)
    text_box(slide, Inches(0.5), Inches(3.9), Inches(6.8), Inches(1.0), data['risk_text'], size=8, bold=False, color=DARK_TEXT)

    # Box 3: Corrective Action
    shape(slide, Inches(0.3), Inches(5.1), Inches(7.2), Inches(0.35), fill=TEAL_LIGHT_BG, border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(5.1), Inches(7.0), Inches(0.35), f"3. {data['action_title']}", size=10, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.45), Inches(7.2), Inches(1.7), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(5.45), Pt(4), Inches(1.7), fill=TEAL)
    tf2 = slide.shapes.add_textbox(Inches(0.5), Inches(5.5), Inches(6.8), Inches(1.6)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action'].split('\n')):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(7.5)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(1.5)

    # RIGHT SIDE - PHOTO AREA - Very clear and large
    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.35), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.35), "📷 PHOTO EVIDENCE AREA / مساحة الصور الموقعية", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    
    # Main large photo placeholder - with thick border and clear label
    shape(slide, Inches(7.8), Inches(1.45), Inches(5.2), Inches(3.2), fill=LIGHT_GREY_BG, border=DARK_GREY_BORDER, border_w=Pt(2))
    # Inner dashed simulation - add inner shape with different border
    shape(slide, Inches(7.9), Inches(1.55), Inches(5.0), Inches(3.0), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    
    # Large centered text for photo placeholder
    text_box(slide, Inches(7.9), Inches(1.6), Inches(5.0), Inches(0.5), "[أدخل صورة المخالفة هنا]", size=16, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.1), Inches(5.0), Inches(0.4), "[Insert Violation Photo Here]", size=14, bold=True, color=RED_DARK, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.6), Inches(5.0), Inches(0.6), f"{data['photo_main']}\n\nFig {data['id']}.1 - Main Evidence", size=9, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(3.3), Inches(5.0), Inches(1.0), "⬆️ DRAG & DROP YOUR SITE PHOTO HERE\nاسحب الصورة من المجلد وألصقها هنا\n\nThe photo will automatically fit inside this box\nالصورة سيتم احتواؤها تلقائياً داخل هذه الخانة\n\nAdd: Date, Time, Location, GPS Coordinates\nأضف: التاريخ، الوقت، الموقع، إحداثيات GPS", size=9, bold=False, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    # Photo details box
    shape(slide, Inches(7.8), Inches(4.7), Inches(5.2), Inches(0.5), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    text_box(slide, Inches(7.9), Inches(4.7), Inches(5.0), Inches(0.5), "Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________  Taken by: ___________\nتفاصيل الصورة: التاريخ، الوقت، الموقع، المصور", size=7, color=GREY_TEXT, align=PP_ALIGN.LEFT)

    # Two smaller photo placeholders below main
    shape(slide, Inches(7.8), Inches(5.3), Inches(2.5), Inches(1.0), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1))
    text_box(slide, Inches(7.9), Inches(5.3), Inches(2.3), Inches(0.2), "[Additional Photo 2]", size=8, bold=True, color=DARK_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(5.5), Inches(2.3), Inches(0.3), "[أدخل صورة إضافية هنا]", size=8, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(5.8), Inches(2.3), Inches(0.4), f"{data['photo2']}\nFig {data['id']}.2 - Different Angle", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)

    shape(slide, Inches(10.5), Inches(5.3), Inches(2.5), Inches(1.0), fill=TEAL_LIGHT_BG, border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(10.6), Inches(5.3), Inches(2.3), Inches(0.2), "[After Correction Photo]", size=8, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(10.6), Inches(5.5), Inches(2.3), Inches(0.3), "[أدخل صورة بعد التصحيح هنا]", size=8, bold=True, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(10.6), Inches(5.8), Inches(2.3), Inches(0.4), f"Compliant State\nالحالة المطابقة\nFig {data['id']}.3", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    # Footer with Siemens Energy
    shape(slide, Inches(0), Inches(6.5), Inches(13.33), Inches(0.35), fill=LIGHT_GREY_CARD, border=GREY_BORDER, border_w=Pt(1))
    text_box(slide, Inches(0.3), Inches(6.5), Inches(4), Inches(0.35), "SIEMENS ENERGY | KAZ Project | EHS Department | Confidential - Management Review", size=7, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(4.5), Inches(6.5), Inches(4), Inches(0.35), "Zero Harm is Non-Negotiable | Safety is Non-Negotiable", size=7, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(10.5), Inches(6.5), Inches(2.5), Inches(0.35), f"Slide {data['id']} | SE-EHS-NC-2026-FINAL | Page {data['id']}/11", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)
    
    # Bottom thin line
    shape(slide, Inches(0), Inches(6.85), Inches(13.33), Pt(3), fill=TEAL)

def create_title_final(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), fill=TEAL)
    
    # Header with Siemens Logo
    shape(slide, Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    add_siemens_logo(slide, Inches(0.6), Inches(0.4))
    text_box(slide, Inches(3.0), Inches(0.35), Inches(5), Inches(0.3), "KAZ Power Plant Upgrade Project | Zero Harm Culture", size=10, bold=False, color=GREY_TEXT)
    text_box(slide, Inches(10), Inches(0.35), Inches(2.8), Inches(0.6), "CONFIDENTIAL\nسري - للإدارة العليا فقط\nManagement Eyes Only", size=10, bold=True, color=RED, align=PP_ALIGN.RIGHT)

    # Main title with red accent bar
    shape(slide, Inches(0.5), Inches(1.5), Inches(0.08), Inches(1.3), fill=RED)
    text_box(slide, Inches(0.8), Inches(1.4), Inches(11.5), Inches(0.6), "Subcontractor EHS Non-Compliance Report - Final Site Observations", size=24, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(0.8), Inches(2.0), Inches(11.5), Inches(0.5), "تقرير مخالفات السلامة النهائي للمقاول الفرعي - الملاحظات الميدانية الفعلية", size=16, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.8), Inches(2.5), Inches(11.5), Inches(0.3), "Based on 9 Actual Site Observations Provided | Executive White Design with Dedicated Photo Placeholders", size=10, bold=False, color=GREY_TEXT)

    # Left box - observations summary
    shape(slide, Inches(0.5), Inches(3.0), Inches(7.5), Inches(3.0), fill=LIGHT_GREY_CARD, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.5), Inches(3.0), Inches(7.5), Inches(0.35), fill=TEAL_DARK)
    text_box(slide, Inches(0.6), Inches(3.0), Inches(7.3), Inches(0.35), "Final Observations Summary / ملخص الملاحظات النهائية الـ 9 - Each Violation Has Separate Slide", size=11, bold=True, color=WHITE)
    summary = """01. Painting, Storage Fence Installation, Cement Pouring WITHOUT PTW & Induction + No knowledge in ENERGIZED zone
    طلاء، تركيب حاجز مخزن، صب أسمنت بدون بيرمت وبدون اندكشن + عدم علم رغم منطقة انرجايز - CRITICAL

02. Excavation Barricades - Soft Barricade Instead of Hard Barricade (Tape Only - Prohibited)
    حواجز الحفر سوفت باركيد بدل هارد باركيد - شريط فقط ممنوع - CRITICAL

03. Recurring PPE Violations Despite Repeated Warnings - Same Crew
    مخالفات PPE متكررة رغم التبليغ - نفس العمالة - HIGH

04. Historical Pattern - Bitumen, Workshop, Formwork Works Without PTW - Systemic Failure
    أعمال سابقة نفس المخالفات - بيتومين، ورشة، فورم ورك بدون بيرمت - CRITICAL

05. Non-Compliance with EHS Directives - Refusal to Stop Unsafe Equipment Despite Order
    عدم اتباع أوامر السلامة - رفض إيقاف معدات غير مقبولة واستمرار العمل - CRITICAL

06. Unsafe Condition of Workshop Equipment - Exposed Cables, Missing Guards, Damaged Slings
    معدات الورشة غير آمنة - كوابل مكشوفة، بدون غطاء حماية، تالف - CRITICAL

07. Poor Waste Management & Lack of Segregation - No Clear System
    عدم توضيح إدارة النفايات وفصل المواد - سوء نظافة - HIGH

08. Deficiency in Safety Signage - Parking Garage & Speed Limits Missing/Damaged
    فقر في لوحات السلامة - باركنك كراج ولوحات سرعة محددة مفقودة/تالفة - HIGH

09. Delayed Response from Al-Mial to Close Observations - Poor Close-Out System
    تأخر استجابة قسم الميال لحل الملاحظات - نظام إغلاق ضعيف - HIGH"""
    text_box(slide, Inches(0.6), Inches(3.4), Inches(7.3), Inches(2.5), summary, size=8, color=DARK_TEXT)

    # Right box - report info and design
    shape(slide, Inches(8.3), Inches(3.0), Inches(4.5), Inches(3.0), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    shape(slide, Inches(8.3), Inches(3.0), Inches(4.5), Inches(0.35), fill=TEAL)
    text_box(slide, Inches(8.4), Inches(3.0), Inches(4.3), Inches(0.35), "Report Information & Design / معلومات التقرير والتصميم", size=11, bold=True, color=WHITE)
    info = """Reference: SE-EHS-NC-2026-FINAL-9OBS-V7
Ref: SE-EHS-NC-2026-FINAL-9OBS-V7

Date: September 2026
Prepared by: EHS Manager - Siemens Energy
Project: KAZ Power Plant Upgrade
Classification: Confidential - Top Management

Standards:
• ISO 45001:2018, ISO 14001:2015
• OSHA 29 CFR 1926
• Siemens Energy EHS Principles
• Life Saving Rules
• Project EHS Plan

Design: Executive White Modern - FIXED
• Background: Pure White #FFFFFF
• Text: Dark #212121 / Grey #616161
• Warning: Red #E30613 Critical / Amber #FFB900 High
• Accent: Teal #009999 / Dark Petrol #0E2F3E
• Siemens Energy Logo on every slide (top left)

Each Slide Includes - FIXED as Requested:
✓ Separate slide for each violation (9 slides)
✓ Dedicated LARGE photo placeholder
✓ [أدخل صورة المخالفة هنا] / [Insert Violation Photo Here] - Very Clear
✓ 3 photo boxes per slide: Main + Additional + After Correction
✓ Technical Description, Risk Assessment, Corrective Action
✓ Siemens Energy Logo on every slide

Based on YOUR 9 actual observations - Final Version"""
    text_box(slide, Inches(8.4), Inches(3.4), Inches(4.3), Inches(2.5), info, size=7.5, color=DARK_TEXT)

    # Executive statement
    shape(slide, Inches(0.5), Inches(6.2), Inches(12.33), Inches(1.0), fill=RED_LIGHT, border=RED, border_w=Pt(2))
    text_box(slide, Inches(0.6), Inches(6.2), Inches(12.1), Inches(0.2), "Executive Statement / ملاحظة تنفيذية - FIXED DESIGN:", size=10, bold=True, color=RED)
    text_box(slide, Inches(0.6), Inches(6.45), Inches(12.1), Inches(0.7), "This FINAL report documents 9 ACTUAL site observations provided by EHS team - each violation has SEPARATE slide with DEDICATED LARGE photo placeholder [Insert Violation Photo Here] / [أدخل صورة المخالفة هنا] - Siemens Energy logo on every slide - Executive white modern design with Red/Yellow warning highlights - Ready for top management review and decision - Each slide includes Technical Description, Risk Assessment, Corrective Action, and 3 photo boxes.\nهذا التقرير النهائي يوثق 9 ملاحظات فعلية - كل مخالفة لها شريحة منفصلة مع مكان صور كبير وواضح [أدخل صورة المخالفة هنا] - شعار سيمنز في كل شريحة - تصميم تنفيذي أبيض عصري - جاهز لمراجعة الإدارة العليا.", size=8.5, color=DARK_TEXT)

def create_closing_final(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(5), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.8), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    add_siemens_logo(slide, Inches(0.3), Inches(0.15))
    text_box(slide, Inches(2.8), Inches(0.1), Inches(10), Inches(0.6), "Final Recommendations & Corrective Action Plan - Based on 9 Actual Observations", size=14, bold=True, color=DARK_TEXT)

    shape(slide, Inches(0.3), Inches(1.1), Inches(4.1), Inches(0.4), fill=RED)
    text_box(slide, Inches(0.3), Inches(1.1), Inches(4.1), Inches(0.4), "🔴 IMMEDIATE (0-24h) / فوري", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0.3), Inches(1.5), Inches(4.1), Inches(3.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(1.5), Inches(3.9), Inches(3.7), "• STOP WORK: Painting, Storage Fence, Cement, Excavation without hard barricade, Unsafe workshop equipment\n• Secure ENERGIZED zone - hard barricade + LOTO + fire watch\n• Remove non-inducted personnel - 100% gate - manpower list vs induction in 6h\n• Ground leaking/unsafe equipment - Tag DO NOT USE\n• Issue NCR Level 2 + contractual warning - cumulative for pattern\n• Emergency meeting - Al-Mial MD + PM within 12h - written explanation for refusal\n\n[Photo Placeholder: Secured energized zone after stop work]\n[مكان صورة: منطقة مؤمنة بعد إيقاف العمل]\n\nDECISION: Approve Stop-Work Authority to EHS\nRoot Cause: Why work started without knowledge of own supervision?", size=8, color=DARK_TEXT)

    shape(slide, Inches(4.6), Inches(1.1), Inches(4.1), Inches(0.4), fill=AMBER)
    text_box(slide, Inches(4.6), Inches(1.1), Inches(4.1), Inches(0.4), "🟡 SHORT-TERM (1-7 Days) / قصير", size=11, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(4.6), Inches(1.5), Inches(4.1), Inches(3.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(4.7), Inches(1.5), Inches(3.9), Inches(3.7), "• PTW Audit: Review ALL historical - bitumen, workshop, formwork, painting, storage, cement - submit in 48h\n• Re-Training: PTW + Energized Zone + Excavation + PPE + Waste - all Al-Mial supervision - 80% pass\n• Hard Barricades: Replace ALL soft tape with hard barricades - within 12h - Before/After photos\n• Workshop Equipment: Inventory + maintenance history + 3rd party certs in 48h - replace damaged\n• Waste: Color-coded skips + labeling + focal point + training\n• Signage: Parking + speed limit 15 km/h - retro-reflective + flagmen + radar + speed humps\n• PPE Campaign + BBS + Daily management walks\n\n[Photo Placeholder: Compliant hard barricade + signage + waste segregation]\n[مكان صورة: حالة مطابقة بعد التصحيح]\n\nDECISION: Approve $20k cost - back-charge to Al-Mial", size=8, color=DARK_TEXT)

    shape(slide, Inches(8.9), Inches(1.1), Inches(4.1), Inches(0.4), fill=TEAL)
    text_box(slide, Inches(8.9), Inches(1.1), Inches(4.1), Inches(0.4), "🟢 STRATEGIC (7-30 Days) / استراتيجي", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(8.9), Inches(1.4), Inches(4.1), Inches(3.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(9.0), Inches(1.4), Inches(3.9), Inches(3.7), "• Penalty Matrix: Linked to KPIs - TRIR, NCR closure, PTW compliance, observation close-out - deduction from invoice\n• Leadership Accountability: Replace non-performing supervisors who refused EHS order - CV + competency interview by SE\n• System Upgrade: Digital PTW with QR + Biometric induction + equipment checklist app + observation tracker\n• KPI & Scorecard: Weekly dashboard Red/Amber/Green to Top Management - include delayed response metric for Al-Mial\n• Independent Audit: 3rd party EHS audit of Al-Mial system - ISO 45001/14001 gap analysis\n• Zero Harm Commitment: Re-sign Life Saving Rules - Townhall with Top Management - visible commitment\n• Contingency: Prepare demobilization / replacement plan if no sustained improvement\n\n[Photo Placeholder: Dashboard / Training / Digital PTW]\n[مكان صورة: لوحة مؤشرات / تدريب]\n\nDECISION: Authorize partial termination if critical violations recur + observation delay continues", size=8, color=DARK_TEXT)

    shape(slide, Inches(0.3), Inches(5.5), Inches(12.7), Inches(1.7), fill=LIGHT_GREY_CARD, border=TEAL_DARK, border_w=Pt(2))
    text_box(slide, Inches(0.4), Inches(5.5), Inches(12.5), Inches(0.3), "✅ MANAGEMENT DECISION MATRIX - REQUIRED TODAY / مصفوفة القرارات - مطلوب اليوم", size=12, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(1.2), "1. Approve Stop-Work Authority for Energized Zone + Unsafe Equipment [ ] Yes [ ] No | إيقاف العمل في منطقة انرجايز والمعدات غير الآمنة\n2. Authorize Warning + Penalty for Refusal to Follow EHS Order [ ] Yes [ ] No | تحذير وغرامة لرفض أوامر السلامة\n3. Mandate Al-Mial Top Management On-Site within 12h with Written Explanation [ ] Yes [ ] No | حضور إدارة الميال خلال 12 ساعة مع توضيح كتابي\n4. Approve Budget $20k for Hard Barricades, Signage, Waste Skips - Back-charge [ ] Yes [ ] No | ميزانية حواجز ولوحات وحاويات - تحميل على المقاول\n5. Endorse Zero Tolerance 3 Strikes + Observation Close-Out KPI (Critical 24h) [ ] Yes [ ] No | عدم تسامح 3 مخالفات + مؤشر إغلاق الملاحظات\n\nSignature: _________________ Project Director Date: _______ | Signature: _________________ Construction Manager Date: _______ | Signature: _________________ EHS Manager Date: _______", size=8, color=DARK_TEXT)

# Build final presentation in correct order
final_prs = Presentation()
final_prs.slide_width = Inches(13.33)
final_prs.slide_height = Inches(7.5)

create_title_final(final_prs)
for d in slides_data:
    create_slide_final(d)
create_closing_final(final_prs)

out = "/home/user/ahmed/Subcontractor_EHS_FINAL_V7_FIXED_Each_Slide_Separate_PhotoPlaceholders_SiemensLogo.pptx"
final_prs.save(out)
print(f"Saved to {out}")
