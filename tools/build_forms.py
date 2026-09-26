#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Companion controlled forms for night concrete pouring - same house style and
same Siemens Energy logo / header / footer as the Method Statement:

  KAZ-EHS-FRM-2026-006  Shift Handover Log (Day <-> Night)
  KAZ-EHS-FRM-2026-007  Night Concrete Pouring - EHS Checklist Pack

Usage:  python3 tools/build_forms.py
"""
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from build_ms_night_concrete_pour import (   # noqa: E402  (shared house style)
    REPO, TMP, USABLE_P, WHITE, BODY, TEAL_DARK, TEAL_MID, BAND, LBL_FILL,
    make_shell, apply_header_footer, prune_orphans, banner, heading, subheading,
    body, bullets, grid, label_table, spacer, note_box, setup_section,
)
from docx import Document   # noqa: E402

BLANK = ""

CHK_COLS = [400, 5400, 520, 520, 2200, 824]
CHK_AL = ["c", None, "c", "c", None, "c"]


def checklist_table(doc, title, items, extra_cols=None, widths=None, aligns=None,
                    size=9, header=None):
    if title:
        subheading(doc, title, size=10.5, space_before=150, fill=BAND)
    hdr = header or ["#", "Verification Item", "Yes", "N/A",
                     "Verification Method", "Initials"]
    rows = [hdr]
    for i, it in enumerate(items, start=1):
        rows.append([str(i), it] + [BLANK] * (len(hdr) - 2))
    grid(doc, widths or CHK_COLS, rows, size=size, header_size=9,
         aligns=aligns or CHK_AL, total=USABLE_P)


def signature_block(doc, title, roles):
    subheading(doc, title, space_before=160)
    rows = [["Role", "Name", "Signature", "Date / Time"]]
    for r in roles:
        rows.append([r, BLANK, BLANK, BLANK])
    grid(doc, [3600, 2400, 2400, 1464], rows, size=9.5, header_size=9.5,
         bolds=[0], total=USABLE_P)


# ============================================================================
# FORM 006 - SHIFT HANDOVER LOG
# ============================================================================
def build_handover():
    tmp = TMP.replace("_ms_tmp", "_frm006_tmp")
    make_shell(tmp, "Shift Handover Log - Day / Night (Night Concrete Pouring)",
               "EHS; Shift Handover; Night Work; Concrete Pouring; KAZ-EHS-FRM-2026-006")
    doc = Document(tmp)
    setup_section(doc.sections[0], landscape=False)

    banner(doc, "CONTROLLED FORM", "SHIFT HANDOVER LOG",
           "DAY \u2194 NIGHT SHIFT HANDOVER \u2014 NIGHT CONCRETE POURING OPERATIONS",
           "KAZ-EHS-FRM-2026-006", "0")

    heading(doc, "SECTION A: FORM IDENTIFICATION", space_before=120)
    label_table(doc, [
        ("Project / Location", "KAZ Power Plant Upgrade Project \u2014 Khur Al-Zubair "
                               "Power Station, Basra, Iraq"),
        ("Companion documents", "KAZ-EHS-MS-2026-004 (Method Statement, Annex A) \u2022 "
                                "KAZ-EHS-SOP-2026-003 (PTW System, \u00a75A)"),
        ("Pour location / area", BLANK),
        ("Date of handover", BLANK),
        ("Handover type", "\u2610 Day \u2192 Night      \u2610 Night \u2192 Day"),
        ("Shift window", "From ______ : ______   to   ______ : ______"),
        ("PTW number(s) & validity", BLANK),
        ("Planned / placed concrete (m\u00b3)", BLANK),
        ("Weather at handover", "Wind ______ km/h   \u2610 Clear  \u2610 Rain  "
                                "\u2610 Dust  \u2610 Lightning"),
        ("Outgoing shift in-charge", BLANK),
        ("Incoming shift in-charge", BLANK),
    ])

    heading(doc, "SECTION B: HANDOVER CHECKLIST")
    body(doc, "Both supervisors walk the area together before completing this form. "
              "Every line requires an outgoing entry and an incoming verification "
              "initial. Any unresolved line is carried to Section D as an open item.",
         size=10)
    items = [
        "Works completed during the shift (locations, volumes, joints formed)",
        "Works outstanding / to be continued by the incoming shift",
        "Active live zones and energized equipment; restricted red zones",
        "Open isolations / LOTO applied and remaining under Isolation Authority control",
        "PTW status, endorsements, restrictions and expiry time",
        "Barricading condition; any temporary removals and their reinstatement plan",
        "Lighting condition; lux values recorded; light-tower positions; backup floodlights",
        "Earthing / grounding connections and 30 mA RCD test status",
        "Open excavations, pits, open edges and rebar-cap condition",
        "Plant left on site, keys, defects and out-of-service equipment",
        "Emergency egress route condition; standby ambulance and medic status",
        "Incidents, near misses, unsafe acts/conditions and corrective actions raised",
        "Environmental issues (washout, spills, waste) and MOTAT records",
        "Instructions, restrictions or changes received from the Area Authority / client",
        "Crew fatigue status: hours worked, breaks taken, anyone stood down",
        "Any other hazard or information the incoming shift must know",
    ]
    rows = [["#", "Handover Item", "Outgoing Shift Entry", "Incoming Verification (initials)"]]
    for i, it in enumerate(items, start=1):
        rows.append([str(i), it, BLANK, BLANK])
    grid(doc, [400, 4300, 3300, 1864], rows, size=9, header_size=9,
         aligns=["c", None, None, "c"], total=USABLE_P)

    heading(doc, "SECTION C: LIVE ZONES, ENERGIZED EQUIPMENT & ISOLATIONS REGISTER")
    rows = [["#", "Zone / Equipment", "Voltage / Status", "Isolation / LOTO Ref",
             "Controlled By", "Remarks"]]
    for _ in range(6):
        rows.append([BLANK] * 6)
    grid(doc, [380, 2400, 1500, 1800, 1800, 1984], rows, size=9, header_size=8.5,
         aligns=["c", None, "c", None, None, None], total=USABLE_P)

    heading(doc, "SECTION D: OPEN ITEMS & CORRECTIVE ACTIONS CARRIED OVER")
    rows = [["#", "Open Item", "Raised By", "Priority", "Owner", "Due", "Status"]]
    for _ in range(6):
        rows.append([BLANK] * 7)
    grid(doc, [380, 3200, 1400, 900, 1400, 1200, 1384], rows, size=9, header_size=8.5,
         aligns=["c", None, None, "c", None, "c", "c"], total=USABLE_P)

    heading(doc, "SECTION E: INCIDENTS / NEAR MISSES THIS SHIFT")
    rows = [["#", "Description", "Report Ref (000480 / incident)", "Immediate Action Taken"]]
    for _ in range(4):
        rows.append([BLANK] * 4)
    grid(doc, [380, 4200, 2400, 2884], rows, size=9, header_size=9,
         aligns=["c", None, None, None], total=USABLE_P)

    signature_block(doc, "SECTION F: HANDOVER SIGNATURES", [
        "Outgoing EHS Supervisor / Officer",
        "Incoming Night EHS Officer / Day EHS Officer",
        "Outgoing Engineering / Civil Supervisor",
        "Incoming Night Shift In-Charge / Day Supervisor",
        "Permit Issuer (noted)",
        "Site Manager (where a Stop-Work or abort occurred)",
        "Concrete Pouring Subcontractor Representative",
    ])

    heading(doc, "SECTION G: INSTRUCTIONS FOR COMPLETION")
    bullets(doc, [
        "One form per shift change. The form travels with the Permit to Work.",
        "Complete physically on site, at the work face, with lighting on for a night "
        "handover \u2014 never from the site office.",
        "The incoming supervisor may refuse to accept the shift and escalate to the "
        "Site Manager if any Section B line is unsafe or incomplete.",
        "Retain the signed form with the PTW and the Daily EHS Report for the project "
        "retention period.",
        "A copy of every completed handover is filed against KAZ-EHS-MS-2026-004 Annex A.",
    ], size=10)

    apply_header_footer(doc, "KAZ-EHS-FRM-2026-006", "0")
    out = os.path.join(REPO, "FRM-KAZ-EHS-FRM-2026-006_Shift Handover Log (Rev.0).docx")
    doc.save(tmp)
    shutil.move(tmp, out)
    prune_orphans(out)
    print("WROTE:", out)
    return out


# ============================================================================
# FORM 007 - NIGHT POUR EHS CHECKLIST PACK
# ============================================================================
C1 = [
    ("A. DOCUMENTATION, PERMITS & CERTIFICATES", [
        "Valid PTW issued/endorsed for this night shift, with all mandatory attachments",
        "Method Statement KAZ-EHS-MS-2026-004 and JSA Tables 10.1/10.2 approved and on site",
        "Concrete pump truck & placing boom \u2014 valid third-party inspection certificate sighted",
        "Pump operator \u2014 valid third-party competency card sighted",
        "All mixer drivers \u2014 valid heavy vehicle licences and third-party induction sighted",
        "Site Medic certification (CPR / trauma / HV shock response) sighted",
        "Ambulance inspection record and equipment check completed",
        "Calibrated lux meter certificate valid; anemometer available",
        "This checklist pack printed, and a clean copy available at the pour site",
    ]),
    ("B. CREW, FATIGUE & BRIEFING", [
        "100 % night crew deployed \u2014 no day-shift worker continuing into the night",
        "Every crew member recorded \u2265 8 h rest (Checklist C9)",
        "Dedicated, fully rested Night EHS Officer present for the whole shift",
        "Pre-shift fitness checks completed for all crew, drivers and operators",
        "Day \u2192 Night handover log (FRM-2026-006) completed and signed",
        "Night Toolbox Talk delivered; attendance signed (Checklist C8)",
        "TBT covered approach distances, live energized zones, evacuation routes",
        "Spotter, Banksman, Permit Holder and Medic named and present; radios tested",
    ]),
    ("C. BARRICADING, REBAR & ACCESS", [
        "Pour area, open edges and pump perimeter enclosed with rigid hard barricades",
        "Reflective tape and warning flashers fitted; illuminated signage at entry",
        "Single controlled entry/exit manned; PTW access control enforced",
        "100 % of exposed vertical/horizontal rebar ends fitted with mushroom caps",
        "Access pathways clear, lit and free of rebar, debris, standing water",
        "Heavy-duty wheel stoppers at mixer discharge positions",
        "Open excavations/pits/edges barricaded, lit, signed; lit ladders provided",
        "Formwork, reinforcement, embedments released by Quality; pour sequence agreed",
    ]),
    ("D. PLANT, PUMP & TRAFFIC", [
        "Pump on firm, level, compacted ground clear of excavations/services",
        "Outriggers 100 % extended on pads; ground bearing confirmed",
        "Boom maximum envelope measured, marked and briefed; clearance verified",
        "Emergency stops, alarms, hydraulics, hoses, clamps inspected; line secured",
        "Mixers one-way routed; holding area defined; speed limit briefed",
        "Reverse alarms, lights, mirrors, brakes functional; baton working",
        "Wind measured < 32 km/h (or manufacturer limit); monitored during shift",
        "Washout pit/skip and spill kit in position; extinguishers at pump & boards",
    ]),
    ("E. ELECTRICAL, GROUNDING & RCD", [
        "Pump body grounding connected directly to the station earth grid",
        "Earthing continuity/resistance measured and recorded (C3)",
        "30 mA RCD on all outlets feeding vibrators, boards and temporary lights",
        "RCD test button operated on every device; trip/reset confirmed and logged",
        "Cables elevated off wet ground and fresh concrete on insulated hangers",
        "Enclosures/connectors IP-rated; nothing standing in water",
        "No damaged cables or taped repairs; defective equipment removed & tagged",
        "Insulated rescue hook and insulating mats at the pour site",
    ]),
    ("F. ILLUMINATION & VISIBILITY", [
        "\u2265 100 lux at the concrete discharge / pouring point (C4)",
        "\u2265 50 lux along access pathways and transit routes (C4)",
        "Towers angled downward \u2014 no operator glare; no dark shadows",
        "Battery-operated backup floodlights staged, charged and tested",
        "High-intensity torches + whistles to Spotter, Banksman, EHS, Medic",
        "Reflective marking on plant, outriggers, lines, barricades, vehicle tails",
        "Class 3 high-visibility attire on 100 % of night personnel",
    ]),
    ("G. PPE & RESCUE GEAR", [
        "Steel-toe rubber boots, chemical/rubber gloves, sealed goggles worn by handlers",
        "Face shields / waterproof aprons / long sleeves for high-splatter tasks",
        "Helmets with chin strap where fall or struck-by risk exists",
        "Hearing protection, FFP3 masks, knee protection available as required",
        "PPE inspected at TBT; damaged/contaminated PPE replaced",
        "Illuminated first-aid kit at pour site; contents & expiry checked",
        "Clean water / eyewash available; MSDS accessible",
    ]),
    ("H. MEDICAL & EMERGENCY", [
        "Standby ambulance at the pour zone for the entire night shift",
        "Dedicated ambulance driver on duty, remaining with the vehicle",
        "Ambulance: oxygen, spinal board, resuscitation kit, working lights/siren",
        "Certified Site Medic on duty on site for the whole shift",
        "Evacuation route clear of all mixers, machinery and obstructions",
        "Route/travel time/receiving point at Al-Zubair General Hospital confirmed",
        "Alarm signal, muster point, contacts and radio channel briefed at TBT",
        "Scenario responses briefed: shock, impalement, struck-by, fall, burn, boom failure",
    ]),
    ("I. ENVIRONMENT & HOUSEKEEPING", [
        "Washout restricted to designated pit/skip away from drains & earth grid",
        "Waste segregation in place; MOTAT records ready",
        "Spill kit available; crew briefed on hydraulic/fuel spill response",
        "Housekeeping completed and maintained; walkways and egress kept clear",
    ]),
]

C2 = [
    "Concrete pump truck \u2014 third-party inspection certificate (no. / expiry)",
    "Placing boom \u2014 third-party inspection / load test certificate (no. / expiry)",
    "Pump operator \u2014 third-party competency card (no. / expiry)",
    "Mixer driver 1 \u2014 heavy vehicle licence + induction (no. / expiry)",
    "Mixer driver 2 \u2014 heavy vehicle licence + induction (no. / expiry)",
    "Mixer driver 3 \u2014 heavy vehicle licence + induction (no. / expiry)",
    "Mixer trucks \u2014 vehicle inspection records (brakes, tyres, lights, reverse alarm)",
    "Immersion vibrators \u2014 portable appliance test tags in date",
    "Temporary distribution boards \u2014 inspection record and IP rating",
    "Light towers \u2014 inspection record; fuel/cable condition",
    "Lux meter \u2014 calibration certificate (no. / due date)",
    "Anemometer \u2014 calibration / function check",
    "Earthing kit \u2014 lead, clamps, continuity tester condition",
    "Insulated rescue hook \u2014 certification / test date",
    "Ambulance \u2014 inspection and roadworthiness record",
    "Site Medic \u2014 certification (CPR / trauma / HV shock) validity",
]

C3 = [
    "Pump truck body grounding connected to the station earth grid (point identified)",
    "Earthing continuity / resistance measured \u2014 record value and limit applied",
    "RCD 1 \u2014 board/ID ______ : test button operated, tripped and reset",
    "RCD 2 \u2014 board/ID ______ : test button operated, tripped and reset",
    "RCD 3 \u2014 board/ID ______ : test button operated, tripped and reset",
    "RCD 4 \u2014 board/ID ______ : test button operated, tripped and reset",
    "All vibrator and tool cables elevated on insulated hangers",
    "No cables, joints or sockets in water, laitance or fresh concrete",
    "Enclosures IP-rated, dry, undamaged; no exposed conductors",
    "No taped repairs or unauthorised extensions; defects removed and tagged",
    "Light towers aimed downward; no glare; no dark shadows in work zone",
    "Insulated mats and rescue hook available at the pour site",
]

C4 = [
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

C5 = [
    "Class 3 high-visibility reflective attire issued to and worn by all personnel",
    "Steel-toe rubber safety boots (EN ISO 20345 S5) worn by all concrete handlers",
    "Heavy-duty chemical / rubber gloves (EN 388 / EN ISO 374) issued",
    "Sealed safety goggles (EN 166) issued; face shields for high-splatter tasks",
    "Helmets (EN 397) with chin strap where fall / struck-by risk exists",
    "Hearing protection available near pump, vibrator and mixer",
    "FFP3 dust masks available for cement handling and dry sweeping",
    "Waterproof trousers / apron and knee protection for finishers",
    "Insulated electrical rescue hook stationed at the pour site",
    "Illuminated first-aid kit stocked, sealed and within expiry at the pour site",
    "Clean water / eyewash station available; barrier cream provided",
    "Contaminated or damaged PPE replaced immediately; spare stock on site",
]

C6 = [
    "Standby ambulance stationed directly at the concrete pouring area",
    "Dedicated standby driver present and remaining with the ambulance",
    "Oxygen cylinder charged and fitted",
    "Spinal board present and accessible",
    "Resuscitation kit present and sealed",
    "Trauma dressings and burns kit present",
    "Emergency lights and siren operational (tested)",
    "Certified Site Medic on duty on site for the whole shift",
    "Stretcher accessible at the pour site",
    "Evacuation route to site gate clear of all plant and obstructions",
    "Route, travel time and receiving point at Al-Zubair General Hospital confirmed",
    "Hospital pre-notified of night works and crew numbers",
    "Emergency contacts board displayed at the pour site and in the ambulance",
    "Alarm signal and muster point briefed; headcount method agreed",
]

C7 = [
    "Pour area perimeter enclosed with rigid hard barricades (no soft tape)",
    "Open edges hard-barricaded with handrails / fencing",
    "Concrete pump perimeter enclosed and signed",
    "Reflective tape fitted to all barricades, plant and outriggers",
    "Warning flashers operating on the perimeter at night",
    "Single controlled entry/exit point established and manned",
    "PTW access control enforced; visitor log and escort rule in use",
    "Illuminated / retro-reflective signage at the entry point",
    "100 % of exposed rebar ends fitted with approved mushroom caps",
    "Spare mushroom caps kept at the pour site for replacements",
    "Heavy-duty wheel stoppers installed at mixer discharge positions",
    "Open excavations / pits barricaded, lit and signed; lit ladders provided",
    "Access pathways and transit routes clear and lit",
]

C8_TOPICS = [
    "Scope, sequence and planned volume of tonight's pour; joint locations",
    "Safe approach distances; location of live energized zones and red zones",
    "Marked boom envelope and the Spotter's stop signal; never stand under boom/hose",
    "Barricaded perimeter, single controlled entry, PTW access control",
    "Mixer routing, one-way circuit, reversing rules, banksman signals, wheel stoppers",
    "Rebar mushroom caps, open edges, excavations, fresh-concrete exclusion",
    "Electrical hazards: 30 mA RCDs, elevated cables, pump grounding, rescue hook",
    "Wet concrete chemical burns: gloves, goggles, boots, immediate washing, eyewash",
    "Illumination limits (100 / 50 lux), glare control, action if lights fail",
    "Emergency evacuation routes, muster point, alarm, ambulance, hospital route",
    "Fatigue, rest breaks, hydration and STOP WORK AUTHORITY",
    "Environmental controls: washout pit, spill kit, waste segregation",
    "Weather limits: wind > 32 km/h, lightning, dust storm \u2014 stop rules",
    "Lessons learned from previous incidents / near misses on the project",
]

C10 = [
    "Pump, boom, delivery lines and vibrators cleaned at the designated washout area",
    "Residual concrete and washout water contained; disposed via approved route",
    "Barricades, towers, boards and cables recovered or left safe, lit, barricaded",
    "Temporary lighting left on over fresh concrete and open edges",
    "Plant parked in designated area; keys controlled",
    "Emergency egress lane restored and kept clear",
    "Waste segregated and removed to approved storage; MOTAT records updated",
    "PTW closed jointly by Permit Holder and Permit Issuer",
    "Site condition, isolations and remaining hazards recorded at closure",
    "Any incident / near miss recorded on form 000480 and in the Daily EHS Report",
    "Night \u2192 Day handover log (FRM-2026-006) completed and signed",
    "Fresh concrete, edges, openings and washout re-inspected at first light",
]


def build_checklists():
    tmp = TMP.replace("_ms_tmp", "_frm007_tmp")
    make_shell(tmp, "Night Concrete Pouring - EHS Checklist Pack",
               "EHS; Checklists; Night Work; Concrete Pouring; KAZ-EHS-FRM-2026-007")
    doc = Document(tmp)
    setup_section(doc.sections[0], landscape=False)

    banner(doc, "CONTROLLED FORM PACK", "NIGHT CONCRETE POURING \u2014 EHS CHECKLIST PACK",
           "Pre-Pour Readiness, Certification, Electrical, Illumination, PPE, Medical, "
           "Barricading, TBT, Crew Fitness & Permit Closure",
           "KAZ-EHS-FRM-2026-007", "0")

    heading(doc, "HOW TO USE THIS PACK", space_before=120)
    bullets(doc, [
        "Complete every checklist physically on site, at night, with the lighting on and "
        "the plant positioned as it will be used during the pour.",
        "Checklists C1\u2013C7 must be fully signed BEFORE any concrete is placed; C8 and C9 "
        "before the crew enters the pour zone; C10 at shift end.",
        "Any \u201cNo\u201d answer is a hold point: the pour shall not start until the item is "
        "closed or the work is re-scheduled to daylight.",
        "The signed pack is filed with the Permit to Work and referenced by "
        "KAZ-EHS-MS-2026-004 Sections 9 and 11.",
        "This pack is issued under the Siemens Energy KAZ Power Plant Upgrade Project EHS "
        "management system; the logo and document control above are part of the record.",
    ], size=10)

    heading(doc, "CHECKLIST C1 \u2014 PRE-POUR NIGHT READINESS (MASTER)")
    for gtitle, items in C1:
        checklist_table(doc, gtitle, items)
    signature_block(doc, "C1 READINESS AUTHORISATION \u2014 POUR MAY / MAY NOT COMMENCE", [
        "Permit Holder / Civil Supervisor",
        "Dedicated Night EHS Officer",
        "Permit Issuer (Siemens Energy)",
        "Area Authority (client / operations, where applicable)",
        "Night Shift In-Charge / Construction Manager",
        "Site Manager (Siemens Energy)",
        "Concrete Pouring Subcontractor \u2014 Authorized Representative",
    ])

    heading(doc, "CHECKLIST C2 \u2014 PLANT & EQUIPMENT CERTIFICATION")
    rows = [["#", "Certificate / Record to be Sighted on Site", "Sighted (Y/N)",
             "No. / Reference", "Expiry", "Verified By"]]
    for i, it in enumerate(C2, start=1):
        rows.append([str(i), it, BLANK, BLANK, BLANK, BLANK])
    grid(doc, [380, 4300, 900, 1700, 1100, 1484], rows, size=9, header_size=8.5,
         aligns=["c", None, "c", None, "c", None], total=USABLE_P)

    heading(doc, "CHECKLIST C3 \u2014 ELECTRICAL, GROUNDING & RCD")
    rows = [["#", "Verification Item", "Result / Reading", "Pass / Fail", "Initials", "Time"]]
    for i, it in enumerate(C3, start=1):
        rows.append([str(i), it, BLANK, BLANK, BLANK, BLANK])
    grid(doc, [380, 4600, 1700, 1000, 1000, 1184], rows, size=9, header_size=8.5,
         aligns=["c", None, None, "c", "c", "c"], total=USABLE_P)

    heading(doc, "CHECKLIST C4 \u2014 ILLUMINATION (LUX) SURVEY")
    body(doc, "Measured with a calibrated light meter at task / walking surface level. "
              "Re-measure after any change in tower position and at least once mid-shift.",
         size=9.5)
    rows = [["#", "Measurement Point", "Required", "Measured (lux)", "Pass / Fail",
             "Action if Fail", "Time"]]
    for i, (pt, req) in enumerate(C4, start=1):
        rows.append([str(i), pt, req, BLANK, BLANK, BLANK, BLANK])
    grid(doc, [380, 3200, 900, 1250, 950, 1900, 1284], rows, size=9, header_size=8.5,
         aligns=["c", None, "c", "c", "c", None, "c"], total=USABLE_P)

    heading(doc, "CHECKLIST C5 \u2014 PPE & RESCUE GEAR")
    checklist_table(doc, None, C5)

    heading(doc, "CHECKLIST C6 \u2014 MEDICAL & EMERGENCY STANDBY")
    checklist_table(doc, None, C6)

    heading(doc, "CHECKLIST C7 \u2014 BARRICADING, REBAR PROTECTION & ACCESS")
    checklist_table(doc, None, C7)

    heading(doc, "CHECKLIST C8 \u2014 NIGHT TOOLBOX TALK RECORD")
    label_table(doc, [
        ("Date / Time of TBT", BLANK),
        ("Pour location / activity", BLANK),
        ("Delivered by (Night EHS Officer)", BLANK),
        ("PTW number", BLANK),
        ("Language(s) used", "\u2610 English   \u2610 Arabic   \u2610 Other: ______"),
    ])
    rows = [["#", "Mandatory Topic \u2014 tick when covered", "Covered (\u2713)"]]
    for i, tp in enumerate(C8_TOPICS, start=1):
        rows.append([str(i), tp, BLANK])
    grid(doc, [400, 7800, 1664], rows, size=9, header_size=9,
         aligns=["c", None, "c"], total=USABLE_P)
    rows = [["#", "Name / Surname", "Company", "Role", "Signature"]]
    for i in range(1, 17):
        rows.append([str(i), BLANK, BLANK, BLANK, BLANK])
    grid(doc, [400, 2800, 2200, 2200, 2264], rows, size=9, header_size=9,
         aligns=["c", None, None, None, None], total=USABLE_P)

    heading(doc, "CHECKLIST C9 \u2014 NIGHT CREW REST & FITNESS REGISTER")
    body(doc, "Minimum 8 hours rest before the shift for every individual. No person may "
              "work more than one shift in any 24-hour period.", size=9.5)
    rows = [["#", "Name / Surname", "Company", "Role on Night Shift",
             "Rest Start / End (\u2265 8 h)", "Fitness (\u2713)", "Certs (\u2713)", "Signature"]]
    for i in range(1, 19):
        rows.append([str(i), BLANK, BLANK, BLANK, BLANK, BLANK, BLANK, BLANK])
    grid(doc, [340, 1880, 1240, 1760, 1520, 880, 1180, 1064], rows, size=8.5,
         header_size=8, aligns=["c", None, None, None, "c", "c", "c", None],
         total=USABLE_P)

    heading(doc, "CHECKLIST C10 \u2014 POST-POUR, HOUSEKEEPING & PERMIT CLOSURE")
    checklist_table(doc, None, C10)
    signature_block(doc, "C10 CLOSURE SIGN-OFF", [
        "Permit Holder / Civil Supervisor",
        "Night EHS Officer",
        "Permit Issuer (Siemens Energy)",
        "Concrete Pouring Subcontractor Representative",
    ])

    apply_header_footer(doc, "KAZ-EHS-FRM-2026-007", "0")
    out = os.path.join(REPO, "FRM-KAZ-EHS-FRM-2026-007_Night Pour EHS Checklist Pack (Rev.0).docx")
    doc.save(tmp)
    shutil.move(tmp, out)
    prune_orphans(out)
    print("WROTE:", out)
    return out


if __name__ == "__main__":
    build_handover()
    build_checklists()
