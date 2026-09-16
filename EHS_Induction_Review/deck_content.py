"""Content master for the KAZ Management & Staff EHS Induction (Rev. 02 — branded edition).

Single source of truth used to render the PPTX (presentation) and the PDF (distribution copy).
All text is professional English. Layout coordinates are in points on a 960 x 540 canvas (16:9).

slide fields
------------
layout : cover | section | bullets | image_right | image_full | table | closing
title  : main title
sub    : subtitle / kicker
bullets: list of (level, text)   level 0 = bullet, 1 = sub-bullet, 2 = plain line
image  : asset filename in assets/img
caption: caption under / over an image
table  : dict(head=[...], rows=[[...]], widths=[...])
notes  : trainer notes (written into PPTX speaker notes)
"""

PREPARED_BY = "Prepared by: Ahmed Al-Mansoury | EHS Manager – Siemens"
DOC_NO = "KAZ-EHS-IND-001"
REV = "Rev. 02"
CLASSIFICATION = "Restricted — © Siemens Energy. Siemens Energy is a trademark licensed by Siemens AG."

TITLE = "Management & Staff EHS Induction"
PROJECT = "KAZ Power Plant Upgrade Project"
SITE = "Khor Al-Zubair Gas Turbine Power Plant · Basrah, Iraq"
CLIENT = "Client: Ministry of Electricity (MoE) & Basrah Gas Company (BGC)"
SUBCONTRACTOR = "Subcontractor: Al-Mial Company"
CONTRACTOR = "Principal Contractor: Siemens Energy"


# ---------------------------------------------------------------------------
# Course structure - six chapters.  A divider slide is generated before the
# slide named in "starts_at", and every slide after it carries its chapter in
# the header, so the deck reads like a corporate training course.
# ---------------------------------------------------------------------------
CHAPTERS = [
    dict(no="01", name="Introduction and Objectives",
         blurb="How this session runs, what you will learn and what is recorded.",
         topics=["Session format", "Learning objectives", "Validity", "Assessment"],
         starts_at="How This Session Runs"),
    dict(no="02", name="EHS Leadership and Responsibilities",
         blurb="The Siemens EHS target and who is accountable for what on this site.",
         topics=["Zero Harm", "EHS Target", "Leadership", "Everyone", "Contractors"],
         starts_at="Siemens EHS Target — Our Vision"),
    dict(no="03", name="The Project, the Site and Daily Rules",
         blurb="Where you are working, how the site is organised and the rules that always apply.",
         topics=["Project scope", "Site layout", "Base camp", "EHS organisation",
                 "General rules", "Stop Work Authority", "Risk assessment", "Daily routine"],
         starts_at="The Project"),
    dict(no="04", name="Permit to Work and Electrical Safety",
         blurb="The control system for hazardous work and the rules that keep the plant safe.",
         topics=["Permit to Work", "Lock-Out / Tag-Out", "Electrical rules", "Arc flash",
                 "Live plant"],
         starts_at="Permit to Work (1)"),
    dict(no="05", name="High-Risk Activities and Occupational Health",
         blurb="The activities that cause the most serious injuries and how each one is controlled.",
         topics=["Work at height", "Lifting", "Excavation", "Confined space", "Hot work",
                 "Tools", "Manual handling", "PPE", "Heat stress", "First aid"],
         starts_at="Work at Height"),
    dict(no="06", name="Environment, Emergency and Security",
         blurb="Protecting the environment, responding to emergencies, security and the consequences of breaking the rules.",
         topics=["Environment", "Waste", "Fire", "Emergency", "Incident reporting",
                 "Security", "Driving", "Penalties"],
         starts_at="Environmental Protection"),
]

SLIDES = [

# ---------------------------------------------------------------- 1. cover
dict(layout="cover", title=TITLE, sub=PROJECT, image="01_cover_switchyard.jpg",
     notes="Welcome participants, confirm fire exits and assembly point for the training room, "
           "and state that attendance is recorded. Introduce yourself and the EHS team."),

# ---------------------------------------------------------------- 2. how this session runs
dict(layout="bullets", title="How This Session Runs", sub="Emergency instructions for the training room",
     bullets=[(0, "If an emergency occurs during this induction: stop, follow the exit signs and go to the nearest assembly point."),
              (0, "Alarm signals: continuous siren = evacuate · intermittent = stand by · all-clear = return to work."),
              (0, "Mobile phones on silent — keep your phone available for emergencies."),
              (0, "Smoking is prohibited except in the designated smoking area, and never in red / restricted zones."),
              (0, "Duration: approx. 75 minutes, including the knowledge check and questions."),
              (2, "Emergency numbers are on slide {{slide:Emergency Contacts}} — know them before you leave this room.")],
     notes="Read the emergency arrangements of the training venue aloud. Confirm the assembly point "
           "location and the site emergency line before starting the content."),

# ---------------------------------------------------------------- 3. contents
dict(layout="contents", title="Contents", sub="Six chapters — about 75 minutes, including the knowledge check",
     notes="Walk through the six chapters quickly. Point out that the operators' own rule set is issued separately and that the knowledge check comes at the end."),

dict(layout="bullets", title="Learning Objectives", sub="At the end of this induction you will be able to:",
     bullets=[(0, "State the Siemens EHS target and the Zero Harm policy, and explain what it means for your work."),
              (0, "Describe your personal responsibilities and the authority you hold to stop unsafe work."),
              (0, "Apply the general site rules and select the correct PPE for your task."),
              (0, "Read a risk assessment, complete an LMRA and know when an SJA is required."),
              (0, "Explain how the Permit to Work and LOTO systems operate on this site."),
              (0, "Recognise the electrical, height, lifting, excavation, confined-space and hot-work hazards of KAZ."),
              (0, "Respond correctly to a fire, an emergency, an injury, a spill or a security event."),
              (0, "Report incidents, near misses and unsafe conditions within the required timeframes.")],
     notes="Link each objective to the participant's own role — managers and staff are expected to lead, "
           "verify and coach, not only comply."),

# ---------------------------------------------------------------- 5. validity & records
dict(layout="bullets", title="Validity, Assessment and Records", sub="This induction is part of your site authorisation",
     bullets=[(0, "Validity: 12 months from the date of attendance; renewal is required on expiry."),
              (0, "Re-induction is required after a serious incident or when the hazards of your work change."),
              (0, "Knowledge check: 10 questions — pass mark 80%. Below 80% you are retrained before site access."),
              (0, "Your attendance and signature are recorded in the site training register."),
              (0, "You must also complete Personal Safety Instruction (PSI) and, for high-risk activities, the Siemens Energy Safe Start session."),
              (0, "The induction card must be carried and shown on request — it does not replace the Permit to Work."),
              (0, "Minimum age for site access: 18 years.")],
     notes="Confirm that supervisors will be checked for valid induction, PSI and task certificates. "
           "Underline that the card is not a work permit."),

# ---------------------------------------------------------------- 6. Siemens EHS target (vision/objective)
dict(layout="image_right", title="Siemens EHS Target — Our Vision", sub="Derived from the KAZ Project EHS Plan (Section 5)",
     bullets=[(0, "Our standard for EHS is excellence — it must exceed customer expectations."),
              (0, "The project will promote a Zero-Harm culture that prevents damage to people, the environment and equipment."),
              (0, "Project target: zero harm — measured as Zero H1 = ZERO (no lost-time injuries)."),
              (0, "Where EHS, progress and finances appear to conflict, EHS always prevails."),
              (0, "Health and safety are conditions of every activity — never a trade-off against schedule."),
              (2, "Zero Harm is not a slogan: it is the way we plan, verify and behave on this site.")],
     image="02_zero_harm_team.jpg",
     caption="Zero Harm — we take care of each other",
     notes="This is the vision slide. Connect the EHS Plan target to the daily reality: every task planned, "
           "every control verified, every person home safely."),

# ---------------------------------------------------------------- 7. zero harm framework
dict(layout="bullets", title="The Zero Harm Framework", sub="Principles · Behaviours · Essentials",
     bullets=[(0, "Principles:"),
              (1, "Zero incidents is achievable."),
              (1, "Health and safety — no compromises."),
              (1, "We take care of each other."),
              (0, "Behaviours:"),
              (1, "Take the time to assess risk — plan and organise every job before you start."),
              (1, "Talk about safety before the work and again during the shift."),
              (1, "If a task cannot be done safely, or conditions change — stop and re-assess."),
              (1, "Follow the rules and procedures, including site-specific rules and PPE."),
              (1, "Intervene when you see risk — never walk by."),
              (0, "Essentials: Stop principle · safe systems of work · competency · risk assessment · permit to work · PPE · emergency preparedness.")],
     notes="Ask the group for one example of intervening on an unsafe act. Reinforce that interventions are "
           "expected and will never be penalised."),

# ---------------------------------------------------------------- 8. leadership responsibilities
dict(layout="bullets", title="Leadership Responsibilities", sub="For managers, engineers and supervisors",
     bullets=[(0, "Own EHS in your area — you are accountable for the people and the tasks you control."),
              (0, "Provide the resources: competent people, correct tools, valid PPE, adequate time."),
              (0, "Verify compliance — check permits, certificates, controls and housekeeping every day."),
              (0, "Lead visibly: safety walks, toolbox talks, safety stand-downs and recognition of good practice."),
              (0, "Make EHS the first agenda item in every formal meeting and record it in the minutes."),
              (0, "Stop any work you consider unsafe, and support anyone who stops work."),
              (0, "Ensure incidents and near misses are reported and investigated — and lessons are applied.")],
     notes="This is the key audience for this induction. Ask each manager to commit to one visible action "
           "within the next 7 days."),

# ---------------------------------------------------------------- 9. everyone's responsibilities
dict(layout="bullets", title="Everyone's Responsibilities", sub="Applies to employees, contractors and visitors",
     bullets=[(0, "Take care of your own health and safety and that of everyone around you."),
              (0, "Comply with site rules, permits, method statements and instructions."),
              (0, "Report hazards, incidents, near misses, unsafe acts and unsafe conditions."),
              (0, "Use and maintain PPE correctly; never work with defective equipment."),
              (0, "Attend inductions, PSI, toolbox talks and task-specific training."),
              (0, "Never work under the influence of alcohol or drugs — random testing is enforced."),
              (0, "Protect the environment: prevent spills, segregate waste, conserve energy and water."),
              (0, "Participate in risk assessments for your own tasks — you know the job best.")],
     notes="Emphasise the duty to report. Remind the group that honest reporting is protected by the "
           "just-culture approach: no blame for reporting."),

# ---------------------------------------------------------------- 10. contractor obligations
dict(layout="bullets", title="Contractor & Subcontractor Obligations", sub="How we work together",
     bullets=[(0, "Each contractor works to its own EHS system, enriched with Siemens Energy, Client and project requirements."),
              (0, "Every contractor appoints a competent Contractor EHS Representative."),
              (0, "Contractors provide: risk assessments, method statements, permits, competency certificates and equipment certificates."),
              (0, "Monthly man-hours and incident data are submitted to the EHS Department."),
              (0, "Siemens Energy may stop work, suspend, exclude or ban individuals for EHS violations — at the contractor's cost."),
              (0, "Where local legal requirements differ from Siemens Energy requirements, the stricter requirement applies.")],
     notes="Remind subcontractor managers of the documentation they must keep current; the EHS team audits it."),

# ---------------------------------------------------------------- 11. project description
dict(layout="bullets", title="The Project", sub=PROJECT,
     bullets=[(0, "Modernisation and expansion of the existing 400/132/11 kV Air Insulated Switchgear (AIS) facilities."),
              (0, "New 400 kV and 132 kV bays: circuit breakers, disconnectors, earthing switches, current and voltage transformers, surge arresters and busbar extensions."),
              (0, "Civil works: foundations, earthing grid enhancement, steel structures, cable trenches and ducts."),
              (0, "Electrical works: control, protection, metering, telecom and SCADA/DAS interfacing."),
              (0, "Activities: installation, testing, commissioning and trial operation."),
              (0, "Duration: 13.5 months. Client: MoE & BGC. Principal Contractor: Siemens Energy."),
              (2, "Critical context: the works are executed inside a LIVE power plant — coordination with MoE Operations is mandatory for every activity.")],
     notes="Stress the live-plant context: the greatest single risk on this project is electrical. "
           "Every activity needs operational coordination."),

# ---------------------------------------------------------------- 12. site layout
dict(layout="bullets", title="Site Layout and Emergency Facilities", sub="Know your location before you start work",
     bullets=[(0, "Your work area, mobilisation and laydown area, warehouses, camp and offices — shown on the site plot plan."),
              (0, "Emergency layer you must know:"),
              (1, "Assembly points (Muster points) for your area."),
              (1, "First-aid points, site clinic and the route to the nearest hospital."),
              (1, "Fire points, extinguishers and eyewash stations."),
              (1, "Emergency exits and the route for fire, ambulance and emergency vehicles."),
              (0, "Red zones: high-voltage areas where access is restricted and controlled."),
              (0, "The plot plan is posted at the site entrance, offices and assembly points — check it on your first day.")],
     notes="Use the actual plot plan from the project drawing set when presenting. Walk the group to their "
           "nearest assembly point after the session, if practical."),

# ---------------------------------------------------------------- 13. organisation
dict(layout="image_right", title="Base Camp and Site Offices", sub="The main office, the camp and the facilities you will use",
     bullets=[
         (0, "**Main office block** — project management, EHS office, document control, meeting rooms and the site clinic."),
         (0, "**Base camp** — accommodation, canteen, prayer room, laundry and recreation areas. Camp rules apply at all times."),
         (0, "**Camp and office safety:** keep escape routes and extinguishers clear; no cooking, no smoking and no unauthorised electrical appliances inside rooms or offices."),
         (0, "**Utilities** — potable water points, generators and UPS areas, waste collection points and the sewage / grey-water system."),
         (0, "**Housekeeping** — bins are provided in every block; report damage, leaks, pests or a faulty air conditioner to the camp supervisor immediately."),
         (0, "**Transport** — only approved site buses and drivers; report to the camp marshal before boarding and never leave site without authorisation."),
         (0, "**Emergency at the camp** — the camp assembly point is marked, and the clinic, first aiders and security gate are manned at all times."),
     ],
     image=None,
     caption="Base camp and main office — insert site photograph",
     notes="Insert the approved photograph of the base camp / main office before issuing this deck (right-click the placeholder and choose Change Picture). Walk the group through the layout: main office, clinic, canteen, laundry and the camp assembly point. Confirm the camp supervisor's name and the camp emergency number during the session."),

dict(layout="bullets", title="Project EHS Organisation", sub="Who is responsible for what",
     bullets=[(0, "Project Director / Site Manager — overall responsibility for EHS performance of the project."),
              (0, "Siemens Energy EHS Manager in Projects (EHSMIP) — governance, audits, verification and stop-work authority."),
              (0, "Contractor EHS Manager and EHS Officers — daily supervision, permits, inspections and training."),
              (0, "Area Authority / MoE Operations — authorises work in live operational areas."),
              (0, "Authorized Electrical Person (AEP) / Isolation Authority — energy isolation, LOTO and energisation."),
              (0, "First aiders, fire wardens and gas testers — appointed and posted in every area."),
              (2, "Rule: No EHS supervision = No work.")],
     notes="Introduce the real people behind each role and their contact numbers. Participants should know "
           "who to call for each situation."),

# ---------------------------------------------------------------- 14. general site rules 1
dict(layout="bullets", title="General Site Rules (1)", sub="Applies from the moment you enter the gate",
     bullets=[(0, "Complete site induction and PSI before starting any work — and carry your induction card."),
              (0, "No work without a valid Permit to Work; no work without EHS supervision."),
              (0, "Complete an LMRA before every task and attend the daily toolbox talk."),
              (0, "Read and understand the risk assessment and method statement for your task — ask if anything is unclear."),
              (0, "Wear the mandatory minimum PPE at all times (slide {{slide:Personal Protective Equipment}})."),
              (0, "Stay within your authorised area — never enter red or restricted zones without authorisation."),
              (0, "Keep your workplace clean and tidy — 'Clean everything, always.'"),
              (0, "Report every near miss, unsafe act, unsafe condition and incident.")],
     notes="Repeat the two absolutes: no permit = no work; no EHS supervision = no work."),

# ---------------------------------------------------------------- 15. general site rules 2
dict(layout="bullets", title="General Site Rules (2)", sub="Working hours, competency and behaviour",
     bullets=[(0, "Working hours: 06:00–15:00, six days per week; subcontractor hours must be approved by Siemens Energy."),
              (0, "Maximum 11 hours daytime / 7.5 hours night work, excluding breaks; one rest day in every seven. No fatigue working or driving."),
              (0, "Only competent, certified persons operate plant, equipment and electrical systems."),
              (0, "No horseplay, no sleeping on site, no unauthorised removal of materials (Material Gate Pass required)."),
              (0, "Zero tolerance: alcohol, drugs, weapons, fighting, harassment and theft."),
              (0, "Photography and video recording on site are prohibited without written authorisation from the Site EHS/Security Manager."),
              (0, "Report to the site medic any illness, injury or medication that could affect your work.")],
     notes="Have the site rule list available as a handout. Point out that hours and rest days are monitored."),

# ---------------------------------------------------------------- 16. stop work authority
dict(layout="bullets", title="Stop Work Authority", sub="You are empowered — and expected — to use it",
     bullets=[(0, "Anyone — employee, contractor or visitor — may stop any task where there is imminent danger, a rule violation, a missing control or a change in conditions."),
              (0, "How to apply it:"),
              (1, "1. Stop the activity."),
              (1, "2. Intervene with the people involved — calmly and respectfully."),
              (1, "3. Escalate to the supervisor or area authority."),
              (1, "4. Verify and implement the required controls."),
              (1, "5. Restart only when authorised."),
              (1, "6. Record the event and give positive feedback."),
              (0, "You will not be punished, prosecuted or harassed for stopping work — this is a Siemens Energy commitment.")],
     notes="This is the single most important empowerment message in the deck. Ask each participant to state "
           "one situation in which they would stop work."),

# ---------------------------------------------------------------- 17. risk management
dict(layout="bullets", title="Risk Management", sub="Risk Assessment · SJA · LMRA · Management of Change",
     bullets=[(0, "Risk assessment (RA): prepared before work starts and communicated to the team. Read it, follow the controls, never work outside its scope."),
              (0, "Safe Job Analysis (SJA): required for high-risk or non-routine tasks; prepared with the crew before work starts."),
              (0, "Last Minute Risk Assessment (LMRA): a short check before every task — What can hurt me? What has changed? Are the controls in place?"),
              (0, "Hierarchy of controls: eliminate · substitute · engineering controls · administrative controls · PPE (PPE is the last line of defence)."),
              (0, "Management of Change: if the method, equipment, people or conditions change, stop work and revise the assessment."),
              (0, "After any incident or high-potential near miss, the risk assessment is reviewed and updated.")],
     notes="Explain the difference between RA (planned), SJA (task) and LMRA (moment). Use a real job from the "
           "site as the example."),

# ---------------------------------------------------------------- 18. daily routine
dict(layout="bullets", title="The Daily Safety Routine", sub="Start of shift · during work · end of shift",
     bullets=[(0, "Start of shift: toolbox talk → LMRA → permit verification → PPE and equipment pre-use checks."),
              (0, "During work: supervision present, barriers and signage in place, housekeeping maintained, no lone working on high-risk tasks."),
              (0, "End of shift: close and endorse permits, restore isolations per the LOTO register, secure the area."),
              (0, "Shift handover for isolated work: joint re-verification of permit conditions, isolation status and zero-energy state, signed by all parties."),
              (0, "Any change of plan during the shift is briefed to the crew and, where required, re-approved.")],
     notes="Walk through a typical day of a subcontractor crew, step by step, so the routine is concrete."),

# ---------------------------------------------------------------- 19. PTW system
dict(layout="image_right", title="Permit to Work (1)", sub="System, roles and principles",
     bullets=[(0, "No permit = no work. The PTW system is administered and controlled by Siemens Energy."),
              (0, "For work on live or operational plant and client-controlled areas, the MoE/BGC permit system and its Area/Isolation Authority take precedence."),
              (0, "Roles: Permit Issuer (SE) · Permit Holder (contractor supervisor) · Area Authority (operations) · Isolation Authority / AEP · EHS Reviewer · Authorized Gas Tester."),
              (0, "Every permit requires: approved JSA / method statement, controls verified on site, briefed working party, valid certificates and inspections."),
              (0, "EHS reviews and audits permits; technical approval belongs to the Area Authority; isolation authority belongs to the AEP."),
              (0, "Never work outside the scope, time or conditions stated on the permit.")],
     image="03_ptw_loto.jpg",
     caption="No permit — no work",
     notes="Explain that the permit is a contract between the issuer and the holder. Removing or bypassing "
           "a permit condition is a dismissible offence."),

# ---------------------------------------------------------------- 20. PTW types
dict(layout="table", title="Permit to Work (2)", sub="Permit types and validity on this project",
     table=dict(head=["Permit", "Typical validity"],
                rows=[["General Cold Work Permit", "7 days with daily endorsement"],
                      ["Electrical Cold Work Permit", "7 days with daily endorsement"],
                      ["Electrical Hot Work & Commissioning Permit", "Daily / each shift"],
                      ["Energy Isolation (LOTO) Permit", "Daily / each shift"],
                      ["Hot Work Permit (welding, cutting, grinding)", "Daily / each shift"],
                      ["Excavation & Trenching Permit (over 30 cm)", "7 days with daily endorsement"],
                      ["Lifting Permit (all lifts; engineered plan for critical lifts)", "Daily / each shift"],
                      ["Work at Height Permit (1.8 m and above)", "Daily / each shift"],
                      ["Confined Space Entry Permit", "Daily / each shift"]],
                widths=[0.62, 0.38]),
     notes="Tell participants where to obtain each permit and who signs it. Confirm the current validity "
           "rules with the project PTW procedure before presenting."),

# ---------------------------------------------------------------- 21. LOTO
dict(layout="bullets", title="Lock-Out / Tag-Out and Energy Isolation", sub="Control every energy source",
     bullets=[(0, "Cover all energy sources: electrical, mechanical, pneumatic, hydraulic, thermal and stored energy."),
              (0, "Sequence: notify → isolate → lock and tag → release stored energy → verify zero energy → work."),
              (0, "Every worker applies their own personal lock; keys are controlled in the lockbox by the AEP."),
              (0, "PTW validity and LOTO validity are managed separately — permit expiry never releases an isolation."),
              (0, "Only the AEP de-isolates, through the controlled de-isolation process and the key register."),
              (0, "Never remove another person's lock or tag — this is a dismissible offence."),
              (0, "Test before touch: prove the tester, prove dead, and prove the tester again.")],
     notes="Use the isolation certificate and lockbox as physical props if available. Explain the personal "
           "lock principle clearly."),

# ---------------------------------------------------------------- 22. electrical safety rules
dict(layout="image_right", title="Electrical Safety Rules", sub="The golden rule of this project",
     bullets=[(0, "Treat every electrical circuit as live until proven dead."),
              (0, "Isolate the source and apply LOTO before starting any work."),
              (0, "Only an Authorized Electrical Person may isolate, de-isolate or energise."),
              (0, "Verify absence of voltage with an approved, calibrated tester — prove the tester before and after."),
              (0, "Use insulated tools rated for the voltage; inspect them before use."),
              (0, "Maintain safe approach distances at all times; respect barriers, interlocks and signage."),
              (0, "Never work on live equipment unless formally authorised and only under an Electrical Hot Work / Commissioning Permit."),
              (2, "GOLDEN RULE: ISOLATE — LOCK OUT — TAG OUT — TEST BEFORE TOUCH.")],
     image="04_arc_flash_ppe.jpg",
     caption="Arc-rated PPE and approach boundary",
     notes="This is the highest-risk topic of the project. Slow down here and check understanding with questions."),

# ---------------------------------------------------------------- 23. arc flash & distances
dict(layout="table", title="Arc Flash and Safe Approach Distances", sub="Boundaries must be approved for this site — do not use generic values",
     table=dict(head=["Nominal voltage", "Minimum approach distance", "Arc flash boundary", "Notes"],
                rows=[["11 kV", "per approved project table", "per arc flash study", "Limited / restricted approach per NFPA 70E and client rules"],
                      ["132 kV", "per approved project table", "per arc flash study", "Client operational rules apply"],
                      ["400 kV", "per approved project table", "per arc flash study", "Highest risk on this site — client operational rules apply"],
                      ["Work within 30 m of HV system", "Agreement with plant owner + physical demarcation", "—", "Mandatory before any activity"]],
                widths=[0.22, 0.28, 0.18, 0.32]),
     notes="IMPORTANT for the presenter: insert the values approved by the Siemens Energy EHSMIP and the client "
           "before delivering this slide. Generic or imperial values from other sites must not be used — this "
           "was a critical finding in the audit of the previous induction."),

# ---------------------------------------------------------------- 24. live plant coordination
dict(layout="bullets", title="Working in a Live Power Plant", sub="Coordination with MoE Operations",
     bullets=[(0, "Work near or on live systems is governed by the client's operational rules in addition to ours."),
              (0, "A nominated Operations Manager is responsible for the installation; a Switching Supervisor controls switching."),
              (0, "No work starts without: agreed method, permit, isolation and the operational authority's confirmation that the plant condition is safe."),
              (0, "Maintain physical demarcation between live and work areas; never breach a barrier or open a panel without authorisation."),
              (0, "Stop all work during thunderstorms — cranes, lifts, scaffolding and installation work are suspended (lightning and induced voltage)."),
              (0, "Report any contact with a conductor, any flash and any damaged insulation immediately."),
              (0, "Never touch an electric-shock casualty before the supply is isolated.")],
     notes="Explain the difference between our responsibility and the client's operational responsibility, and "
           "how both are coordinated daily."),

# ---------------------------------------------------------------- 25. work at height
dict(layout="image_right", title="Work at Height", sub="Fall protection from 1.8 m",
     bullets=[(0, "Fall protection is required from 1.8 m, or wherever a fall hazard exists."),
              (0, "Use approved scaffolding, work platforms or MEWPs — with guardrails and toe boards."),
              (0, "Harness with double lanyard and shock absorber, 100% tie-off to a certified anchor point (above 22 kN)."),
              (0, "Scaffolding: erected and altered only by certified scaffolders; green tag valid; inspected before each shift and after any alteration. Never work on a red-tagged scaffold."),
              (0, "Ladders: short-duration work only, secured, three points of contact, no loads carried."),
              (0, "Barricade the area below, tether tools, protect openings and edges, and never work at several levels without protection.")],
     image="05_work_at_height.jpg",
     caption="100% tie-off to certified anchor",
     notes="Ask who has a valid working-at-height certificate. Verify harness inspection records."),

# ---------------------------------------------------------------- 26. lifting
dict(layout="image_right", title="Lifting Operations", sub="Every lift is planned and permitted",
     bullets=[(0, "All lifts require a permit; critical lifts require a detailed engineered lifting plan."),
              (0, "Critical lifts include: loads above 75% of rated capacity, multi-crane lifts, lifts over live plant, and non-routine rigging."),
              (0, "Only certified cranes and lifting tackle (third-party certificates), competent operators, riggers and banksmen."),
              (0, "Outriggers fully extended on timber pads on firm, verified ground."),
              (0, "Exclusion zone barricaded — nobody under the load, no side pulls, no shock loading."),
              (0, "No lifting when wind exceeds 32 km/h or in poor visibility."),
              (0, "Taglines for load control; the banksman is the only person who signals the operator.")],
     image="06_lifting_operation.jpg",
     caption="Exclusion zone · taglines · banksman",
     notes="Lifting is a major hazard on this project because of the transformer and steel structure erection. "
           "Confirm the wind limit with the project lifting procedure."),

# ---------------------------------------------------------------- 27. excavation
dict(layout="image_right", title="Excavation and Trenching", sub="Permit required beyond 30 cm",
     bullets=[(0, "A permit is required for any ground penetration deeper than 30 cm."),
              (0, "Obtain the signed underground utility clearance before digging; hand-dig trial pits near services and the earthing grid."),
              (0, "Protective systems — shoring, benching or sloping — are mandatory for excavations deeper than 1.2 m."),
              (0, "Spoil heaps set back at least 1.0 m from the edge; materials and plant kept away from the crest."),
              (0, "Barricade, sign and light all excavations; access ladders every 7.5 m for emergency egress."),
              (0, "Inspect after rain or any ground movement; never enter an unprotected trench."),
              (0, "A competent banksman controls plant movement around the excavation.")],
     image="07_excavation.jpg",
     caption="Shoring · barriers · access ladder",
     notes="Highlight that excavations can also be confined spaces depending on risk assessment — but not "
           "automatically."),

# ---------------------------------------------------------------- 28. confined space
dict(layout="image_right", title="Confined Space Entry", sub="Permit · gas test · attendant · rescue plan",
     bullets=[(0, "Entry only with a Confined Space Entry Permit, trained entrants and attendants, and a rescue plan in place."),
              (0, "Atmospheric testing before entry and continuously during the work: oxygen, LEL, H2S and CO, by an Authorized Gas Tester."),
              (0, "Ventilation and isolation (LOTO) in place before entry."),
              (0, "The attendant remains at the entry at all times and maintains communication with the entrant."),
              (0, "Never enter to rescue without breathing apparatus and the rescue team — most confined-space fatalities are would-be rescuers."),
              (0, "Classification as a confined space is based on risk assessment of the actual conditions.")],
     image="08_confined_space.jpg",
     caption="Gas monitoring · attendant · tripod",
     notes="Emphasise the rescue rule: the instinct to help is the biggest killer. Call the rescue team."),

# ---------------------------------------------------------------- 29. hot work
dict(layout="image_right", title="Hot Work and Fire Watch", sub="Sparks and flames are controlled, permitted activities",
     bullets=[(0, "Hot Work Permit required for welding, cutting, grinding, soldering, thermal spraying or any activity producing sparks, flames or heat above 100 °C."),
              (0, "Certified welder or cutter; PPE and screens to protect others."),
              (0, "Fire blanket containment and at least 10 m clearance of combustible materials."),
              (0, "A dedicated Fire Watch is present during the work and for 30 minutes after it ends."),
              (0, "Minimum two extinguishers and a fire blanket at the work site."),
              (0, "Gas cylinders secured upright, caps fitted, oxygen and fuel gas stored separately."),
              (0, "Gas test required near fuel, chemical or potentially flammable areas — never hot work in a flammable atmosphere.")],
     image="09_hot_work.jpg",
     caption="Fire watch · containment · 10 m clearance",
     notes="Confirm the fire-watch duration and extinguisher requirements from the project PTW procedure."),

# ---------------------------------------------------------------- 30. tools
dict(layout="bullets", title="Hand and Power Tools", sub="The right tool, inspected, used correctly",
     bullets=[(0, "Use the right tool for the job; inspect before use; remove defective tools from service immediately."),
              (0, "Portable electrical tools: 30 mA RCD protection, undamaged armoured cables, earthing verified, no taped joints."),
              (0, "Angle grinders: guard in place, correct disc for the material, never over-speed a disc."),
              (0, "Wear a full face shield (EN 166) plus safety glasses and gloves for grinding and cutting."),
              (0, "Secure the workpiece; never cut towards your body; keep hands clear of the cutting line."),
              (0, "No loose clothing, no jewellery, no rings; tie back long hair."),
              (0, "Never use compressed air to clean clothing or skin.")],
     notes="This is a common source of hand injuries. Ask the group which tools are most frequently "
           "misused on site."),

# ---------------------------------------------------------------- 31. manual handling
dict(layout="bullets", title="Manual Handling and Ergonomics", sub="Plan the lift, not just the job",
     bullets=[(0, "Assess the load: weight, size, grip, centre of gravity, distance and destination."),
              (0, "Use mechanical aids or ask for help — team lifts need a leader and clear commands."),
              (0, "Lift technique: bend the knees, keep the load close, no twisting while lifting."),
              (0, "Plan breaks and vary tasks to reduce repetition and strain."),
              (0, "Report pain early — discomfort today becomes an injury tomorrow."),
              (0, "Consider the environment: heat, uneven ground and confined spaces all increase the risk.")],
     notes="Tie this to the heat stress session: dehydration and fatigue make manual handling injuries "
           "far more likely."),

# ---------------------------------------------------------------- 32. PPE
dict(layout="table", title="Personal Protective Equipment", sub="Mandatory minimum on site — additional PPE per task risk assessment",
     table=dict(head=["PPE item", "Requirement on this site"],
                rows=[["Safety helmet", "Mandatory — with chin strap; replaced after impact"],
                      ["Flame-retardant coverall", "Mandatory — visibility class 1 (high-visibility)"],
                      ["Safety footwear", "Mandatory — S3 / S1P, toe protection and puncture resistance"],
                      ["Safety glasses", "Mandatory — ANSI/EN certified"],
                      ["Protective gloves", "Mandatory — task-specific (electrical, chemical, cut-resistant)"],
                      ["Hearing protection", "Mandatory above 85 dB(A) or where indicated"],
                      ["Arc-rated PPE", "For electrical work — Category 2/4, face shield, insulated gloves"],
                      ["Full-body harness", "Working at height — double lanyard with shock absorber"],
                      ["Respiratory protection", "Where dust or vapour is identified — face-fit tested (e.g. FFP3)"],
                      ["High-visibility vest", "Mandatory in areas with mobile plant and vehicle movement"]],
                widths=[0.34, 0.66]),
     notes="PPE is the last line of defence. Remind everyone that PPE is provided free of charge and damaged "
           "PPE must be replaced, not repaired."),

# ---------------------------------------------------------------- 33. heat stress
dict(layout="image_right", title="Heat Stress Management", sub="Our most predictable environmental hazard",
     bullets=[(0, "Heat stress is managed by WBGT monitoring — monitoring starts when ambient temperature exceeds 40 °C."),
              (0, "Published work / rest cycles apply for light, moderate and heavy work; follow them."),
              (0, "Acclimatisation: new arrivals and anyone returning from leave start with reduced workload and increased rest."),
              (0, "Drink water regularly — about 1.0 to 1.2 litres per hour for light work in high heat. Use oral rehydration salts."),
              (0, "Shaded, ventilated rest areas and chilled drinking water are provided — use them."),
              (0, "Lone working is prohibited on high-risk tasks — use the buddy system."),
              (0, "Warning signs: headache, dizziness, cramps, nausea, confusion, hot dry skin, fainting. Heat stroke is a medical emergency — cool the casualty and call the medic immediately.")],
     image="10_heat_stress.jpg",
     caption="Shade · hydration · work-rest cycles",
     notes="Confirm the current WBGT thresholds, work-rest tables and Ramadan arrangements before presenting. "
           "This was a critical finding in the audit of the previous induction."),

# ---------------------------------------------------------------- 34. first aid
dict(layout="bullets", title="First Aid and Medical Emergency", sub="Know your nearest help",
     bullets=[(0, "Know: your nearest first-aid point, the site medic, the emergency number and the route to the nearest hospital."),
              (0, "Do not move a casualty unless there is immediate danger to life."),
              (0, "Electric shock: do not touch the casualty before the supply is isolated."),
              (0, "Keep the airway open; start CPR / AED if the casualty is not breathing and you are trained."),
              (0, "Preserve the incident scene; do not disturb equipment involved in the incident."),
              (0, "Meet the ambulance at the main gate and escort it to the casualty."),
              (0, "Report every injury — including first-aid cases — to the EHS Department and the site medic."),
              (2, "Fitness for work: a valid medical certificate is required; report any condition or medication that affects your ability to work safely.")],
     notes="Confirm the number of trained first aiders in each shift and area, and the current site clinic "
           "arrangements."),

# ---------------------------------------------------------------- 35. environment
dict(layout="bullets", title="Environmental Protection", sub="Prevent pollution — it is part of every task",
     bullets=[(0, "No spill of oil, fuel, chemical or concrete washwater to soil, drains or watercourses."),
              (0, "Use drip trays, secondary containment, funnels and spill kits; keep spill kits stocked and accessible."),
              (0, "Dust suppression on unpaved roads and during cutting, grinding or excavation."),
              (0, "Control noise and vibration through plant selection and work methods; respect site noise limits."),
              (0, "Chemicals: no substance on site without a Safety Data Sheet (SDS) and registration in the electronic substance register."),
              (0, "Asbestos and PCBs are prohibited on site — report any suspected material immediately."),
              (0, "Protect flora, fauna and heritage: no uncontrolled clearing, no hunting, no littering."),
              (0, "Report environmental incidents (spills, leaks, uncontrolled emissions) immediately — they are EHS incidents.")],
     notes="Explain where spill kits and SDS are kept, and how the electronic substance register is used."),

# ---------------------------------------------------------------- 36. waste
dict(layout="image_right", title="Waste Management", sub="Segregate at source — every day",
     bullets=[(0, "Segregate waste: general · recyclable (metal, cable, paper, plastic) · hazardous · construction and demolition waste."),
              (0, "Hazardous waste (oils, filters, batteries, chemicals, e-waste) is stored in labelled, covered, bunded areas."),
              (0, "Hazardous waste leaves site only with MOTAT records and a licensed transporter."),
              (0, "Never burn waste, never mix hazardous with general waste, never dump into trenches or open ground."),
              (0, "Material and waste removal requires a Material Gate Pass."),
              (0, "Keep waste areas tidy, marked and away from drains; report overflowing or damaged bins.")],
     image="12_waste_management.jpg",
     caption="Segregated bins · bunded store",
     notes="Show the actual bin colours and locations on site. Ask each supervisor to confirm their crews know "
           "where the hazardous waste store is."),

# ---------------------------------------------------------------- 37. fire
dict(layout="image_right", title="Fire Safety and Firefighting", sub="Classes, extinguishers and the PASS method",
     bullets=[(0, "Class A — wood, paper, cloth, rubber, plastics → water or foam (DCP is acceptable but less effective)."),
              (0, "Class B — gasoline, kerosene, paint, propane → foam, DCP, CO2."),
              (0, "Class C — energised electrical equipment → CO2 or clean agent, plus de-energise the supply. Never use water."),
              (0, "Class D — metals such as magnesium, titanium and sodium → special metal powder (graphite or sodium chloride). Never water or standard DCP."),
              (0, "Class K / F — cooking oils and fats → wet chemical (camp kitchen)."),
              (0, "PASS method: Pull the pin · Aim at the base of the fire · Squeeze the handle · Sweep side to side."),
              (0, "For CO2 extinguishers hold the insulated horn, never the metal parts."),
              (0, "Only fight a small fire if you are trained and have a clear escape route. Otherwise raise the alarm and evacuate. Always alarm first.")],
     image="13_fire_extinguisher.jpg",
     caption="Aim at the base · keep an escape route",
     notes="Correct the previous deck's error: Class D requires special metal powder, not DCP. Never use water "
           "on energised electrical equipment."),

# ---------------------------------------------------------------- 38. evacuation
dict(layout="image_right", title="Emergency Response and Evacuation", sub="Alarm · assembly · head count",
     bullets=[(0, "Alarm signals: continuous siren = evacuate · intermittent = stand by · all-clear = return to work."),
              (0, "On evacuation: stop the job, isolate your equipment, proceed to the nearest assembly point."),
              (0, "Report to your area marshal for the head count; report any missing person immediately."),
              (0, "Do not use vehicles; do not block emergency routes; keep the gate clear for emergency services."),
              (0, "Do not leave or return to the area until instructed by the emergency controller."),
              (0, "For security incidents (threat, intrusion, civil disturbance) follow the same assembly principle and the security team's instructions."),
              (0, "Fire wardens sweep their areas, including visitors and subcontractor crews.")],
     image="11_emergency_assembly.jpg",
     caption="Assembly point · roll call · clear routes",
     notes="Confirm the assembly point locations and the identity of area marshals for the audience's areas."),

# ---------------------------------------------------------------- 39. emergency classification
dict(layout="table", title="Emergency Classification and Notification", sub="Know the level — and who to tell",
     table=dict(head=["Level", "Description", "Notification"],
                rows=[["Level 1", "Minor injury or health effect, minor damage or minor environmental impact, manageable on site with site resources.",
                         "Supervisor → EHS → Site Manager (verbal immediately, written within 24 hours)"],
                      ["Level 2", "Significant event manageable locally with support from the Company or local authorities — major injury, local damage, local impact.",
                         "EHS → SE EHSMIP → Project Director + Client (immediately)"],
                      ["Level 3", "Major event or crisis beyond the contractor's control, with broader implications, requiring Company / corporate resources.",
                         "Project Director → SE Corporate / Regional Hub + Client (immediately)"],
                      ["Escalate immediately", "Life-threatening injury, fatality, high-potential incident, major environmental release, security crisis — never treated as a locally managed event.",
                         "Verbal immediately · written within 24 hours · investigation report within 7 days"]],
                widths=[0.16, 0.47, 0.37]),
     notes="Correct the previous deck's error: a fatality is never a Level 2 locally-managed event. Confirm the "
           "current notification matrix with the project notification plan."),

# ---------------------------------------------------------------- 40. emergency contacts
dict(layout="table", title="Emergency Contacts", sub="Posted at every assembly point, office and camp",
     table=dict(head=["Contact", "Number / channel"],
                rows=[["Site EHS control room / emergency line", "to be inserted (24/7)"],
                      ["UHF radio", "Channel 1"],
                      ["Site medic / first aiders", "to be inserted"],
                      ["Contractor emergency numbers", "0785-285-7404 · 0786-162-9706"],
                      ["Siemens Energy emergency number", "+90 533 625 9494"],
                      ["Ambulance (Iraq)", "122"],
                      ["Civil Defence / Fire (Iraq)", "115"],
                      ["Police (Iraq)", "104"],
                      ["Nearest hospital / medevac", "to be inserted"]],
                widths=[0.55, 0.45]),
     notes="IMPORTANT: replace the placeholder entries with live numbers and verify them weekly. The previous "
           "deck quoted 911, which is not an official emergency number in Iraq."),

# ---------------------------------------------------------------- 41. incident reporting
dict(layout="bullets", title="Incident, Near Miss and Unsafe Condition Reporting", sub="Report immediately — reporting is protected",
     bullets=[(0, "Report immediately: all injuries (including first aid), near misses, unsafe acts, unsafe conditions, property damage, environmental releases and security events."),
              (0, "Report through your supervisor and the EHS Department; incidents are recorded in the Siemens Energy EHS reporting tool (Enablon)."),
              (0, "Timelines:"),
              (1, "Verbal notification — immediately."),
              (1, "Written notification — within 24 hours."),
              (1, "Investigation report — within 7 days (unless an extension is agreed)."),
              (1, "Corrective action status — reported weekly until closed."),
              (0, "A formal investigation team is formed by the Contractor, the Company and subcontractors; root causes and actions are documented."),
              (0, "Never withhold information, alter the scene or delay reporting. Honest reporting carries no blame.")],
     notes="Give a practical example of a near miss report that prevented an accident. Encourage reporting "
           "rather than hiding."),

# ---------------------------------------------------------------- 42. lessons learned
dict(layout="bullets", title="Learning from Incidents", sub="Every event makes the site safer",
     bullets=[(0, "Lessons Learned and safety alerts are issued after incidents and high-potential near misses."),
              (0, "Each lesson triggers a risk assessment review, a toolbox talk and verification that controls are effective."),
              (0, "Siemens Energy may conduct Eye on Safety reviews and its own investigations."),
              (0, "If you are involved in or witness a serious incident: preserve the scene, cooperate fully and provide a written statement."),
              (0, "Share your experience — it prevents the next accident.")],
     notes="Show one recent lesson learned from the project or from Siemens Energy safety alerts."),

# ---------------------------------------------------------------- 43. security
dict(layout="image_right", title="Security and Access Control", sub="Everyone is searched — everyone is protected",
     bullets=[(0, "Site and camps are protected by the Iraqi Police and site security; persons, belongings, vehicles and deliveries may be searched on entry and exit."),
              (0, "Wear your Site ID / Gate Pass visibly at all times; access without it is refused."),
              (0, "Visitors: pre-approval 24 hours in advance, photo ID deposited at the gate, escorted by a host at all times; no work, no plant operation, no photography."),
              (0, "Report immediately: unauthorised access, theft, suspicious persons or vehicles, security breaches or threats."),
              (0, "Lost or damaged ID card must be reported; a replacement fee applies."),
              (0, "Security personnel are authorised to carry weapons — cooperate calmly and follow their instructions."),
              (0, "Never resist or interfere with security checks; the minimum age for site access is 18 years.")],
     image="14_security_journey.jpg",
     caption="Access control · vehicle search",
     notes="Cover the visitor process and the material gate pass. Security incidents follow the same reporting "
           "route as EHS incidents."),

# ---------------------------------------------------------------- 44. driving
dict(layout="bullets", title="Driving Safety and Journey Management", sub="The road is our highest-risk journey",
     bullets=[(0, "Seat belts for all occupants, in every vehicle, at all times."),
              (0, "Valid Iraqi driving licence for the vehicle class; only authorised drivers and vehicles."),
              (0, "Site speed limit: 5 km/h. Obey traffic signs, signals and marshals."),
              (0, "No mobile phone use while driving — including hands-free. No other distractions."),
              (0, "Reverse parking only in designated areas; reversing requires a trained banksman."),
              (0, "Daily pre-use checks and valid third-party inspection certificates for all vehicles."),
              (0, "Journey management: plan the route, rest at least every two hours, no night driving without approval, no driving when tired."),
              (0, "Driving is restricted in adverse weather — fog, heavy rain and dust storms.")],
     notes="Adapt the route and rest requirements to the actual Basrah–site conditions. Report all road "
           "incidents immediately."),

# ---------------------------------------------------------------- 45. violations
dict(layout="table", title="Violations and Penalties", sub="Consequences are applied consistently",
     table=dict(head=["Violation", "1st offence", "2nd offence", "3rd offence / critical"],
                rows=[["Speeding (over 5 km/h)", "Warning + 2-day vehicle ban", "1-week vehicle and driver ban", "Permanent site ban"],
                      ["PPE non-compliance", "Work stoppage + written warning", "Contractor fine", "ID cancellation and site removal"],
                      ["Red zone / LOTO breach", "Work stoppage + 1-week suspension", "—", "Permanent site expulsion"],
                      ["Lost or damaged ID card", "Replacement fee", "—", "—"],
                      ["Positive alcohol / drugs test", "Immediate permanent site ban", "—", "—"],
                      ["Unauthorised material removal", "Confiscation and investigation", "—", "—"],
                      ["Unauthorised photography", "Media deletion and warning", "—", "Permanent site ban"]],
                widths=[0.28, 0.26, 0.24, 0.22]),
     notes="Confirm the current penalty matrix with the site security instruction before presenting; penalties "
           "must be applied consistently to employees and contractors."),

# ---------------------------------------------------------------- 46. knowledge check
dict(layout="bullets", title="Knowledge Check", sub="10 questions — pass mark 80%",
     bullets=[(0, "1. What is the site speed limit, and when is a banksman required?"),
              (0, "2. At what height does fall protection become mandatory?"),
              (0, "3. When is a protective system required for an excavation?"),
              (0, "4. Which gases are tested before confined space entry?"),
              (0, "5. What is the maximum wind speed for lifting operations?"),
              (0, "6. Which extinguisher would you use on an energised electrical panel fire?"),
              (0, "7. What are the reporting deadlines for an incident?"),
              (0, "8. What must you do before touching any electrical conductor?"),
              (0, "9. What are the official emergency numbers in Iraq?"),
              (0, "10. What happens if you stop work because you believe it is unsafe?")],
     notes="Run the check individually, mark it immediately, and retrain anyone below 80% before site access. "
           "Keep the answer sheets in the training register."),

# ---------------------------------------------------------------- 47. card & acknowledgement
dict(layout="bullets", title="Induction Card and Acknowledgement", sub="Before you leave this room",
     bullets=[(0, "Collect your numbered induction card — valid 12 months, must be carried and shown on request."),
              (0, "Sign the attendance and acknowledgement register."),
              (0, "The card does not replace the Permit to Work, PSI, Safe Start or task-specific training."),
              (0, "Acknowledgement: 'I attended the KAZ Management & Staff EHS Induction, I understood its content, and I accept that compliance with safety rules, permits and Stop Work Authority is a condition of my access to this site.'"),
              (0, "Report any content gaps or unclear points to the EHS Department — this induction is reviewed regularly.")],
     notes="Collect signatures before participants leave. File the register with the training records."),

# ---------------------------------------------------------------- 48. closing
dict(layout="closing", title="Everyone Goes Home Safely — Every Day",
     sub="Zero Harm · We take care of each other",
     bullets=[(0, "Questions? Contact the EHS Department — UHF Channel 1 / Site EHS control room."),
              (0, "Thank you for your attention and for the work you do safely.")],
     notes="Close positively, thank the participants, and remind them of the induction card and the "
           "acknowledgement signature."),
]

# visual asset register - every illustration is now delivered
VISUAL_REGISTER = [
    ("Assembly point / head count", "11_emergency_assembly.jpg",
     "Site emergency assembly point: workers in high-visibility vests and hard hats standing in a marked muster "
     "area beside a green assembly-point sign pole, a supervisor with a clipboard conducting a head count, "
     "emergency vehicle access kept clear. No text or logos."),
    ("Waste segregation", "12_waste_management.jpg",
     "Clean construction site waste segregation station: four colour-coded industrial bins (green, blue, yellow, "
     "red) under a covered bunded shelter, the red hazardous-waste bin on a spill containment pallet, sorted "
     "cable and metal offcuts nearby. No text or logos."),
    ("Firefighting / PASS", "13_fire_extinguisher.jpg",
     "A worker in high-visibility vest demonstrating the PASS method with a CO2 fire extinguisher aimed at the "
     "base of a small controlled training fire, second worker standing by with a second extinguisher and radio. "
     "No text or logos."),
    ("Access control / journey", "14_security_journey.jpg",
     "Site main gate access control: security guard checking a vehicle pass at the barrier with a mirror trolley, "
     "clearly marked pedestrian walkway and speed limit signs, employees showing ID cards. No text or logos."),
]
