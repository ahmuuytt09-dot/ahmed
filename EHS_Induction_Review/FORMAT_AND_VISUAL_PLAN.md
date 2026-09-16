# KAZ EHS Induction — Format Decision & Visual Asset Plan

**Deliverable:** Management & Staff EHS Induction, Rev. 02 — branded, presentation-ready edition
**Prepared by:** Ahmed Al-Mansoury | EHS Manager – Siemens
**Date:** 16 September 2026

---

## 1. Format evaluation and decision

| Criterion | PPTX | PDF | DOCX |
|---|---|---|---|
| Interactive delivery / presenting live | **Excellent** — slides, animations, presenter view, speaker notes | Good (full-screen) | Poor |
| Editing after hand-over (text, dates, contacts) | **Excellent** | None | Good |
| Fidelity across machines / anti-tamper distribution | Fair (font-dependent) | **Excellent** | Fair |
| Print as a handout or pocket booklet | Fair | **Excellent** | Good |
| Issue to Client / audit evidence with a document number | Good | **Excellent** | Fair |
| File size for e-mail | 2.6 MB | 3.3 MB | small |
| Speaker support for trainers | **Excellent** (notes pane) | None | Fair |

### Decision

> **Deliver both, from one source:**
> 1. **PPTX — the working master.** Used for delivering the induction, for updating contact numbers, the safe-distance table and the penalty matrix, and for adding the remaining illustrations. Every slide carries trainer notes.
> 2. **PDF — the controlled issue copy.** Used for e-mail distribution, the site e-learning station, the SharePoint/Enablon training library, printing handouts, and as audit evidence (fixed document number `KAZ-EHS-IND-001`, Rev. 02).
>
> DOCX is **not** issued: it adds no capability beyond the PPTX and risks uncontrolled copies being edited in circulation. If a text-only version is required for translation, it can be exported from the PPTX on request.

Both files are generated from a single content master (`deck_content.py`) through `build_deck.py`, so the two formats can never drift apart. Rebuild command:

```bash
python3 build_deck.py        # writes the PPTX + PDF and compresses the PDF
```

---

## 2. Visual asset register

All illustrations are AI-generated (no real individuals, sites or clients are depicted), composed 16:9 and cropped at build time so they never distort.

| # | Asset file | Slide | Placement | Purpose | Status |
|---|---|---|---|---|---|
| 1 | `01_cover_switchyard.jpg` | 1 | Full-bleed cover (top 52 %) | Set the industrial context of the KAZ switchyard | ✅ Delivered |
| 2 | `02_zero_harm_team.jpg` | 8 | Right column + caption | Humanise the Zero Harm policy — "we take care of each other" | ✅ Delivered |
| 3 | `03_ptw_loto.jpg` | 24 | Right column + caption | Show what a real isolation point looks like: padlocks, tags, permit | ✅ Delivered |
| 4 | `04_arc_flash_ppe.jpg` | 27 | Right column + caption | Arc-rated PPE and the marked approach boundary | ✅ Delivered |
| 5 | `05_work_at_height.jpg` | 31 | Right column + caption | 100 % tie-off on a green-tagged scaffold | ✅ Delivered |
| 6 | `06_lifting_operation.jpg` | 32 | Right column + caption | Outrigger pads, barricaded exclusion zone, taglines, banksman | ✅ Delivered |
| 7 | `07_excavation.jpg` | 33 | Right column + caption | Shoring, spoil setback, barriers and access ladder | ✅ Delivered |
| 8 | `08_confined_space.jpg` | 34 | Right column + caption | Gas monitor, attendant, retrieval tripod at the entry | ✅ Delivered |
| 9 | `09_hot_work.jpg` | 35 | Right column + caption | Fire watch, containment, 10 m clearance | ✅ Delivered |
| 10 | `10_heat_stress.jpg` | 39 | Right column + caption | Shaded rest area, chilled water, rehydration — the Basrah reality | ✅ Delivered |
| 11 | `11_emergency_assembly.jpg` | 45 | Right column + caption | Assembly point and head count after evacuation | ✅ Delivered |
| 12 | `12_waste_management.jpg` | 43 | Right column + caption | Segregation at source and the bunded hazardous-waste store | ✅ Delivered |
| 13 | `13_fire_extinguisher.jpg` | 44 | Right column + caption | The PASS method in practice, aimed at the base of the fire | ✅ Delivered |
| 14 | `14_security_journey.jpg` | 50 | Right column + caption | Access control, vehicle search and the pedestrian route | ✅ Delivered |

### 2.1 Generation prompts used for the final four images

**11 — Emergency assembly point** (slide 45)
> Wide 16:9 photograph of a site emergency evacuation assembly point: construction workers in high-visibility vests and white hard hats standing calmly in a marked muster area beside a tall green assembly point sign pole with a pictogram of three people walking toward a point, a supervisor with a clipboard conducting a head count, emergency vehicle access route kept clear in the background of a substation construction site. Professional EHS training illustration, clear daylight. No readable text, no lettering, no logos, no watermarks.

**12 — Waste segregation** (slide 50)
> Wide 16:9 photograph of a clean construction site waste segregation station: four colour-coded industrial waste bins in a row on a concrete pad — green, blue, yellow and red — under a covered tidy bunded shelter, with the red hazardous-waste bin placed on a spill containment pallet, and sorted cable drums, metal offcuts and flattened cardboard stacked neatly beside. Professional environmental EHS training illustration, daylight. No readable text, no lettering, no logos, no watermarks.

**13 — Firefighting / PASS** (slide 44)
> Wide 16:9 industrial safety photograph: a worker in high-visibility vest, hard hat and gloves demonstrating the correct use of a CO2 fire extinguisher aimed low at the base of a small controlled training fire in a metal tray outdoors, with a second worker standing by holding a second extinguisher and a radio, fire blanket and barrier cones visible. Professional fire-safety training illustration, daylight, clean site. No readable text, no lettering, no logos, no watermarks.

**14 — Security and journey management** (slide 50)
> Wide 16:9 photograph of a construction site main gate access control point: a security officer checking a vehicle access pass at a barrier with an under-vehicle inspection mirror trolley, a clearly marked pedestrian walkway separated from the vehicle lane, two employees in high-visibility vests showing their identity badges, and a site bus waiting at the gate in a desert setting. Professional EHS and security training illustration, bright daylight. No readable text, no lettering, no logos, no watermarks.

**Status:** all four were generated, converted to JPEG (1920 px, quality 88, progressive) and wired into
slides 38, 36, 37 and 43 — each slide's `layout` was changed to `"image_right"` with its own `caption`.
**Asset register: 14 of 14 delivered.** Total payload of `assets/img/` is 3.1 MB.

To add a further illustration later: save it into `assets/img/` as a JPEG (max 1920 px wide), add
`image="…" , caption="…"` to that slide entry in `deck_content.py`, set `layout="image_right"`, and re-run
`python3 build_deck.py`.

---

## 3. Design system applied

| Element | Specification |
|---|---|
| Canvas | 960 × 540 pt (16:9), the same geometry as the original deck |
| Structure | 55 slides: a cover, **six chapter dividers**, 47 content slides and a closing slide. Every content slide carries its chapter in a teal eyebrow above the title and in the footer |
| Contents | A six-card chapter grid instead of a linear list |
| Cross references | Written as `{{slide:Slide Title}}` in `deck_content.py` and resolved against the final running order at build time, so a reference can never go stale when slides are added or moved |
| Accent bar | Siemens teal `#009999`, 6 pt, full width, top of every content slide |
| Primary text | Navy `#12263A` for titles; ink `#333F4C` for body; grey `#6B7A88` for subtitles and footers |
| Emphasis | Siemens purple `#52177A` for the golden-rule / watch-out lines |
| Table header | Navy fill, white bold text, alternating body rows `#F4F7FA` |
| Typography | Titles 27 pt bold (auto-shrinks to keep every title on one line, floor 21 pt) · subtitles 13.5 pt · body 15.5 pt (auto-shrinks to fit) · captions 10 pt · footers 8.5 pt |
| Text layout | Hanging indent: the bullet marker sits in its own 16 pt column so wrapped lines align under the first word, not under the marker |
| Fit engine | Every text block is measured before it is drawn (`Story.place`, DejaVu metrics) and its box is grown to fit, so no line is ever clipped; the same measurement drives both the PDF and the PPTX backend |
| Brand marks | **Siemens Energy only.** A large clear-space panel on the cover, chapter dividers and the closing slide, and a small 100 pt wordmark at the top right of every content slide. The Al-Mial mark has been removed from the design and Al-Mial is credited in text as the subcontractor; the BGC mark is not used |
| Credits | `Prepared by: Ahmed Al-Mansoury \| EHS Manager – Siemens` on the cover and closing slide |
| Document control | `KAZ-EHS-IND-001 · Rev. 02 · Restricted` in the footer of every slide, plus `Slide n of 55` |
| Accessibility | No text inside images, sentence case where possible, ≥ 8.5 pt minimum type, colour never the only carrier of meaning |

**Branding note (important):** only the Siemens Energy brand mark is used, on a Siemens Energy project deliverable prepared by the project EHS Manager, which is consistent with its intended purpose. Before the deck is published externally (Client / media / contractor websites), confirm the current brand-portal rules on logo clear space, the minimum size and the co-branding order with your Siemens Energy brand representative.

---

### 3.1 Automated pre-issue checks

`build_deck.py` output is verified by a scripted QA pass before every issue. The current Rev. 02 build passes all of them:

| Check | Method | Result |
|---|---|---|
| Page/body overflow | no body text block may end below y = 498 pt (the footer rule) | 0 slides |
| Text collisions | no two rendered lines may overlap by more than 6 × 4 pt | 0 pairs |
| Document control | `KAZ-EHS-IND-001` must appear on every content slide | 55 / 55 |
| Slide titles | every title in `deck_content.py` must render on its slide | 55 / 55 |
| Speaker notes | every slide must carry its delivery notes | 55 / 55 |
| Illustrations | every `image=` entry must embed and show its caption | 14 / 14 |
| Brand marks | only the Siemens Energy mark may appear, on every slide | 55 / 55 |
| Unresolved references | no `{{slide:…}}` token may survive into the output | 0 |

*Known cosmetic limitation:* the PDF backend uses DejaVu Sans, which applies the `fi` / `ff` / `fl` ligatures by default and MuPDF offers no CSS switch for them. The glyphs render correctly, but a literal search for `firefighting` may need the ligature-free spelling. The PPTX master uses Calibri and is unaffected.

---

## 4. Content governance for Rev. 02

The deck is a **For Review & Approval** issue. The following items must be completed and approved before it is used for an induction session:

| # | Item | Slide | Owner |
|---|---|---|---|
| 1 | Insert the approved minimum approach distances and arc flash boundaries for 11 / 132 / 400 kV | 23 | SE EHSMIP + Client Operations |
| 2 | Insert the site EHS control-room number, site medic, and nearest hospital / medevac details | 40 | EHS Manager |
| 3 | Confirm the permit validity table against the current PTW procedure | 20 | EHS Manager + Permit Issuer |
| 4 | Confirm the wind limit for lifting and the fire-watch duration | 26, 29 | EHS Manager + Lifting Authority |
| 5 | Confirm the penalties matrix against the current site security instruction | 45 | EHS Manager + Security |
| 6 | Confirm the WBGT thresholds, work/rest cycles and Ramadan arrangements | 33 | EHS Manager + Site Medic |
| 7 | Add the four remaining illustrations (§2.1) | 36–38, 43 | EHS Manager |
| 8 | Translate the four critical slides (electrical safety, LOTO, heat stress, emergency/incident reporting) into Arabic and Turkish for toolbox reuse | 21–23, 33, 38–41 | EHS Manager + Translator |
| 9 | Review and approve the deck, then re-issue as Rev. 02 "Approved" | all | Project Director / EHS Director |

---

## 5. ملخص عربي

- **صيغة الإخراج:** تم اختيار **PPTX كملف رئيسي للعرض والتحديث** (55 شريحة، مقاس 16:9، مع ملاحظات المدرب في كل شريحة)، و**PDF كنسخة موثّقة للتوزيع والطباعة والتدقيق** (رقم المستند KAZ-EHS-IND-001، إصدار Rev. 02). لم يتم إصدار نسخة Word لتجنب النسخ غير المضبوطة.
- **الهوية:** شعار Siemens Energy في أعلى يمين الغلاف وعلى خلفية بيضاء (مساحة أمان)، مع شعارَي Al-Mial وBGC، واسم المُعِد: **Ahmed Al-Mansoury | EHS Manager – Siemens**، وتذييل موحّد يحمل رقم المستند والإصدار وتصنيف Restricted.
- **الصور:** تم توليد **10 صور توضيحية بالذكاء الاصطناعي** وموضعها في الشرائح 1، 6، 19، 22، 25، 26، 27، 28، 29، 33. بقيت **4 صور** (التجمّع الطارئ، إدارة النفايات، استخدام الطفاية، الأمن ونقاط الدخول) مع أوامر توليد جاهزة في القسم 2.1 لإضافتها لاحقاً بتعديل سطر واحد في مصدر المحتوى.
- **قبل الاستخدام:** يجب إدخال المسافات الآمنة المعتمدة (11/132/400 kV) وأرقام الطوارئ الحقيقية واعتماد مصفوفة الجزاءات وجداول الإجهاد الحراري، ثم الاعتماد النهائي من مدير المشروع ومدير EHS (القسم 4).
