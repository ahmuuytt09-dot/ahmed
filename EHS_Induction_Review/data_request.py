#!/usr/bin/env python3
"""Data request for the KAZ EHS induction (Rev. 03).

Single source of truth for the information the induction still needs.  It both
*generates* the fill-in form and *reads* it back, so the two can never drift
apart.

    python3 data_request.py build   -> writes DATA_REQUEST.md and the .xlsx
    python3 data_request.py read    -> reports which fields are still empty

Two files are produced so you can use whichever is easier:

  DATA_REQUEST.md                     editable directly in the GitHub web editor
  KAZ_Induction_Data_Request.xlsx     download, fill in Excel, upload back

The reader prefers a filled .xlsx cell, then the markdown value.
Values containing a pipe or a trailing placeholder are ignored.

SECURITY NOTE: this repository is currently PUBLIC.  Everything written into
these files is visible to anyone on the internet.  See the warning at the top
of DATA_REQUEST.md before entering live emergency or security numbers.
"""
import os
import re
import sys

SHEET = "Induction Data"
XLSX = "KAZ_Induction_Data_Request.xlsx"
MD = "DATA_REQUEST.md"

# ---------------------------------------------------------------------------
# The requested information.  (id, group, label_ar, label_en, help)
# ---------------------------------------------------------------------------
FIELDS = [
    # A - emergency and rescue
    ("A1", "A", "رقم غرفة الطوارئ / خط الطوارئ في الموقع 24 ساعة",
     "Site emergency control room / 24-7 emergency line",
     "مثال: 0785-000-0000 — مفتوح 24/7"),
    ("A2", "A", "قناة الراديو UHF للطوارئ",
     "UHF radio channel for emergencies", "مثال: Channel 1"),
    ("A3", "A", "رقم المسعف / نقطة الإسعاف وموقعها",
     "Site medic / first aid post — number and location",
     "المبنى أو الموقع داخل المعسكر"),
    ("A4", "A", "مواقع أجهزة الصدمات الكهربائية AED",
     "AED locations (all units)",
     "مطلوب في بطاقة الطوارئ المعتمدة — اذكر كل المواقع"),
    ("A5", "A", "المستشفى الأقرب ورقمه",
     "Nearest hospital — name and phone", ""),
    ("A6", "A", "مستشفى الحروق",
     "Hospital for burns", "حقل إلزامي في بطاقة الطوارئ"),
    ("A7", "A", "مستشفى العيون",
     "Hospital for eye injuries", "حقل إلزامي في بطاقة الطوارئ"),
    ("A8", "A", "مستشفى حالات البتر / الجراحة",
     "Hospital for detached body parts / surgery",
     "حقل إلزامي في بطاقة الطوارئ"),
    ("A9", "A", "ترتيب الإخلاء الطبي (طائرة / سيارة) وجهة الاتصال",
     "Medevac arrangement (air / road) and contact",
     "من ينظم الإخلاء: Siemens Field Service أم العميل"),
    ("A10", "A", "نقاط التجمع ومواقعها",
     "Assembly points — locations",
     "الموقع الرئيسي + نقاط المعسكر والمكاتب"),
    ("A11", "A", "رقم سيارة الإسعاف في الموقع",
     "On-site ambulance number", ""),

    # B - electrical safety data
    ("B1", "B", "المسافات الآمنة والقوس الكهربائي — 11 ك.ف",
     "Approach distance and arc flash — 11 kV",
     "الحد الأدنى للمسافة | حد القوس | الطاقة الحادثة | فئة PPE"),
    ("B2", "B", "المسافات الآمنة والقوس الكهربائي — 132 ك.ف",
     "Approach distance and arc flash — 132 kV", "نفس التنسيق"),
    ("B3", "B", "المسافات الآمنة والقوس الكهربائي — 400 ك.ف",
     "Approach distance and arc flash — 400 kV", "نفس التنسيق"),
    ("B4", "B", "مرجع دراسة القوس الكهربائي (الرقم والتاريخ)",
     "Arc flash study reference and date",
     "مثال: SE Arc Flash Study KAZ-2026-014، تاريخ 2026-03"),

    # C - hazardous gases
    ("C1", "C", "هل يوجد غاز عزل المفاتيح SF6 في النطاق؟",
     "Is SF6 present in the scope?",
     "اكتب: نعم / لا — وإذا نعم اذكر المعدات والجهة المسؤولة"),
    ("C2", "C", "إجراء التعامل مع SF6 وإعادة التعبئة",
     "SF6 handling and refill procedure", "اسم الإجراء أو رقمه"),
    ("C3", "C", "هل يوجد H2S أو غازات سامة في المنطقة؟",
     "Is H2S or any toxic gas present?",
     "نعم / لا — وإذا نعم: حد التعرض بالـ ppm ومناطق الخطر"),
    ("C4", "C", "غازات أخرى (ميثان، أول أكسيد الكربون، نقص أكسجين)",
     "Other gases (methane, CO, oxygen deficiency)", ""),
    ("C5", "C", "أجهزة كشف الغاز الشخصية — النوع والعدد",
     "Personal gas detectors — type and quantity", ""),

    # D - NDT / radiography
    ("D1", "D", "هل يوجد تصوير إشعاعي أو NDT في النطاق؟",
     "Is radiography / NDT in the scope?", "نعم / لا"),
    ("D2", "D", "إن وُجد: الشركة المنفذة والترخيص وإجراء المنطقة",
     "If yes: contractor, licence and area control procedure", ""),

    # E - site layout and photographs
    ("E1", "E", "صورة البيس كمب / المكتب الرئيسي — اسم ملف الصورة",
     "Base camp / main office photograph — file name",
     "ارفع الصورة إلى المجلد ثم اكتب اسم الملف، مثال: basecamp.jpg"),
    ("E2", "E", "مخطط الموقع — اسم ملف المخطط",
     "Site plot plan — file name",
     "يجب أن يُظهر المكاتب والمخيم والعيادة ونقاط التجمع وطرق الطوارئ"),
    ("E3", "E", "مواقع العيادة / المطعم / المغسلة / المصلى",
     "Clinic / canteen / laundry / prayer room locations", ""),
    ("E4", "E", "حدود مناطق الخطر الحمراء وقيود الدخول",
     "Red zone boundaries and access restrictions", ""),

    # F - reporting route
    ("F1", "F", "الإجراء المعتمد للإبلاغ: بطاقة السلامة أم Enablon أم الاثنان؟",
     "Approved reporting route: Safety Card, Enablon, or both?",
     "يجب توحيد الإجراء بين الإندكشن وبطاقة الطوارئ"),
    ("F2", "F", "رابط أو رقم نموذج Enablon",
     "Enablon form link or number", ""),
    ("F3", "F", "مهلة تعبئة بطاقة السلامة الورقية",
     "Deadline for completing the paper Safety Card",
     "بطاقة الطوارئ تقول: في نفس اليوم"),
    ("F4", "F", "جهة استلام التقرير والتحقيق",
     "Who receives the report and investigates", ""),

    # G - language, training and logistics
    ("G1", "G", "اللغات المطلوبة للإندكشن",
     "Required induction languages",
     "عربي / تركي / كردي / أخرى — البند 3.8 من SE Instruction يُلزم بذلك"),
    ("G2", "G", "الشرائح التي تريد ترجمتها",
     "Which slides to translate",
     "افتراضي: 15 شريحة حرجة (القواعد، الكهرباء، الارتفاع، الرفع، الحفر، "
     "الأماكن المغلقة، الحرائق، الطوارئ، الأمن، العقوبات)"),
    ("G3", "G", "عدد المشاركين المتوقع في الجلسة",
     "Expected number of participants per session", ""),
    ("G4", "G", "مدة الجلسة المعتمدة بالدقائق",
     "Approved session duration in minutes",
     "الملف الحالي يقول 75 دقيقة — غير كافٍ لـ 47 شريحة محتوى"),
    ("G5", "G", "هل توجد شاشة/بروجكتر ووسائط بصرية في قاعة التدريب؟",
     "Projector / screen and visual media available in the training room?", ""),

    # H - security and camp
    ("H1", "H", "رقم غرفة الأمن الرئيسية",
     "Main security control room number", ""),
    ("H2", "H", "إجراء الزوار المعتمد",
     "Approved visitor procedure", ""),
    ("H3", "H", "مشرف المعسكر — الاسم والرقم",
     "Camp supervisor — name and number", ""),
    ("H4", "H", "رقم إدارة EHS في الموقع",
     "Site EHS department number", ""),

    # I - approval
    ("I1", "I", "مدير المشروع — الاسم",
     "Project Manager — name", "لصفحة الاعتماد"),
    ("I2", "I", "مدير EHS — الاسم",
     "EHS Manager — name", "لصفحة الاعتماد"),
    ("I3", "I", "تاريخ الاعتماد المطلوب",
     "Target approval date", ""),
    ("I4", "I", "رقم الإصدار المطلوب",
     "Issue revision to publish", "افتراضي: Rev. 03"),
    ("I5", "I", "هل تريد صفحة اعتماد وتوقيعات وسجل إصدارات؟",
     "Add an approval page, signatures and revision history?", "نعم / لا"),

    # J - anything else
    ("J1", "J", "مخاطر أو مواضيع خاصة أخرى يجب إضافتها",
     "Any other hazards or topics that must be added", ""),
    ("J2", "J", "ملاحظات حرة",
     "Any other notes", ""),
]

GROUPS = {
    "A": ("الطوارئ والإنقاذ الطبي", "Emergency and rescue medical"),
    "B": ("بيانات السلامة الكهربائية المعتمدة", "Approved electrical safety data"),
    "C": ("الغازات الخطرة", "Hazardous gases"),
    "D": ("التصوير الإشعاعي و NDT", "Radiography and NDT"),
    "E": ("الموقع والصور والمخطط", "Site layout, photographs and plot plan"),
    "F": ("مسار الإبلاغ", "Incident reporting route"),
    "G": ("اللغة والتدريب واللوجستيات", "Language, training and logistics"),
    "H": ("الأمن والمعسكر", "Security and camp"),
    "I": ("الاعتماد", "Approval"),
    "J": ("أي شيء آخر", "Anything else"),
}

HEADER = """# نموذج بيانات الإندكشن · Induction Data Request

**المستند:** KAZ-EHS-IND-001 · إصدار الإندكشن المطلوب: Rev. 03
**الغرض:** المعلومات الناقصة التي يجب تعبئتها قبل إصدار Rev. 03 المعتمد.
**How to use:** اكتب القيمة بعد `**VALUE:**` مباشرة ثم احفظ (Commit). اتركها فارغة
إذا لم تتوفر بعد. يمكنك أيضاً تحميل ملف Excel `{xlsx}` وتعبئته ورفعه.

---

## ⚠️ تحذير أمني مهم — اقرأه قبل التعبئة

هذا المستودع **عام (Public)** على GitHub، أي أن كل ما تكتبه هنا **يراه أي شخص على
الإنترنت**. لذلك:

| نوع البيانات | التوصية |
|---|---|
| أرقام الطوارئ والمسعف والمستشفيات وأجهزة AED ومواقع نقاط التجمع والأمن | **لا تكتبها هنا** إذا كان المستودع عاماً |
| بيانات عامة (أسماء مسؤولين، مدد، لغات، نعم/لا، مراجع دراسات) | آمنة نسبياً — يمكن كتابتها |

**الحلول المتاحة — اختر واحداً:**

1. **اجعل المستودع خاصاً** (خاص أوصي به): من صفحة المستودع → Settings → عام →
   غيّر الرؤية إلى Private. بعدها عُدّ إلى هذا الملف واكتب كل شيء فيه بأمان.
2. **اتركها فارغة هنا** وأرسل الأرقام الحساسة لي مباشرة في المحادثة — سأدمجها في
   الملف النهائي ولن تُنشر في المستودع العام.
3. **اكتبها بصيغة مؤقتة** مثل `xxx-xxx-xxxx` لحين تحويل المستودع إلى خاص.

البنود الحساسة مُعلّمة بـ 🔒 في الأسفل. لن أصدر Rev. 03 بأرقام وهمية — إما أرقام
حقيقية معتمدة أو تبقى خانة «to be inserted» في نسخة المراجعة.

---

## الحالة
"""


def build():
    # ---------------- markdown ----------------
    out = [HEADER.format(xlsx=XLSX), "",
           "| المجموعة | الموضوع | عدد الحقول |", "|---|---|---|"]
    for g, (ar, en) in GROUPS.items():
        n = sum(1 for f in FIELDS if f[1] == g)
        out.append(f"| {g} | {ar} — {en} | {n} |")
    out.append("")
    out.append(f"**الإجمالي: {len(FIELDS)} حقلاً.** اكتب القيمة بعد `**VALUE:**` "
               "واحفظ التغيير.")
    out.append("")

    sensitive = {"A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "A10",
                 "A11", "H1", "H3", "B1", "B2", "B3", "E2", "E4"}
    for g, (ar, en) in GROUPS.items():
        out.append("---")
        out.append("")
        out.append(f"## {g} · {ar} — {en}")
        out.append("")
        for fid, fg, lar, len_, help_ in FIELDS:
            if fg != g:
                continue
            lock = " 🔒" if fid in sensitive else ""
            out.append(f"### {fid} · {lar}{lock}")
            out.append(f"*{len_}*")
            if help_:
                out.append(f"`{help_}`")
            out.append("")
            out.append("**VALUE:**")
            out.append("")

    out.append("---")
    out.append("")
    out.append("## كيف أقرأ البيانات")
    out.append("")
    out.append("بعد تعبئة النموذج أخبرني، أو شغّل الأمر التالي لمعرفة ما تبقّى:")
    out.append("")
    out.append("```bash")
    out.append("python3 data_request.py read")
    out.append("```")
    out.append("")
    with open(MD, "w", encoding="utf-8") as fh:
        fh.write("\n".join(out))
    print(f"{MD}: {len(FIELDS)} fields written")

    # ---------------- excel ----------------
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        print("openpyxl not installed - skipped the .xlsx")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws.sheet_view.rightToLeft = True

    ws["A1"] = "نموذج بيانات الإندكشن — KAZ Management & Staff EHS Induction (Rev. 03)"
    ws["A1"].font = Font(bold=True, size=13, color="0F3D5C")
    ws["A2"] = ("🔒 تنبيه: هذا المستودع عام — لا تكتب أرقام الطوارئ أو الأمن الحقيقية "
                "قبل تحويله إلى خاص")
    ws["A2"].font = Font(bold=True, size=10, color="B00020")
    ws["A3"] = ("اكتب القيمة في العمود C. اتركها فارغة إذا لم تتوفر. "
                "بعد الانتهاء ارفع الملف إلى المستودع.")
    ws["A3"].font = Font(size=9, italic=True, color="555555")

    head = ["المعرّف", "الحقل · Field", "القيمة · VALUE", "ملاحظات · Notes"]
    for j, h in enumerate(head, 1):
        c = ws.cell(row=5, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor="0F3D5C")
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 22

    thin = Side(style="thin", color="C3CED8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    r = 6
    for g, (ar, en) in GROUPS.items():
        gc = ws.cell(row=r, column=1, value=f"{g} · {ar} — {en}")
        gc.font = Font(bold=True, size=10, color="0F3D5C")
        gc.fill = PatternFill("solid", fgColor="E7EEF3")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        gc.alignment = Alignment(horizontal="right")
        r += 1
        for fid, fg, lar, len_, help_ in FIELDS:
            if fg != g:
                continue
            ws.cell(row=r, column=1, value=fid).font = Font(bold=True, size=9)
            ws.cell(row=r, column=2, value=f"{lar}\n{len_}").font = Font(size=9)
            ws.cell(row=r, column=3, value=None).font = Font(size=9, color="0B5CAD")
            ws.cell(row=r, column=4, value=help_).font = Font(size=8, color="6B7A88")
            for col in range(1, 5):
                cc = ws.cell(row=r, column=col)
                cc.border = border
                cc.alignment = Alignment(wrap_text=True, vertical="top", horizontal="right")
            ws.row_dimensions[r].height = 30
            r += 1

    for col, w in zip("ABCD", (9, 46, 34, 34)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A6"
    wb.save(XLSX)
    print(f"{XLSX}: {r - 6} rows written")


# ---------------------------------------------------------------------------
def parse_md(path):
    """Return {field_id: value} from the markdown form."""
    if not os.path.exists(path):
        return {}
    values, current = {}, None
    for line in open(path, encoding="utf-8"):
        m = re.match(r"^###\s+([A-J]\d+)\s+·", line)
        if m:
            current = m.group(1)
            continue
        if current and line.startswith("**VALUE:**"):
            v = line.split("**VALUE:**", 1)[1].strip()
            if v:
                values[current] = v
            current = None
    return values


def parse_xlsx(path):
    """Return {field_id: value} from the spreadsheet."""
    if not os.path.exists(path):
        return {}
    try:
        from openpyxl import load_workbook
    except ImportError:
        return {}
    wb = load_workbook(path, data_only=True)
    ws = wb[SHEET] if SHEET in wb.sheetnames else wb.active
    values = {}
    for row in ws.iter_rows(min_row=6, max_col=3):
        fid = row[0].value
        if isinstance(fid, str) and re.fullmatch(r"[A-J]\d+", fid.strip()):
            v = row[2].value
            if v is not None and str(v).strip():
                values[fid.strip()] = str(v).strip()
    return values


def read():
    md, xl = parse_md(MD), parse_xlsx(XLSX)
    merged = dict(md)
    merged.update(xl)                       # a filled Excel cell wins
    print(f"\n  sources: {MD} ({len(md)}) · {XLSX} ({len(xl)})")
    print(f"  {'':4s} {'FIELD':52s} STATUS")
    print("  " + "-" * 78)
    done = 0
    for g, (ar, en) in GROUPS.items():
        rows = [f for f in FIELDS if f[1] == g]
        filled = [f for f in rows if f[0] in merged]
        print(f"\n  {g} · {ar}  —  {len(filled)}/{len(rows)}")
        for fid, fg, lar, len_, help_ in rows:
            ok = fid in merged
            done += ok
            mark = "✔ " + merged[fid][:34] if ok else "·  (فارغ)"
            print(f"     {fid:4s} {lar[:44]:46s} {mark}")
    print("\n  " + "=" * 78)
    print(f"  المجموع: {done} / {len(FIELDS)} حقلاً معبّأ "
          f"({done * 100 // len(FIELDS)}%)")
    if done == 0:
        print("  → النموذج فارغ. املأه ثم أخبرني أو شغّل read مرة أخرى.")
    elif done < len(FIELDS):
        print(f"  → تبقّى {len(FIELDS) - done} حقلاً.")
    else:
        print("  → اكتمل النموذج — جاهز لإصدار Rev. 03.")
    print()
    return merged


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else "read"
    {"build": build, "read": read}[cmd]()
