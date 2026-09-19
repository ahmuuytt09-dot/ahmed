#!/usr/bin/env python3
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Executive Clean Design - White / Light Grey
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY_BG = RGBColor(0xF5, 0xF5, 0xF7)  # #F5F5F7
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

def add_para(tf, txt, size=10, bold=False, color=DARK_TEXT, space=Pt(4), align=PP_ALIGN.LEFT):
    p = tf.add_paragraph()
    p.text = txt
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Calibri"
    p.space_after = space
    p.alignment = align
    return p

# Data for slides as requested
slides_data = [
    {
        "id": "01",
        "title_en": "PPE Violations",
        "title_ar": "مخالفات معدات الوقاية الشخصية",
        "desc": [
            "• عدم التزام كوادر المقاول بارتداء معدات الوقاية الأساسية: خوذة سلامة (Safety Helmet)، نظارات واقية (Safety Glasses)، أحذية سلامة (Safety Boots).",
            "• عدم استخدام حزام الأمان (Full Body Harness) عند العمل على ارتفاعات تزيد عن 1.8 متر.",
            "• ارتداء جزئي أو غير صحيح لمعدات الوقاية - خوذة بدون حزام ذقن، سترة عاكسة غير مغلقة.",
            "• الإشراف متواجد في الموقع لكن لا يطبق إجراءات السلامة - فشل إشرافي واضح."
        ],
        "risk": "HIGH",
        "risk_text": "خطر إصابات الرأس، العين، السقوط من الارتفاع، والانزلاق. زيادة مؤشر LTIR. فشل في تطبيق القاعدة الذهبية لإنقاذ الحياة رقم 3 (PPE).",
        "risk_en": "Risk of head injury, eye injury, fall from height, slip/trip. Direct violation of Life Saving Rule #3. Increases LTIFR/TRIR and indicates weak safety culture.",
        "action": [
            "1. تصحيح فوري في الموقع + توعية (Toolbox Talk) موثقة بالصور",
            "2. إصدار بطاقات ملاحظة سلامة (Observation Cards) وإشعار مخالفة رسمي",
            "3. حملة التزام يومية بمعدات الوقاية - فحص بداية المناوبة + فحص عشوائي",
            "4. ربط الالتزام بمؤشرات الأداء والدفعات المالية للمقاول",
            "5. تطبيق سياسة 3 مخالفات = استبعاد العامل والمشرف"
        ],
        "photo_label": "مخالفة عدم ارتداء الخوذة / النظارات / حزام الأمان"
    },
    {
        "id": "02",
        "title_en": "Working Without PTW",
        "title_ar": "العمل بدون تصريح عمل",
        "desc": [
            "• المباشرة بأنشطة عالية الخطورة (أعمال ساخنة Hot Work، أماكن مغلقة Confined Space، أعمال كهربائية) دون إصدار أو تفعيل تصريح العمل (Permit to Work).",
            "• عدم مراجعة تقييم المخاطر للمهمة (Task Risk Assessment) وعدم التحقق من عزل الطاقة (Energy Isolation).",
            "• سجل تصاريح العمل غير محدث، والتصاريح غير معروضة في موقع العمل كما هو مطلوب.",
            "• مخالفة مباشرة لإجراءات Siemens Energy وتعليمات المشروع - فقدان كامل لحاجز التحكم الأساسي."
        ],
        "risk": "CRITICAL",
        "risk_text": "خطر الوفاة، انفجار/حريق، إطلاق طاقة خطرة غير متحكم بها، وتعارض العمليات المتزامنة (SIMOPS). يعتبر مؤشر رئيسي لحادث جسيم.",
        "risk_en": "Potential fatality, fire/explosion, uncontrolled release of hazardous energy, SIMOPS conflict. Total loss of primary control barrier - leading indicator of major accident. Breach of ISO 45001 Cl 8.1.2.",
        "action": [
            "1. أمر إيقاف عمل فوري (STOP WORK) لجميع الأنشطة بدون PTW حتى التحقق من الالتزام",
            "2. إعادة تدريب إلزامي لجميع المشرفين على نظام PTW خلال 24 ساعة مع اختبار كفاءة",
            "3. إصدار تقرير عدم مطابقة رسمي (NCR) وخطاب تحذير تعاقدي مستوى 1",
            "4. تدقيق يومي لنظام PTW لمدة 14 يوم من قبل قسم EHS",
            "5. تطبيق مبدأ المساءلة - استبعاد المشرفين غير الملتزمين عند التكرار"
        ],
        "photo_label": "نشاط بدون تصريح عمل معروض / سجل التصاريح فارغ"
    },
    {
        "id": "03",
        "title_en": "Hard Barricades for Excavations",
        "title_ar": "عدم توفير الحواجز الصلبة لأعمال الحفر",
        "desc": [
            "• ترك مناطق الحفر بعمق أكثر من 1.2 متر بدون حواجز صلبة (Hard Barricades) - استخدام شريط تحذيري فقط وهو ممنوع للحفر العميق.",
            "• عدم توفير حماية الحافة، عدم وجود سلم آمن للدخول/الخروج، وتكديس ناتج الحفر على بعد أقل من 1 متر من الحافة.",
            "• عدم وجود علامات عاكسة أو إضاءة ليلية - خطر السقوط في الظلام.",
            "• مخالفة لمعايير OSHA 1926 Subpart P ومعيار Siemens SE-GEN-EHS-007 الخاص بأعمال الحفر."
        ],
        "risk": "CRITICAL",
        "risk_text": "خطر السقوط المميت داخل الحفرة، انهيار التربة والانطمار، سقوط الأشخاص والمركبات. نشاط عالي الخطورة مصنف كخطر حادث جسيم (MAH).",
        "risk_en": "High potential for fatal fall into excavation, cave-in engulfment, personnel/vehicle fall. Excavation is High-Risk activity with fatality history. Regulatory prohibition risk.",
        "action": [
            "1. توفير حواجز صلبة فورية - أنابيب سقالات / حواجز خرسانية حول جميع الحفر المفتوحة",
            "2. توفير سلالم آمنة كل 7.5 متر وإبعاد ناتج الحفر 1.5 متر على الأقل مع ألواح حماية",
            "3. إعادة التحقق من تصريح الحفر مع فحص جيوتقني وفحص الشخص المختص",
            "4. تركيب علامات عاكسة وإضاءة شمسية وامضة للساعات الليلية",
            "5. تقديم قائمة فحص يومية من الشخص المختص قبل بدء العمل - موقعة ومحققة من EHS"
        ],
        "photo_label": "حفرة بدون حاجز صلب - شريط فقط / تكديس ناتج حفر قريب من الحافة"
    },
    {
        "id": "04",
        "title_en": "Parking & Speed Limits",
        "title_ar": "مخالفات العلامات والسرعة داخل المحطة",
        "desc": [
            "• عدم الالتزام باللوحات الإرشادية - لوحات تحديد السرعة مفقودة أو تالفة، عدم وجود علامات أماكن الركن المخصصة.",
            "• تجاوز السرعة المحددة داخل المحطة (15 كم/ساعة) - تم رصد سرعة 30 كم/ساعة بجهاز قياس السرعة.",
            "• الركن في غير الأماكن المخصصة - إعاقة مسارات الطوارئ والمشاة.",
            "• عدم وجود رجل توجيه (Flagman) في التقاطعات الخطرة والنقاط العمياء - واجهة مركبات-مشاة غير مُدارة."
        ],
        "risk": "HIGH",
        "risk_text": "خطر تصادم مركبة-مشاة، تلف معدات، ووفاة محتملة. حوادث المركبات من أهم 3 أسباب وفيات في الإنشاءات. إعاقة وصول مركبات الطوارئ.",
        "risk_en": "High risk of vehicle-pedestrian collision, equipment damage, fatality. Vehicle incidents are top 3 fatality causes in construction. Emergency vehicle access may be impeded. Liability increases.",
        "action": [
            "1. إعادة تركيب جميع علامات المرور حسب خطة إدارة المرور المعتمدة - بمواصفات عاكسة",
            "2. إنشاء مسارات مشاة مخصصة بفصل صلب - حواجز Jersey + طلاء أخضر للممرات",
            "3. نشر رجال توجيه مدربين وتطبيق مراقبة السرعة - فحص بالرادار مرتين يومياً",
            "4. حملة توعية مرورية وإعادة تعريف سلامة لجميع السائقين",
            "5. فحص يومي مشترك بين EHS والخدمات اللوجستية لضوابط المرور - قائمة فحص موقعة"
        ],
        "photo_label": "لوحة سرعة مفقودة / ركن عشوائي / عدم وجود فصل مسار مشاة"
    },
    {
        "id": "05",
        "title_en": "Waste Management & Segregation",
        "title_ar": "سوء إدارة النفايات وفصلها",
        "desc": [
            "• تراكم النفايات الموقعية وعدم فرز الحاويات حسب النوع - نفايات خطرة، خشب، بلاستيك، أنقاض مختلطة في حاوية واحدة.",
            "• عدم وجود ملصقات واضحة أو ترميز لوني أو مرجع MSDS لحاويات النفايات الخطرة.",
            "• رمي خرق ملوثة بالزيوت وفلاتر في النفايات العامة - مخالفة بيئية وخطر حريق.",
            "• مخالفة لمعيار ISO 14001:2015 البند 8.1 وخطة إدارة النفايات بالمشروع ولوائح حماية البيئة المحلية."
        ],
        "risk": "HIGH",
        "risk_text": "خطر تلوث التربة والمياه الجوفية، خطر حريق من خلط نفايات غير متوافقة، غرامة تنظيمية، ضرر لسمعة Siemens Energy، وفشل مؤشر الاستدامة.",
        "risk_en": "Soil/groundwater pollution, fire hazard from incompatible waste mixing, regulatory fine, reputational damage, sustainability KPI failure. Cross-contamination prevents recycling.",
        "action": [
            "1. توفير حاويات نفايات مرمزة لونياً وملصقة حسب معيار SE: أسود-عامة، أخضر-قابلة للتدوير، أحمر-خطرة مع ملصقات HAZCHEM",
            "2. تنظيف وفرز صحيح للنفايات المختلطة الحالية بواسطة فريق مختص تحت إشراف EHS",
            "3. تدريب على إدارة النفايات - التركيز على التعامل مع النفايات الخطرة ونظام البيان",
            "4. تعيين مسؤول نفايات من المقاول - مسؤول عن النظافة اليومية",
            "5. تدقيق نفايات أسبوعي - تتبع الكميات وبيانات التخلص - تقرير شهري"
        ],
        "photo_label": "حاوية نفايات مختلطة - خطرة + عامة / خرق زيتية في النفايات العامة"
    },
    {
        "id": "06",
        "title_en": "Unsafe Workshop Equipment",
        "title_ar": "استخدام معدات وأدوات غير آمنة",
        "desc": [
            "• استخدام معدات ورشة تالفة أو غير مفحوصة - كوابل كهربائية مكشوفة، أجهزة بدون أغطية حماية (Guard) مثل جلخ/صاروخية.",
            "• عدم وجود سجل فحص وصيانة وقائية - قوائم الفحص اليومي مزورة أو غير معبأة - نفس التوقيع لمدة 7 أيام.",
            "• عدم توفير صواني منع التسريب (Drip Trays) - تلوث تربة تحت منطقة وقوف المعدات.",
            "• مخالفة لمعايير سلامة المعدات PUWER ومعيار معدات Siemens Energy ومتطلبات حماية البيئة."
        ],
        "risk": "CRITICAL",
        "risk_text": "خطر صعقة كهربائية، حريق/انفجار، انزلاق، تلوث تربة، فشل معدات يؤدي لفقدان السيطرة. انهيار نظام إدارة الصيانة.",
        "risk_en": "Electric shock, fire/explosion, slip hazard, soil pollution, equipment failure loss of control. Breakdown of maintenance management system. Breach of ISO 14001 and PUWER - major incident potential.",
        "action": [
            "1. إيقاف فوري للمعدات المتسربة/التالفة - وضع بطاقة 'ممنوع الاستخدام - حجز EHS' وسحب المفاتيح",
            "2. حفر التربة الملوثة والتخلص منها كنفايات خطرة - تنظيف المنطقة بمواد ماصة",
            "3. تقديم سجل صيانة كامل وشهادات فحص طرف ثالث خلال 48 ساعة - بدون شهادة، ممنوع التشغيل",
            "4. تطبيق قائمة فحص يومية إلزامية - محققة من EHS/الميكانيكا مع صورة تثبت عدم وجود تسريب",
            "5. توفير صواني منع تسريب ومجموعات مكافحة انسكاب في جميع مناطق وقوف المعدات - تغطية 100%"
        ],
        "photo_label": "كابل مكشوف / جهاز بدون غطاء حماية Guard / تسريب زيت فعال"
    },
    {
        "id": "07",
        "title_en": "Uncoordinated Work & Lack of Induction",
        "title_ar": "عدم التنسيق والعمل بدون توجيه قسم EHS",
        "desc": [
            "• إدخال عمالة للموقع دون اجتياز تعريف السلامة (EHS Induction) - تم العثور على 8-12 شخص يعملون بدون كارت تعريف.",
            "• بدء الأعمال دون التنسيق المسبق مع قسم EHS - عدم وجود سجل دخول، سجل التعريف غير مكتمل وغير مطابق لعدد الأفراد في البوابة.",
            "• محاضرات صندوق الأدوات (Toolbox Talks) غير منفذة أو بدون توثيق / كشف حضور / صلة بالمهمة.",
            "• إغلاق تقارير عدم مطابقة سابقة بدون دليل أو تحقق - تزوير إغلاق - تجاهل متعمد لنظام إدارة EHS بالمشروع."
        ],
        "risk": "CRITICAL",
        "risk_text": "عمال غير مدركين لمخاطر الموقع، نقاط التجمع للطوارئ، إجراءات النداء، وقواعد إنقاذ الحياة. احتمال عالي لأفعال غير آمنة وتأخر استجابة طوارئ وتوقف مشروع من السلطات.",
        "risk_en": "Workers unaware of site-specific hazards, emergency assembly points, muster procedures, life-saving rules. High probability of unsafe acts, delayed emergency response, regulatory non-compliance leading to project stoppage. Breach of ISO 45001 Cl 7.2/7.3.",
        "action": [
            "1. إخراج فوري لجميع الأفراد غير الحاصلين على تعريف السلامة من الموقع - بمرافقة الأمن",
            "2. تطبيق فحص 100% في البوابة - تحقق مشترك أمن + EHS - بدون كارت، ممنوع الدخول - من المناوبة القادمة",
            "3. تقديم قائمة كاملة بالقوى العاملة مقابل سجلات التعريف خلال 12 ساعة مع خطة سد الفجوة",
            "4. إدخال نظام تتبع تعريف ببصمة/QR مرتبط بالتحكم في الدخول + سجل توجيهات EHS يومي موقع من مدير المقاول",
            "5. تحذير رسمي وتحميل تكلفة إعادة جلسات التعريف + عقد اجتماع قيادة EHS طارئ مع جميع المراقبين - التزام موثق"
        ],
        "photo_label": "عمال بدون كارت تعريف / سجل بوابة غير مطابق / سجل TBT فارغ"
    },
]

def create_slide(data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)

    # Top thin line
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    # Header
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.75), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    # ID badge
    shape(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), fill=TEAL_DARK)
    text_box(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), data['id'], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    # Title AR + EN
    text_box(slide, Inches(1.0), Inches(0.1), Inches(8), Inches(0.3), f"{data['title_ar']} - {data['title_en']}", size=16, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(1.0), Inches(0.4), Inches(8), Inches(0.3), f"Observation # {data['id']} | تقييم الالتزام والمخاطر التشغيلية", size=9, bold=False, color=GREY_TEXT)

    # Risk badge
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), fill=risk_bg, border=risk_color, border_w=Pt(1.5))
    text_box(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), f"RISK: {data['risk']} | مستوى الخطورة: {data['risk']}", size=11, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    # LEFT SIDE - Content (7.5 inch width)
    # Box 1: Description
    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.3), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.3), "1. الوصف الفني للمخالفة / Technical Description", size=10, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.4), Inches(7.2), Inches(2.0), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    tf = slide.shapes.add_textbox(Inches(0.4), Inches(1.45), Inches(7.0), Inches(1.9)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc']):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(9.5)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(3)

    # Box 2: Risk
    shape(slide, Inches(0.3), Inches(3.6), Inches(7.2), Inches(0.3), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.6), Inches(7.0), Inches(0.3), f"2. تقييم المخاطر المترتبة / Risk Assessment - {data['risk']}", size=10, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.9), Inches(7.2), Inches(1.0), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(3.95), Inches(7.0), Inches(0.45), data['risk_text'], size=9, bold=False, color=DARK_TEXT)
    text_box(slide, Inches(0.4), Inches(4.4), Inches(7.0), Inches(0.45), data['risk_en'], size=8, bold=False, color=GREY_TEXT)

    # Box 3: Corrective Action
    shape(slide, Inches(0.3), Inches(5.1), Inches(7.2), Inches(0.3), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(5.1), Inches(7.0), Inches(0.3), "3. الإجراء التصحيحي المطلوب / Required Corrective Action (Contractor)", size=10, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.4), Inches(7.2), Inches(1.7), fill=WHITE, border=GREY_BORDER)
    tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(5.45), Inches(7.0), Inches(1.6)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(8.8)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    # RIGHT SIDE - Photo Placeholder (5.3 inch width)
    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.3), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.3), "📷 مساحة الصور الموقعية / Photo Evidence Area", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    
    # Main photo box
    shape(slide, Inches(7.8), Inches(1.4), Inches(5.2), Inches(4.2), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1.5))
    # Inner dashed effect with text
    text_box(slide, Inches(7.9), Inches(1.5), Inches(5.0), Inches(0.4), f"[أدخل صورة المخالفة هنا] / [Insert Violation Photo Here]", size=12, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.0), Inches(5.0), Inches(0.6), f"{data['photo_label']}\nFig {data['id']}.1 - {data['title_en']}", size=9, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.8), Inches(5.0), Inches(1.0), "⬆️ اسحب الصورة وألصقها هنا\nDrag & Drop Your Site Photo Here\n\nسيتم احتواء الصورة تلقائياً داخل هذه الخانة\nPhoto will auto-fit inside this box\n\nأضف: التاريخ، الوقت، الموقع، إحداثيات GPS\nAdd: Date, Time, Location, GPS", size=10, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    
    # Photo details box below
    shape(slide, Inches(7.8), Inches(5.7), Inches(5.2), Inches(0.6), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(5.7), Inches(5.0), Inches(0.6), f"Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________  Taken by: ___________\nتفاصيل الصورة: التاريخ، الوقت، الموقع، المصور\nStandard Ref: ISO 45001, OSHA, SE EHS Requirements", size=7.5, color=GREY_TEXT)

    # Second photo small box
    shape(slide, Inches(7.8), Inches(6.4), Inches(2.5), Inches(0.7), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(6.4), Inches(2.3), Inches(0.7), "[صورة إضافية 2]\n[Additional Photo 2]\nFig X.2 - Different Angle", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    
    shape(slide, Inches(10.5), Inches(6.4), Inches(2.5), Inches(0.7), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(10.6), Inches(6.4), Inches(2.3), Inches(0.7), "[صورة بعد التصحيح]\n[After Correction Photo]\nFig X.3 - Compliant State", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)

    # Footer
    shape(slide, Inches(0), Inches(7.2), Inches(13.33), Pt(1), fill=GREY_BORDER)
    text_box(slide, Inches(0.3), Inches(7.25), Inches(6), Inches(0.2), "Siemens Energy | KAZ Project | EHS Department | Confidential - Management Review", size=7, color=GREY_TEXT)
    text_box(slide, Inches(10.5), Inches(7.25), Inches(2.5), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

# TITLE SLIDE
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), fill=TEAL)
# Logo area
shape(slide, Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8), fill=WHITE, border=GREY_BORDER)
text_box(slide, Inches(0.6), Inches(0.35), Inches(3), Inches(0.3), "SIEMENS ENERGY", size=14, bold=True, color=TEAL_DARK)
text_box(slide, Inches(0.6), Inches(0.65), Inches(3), Inches(0.3), "KAZ Power Plant Upgrade Project", size=9, color=GREY_TEXT)
text_box(slide, Inches(10), Inches(0.35), Inches(2.8), Inches(0.6), "CONFIDENTIAL\nسري - للإدارة فقط", size=10, bold=True, color=RED, align=PP_ALIGN.RIGHT)

# Main title
shape(slide, Inches(0.5), Inches(1.5), Inches(0.08), Inches(1.2), fill=RED)
text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.7), "تقرير مخالفات السلامة للعمليات الميدانية للمقاول الفرعي", size=26, bold=True, color=DARK_TEXT)
text_box(slide, Inches(0.8), Inches(2.2), Inches(11.5), Inches(0.5), "Subcontractor EHS Non-Compliance Report", size=20, bold=True, color=TEAL_DARK)

text_box(slide, Inches(0.8), Inches(2.9), Inches(11.5), Inches(0.4), "تقييم الالتزام والمخاطر التشغيلية بالموقع | Operational Compliance & Risk Assessment", size=13, bold=False, color=GREY_TEXT)

# Info boxes
shape(slide, Inches(0.5), Inches(3.5), Inches(5.5), Inches(2.2), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
text_box(slide, Inches(0.6), Inches(3.5), Inches(5.3), Inches(0.3), "معلومات التقرير / Report Information", size=11, bold=True, color=TEAL_DARK)
info = """المرجع: SE-EHS-NC-2026-EXEC-AR
Reference: SE-EHS-NC-2026-EXEC-AR

التاريخ: سبتمبر 2026 | Date: September 2026
إعداد: مدير السلامة - Siemens Energy
Prepared by: EHS Manager - Siemens Energy

المعايير: ISO 45001:2018, ISO 14001:2015, OSHA 29 CFR 1926
Standards: Siemens Energy EHS Principles & Life Saving Rules

التصنيف: سري - للمراجعة الإدارية واتخاذ القرار
Classification: Confidential - Management Review & Decision

التوزيع: مدير المشروع، مدير الإنشاءات، المدير التجاري، إدارة المقاول الفرعي"""
text_box(slide, Inches(0.6), Inches(3.8), Inches(5.3), Inches(1.8), info, size=9, color=DARK_TEXT)

# Design guide
shape(slide, Inches(6.5), Inches(3.5), Inches(6.3), Inches(2.2), fill=WHITE, border=TEAL, border_w=Pt(1.5))
text_box(slide, Inches(6.6), Inches(3.5), Inches(6.1), Inches(0.3), "تعليمات التصميم والاستخدام / Design & Usage Instructions", size=11, bold=True, color=TEAL_DARK)
guide = """التصميم: تنفيذي وعصري - خلفية بيضاء نقية #FFFFFF مع إبراز حرج أحمر/أصفر
Design: Executive & Modern - Pure White #FFFFFF background with Red/Yellow warning highlights

هيكلة كل شريحة / Each Slide Structure:
1. عنوان المخالفة (عربي + إنجليزي) - Title (Arabic + English)
2. الوصف الفني للمخالفة - Technical Description (4-5 points)
3. تقييم المخاطر المترتبة - Risk Assessment (Arabic + English)
4. الإجراء التصحيحي المطلوب - Corrective Action (5 points)
5. مساحة صور مخصصة واضحة - Dedicated Photo Area [أدخل صورة المخالفة هنا]

كيفية إضافة الصور / How to Add Photos:
• اسحب الصورة من مجلد الصور وألصقها مباشرة في الخانة اليمنى الكبيرة
• Drag & drop your site photo into the large right-side box
• الخانة ستحتوي الصورة تلقائياً - Box will auto-fit photo
• أضف صورتين: عامة + تفصيلية + صورة بعد التصحيح
• Add 2 photos: overview + close-up + after-correction photo

الألوان / Colors:
• خلفية: أبيض نقي #FFFFFF / رمادي فاتح #F5F5F7
• نص: أسود داكن #212121 / رمادي #616161
• تحذير: أحمر #E30613 (حرج) / أصفر #FFB900 (عالي)
• إبراز: تيفاني #009999 / أزرق غامق #0E2F3E"""
text_box(slide, Inches(6.6), Inches(3.8), Inches(6.1), Inches(1.8), guide, size=8, color=DARK_TEXT)

# Executive statement
shape(slide, Inches(0.5), Inches(6.0), Inches(12.33), Inches(0.9), fill=RED_LIGHT, border=RED, border_w=Pt(1.5))
text_box(slide, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.2), "ملاحظة تنفيذية / Executive Statement:", size=10, bold=True, color=RED)
text_box(slide, Inches(0.6), Inches(6.25), Inches(12.1), Inches(0.6), "هذا التقرير يوثق مخالفات EHS حرجة وجوهرية لوحظت أثناء عمليات التفتيش الميدانية. النتائج تشير إلى انهيار أساسي في نظام إدارة السلامة للمقاول الفرعي وتهديد مباشر لالتزامنا بـ Zero Harm وجدول المشروع والامتثال القانوني. مطلوب تدخل إداري فوري.\nThis report documents critical, systemic EHS non-compliances observed during site inspections. Findings indicate fundamental breakdown of subcontractor EHS management system and direct threat to Zero Harm commitment, project schedule, and legal compliance. Immediate management intervention required.", size=8.5, color=DARK_TEXT)

# CLOSING SLIDE
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_bg(slide, WHITE)
shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.7), fill=WHITE, border=GREY_BORDER)
text_box(slide, Inches(0.5), Inches(0.1), Inches(12), Inches(0.6), "التوصيات الخصامية وخطة التصحيح / Final Recommendations & Corrective Action Plan", size=18, bold=True, color=DARK_TEXT)

# 3 columns
# Immediate
shape(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), fill=RED)
text_box(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), "🔴 فوري (0-24 ساعة) / IMMEDIATE (0-24h)", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(0.3), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
text_box(slide, Inches(0.4), Inches(1.4), Inches(3.9), Inches(3.8), "• أمر إيقاف عمل فوري لجميع المخالفات الحرجة (PTW، الحفر، الديزل، المعدات المتسربة)\nSTOP WORK for all critical violations\n\n• تأمين المناطق الخطرة - حواجز صلبة + مراقبة حريق + منطقة استبعاد 15م\nSecure hazardous areas - hard barricade + fire watch\n\n• إخراج جميع الأفراد بدون تعريف سلامة - فحص 100% في البوابة\nRemove non-inducted personnel - 100% gate check\n\n• إيقاف المعدات المتسربة - بطاقة ممنوع الاستخدام\nGround leaking equipment - Tag Out\n\n• إصدار تقرير عدم مطابقة رسمي مستوى 1 و 2 + خطاب تحذير تعاقدي\nIssue NCR Level 1 & 2 + warning letter\n\n• عقد اجتماع EHS طارئ - حضور مدير المقاول خلال 24 ساعة\nEmergency meeting - Subcontractor MD within 24h\n\n[مكان صورة: منطقة مؤمنة بعد إيقاف العمل]\n[Photo: Secured area after stop work]", size=8.5, color=DARK_TEXT)

# Short
shape(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), fill=AMBER)
text_box(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), "🟡 قصير المدى (1-7 أيام) / SHORT-TERM (1-7 Days)", size=11, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
shape(slide, Inches(4.6), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
text_box(slide, Inches(4.7), Inches(1.4), Inches(3.9), Inches(3.8), "• تدقيق كامل لنظام PTW وإعادة تدريب - جميع المشرفين معتمدين - 80% نجاح\nFull PTW Audit + Re-Training - 80% pass mark\n\n• خزن ديزل مطابق - حوض 110% + أرضية غير منفذة + طفايات حريق\nCompliant diesel storage - 110% bund + extinguishers\n\n• إعادة علامات المرور - عاكسة + فصل مسار مشاة + رجال توجيه + رادار سرعة\nTraffic signs - retro-reflective + pedestrian segregation + flagmen\n\n• نظام نفايات - حاويات مرمزة لونياً + ملصقات + مسؤول نفايات\nWaste - color-coded skips + labeling + focal point\n\n• التزام حفر - حواجز صلبة + سلالم كل 7.5م + فحص شخص مختص يومي\nExcavation - hard barricades + ladders + competent person checks\n\n• حملة PPE + برنامج BBS - فحص يومي + لوحة مؤشرات\nPPE Campaign + BBS Launch\n\n[مكان صورة: حالة مطابقة بعد التصحيح]\n[Photo: Compliant state after correction]", size=8.5, color=DARK_TEXT)

# Strategic
shape(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), fill=TEAL)
text_box(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), "🟢 استراتيجي (7-30 يوم) / STRATEGIC (7-30 Days)", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
shape(slide, Inches(8.9), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
text_box(slide, Inches(9.0), Inches(1.4), Inches(3.9), Inches(3.8), "• تطبيق تعاقدي: مصفوفة غرامات مرتبطة بمؤشرات EHS (TRIR، إغلاق NCR، التزام PTW)\nContractual Enforcement: Penalty matrix linked to KPIs\n\n• مساءلة قيادية: استبدال مشرفين غير ملتزمين - فحص سيرة + مقابلة كفاءة\nLeadership Accountability: Replace non-performing supervisors\n\n• ترقية نظام: تتبع تعريف ببصمة/QR + نظام PTW رقمي + تطبيق قائمة فحص معدات\nSystem Upgrade: QR/Biometric + digital PTW + checklist app\n\n• مؤشرات ولوحة: لوحة متابعة أسبوعية أحمر/أصفر/أخضر للإدارة العليا\nKPI & Scorecard: Weekly dashboard Red/Amber/Green\n\n• تدقيق مستقل: تدقيق EHS طرف ثالث لنظام إدارة المقاول - تحليل فجوة ISO\nIndependent Audit: 3rd party EHS audit\n\n• التزام Zero Harm: إعادة توقيع قواعد إنقاذ الحياة - اجتماع عام مع الإدارة العليا\nZero Harm Commitment: Re-sign Life Saving Rules - Townhall\n\n[مكان صورة: لوحة مؤشرات / تدريب]\n[Photo: Dashboard / Training]", size=8.5, color=DARK_TEXT)

# Decision matrix
shape(slide, Inches(0.3), Inches(5.5), Inches(12.7), Inches(1.7), fill=LIGHT_GREY_CARD, border=TEAL_DARK, border_w=Pt(2))
text_box(slide, Inches(0.4), Inches(5.5), Inches(12.5), Inches(0.3), "✅ مصفوفة قرارات الإدارة المطلوبة اليوم / Management Decision Matrix - Required Today", size=12, bold=True, color=TEAL_DARK)
text_box(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(1.2), "1. الموافقة على سلطة إيقاف العمل الفوري لقسم EHS للمخالفات الحرجة [ ] نعم [ ] لا  |  Approve IMMEDIATE Stop-Work Authority to EHS for Critical Violations [ ] Yes [ ] No\n\n2. تفويض التحذير التعاقدي الرسمي وتطبيق الغرامات [ ] نعم [ ] لا  |  Authorize Formal Contractual Warning & Penalty Enforcement [ ] Yes [ ] No\n\n3. إلزام حضور الإدارة العليا للمقاول في الموقع خلال 48 ساعة [ ] نعم [ ] لا  |  Mandate Subcontractor Top Management Presence On-Site within 48h [ ] Yes [ ] No\n\n4. الموافقة على ميزانية ضوابط EHS الإضافية (خزن محمي، علامات، حواجز) - تحميل على المقاول [ ] نعم [ ] لا  |  Approve Budget for Additional EHS Controls - Back-charge to Subcontractor [ ] Yes [ ] No\n\n5. اعتماد سياسة عدم التسامح: 3 مخالفات = استبعاد - إعلان في اجتماع عام [ ] نعم [ ] لا  |  Endorse Zero Tolerance Policy: 3 Strikes = Removal - Communicate in Townhall [ ] Yes [ ] No\n\nالتوقيع: _________________  مدير المشروع  التاريخ: _______  |  التوقيع: _________________  مدير الإنشاءات  التاريخ: _______  |  التوقيع: _________________  مدير السلامة  التاريخ: _______", size=8.5, color=DARK_TEXT)

# Create slides for each violation
for data in slides_data:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.75), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), fill=TEAL_DARK)
    text_box(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), data['id'], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(1.0), Inches(0.1), Inches(8), Inches(0.3), f"{data['title_ar']} - {data['title_en']}", size=15, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(1.0), Inches(0.4), Inches(8), Inches(0.3), f"Observation # {data['id']} | تقييم الالتزام والمخاطر", size=9, bold=False, color=GREY_TEXT)
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), fill=risk_bg, border=risk_color, border_w=Pt(1.5))
    text_box(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), f"RISK: {data['risk']} | {data['risk']}", size=11, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)

    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.3), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.3), "1. الوصف الفني للمخالفة / Technical Description", size=10, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.4), Inches(7.2), Inches(1.9), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    tf = slide.shapes.add_textbox(Inches(0.4), Inches(1.45), Inches(7.0), Inches(1.8)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc']):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(9)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2.5)

    shape(slide, Inches(0.3), Inches(3.5), Inches(7.2), Inches(0.3), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.5), Inches(7.0), Inches(0.3), f"2. تقييم المخاطر / Risk Assessment - {data['risk']}", size=10, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.8), Inches(7.2), Inches(0.9), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(3.85), Inches(7.0), Inches(0.4), data['risk_text'], size=8.5, bold=False, color=DARK_TEXT)
    text_box(slide, Inches(0.4), Inches(4.25), Inches(7.0), Inches(0.4), data['risk_en'], size=7.5, bold=False, color=GREY_TEXT)

    shape(slide, Inches(0.3), Inches(4.9), Inches(7.2), Inches(0.3), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(4.9), Inches(7.0), Inches(0.3), "3. الإجراء التصحيحي / Corrective Action", size=10, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.2), Inches(7.2), Inches(1.9), fill=WHITE, border=GREY_BORDER)
    tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(5.25), Inches(7.0), Inches(1.8)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(8.5)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)

    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.3), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.3), "📷 مساحة الصور / Photo Evidence Area", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(1.4), Inches(5.2), Inches(4.0), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1.5))
    text_box(slide, Inches(7.9), Inches(1.5), Inches(5.0), Inches(0.4), "[أدخل صورة المخالفة هنا]\n[Insert Violation Photo Here]", size=12, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.1), Inches(5.0), Inches(0.5), f"{data['photo_label']}\nFig {data['id']}.1", size=9, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.8), Inches(5.0), Inches(1.0), "⬆️ اسحب الصورة وألصقها هنا\nDrag & Drop Your Site Photo Here\n\nAdd: Date, Time, Location, GPS\nأضف: التاريخ، الوقت، الموقع", size=10, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(5.5), Inches(5.2), Inches(0.6), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(5.5), Inches(5.0), Inches(0.6), f"Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________  Taken by: ___________\nتفاصيل الصورة: التاريخ، الوقت، الموقع، المصور", size=7.5, color=GREY_TEXT)
    shape(slide, Inches(7.8), Inches(6.2), Inches(2.5), Inches(0.9), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(6.2), Inches(2.3), Inches(0.9), "[صورة إضافية 2]\n[Additional Photo 2]\nFig X.2\nDifferent Angle", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(10.5), Inches(6.2), Inches(2.5), Inches(0.9), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL)
    text_box(slide, Inches(10.6), Inches(6.2), Inches(2.3), Inches(0.9), "[صورة بعد التصحيح]\n[After Correction]\nFig X.3\nCompliant State", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0), Inches(7.2), Inches(13.33), Pt(1), fill=GREY_BORDER)
    text_box(slide, Inches(0.3), Inches(7.25), Inches(6), Inches(0.2), "Siemens Energy | KAZ Project | EHS | Confidential", size=7, color=GREY_TEXT)
    text_box(slide, Inches(10.5), Inches(7.25), Inches(2.5), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

# Reorder slides: Title first, then 7 violations, then closing - need to recreate order correctly
# Actually we created closing before violations loop, need to fix order by creating new presentation in correct order
# For simplicity, save as is and we will reorder by recreating - but for now we have title, closing, then violations
# Let's recreate properly ordered version

prs2 = Presentation()
prs2.slide_width = Inches(13.33)
prs2.slide_height = Inches(7.5)

# Copy slides in correct order: 0=title, 1=closing is wrong, so we need to generate again correctly ordered
# Instead generate final file with correct order using same functions

def create_title_slide(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Inches(0.08), fill=TEAL)
    shape(slide, Inches(0.5), Inches(0.3), Inches(12.33), Inches(0.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(0.35), Inches(3), Inches(0.3), "SIEMENS ENERGY", size=14, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.6), Inches(0.65), Inches(3), Inches(0.3), "KAZ Power Plant Upgrade Project", size=9, color=GREY_TEXT)
    text_box(slide, Inches(10), Inches(0.35), Inches(2.8), Inches(0.6), "CONFIDENTIAL\nسري - للإدارة فقط", size=10, bold=True, color=RED, align=PP_ALIGN.RIGHT)
    shape(slide, Inches(0.5), Inches(1.5), Inches(0.08), Inches(1.2), fill=RED)
    text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.7), "تقرير مخالفات السلامة للعمليات الميدانية للمقاول الفرعي", size=26, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(0.8), Inches(2.2), Inches(11.5), Inches(0.5), "Subcontractor EHS Non-Compliance Report", size=20, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.8), Inches(2.9), Inches(11.5), Inches(0.4), "تقييم الالتزام والمخاطر التشغيلية بالموقع | Operational Compliance & Risk Assessment", size=13, bold=False, color=GREY_TEXT)
    shape(slide, Inches(0.5), Inches(3.5), Inches(5.5), Inches(2.2), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    text_box(slide, Inches(0.6), Inches(3.5), Inches(5.3), Inches(0.3), "معلومات التقرير / Report Information", size=11, bold=True, color=TEAL_DARK)
    info = """المرجع: SE-EHS-NC-2026-EXEC-AR
Reference: SE-EHS-NC-2026-EXEC-AR

التاريخ: سبتمبر 2026 | Date: September 2026
إعداد: مدير السلامة - Siemens Energy
Prepared by: EHS Manager - Siemens Energy

المعايير: ISO 45001:2018, ISO 14001:2015, OSHA 29 CFR 1926
Standards: Siemens Energy EHS Principles & Life Saving Rules

التصنيف: سري - للمراجعة الإدارية واتخاذ القرار
Classification: Confidential - Management Review & Decision"""
    text_box(slide, Inches(0.6), Inches(3.8), Inches(5.3), Inches(1.8), info, size=9, color=DARK_TEXT)
    shape(slide, Inches(6.5), Inches(3.5), Inches(6.3), Inches(2.2), fill=WHITE, border=TEAL, border_w=Pt(1.5))
    text_box(slide, Inches(6.6), Inches(3.5), Inches(6.1), Inches(0.3), "تعليمات التصميم / Design Instructions", size=11, bold=True, color=TEAL_DARK)
    guide = """التصميم: تنفيذي وعصري - خلفية بيضاء نقية #FFFFFF مع إبراز حرج أحمر/أصفر
Design: Executive & Modern - Pure White #FFFFFF with Red/Yellow warning

هيكلة كل شريحة / Each Slide:
1. عنوان المخالفة (عربي + إنجليزي)
2. الوصف الفني - Technical Description
3. تقييم المخاطر - Risk Assessment
4. الإجراء التصحيحي - Corrective Action
5. مساحة صور مخصصة - Photo Area [أدخل صورة المخالفة هنا]

كيفية إضافة الصور / How to Add Photos:
• اسحب الصورة وألصقها في الخانة اليمنى الكبيرة
• Drag & drop site photo into large right box
• أضف صورتين: عامة + تفصيلية + بعد التصحيح"""
    text_box(slide, Inches(6.6), Inches(3.8), Inches(6.1), Inches(1.8), guide, size=8.5, color=DARK_TEXT)
    shape(slide, Inches(0.5), Inches(6.0), Inches(12.33), Inches(0.9), fill=RED_LIGHT, border=RED, border_w=Pt(1.5))
    text_box(slide, Inches(0.6), Inches(6.0), Inches(12.1), Inches(0.2), "ملاحظة تنفيذية / Executive Statement:", size=10, bold=True, color=RED)
    text_box(slide, Inches(0.6), Inches(6.25), Inches(12.1), Inches(0.6), "هذا التقرير يوثق مخالفات EHS حرجة وجوهرية. النتائج تشير إلى انهيار أساسي في نظام إدارة السلامة للمقاول الفرعي وتهديد مباشر لالتزام Zero Harm. مطلوب تدخل إداري فوري.\nThis report documents critical, systemic EHS non-compliances. Findings indicate fundamental breakdown of subcontractor EHS management system and direct threat to Zero Harm commitment. Immediate intervention required.", size=8.5, color=DARK_TEXT)

def create_violation_slide(prs_obj, data):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.75), fill=WHITE, border=GREY_BORDER, border_w=Pt(1))
    shape(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), fill=TEAL_DARK)
    text_box(slide, Inches(0.3), Inches(0.15), Inches(0.6), Inches(0.5), data['id'], size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(1.0), Inches(0.1), Inches(8), Inches(0.3), f"{data['title_ar']} - {data['title_en']}", size=15, bold=True, color=DARK_TEXT)
    text_box(slide, Inches(1.0), Inches(0.4), Inches(8), Inches(0.3), f"Observation # {data['id']} | تقييم الالتزام", size=9, bold=False, color=GREY_TEXT)
    risk_color = RED if data['risk'] == "CRITICAL" else AMBER
    risk_bg = RED_LIGHT if data['risk'] == "CRITICAL" else AMBER_LIGHT
    risk_text_color = RED if data['risk'] == "CRITICAL" else RGBColor(0x7A, 0x5A, 0x00)
    shape(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), fill=risk_bg, border=risk_color, border_w=Pt(1.5))
    text_box(slide, Inches(10.0), Inches(0.15), Inches(3.0), Inches(0.5), f"RISK: {data['risk']}", size=11, bold=True, color=risk_text_color, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0.3), Inches(1.1), Inches(7.2), Inches(0.3), fill=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(1.1), Inches(7.0), Inches(0.3), "1. الوصف الفني / Technical Description", size=10, bold=True, color=WHITE)
    shape(slide, Inches(0.3), Inches(1.4), Inches(7.2), Inches(1.9), fill=LIGHT_GREY_CARD, border=GREY_BORDER)
    tf = slide.shapes.add_textbox(Inches(0.4), Inches(1.45), Inches(7.0), Inches(1.8)).text_frame
    tf.word_wrap = True
    for i, line in enumerate(data['desc']):
        if i==0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(9)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2.5)
    shape(slide, Inches(0.3), Inches(3.5), Inches(7.2), Inches(0.3), fill=risk_bg, border=risk_color, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(3.5), Inches(7.0), Inches(0.3), f"2. تقييم المخاطر / Risk Assessment - {data['risk']}", size=10, bold=True, color=risk_text_color)
    shape(slide, Inches(0.3), Inches(3.8), Inches(7.2), Inches(0.9), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(3.85), Inches(7.0), Inches(0.4), data['risk_text'], size=8.5, bold=False, color=DARK_TEXT)
    text_box(slide, Inches(0.4), Inches(4.25), Inches(7.0), Inches(0.4), data['risk_en'], size=7.5, bold=False, color=GREY_TEXT)
    shape(slide, Inches(0.3), Inches(4.9), Inches(7.2), Inches(0.3), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL, border_w=Pt(1))
    text_box(slide, Inches(0.4), Inches(4.9), Inches(7.0), Inches(0.3), "3. الإجراء التصحيحي / Corrective Action", size=10, bold=True, color=TEAL_DARK)
    shape(slide, Inches(0.3), Inches(5.2), Inches(7.2), Inches(1.9), fill=WHITE, border=GREY_BORDER)
    tf2 = slide.shapes.add_textbox(Inches(0.4), Inches(5.25), Inches(7.0), Inches(1.8)).text_frame
    tf2.word_wrap = True
    for i, line in enumerate(data['action']):
        if i==0:
            p = tf2.paragraphs[0]
        else:
            p = tf2.add_paragraph()
        p.text = line
        p.font.size = Pt(8.5)
        p.font.color.rgb = DARK_TEXT
        p.font.name = "Calibri"
        p.space_after = Pt(2)
    shape(slide, Inches(7.8), Inches(1.1), Inches(5.2), Inches(0.3), fill=DARK_TEXT)
    text_box(slide, Inches(7.9), Inches(1.1), Inches(5.0), Inches(0.3), "📷 مساحة الصور / Photo Evidence", size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(1.4), Inches(5.2), Inches(4.0), fill=LIGHT_GREY_BG, border=GREY_BORDER, border_w=Pt(1.5))
    text_box(slide, Inches(7.9), Inches(1.5), Inches(5.0), Inches(0.4), "[أدخل صورة المخالفة هنا]\n[Insert Violation Photo Here]", size=12, bold=True, color=RED, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.1), Inches(5.0), Inches(0.5), f"{data['photo_label']}\nFig {data['id']}.1", size=9, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    text_box(slide, Inches(7.9), Inches(2.8), Inches(5.0), Inches(1.0), "⬆️ اسحب الصورة وألصقها هنا\nDrag & Drop Your Site Photo Here\n\nAdd: Date, Time, Location, GPS", size=10, bold=False, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(7.8), Inches(5.5), Inches(5.2), Inches(0.6), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(5.5), Inches(5.0), Inches(0.6), f"Photo Details: Date: ___/___/___  Time: ___:___  Location: ___________\nتفاصيل الصورة: التاريخ، الوقت، الموقع، المصور", size=7.5, color=GREY_TEXT)
    shape(slide, Inches(7.8), Inches(6.2), Inches(2.5), Inches(0.9), fill=LIGHT_GREY_BG, border=GREY_BORDER)
    text_box(slide, Inches(7.9), Inches(6.2), Inches(2.3), Inches(0.9), "[صورة إضافية 2]\n[Additional Photo 2]\nFig X.2", size=7, color=GREY_TEXT, align=PP_ALIGN.CENTER)
    shape(slide, Inches(10.5), Inches(6.2), Inches(2.5), Inches(0.9), fill=RGBColor(0xE8,0xF5,0xE9), border=TEAL)
    text_box(slide, Inches(10.6), Inches(6.2), Inches(2.3), Inches(0.9), "[صورة بعد التصحيح]\n[After Correction]\nFig X.3", size=7, color=TEAL_DARK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0), Inches(7.2), Inches(13.33), Pt(1), fill=GREY_BORDER)
    text_box(slide, Inches(0.3), Inches(7.25), Inches(6), Inches(0.2), "Siemens Energy | KAZ Project | EHS | Confidential", size=7, color=GREY_TEXT)
    text_box(slide, Inches(10.5), Inches(7.25), Inches(2.5), Inches(0.2), f"Slide {data['id']} | SE-EHS-NC-2026", size=7, color=GREY_TEXT, align=PP_ALIGN.RIGHT)

def create_closing_slide(prs_obj):
    slide = prs_obj.slides.add_slide(prs_obj.slide_layouts[6])
    set_bg(slide, WHITE)
    shape(slide, Inches(0), Inches(0), Inches(13.33), Pt(4), fill=TEAL)
    shape(slide, Inches(0), Inches(0.05), Inches(13.33), Inches(0.7), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.5), Inches(0.1), Inches(12), Inches(0.6), "التوصيات الخصامية وخطة التصحيح / Final Recommendations & Corrective Action Plan", size=18, bold=True, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), fill=RED)
    text_box(slide, Inches(0.3), Inches(1.0), Inches(4.1), Inches(0.4), "🔴 فوري (0-24h) / IMMEDIATE", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(0.3), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(0.4), Inches(1.4), Inches(3.9), Inches(3.8), "• STOP WORK all critical violations\n• Secure areas - hard barricade + fire watch\n• Remove non-inducted personnel - 100% gate\n• Ground leaking equipment - Tag Out\n• Issue NCR Level 1 & 2 + warning letter\n• Emergency meeting - MD within 24h\n\n[Photo: Secured area after stop work]\n[مكان صورة: منطقة مؤمنة]\n\nDECISION: Approve Stop-Work Authority", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), fill=AMBER)
    text_box(slide, Inches(4.6), Inches(1.0), Inches(4.1), Inches(0.4), "🟡 قصير (1-7 أيام) / SHORT-TERM", size=11, bold=True, color=BLACK, align=PP_ALIGN.CENTER)
    shape(slide, Inches(4.6), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(4.7), Inches(1.4), Inches(3.9), Inches(3.8), "• PTW Audit + Re-Training - 80% pass\n• Diesel storage - 110% bund + extinguishers\n• Traffic signs - retro-reflective + flagmen\n• Waste - color-coded skips + focal point\n• Excavation - hard barricades + ladders\n• PPE Campaign + BBS Launch\n• EHS Directives Log - daily sign-off\n\n[Photo: Compliant state after correction]\n[مكان صورة: حالة مطابقة]\n\nDECISION: Approve $15k cost - back-charge", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), fill=TEAL)
    text_box(slide, Inches(8.9), Inches(1.0), Inches(4.1), Inches(0.4), "🟢 استراتيجي (7-30 يوم) / STRATEGIC", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    shape(slide, Inches(8.9), Inches(1.4), Inches(4.1), Inches(3.8), fill=WHITE, border=GREY_BORDER)
    text_box(slide, Inches(9.0), Inches(1.4), Inches(3.9), Inches(3.8), "• Penalty matrix linked to KPIs (TRIR, NCR)\n• Replace non-performing supervisors\n• QR/Biometric induction + digital PTW app\n• Weekly dashboard Red/Amber/Green\n• 3rd party EHS audit ISO 45001/14001\n• Re-sign Life Saving Rules - Townhall\n• Demobilization / replacement contingency\n\n[Photo: Dashboard / Training]\n[مكان صورة: لوحة مؤشرات / تدريب]\n\nDECISION: Authorize partial termination if recur", size=8.5, color=DARK_TEXT)
    shape(slide, Inches(0.3), Inches(5.5), Inches(12.7), Inches(1.7), fill=LIGHT_GREY_CARD, border=TEAL_DARK, border_w=Pt(2))
    text_box(slide, Inches(0.4), Inches(5.5), Inches(12.5), Inches(0.3), "✅ مصفوفة القرارات / Management Decision Matrix - Required Today", size=12, bold=True, color=TEAL_DARK)
    text_box(slide, Inches(0.4), Inches(5.85), Inches(12.5), Inches(1.2), "1. Approve Stop-Work Authority [ ] Yes [ ] No | الموافقة على إيقاف العمل الفوري [ ] نعم [ ] لا\n2. Authorize Warning & Penalty [ ] Yes [ ] No | تفويض التحذير والغرامات [ ] نعم [ ] لا\n3. Mandate Top Mgmt Presence in 48h [ ] Yes [ ] No | إلزام حضور الإدارة العليا خلال 48س [ ] نعم [ ] لا\n4. Approve Budget - Back-charge [ ] Yes [ ] No | الموافقة على الميزانية - تحميل على المقاول [ ] نعم [ ] لا\n5. Endorse Zero Tolerance 3 Strikes [ ] Yes [ ] No | اعتماد عدم التسامح 3 مخالفات = استبعاد [ ] نعم [ ] لا\n\nSignature: _________________ Project Director Date: _______ | Signature: _________________ Construction Manager Date: _______", size=8.5, color=DARK_TEXT)

# Build final correctly ordered presentation
final_prs = Presentation()
final_prs.slide_width = Inches(13.33)
final_prs.slide_height = Inches(7.5)
create_title_slide(final_prs)
for data in slides_data:
    create_violation_slide(final_prs, data)
create_closing_slide(final_prs)

out = "/home/user/ahmed/Subcontractor_EHS_Executive_White_7Slides_AR_EN.pptx"
final_prs.save(out)
print(f"Saved final executive white version to {out}")
