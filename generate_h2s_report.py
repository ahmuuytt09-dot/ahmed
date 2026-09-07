from pathlib import Path
from html import escape
from datetime import date

import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "reports" / "h2s-safety-report-ar.pdf"
FONT = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

pdfmetrics.registerFont(TTFont("DejaVu", FONT))
pdfmetrics.registerFont(TTFont("DejaVu-Bold", FONT_BOLD))

NAVY = colors.HexColor("#0B1F33")
BLUE = colors.HexColor("#0E5A8A")
TEAL = colors.HexColor("#0B7A75")
AMBER = colors.HexColor("#D97706")
RED = colors.HexColor("#B42318")
INK = colors.HexColor("#17202A")
MUTED = colors.HexColor("#5B6770")
PALE_BLUE = colors.HexColor("#EAF4FA")
PALE_TEAL = colors.HexColor("#E8F6F4")
PALE_AMBER = colors.HexColor("#FFF4DB")
PALE_RED = colors.HexColor("#FDECEA")
GRID = colors.HexColor("#D6E0E7")


def rtl(value: str) -> str:
    """Shape Arabic and reorder it for display in a left-to-right PDF engine."""
    return get_display(arabic_reshaper.reshape(value))


def p(text: str, style: ParagraphStyle, *, shape: bool = True) -> Paragraph:
    value = rtl(text) if shape else text
    return Paragraph(escape(value).replace("\n", "<br/>"), style)


def rich(text: str, style: ParagraphStyle) -> Paragraph:
    """Paragraph helper for carefully controlled markup in already-shaped text."""
    return Paragraph(text, style)


styles = getSampleStyleSheet()
body = ParagraphStyle(
    "ArabicBody",
    parent=styles["BodyText"],
    fontName="DejaVu",
    fontSize=10.2,
    leading=17,
    textColor=INK,
    alignment=TA_RIGHT,
    spaceAfter=5,
)
body_small = ParagraphStyle(
    "ArabicBodySmall",
    parent=body,
    fontSize=8.8,
    leading=13,
    textColor=MUTED,
)
section = ParagraphStyle(
    "Section",
    parent=body,
    fontName="DejaVu-Bold",
    fontSize=15,
    leading=21,
    textColor=NAVY,
    spaceBefore=10,
    spaceAfter=7,
)
subsection = ParagraphStyle(
    "Subsection",
    parent=body,
    fontName="DejaVu-Bold",
    fontSize=11,
    leading=16,
    textColor=BLUE,
    spaceBefore=5,
    spaceAfter=3,
)
small_right = ParagraphStyle(
    "SmallRight",
    parent=body,
    fontSize=8,
    leading=11,
    textColor=MUTED,
)
cell = ParagraphStyle(
    "Cell",
    parent=body,
    fontSize=8.2,
    leading=12,
    spaceAfter=0,
)
cell_bold = ParagraphStyle(
    "CellBold",
    parent=cell,
    fontName="DejaVu-Bold",
    textColor=NAVY,
)
cell_center = ParagraphStyle(
    "CellCenter",
    parent=cell,
    alignment=TA_CENTER,
)
cell_header = ParagraphStyle(
    "CellHeader",
    parent=cell,
    fontName="DejaVu-Bold",
    fontSize=8.6,
    leading=12,
    textColor=colors.white,
    alignment=TA_RIGHT,
)
callout = ParagraphStyle(
    "Callout",
    parent=body,
    fontSize=10.3,
    leading=17,
    textColor=NAVY,
)
title = ParagraphStyle(
    "Title",
    parent=body,
    fontName="DejaVu-Bold",
    fontSize=24,
    leading=32,
    alignment=TA_CENTER,
    textColor=colors.white,
    spaceAfter=9,
)
subtitle = ParagraphStyle(
    "Subtitle",
    parent=body,
    fontSize=12,
    leading=19,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#D8EEF8"),
)
cover_meta = ParagraphStyle(
    "CoverMeta",
    parent=body,
    fontSize=9,
    leading=15,
    alignment=TA_CENTER,
    textColor=colors.HexColor("#B8CBD8"),
)
source_title = ParagraphStyle(
    "SourceTitle",
    parent=body,
    fontName="DejaVu-Bold",
    fontSize=8.8,
    leading=13,
    textColor=NAVY,
    spaceAfter=1,
)
source_url = ParagraphStyle(
    "SourceUrl",
    parent=styles["BodyText"],
    fontName="DejaVu",
    fontSize=7.2,
    leading=10,
    textColor=BLUE,
    alignment=TA_LEFT,
    wordWrap="CJK",
)


def section_heading(text: str):
    return [p(text, section), Spacer(1, 1 * mm)]


def bullet(text: str, accent=TEAL):
    # The bullet is intentionally kept at the visual right edge by using an RTL paragraph.
    return Table(
        [[p("•  " + text, body)]],
        colWidths=[170 * mm],
        style=TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("LINEBEFORE", (0, 0), (0, -1), 2, accent),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        ),
    )


def info_box(title_text: str, content: str, color, background):
    return Table(
        [[p(title_text, ParagraphStyle("box-title", parent=subsection, textColor=color, spaceAfter=2)), p(content, callout)]],
        colWidths=[42 * mm, 128 * mm],
        style=TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), background),
                ("BOX", (0, 0), (-1, -1), 0.7, color),
                ("LINEBEFORE", (0, 0), (0, -1), 4, color),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        ),
    )


def header_footer(canvas, doc):
    canvas.saveState()
    width, height = A4
    # Thin header rule on content pages.
    if doc.page > 1:
        canvas.setStrokeColor(GRID)
        canvas.setLineWidth(0.5)
        canvas.line(20 * mm, height - 14 * mm, width - 20 * mm, height - 14 * mm)
        canvas.setFont("DejaVu", 7.5)
        canvas.setFillColor(MUTED)
        canvas.drawRightString(width - 20 * mm, height - 11 * mm, rtl("تقرير السلامة | كبريتيد الهيدروجين H2S"))
        canvas.drawString(20 * mm, 10 * mm, "H2S • SAFETY REPORT")
        canvas.drawRightString(width - 20 * mm, 10 * mm, rtl(f"صفحة {doc.page}"))
    canvas.restoreState()


def make_pdf():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc = BaseDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=22 * mm,
        bottomMargin=17 * mm,
        title="أضرار كبريتيد الهيدروجين H2S",
        author="Arena.ai Agent",
        subject="تقرير توعوي عن مخاطر كبريتيد الهيدروجين والوقاية منه",
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates([PageTemplate(id="main", frames=[frame], onPage=header_footer)])

    story = []

    # Cover
    story.append(
        Table(
            [[p("تقرير توعوي", cover_meta)], [p("أضرار كبريتيد الهيدروجين", title)], [p("H₂S | المخاطر الصحية والوقاية والاستجابة للطوارئ", subtitle)], [Spacer(1, 8 * mm)], [p("مرجع عملي مختصر للعاملين والمشرفين وأفراد المجتمع", cover_meta)]],
            colWidths=[170 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                    ("BOX", (0, 0), (-1, -1), 0, NAVY),
                    ("LEFTPADDING", (0, 0), (-1, -1), 16),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 16),
                    ("TOPPADDING", (0, 0), (-1, 0), 14),
                    ("BOTTOMPADDING", (0, -1), (-1, -1), 18),
                    ("TOPPADDING", (0, 1), (-1, -2), 4),
                    ("BOTTOMPADDING", (0, 1), (-1, -2), 4),
                ]
            ),
        )
    )
    story.append(Spacer(1, 8 * mm))
    story.append(
        info_box(
            "تحذير حرج",
            "قد يؤدي التعرض لتركيز مرتفع من H2S إلى فقدان الوعي والانهيار خلال أنفاس قليلة ثم الوفاة خلال دقائق. لا تعتمد على الرائحة، ولا تدخل منطقة مشتبه بها لإنقاذ شخص من دون تدريب ومعدات تنفس مناسبة.",
            RED,
            PALE_RED,
        )
    )
    story.append(Spacer(1, 5 * mm))
    story.append(
        Table(
            [[p("إعداد التقرير", cell_bold), p("مبني على إرشادات OSHA وNIOSH وATSDR الرسمية", cell)], [p("تاريخ الإصدار", cell_bold), p("سبتمبر 2026", cell)], [p("نطاق الاستخدام", cell_bold), p("التوعية العامة؛ يجب تطبيق تشريعات وإجراءات الموقع المحلية وبيانات السلامة SDS", cell)]],
            colWidths=[38 * mm, 132 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), PALE_BLUE),
                    ("GRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            ),
        )
    )
    story.append(PageBreak())

    story += section_heading("1. الملخص التنفيذي")
    story.append(
        p(
            "كبريتيد الهيدروجين H2S غاز عديم اللون، شديد السمية وقابل للاشتعال، وقد يتولد عند تحلل المواد العضوية في غياب الأكسجين. يوجد في شبكات الصرف الصحي ومحطات المعالجة وخزانات السماد وبعض عمليات النفط والغاز والتكرير واللب والورق والتعدين. الخطر الأكبر هو أن المستويات العالية قد تعطل حاسة الشم بسرعة؛ لذلك فإن عدم شم رائحة البيض الفاسد لا يعني أن الجو آمن.",
            body,
        )
    )
    story.append(
        info_box(
            "قاعدة الإنقاذ",
            "إذا انطلق إنذار H2S أو ظهرت أعراض أو كان مصدر الغاز غير معروف: اخرج فوراً إلى هواء نقي من جهة عكس الريح أو عمودياً عليها، نبّه الآخرين، واتصل بالطوارئ. لا تحاول الإنقاذ بالدخول من دون جهاز تنفس مستقل موجب الضغط وفريق مدرّب.",
            AMBER,
            PALE_AMBER,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(p("أهم الحقائق الرقمية", subsection))
    facts = [
        [p("100 ppm", ParagraphStyle("factnum", parent=cell_bold, fontSize=15, textColor=RED, alignment=TA_CENTER)), p("حد التركيز المصنف خطرًا مباشرًا على الحياة والصحة (IDLH) وفق NIOSH.", cell_center)],
        [p("10 ppm", ParagraphStyle("factnum2", parent=cell_bold, fontSize=15, textColor=BLUE, alignment=TA_CENTER)), p("قيمة NIOSH الموصى بها كسقف لمدة 10 دقائق؛ وليست تصريحاً بتجاوز ضوابط الموقع.", cell_center)],
        [p("20 ppm", ParagraphStyle("factnum3", parent=cell_bold, fontSize=15, textColor=AMBER, alignment=TA_CENTER)), p("حد السقف العام القابل للإنفاذ لدى OSHA في الصناعة العامة، مع استثناءات محددة للذروة.", cell_center)],
    ]
    story.append(
        Table(
            facts,
            colWidths=[35 * mm, 135 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), PALE_RED),
                    ("BACKGROUND", (0, 1), (-1, 1), PALE_BLUE),
                    ("BACKGROUND", (0, 2), (-1, 2), PALE_AMBER),
                    ("BOX", (0, 0), (-1, -1), 0.5, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            ),
        )
    )

    story += section_heading("2. ما هو H2S وأين يوجد؟")
    story.append(
        p(
            "H2S مركب يحتوي على الهيدروجين والكبريت. يكون عادة غازاً في درجة الحرارة المحيطة، وقد يُشحن كسائل غاز مضغوط. رائحته المميزة قد تشبه البيض الفاسد عند المستويات المنخفضة، لكن قابلية اكتشافها تختلف بين الأشخاص وتتلاشى سريعاً عند التعرض المستمر أو المرتفع.",
            body,
        )
    )
    sources = [
        "الأماكن المغلقة: الآبار، غرف التفتيش، الخزانات، الحفر، خطوط الصرف وأحواض المعالجة.",
        "المصادر العضوية: مياه الصرف، السماد، مخلفات الحيوانات، المكبات وتحلل المواد العضوية.",
        "القطاعات الصناعية: إنتاج النفط والغاز، التكرير، البتروكيماويات، اللب والورق، التعدين وبعض الصناعات الغذائية.",
        "خصائص إضافية: الغاز قابل للاشتعال؛ حد الاشتعال الأدنى في الهواء نحو 4.3% والحـد الأعلى نحو 45%، وقد يتجمع في الأماكن المنخفضة لأن كثافته أكبر من كثافة الهواء.",
    ]
    for item in sources:
        story.append(bullet(item, BLUE))
    story.append(Spacer(1, 3 * mm))
    story.append(
        info_box(
            "لا تختبر الغاز بحاسة الشم",
            "تبدأ الرائحة عند تراكيز شديدة الانخفاض لدى بعض الأشخاص، لكن حاسة الشم قد تتعب أو تُشل عند نحو 100 ppm. لذلك لا يجوز استخدام الرائحة ككاشف أو كدليل على انخفاض التركيز.",
            TEAL,
            PALE_TEAL,
        )
    )

    story += section_heading("3. الأضرار الصحية حسب مستوى التعرض")
    story.append(
        p(
            "الجدول التالي إرشادي؛ التأثير الفعلي يتأثر بالتركيز والزمن والجهد البدني والحالة الصحية. لا ينبغي تفسير أي صف على أنه وقت آمن للتعرض. عند الشك، اعتبر الجو خطراً واخرج منه.",
            body_small,
        )
    )
    health_rows = [
        [p("التركيز التقريبي", cell_header), p("أمثلة على الأعراض أو التأثيرات", cell_header)],
        [p("0.01–1.5 ppm", cell_bold), p("قد تُلاحظ الرائحة لدى بعض الأشخاص؛ عتبة الشم متغيرة جداً وليست وسيلة حماية.", cell)],
        [p("2–5 ppm", cell_bold), p("مع التعرض المطول: غثيان، دموع، صداع أو اضطراب؛ وقد تتفاقم مشكلات مجرى الهواء لدى بعض المصابين بالربو.", cell)],
        [p("20 ppm", cell_bold), p("قد يظهر تعب، فقدان شهية، صداع، تهيج أو ضعف في الذاكرة.", cell)],
        [p("50–100 ppm", cell_bold), p("تهيج العينين والجهاز التنفسي؛ قد يظهر ما يسمى عين الغاز.", cell)],
        [p("100 ppm", cell_bold), p("سعال وتهيج للعينين، ثم فقدان الشم بسرعة؛ قد يحدث تغير في التنفس ونعاس. هذا المستوى مصنف IDLH.", cell)],
        [p("200–300 ppm", cell_bold), p("التهاب وتهيج واضحان للعينين والجهاز التنفسي؛ وقد تحدث وذمة رئوية مع التعرض المطول.", cell)],
        [p("500–700 ppm", cell_bold), p("ترنح وانهيار خلال دقائق محتملة، وإصابة خطيرة للعينين؛ قد تحدث الوفاة.", cell)],
        [p("700–1000+ ppm", cell_bold), p("فقدان وعي أو انهيار سريع جداً خلال نفس أو نفسَين، توقف التنفس والوفاة خلال دقائق محتملة.", cell)],
    ]
    story.append(
        Table(
            health_rows,
            colWidths=[38 * mm, 132 * mm],
            repeatRows=1,
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                    ("BACKGROUND", (0, 1), (-1, 4), colors.HexColor("#F7FBFD")),
                    ("BACKGROUND", (0, 5), (-1, 5), PALE_AMBER),
                    ("BACKGROUND", (0, 6), (-1, 7), PALE_RED),
                    ("BACKGROUND", (0, 8), (-1, 8), colors.HexColor("#F9D8D5")),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            ),
        )
    )
    story.append(Spacer(1, 3 * mm))
    story.append(
        p(
            "أعضاء الجسم الأكثر تأثراً: العينان والجهاز التنفسي والجهاز العصبي المركزي. وقد تستمر بعد التعرض الشديد بعض المشكلات مثل الصداع وضعف الانتباه والذاكرة أو الوظائف الحركية. المصابون بالربو أو بأمراض تنفسية قد يكونون أكثر حساسية.",
            body,
        )
    )

    story += section_heading("4. متى يصبح الموقف طارئاً؟")
    story.append(p("اعتبر الحالة طارئة في أي من الآتي:", subsection))
    for item in [
        "صدور إنذار كاشف الغاز، أو قراءة غير طبيعية، أو تعذر التحقق من الجو.",
        "وجود شخص منهار أو مشوش أو يعاني صعوبة في التنفس قرب فتحة خزان أو غرفة تفتيش أو مكان منخفض.",
        "ظهور حرقة في العين أو الحلق، سعال، صداع، دوار، غثيان، ضعف، نعاس أو ارتباك.",
        "اشتباه تسرب أو رائحة كبريتية في مكان مغلق أو قرب مصدر قابل للاشتعال.",
    ]:
        story.append(bullet(item, RED))
    story.append(Spacer(1, 3 * mm))
    story.append(info_box("تذكر", "الضحية الأولى قد تتبعها ضحايا ثانويات عندما يدخل شخص غير مدرب لإنقاذها. الإنقاذ من جو IDLH مهمة فريق طوارئ مجهز، وليس مهمة زميل منفرد.", RED, PALE_RED))

    story.append(PageBreak())
    story += section_heading("5. الاستجابة الآمنة للتعرض أو التسرب")
    response_rows = [
        [p("الخطوة", cell_header), p("ما يجب فعله", cell_header)],
        [p("1. اخرج", cell_bold), p("غادر فوراً إلى هواء نقي من جهة عكس اتجاه الريح أو عمودياً عليها. ابتعد عن المناطق المنخفضة، ولا تعُد لجلب أدوات أو أشخاص.", cell)],
        [p("2. نبّه واعزل", cell_bold), p("أطلق الإنذار، امنع الدخول، واتصل بالطوارئ أو فريق الاستجابة. لا تشغّل مفاتيح أو معدات قد تولّد شرراً قرب تسرب محتمل.", cell)],
        [p("3. لا تنقذ منفرداً", cell_bold), p("لا تدخل لإنقاذ شخص من دون تصريح، مراقبة جوية، خطة إنقاذ، فريق مدرب وجهاز تنفس مستقل موجب الضغط. لا تعتمد على قناع غبار أو كمامة عادية.", cell)],
        [p("4. إسعاف من مكان آمن", cell_bold), p("اطلب الإسعاف. ينفذ المدربون فقط الإنعاش القلبي الرئوي واستخدام مزيل الرجفان وتقديم الأكسجين وفق البروتوكول المحلي وبعد تأمين المكان.", cell)],
        [p("5. تلوث العين أو الجلد", cell_bold), p("في حال ملامسة سائل H2S أو الغاز المضغوط، اغسل العينين أو الجلد بالماء الوفير واطلب تقييماً طبياً. قد يسبب السائل أذى بارداً شبيهاً بقضمة الصقيع.", cell)],
    ]
    story.append(
        Table(
            response_rows,
            colWidths=[38 * mm, 132 * mm],
            repeatRows=1,
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), RED),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FFF9F8")),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            ),
        )
    )
    story.append(Spacer(1, 5 * mm))
    story += section_heading("6. الوقاية والسيطرة على الخطر")
    story.append(p("يُفضّل تطبيق هرم السيطرة على المخاطر بالترتيب الآتي، مع عدم اعتبار معدات الوقاية بديلاً عن الضوابط الهندسية:", body))
    controls = [
        ("الإزالة أو المنع", "أوقف العملية أو صمّمها بحيث لا يتولد H2S أو لا يصل إلى العاملين؛ عالج المصادر العضوية والتسربات."),
        ("الضوابط الهندسية", "استخدم تهوية مستمرة مناسبة للمناطق المغلقة، وأنظمة كشف وإنذار معايرة، وشفطاً موضعياً ومعدات مقاومة للانفجار وشرراً أقل."),
        ("الضوابط الإدارية", "تحليل مخاطر المهمة، تصريح دخول الأماكن المحصورة، تدريب، حارس خارج المكان، نظام رفيق، لوحات تحذير، خطط طوارئ وتمارين دورية."),
        ("المراقبة", "اختبر الجو من الخارج قبل الدخول وبشكل مستمر أثناء العمل؛ افحص الأكسجين والغازات القابلة للاشتعال وH2S وباقي الملوثات ذات الصلة."),
        ("معدات الوقاية", "استخدم حماية العين والوجه والملابس المناسبة. عند 100 ppm أو أكثر، أو إذا كان التركيز مجهولاً، يلزم جهاز SCBA موجب الضغط أو جهاز هواء مزود مع مصدر هواء احتياطي وفق تقييم مختص."),
    ]
    control_rows = [[p("طبقة السيطرة", cell_header), p("التطبيق العملي", cell_header)]]
    for a, b in controls:
        control_rows.append([p(a, cell_bold), p(b, cell)])
    story.append(
        Table(
            control_rows,
            colWidths=[42 * mm, 128 * mm],
            repeatRows=1,
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), TEAL),
                    ("BACKGROUND", (0, 1), (-1, -1), PALE_TEAL),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 7),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                    ("TOPPADDING", (0, 0), (-1, -1), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ]
            ),
        )
    )

    story += section_heading("7. حدود التعرض: كيف تُقرأ؟")
    story.append(
        p(
            "حدود التعرض تختلف باختلاف الدولة والقطاع وطريقة القياس. الأرقام أدناه مرجع أمريكي من المصادر الرسمية المذكورة، وليست بديلاً عن النظام المحلي أو تقييم مختص. ppm تعني جزءاً من المليون؛ وTWA متوسط زمني؛ وCeiling سقف لا ينبغي تجاوزه؛ وIDLH تركيز خطر مباشرة على الحياة أو الصحة.",
            body,
        )
    )
    limits = [
        [p("الجهة أو القيمة", cell_header), p("الرقم", cell_header), p("المعنى المختصر", cell_header)],
        [p("NIOSH REL", cell_bold), p("10 ppm / 10 دقائق", cell_center), p("سقف موصى به لمدة عشر دقائق.", cell)],
        [p("NIOSH IDLH", cell_bold), p("100 ppm", cell_center), p("خطر مباشر على الحياة أو الصحة؛ لا دخول إلا بإجراءات ومعدات IDLH.", cell)],
        [p("OSHA صناعة عامة", cell_bold), p("20 ppm سقف", cell_center), p("حد إنفاذ عام؛ توجد قاعدة ذروة محددة بشروط صارمة تصل إلى 50 ppm لمدة 10 دقائق.", cell)],
        [p("OSHA إنشاء وبناء / أحواض سفن", cell_bold), p("10 ppm TWA", cell_center), p("متوسط ثماني ساعات بحسب الجداول المشار إليها في OSHA.", cell)],
        [p("ACGIH TLV", cell_bold), p("1 ppm TWA / 5 ppm STEL", cell_center), p("قيم إرشادية مهنية؛ يجب الرجوع إلى الإصدار الحالي والترخيص المحلي.", cell)],
    ]
    story.append(
        Table(
            limits,
            colWidths=[43 * mm, 43 * mm, 84 * mm],
            repeatRows=1,
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F7FBFD")),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            ),
        )
    )

    story += section_heading("8. قائمة تحقق سريعة قبل العمل")
    checklist = [
        "هل حُددت كل مصادر H2S المحتملة؟",
        "هل فُحصت الأجواء من الخارج بجهاز معاير، بما في ذلك الأكسجين وقابلية الاشتعال؟",
        "هل يعمل الإنذار والمراقبة المستمرة والتهوية؟",
        "هل صدر تصريح دخول للأماكن المحصورة، مع حارس وخطة إنقاذ واتصال؟",
        "هل يعرف الجميع مسار الهروب ونقطة التجمع واتجاه الريح؟",
        "هل تتوافر معدات التنفس المناسبة ويعرف الفريق حدود استخدامها؟",
        "هل تم تدريب العاملين على أن الرائحة ليست إنذاراً موثوقاً؟",
    ]
    checklist_table = [[p("☐  " + item, body)] for item in checklist]
    story.append(
        Table(
            checklist_table,
            colWidths=[170 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), PALE_BLUE),
                    ("BOX", (0, 0), (-1, -1), 0.6, BLUE),
                    ("LINEBELOW", (0, 0), (-1, -2), 0.3, GRID),
                    ("LEFTPADDING", (0, 0), (-1, -1), 9),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            ),
        )
    )

    story += section_heading("9. الخلاصة")
    story.append(
        p(
            "H2S ليس مجرد رائحة مزعجة؛ إنه غاز سام قابل للاشتعال يمكن أن يسبب تهيجاً وإصابة عصبية وتنفسية وانهياراً سريعاً. الحماية الفعالة تبدأ بمنع التولد والتسرب، ثم القياس المستمر والتهوية وتصريح العمل والتدريب، مع تجهيز إنقاذ لا يعرّض المنقذين للخطر. عند الشك: اخرج، حذّر، واتصل بالطوارئ.",
            body,
        )
    )
    story.append(Spacer(1, 4 * mm))
    story.append(info_box("تنبيه طبي وسلامة", "هذا التقرير مادة توعوية عامة، ولا يحدد صلاحية مكان عمل بعينه ولا يستبدل تقييم مخاطر أو نشرة بيانات السلامة أو تعليمات الجهات المختصة. في حالة التعرض اتصل بخدمات الطوارئ المحلية فوراً.", NAVY, PALE_BLUE))

    story.append(PageBreak())
    story += section_heading("المراجع الرسمية")
    story.append(p("تمت مراجعة الروابط في سبتمبر 2026. يُرجى الرجوع إلى أحدث نسخة من المتطلبات المحلية قبل اعتماد أي إجراء أو حد تعرض.", body_small))
    refs = [
        ("OSHA — Hydrogen Sulfide: Hazards", "https://www.osha.gov/hydrogen-sulfide/hazards"),
        ("OSHA — Hydrogen Sulfide: Evaluating and Controlling Exposure", "https://www.osha.gov/hydrogen-sulfide/evaluating-controlling-exposure"),
        ("OSHA — Hydrogen Sulfide: Standards", "https://www.osha.gov/hydrogen-sulfide/standards"),
        ("NIOSH / CDC — Pocket Guide: Hydrogen Sulfide", "https://www.cdc.gov/niosh/npg/npgd0337.html"),
        ("NIOSH / CDC — Low Level Exposure to Hydrogen Sulfide: A Review", "https://stacks.cdc.gov/view/cdc/131722"),
        ("ATSDR / CDC — Toxicological Profile for Hydrogen Sulfide and Carbonyl Sulfide", "https://www.atsdr.cdc.gov/toxprofiles/tp114-c2.pdf"),
    ]
    ref_rows = []
    for name, url in refs:
        ref_rows.append([p(name, source_title), Paragraph(f'<link href="{url}" color="#0E5A8A">{escape(url)}</link>', source_url)])
    story.append(
        Table(
            ref_rows,
            colWidths=[60 * mm, 110 * mm],
            style=TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7FBFD")),
                    ("BOX", (0, 0), (-1, -1), 0.6, GRID),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, GRID),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            ),
        )
    )
    story.append(Spacer(1, 10 * mm))
    story.append(p("نهاية التقرير", ParagraphStyle("end", parent=body, alignment=TA_CENTER, fontName="DejaVu-Bold", textColor=BLUE)))

    doc.build(story)
    print(f"Wrote {OUT} ({OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    make_pdf()
