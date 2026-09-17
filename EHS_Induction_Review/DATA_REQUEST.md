# نموذج بيانات الإندكشن · Induction Data Request

**المستند:** KAZ-EHS-IND-001 · إصدار الإندكشن المطلوب: Rev. 03
**الغرض:** المعلومات الناقصة التي يجب تعبئتها قبل إصدار Rev. 03 المعتمد.
**How to use:** اكتب القيمة بعد `**VALUE:**` مباشرة ثم احفظ (Commit). اتركها فارغة
إذا لم تتوفر بعد. يمكنك أيضاً تحميل ملف Excel `KAZ_Induction_Data_Request.xlsx` وتعبئته ورفعه.

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


| المجموعة | الموضوع | عدد الحقول |
|---|---|---|
| A | الطوارئ والإنقاذ الطبي — Emergency and rescue medical | 11 |
| B | بيانات السلامة الكهربائية المعتمدة — Approved electrical safety data | 4 |
| C | الغازات الخطرة — Hazardous gases | 5 |
| D | التصوير الإشعاعي و NDT — Radiography and NDT | 2 |
| E | الموقع والصور والمخطط — Site layout, photographs and plot plan | 4 |
| F | مسار الإبلاغ — Incident reporting route | 4 |
| G | اللغة والتدريب واللوجستيات — Language, training and logistics | 5 |
| H | الأمن والمعسكر — Security and camp | 4 |
| I | الاعتماد — Approval | 5 |
| J | أي شيء آخر — Anything else | 2 |

**الإجمالي: 46 حقلاً.** اكتب القيمة بعد `**VALUE:**` واحفظ التغيير.

---

## A · الطوارئ والإنقاذ الطبي — Emergency and rescue medical

### A1 · رقم غرفة الطوارئ / خط الطوارئ في الموقع 24 ساعة 🔒
*Site emergency control room / 24-7 emergency line*
`مثال: 0785-000-0000 — مفتوح 24/7`

**VALUE:**

### A2 · قناة الراديو UHF للطوارئ 🔒
*UHF radio channel for emergencies*
`مثال: Channel 1`

**VALUE:**

### A3 · رقم المسعف / نقطة الإسعاف وموقعها 🔒
*Site medic / first aid post — number and location*
`المبنى أو الموقع داخل المعسكر`

**VALUE:**

### A4 · مواقع أجهزة الصدمات الكهربائية AED 🔒
*AED locations (all units)*
`مطلوب في بطاقة الطوارئ المعتمدة — اذكر كل المواقع`

**VALUE:**

### A5 · المستشفى الأقرب ورقمه 🔒
*Nearest hospital — name and phone*

**VALUE:**

### A6 · مستشفى الحروق 🔒
*Hospital for burns*
`حقل إلزامي في بطاقة الطوارئ`

**VALUE:**

### A7 · مستشفى العيون 🔒
*Hospital for eye injuries*
`حقل إلزامي في بطاقة الطوارئ`

**VALUE:**

### A8 · مستشفى حالات البتر / الجراحة 🔒
*Hospital for detached body parts / surgery*
`حقل إلزامي في بطاقة الطوارئ`

**VALUE:**

### A9 · ترتيب الإخلاء الطبي (طائرة / سيارة) وجهة الاتصال 🔒
*Medevac arrangement (air / road) and contact*
`من ينظم الإخلاء: Siemens Field Service أم العميل`

**VALUE:**

### A10 · نقاط التجمع ومواقعها 🔒
*Assembly points — locations*
`الموقع الرئيسي + نقاط المعسكر والمكاتب`

**VALUE:**

### A11 · رقم سيارة الإسعاف في الموقع 🔒
*On-site ambulance number*

**VALUE:**

---

## B · بيانات السلامة الكهربائية المعتمدة — Approved electrical safety data

### B1 · المسافات الآمنة والقوس الكهربائي — 11 ك.ف 🔒
*Approach distance and arc flash — 11 kV*
`الحد الأدنى للمسافة | حد القوس | الطاقة الحادثة | فئة PPE`

**VALUE:**

### B2 · المسافات الآمنة والقوس الكهربائي — 132 ك.ف 🔒
*Approach distance and arc flash — 132 kV*
`نفس التنسيق`

**VALUE:**

### B3 · المسافات الآمنة والقوس الكهربائي — 400 ك.ف 🔒
*Approach distance and arc flash — 400 kV*
`نفس التنسيق`

**VALUE:**

### B4 · مرجع دراسة القوس الكهربائي (الرقم والتاريخ)
*Arc flash study reference and date*
`مثال: SE Arc Flash Study KAZ-2026-014، تاريخ 2026-03`

**VALUE:**

---

## C · الغازات الخطرة — Hazardous gases

### C1 · هل يوجد غاز عزل المفاتيح SF6 في النطاق؟
*Is SF6 present in the scope?*
`اكتب: نعم / لا — وإذا نعم اذكر المعدات والجهة المسؤولة`

**VALUE:**

### C2 · إجراء التعامل مع SF6 وإعادة التعبئة
*SF6 handling and refill procedure*
`اسم الإجراء أو رقمه`

**VALUE:**

### C3 · هل يوجد H2S أو غازات سامة في المنطقة؟
*Is H2S or any toxic gas present?*
`نعم / لا — وإذا نعم: حد التعرض بالـ ppm ومناطق الخطر`

**VALUE:**

### C4 · غازات أخرى (ميثان، أول أكسيد الكربون، نقص أكسجين)
*Other gases (methane, CO, oxygen deficiency)*

**VALUE:**

### C5 · أجهزة كشف الغاز الشخصية — النوع والعدد
*Personal gas detectors — type and quantity*

**VALUE:**

---

## D · التصوير الإشعاعي و NDT — Radiography and NDT

### D1 · هل يوجد تصوير إشعاعي أو NDT في النطاق؟
*Is radiography / NDT in the scope?*
`نعم / لا`

**VALUE:**

### D2 · إن وُجد: الشركة المنفذة والترخيص وإجراء المنطقة
*If yes: contractor, licence and area control procedure*

**VALUE:**

---

## E · الموقع والصور والمخطط — Site layout, photographs and plot plan

### E1 · صورة البيس كمب / المكتب الرئيسي — اسم ملف الصورة
*Base camp / main office photograph — file name*
`ارفع الصورة إلى المجلد ثم اكتب اسم الملف، مثال: basecamp.jpg`

**VALUE:**

### E2 · مخطط الموقع — اسم ملف المخطط 🔒
*Site plot plan — file name*
`يجب أن يُظهر المكاتب والمخيم والعيادة ونقاط التجمع وطرق الطوارئ`

**VALUE:**

### E3 · مواقع العيادة / المطعم / المغسلة / المصلى
*Clinic / canteen / laundry / prayer room locations*

**VALUE:**

### E4 · حدود مناطق الخطر الحمراء وقيود الدخول 🔒
*Red zone boundaries and access restrictions*

**VALUE:**

---

## F · مسار الإبلاغ — Incident reporting route

### F1 · الإجراء المعتمد للإبلاغ: بطاقة السلامة أم Enablon أم الاثنان؟
*Approved reporting route: Safety Card, Enablon, or both?*
`يجب توحيد الإجراء بين الإندكشن وبطاقة الطوارئ`

**VALUE:**

### F2 · رابط أو رقم نموذج Enablon
*Enablon form link or number*

**VALUE:**

### F3 · مهلة تعبئة بطاقة السلامة الورقية
*Deadline for completing the paper Safety Card*
`بطاقة الطوارئ تقول: في نفس اليوم`

**VALUE:**

### F4 · جهة استلام التقرير والتحقيق
*Who receives the report and investigates*

**VALUE:**

---

## G · اللغة والتدريب واللوجستيات — Language, training and logistics

### G1 · اللغات المطلوبة للإندكشن
*Required induction languages*
`عربي / تركي / كردي / أخرى — البند 3.8 من SE Instruction يُلزم بذلك`

**VALUE:**

### G2 · الشرائح التي تريد ترجمتها
*Which slides to translate*
`افتراضي: 15 شريحة حرجة (القواعد، الكهرباء، الارتفاع، الرفع، الحفر، الأماكن المغلقة، الحرائق، الطوارئ، الأمن، العقوبات)`

**VALUE:**

### G3 · عدد المشاركين المتوقع في الجلسة
*Expected number of participants per session*

**VALUE:**

### G4 · مدة الجلسة المعتمدة بالدقائق
*Approved session duration in minutes*
`الملف الحالي يقول 75 دقيقة — غير كافٍ لـ 47 شريحة محتوى`

**VALUE:**

### G5 · هل توجد شاشة/بروجكتر ووسائط بصرية في قاعة التدريب؟
*Projector / screen and visual media available in the training room?*

**VALUE:**

---

## H · الأمن والمعسكر — Security and camp

### H1 · رقم غرفة الأمن الرئيسية 🔒
*Main security control room number*

**VALUE:**

### H2 · إجراء الزوار المعتمد
*Approved visitor procedure*

**VALUE:**

### H3 · مشرف المعسكر — الاسم والرقم 🔒
*Camp supervisor — name and number*

**VALUE:**

### H4 · رقم إدارة EHS في الموقع
*Site EHS department number*

**VALUE:**

---

## I · الاعتماد — Approval

### I1 · مدير المشروع — الاسم
*Project Manager — name*
`لصفحة الاعتماد`

**VALUE:**

### I2 · مدير EHS — الاسم
*EHS Manager — name*
`لصفحة الاعتماد`

**VALUE:**

### I3 · تاريخ الاعتماد المطلوب
*Target approval date*

**VALUE:**

### I4 · رقم الإصدار المطلوب
*Issue revision to publish*
`افتراضي: Rev. 03`

**VALUE:**

### I5 · هل تريد صفحة اعتماد وتوقيعات وسجل إصدارات؟
*Add an approval page, signatures and revision history?*
`نعم / لا`

**VALUE:**

---

## J · أي شيء آخر — Anything else

### J1 · مخاطر أو مواضيع خاصة أخرى يجب إضافتها
*Any other hazards or topics that must be added*

**VALUE:**

### J2 · ملاحظات حرة
*Any other notes*

**VALUE:**

---

## كيف أقرأ البيانات

بعد تعبئة النموذج أخبرني، أو شغّل الأمر التالي لمعرفة ما تبقّى:

```bash
python3 data_request.py read
```
