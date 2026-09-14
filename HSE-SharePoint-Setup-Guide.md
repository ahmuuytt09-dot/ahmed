# دليل إنشاء موقع SharePoint لإدارة السلامة والصحة المهنية (HSE)

> **نطاق الدليل:** Microsoft 365 – SharePoint Online (الواجهة الحديثة – Modern Experience).
> يغطي الدليل الأقسام الأربعة المطلوبة: (1) إنشاء الموقع والصلاحيات، (2) مكتبات المستندات والأعمدة المخصصة، (3) التنبيهات التلقائية عبر Power Automate، (4) تصميم الصفحة الرئيسية.

---

## القسم الأول: إنشاء الموقع وضبط الخصوصية والصلاحيات

### 1-1. متطلبات ما قبل البدء (مهم)

1. **حساب Microsoft 365** بإذن إنشاء مواقع SharePoint. إذا لم يظهر لك زر "Create site" فإن مسؤول المؤسسة (Tenant Admin) قيّد الإنشاء — اطلب منه إنشاء الموقع وتسليمك صلاحية المالك.
2. **المشاركة الخارجية (External Sharing) يجب أن تكون مفعّلة** لأن جزءاً كبيراً من الأعضاء عناوينهم خارجية (Gmail / siemens-energy.com / al-mialiq.com إذا كان موقعكم مستضافاً على مستأجر آخر). التفاصيل في الخطوة 1-3.
3. حسابات الضيوف الخارجيين ستتحقق من هويتها عبر **رمز تحقق يُرسل للبريد (One-Time Passcode)** أو حساب Microsoft المرتبط بنفس البريد — لا يحتاجون ترخيص Microsoft 365 خاصاً بهم للوصول كضيوف.

### 1-2. خطوات إنشاء الموقع

| # | الخطوة | التفاصيل |
|---|--------|----------|
| 1 | ادخل إلى **office.com** وسجّل الدخول بحساب العمل | ثم افتح تطبيق **SharePoint** من قائمة التطبيقات (⋮⋮ أعلى اليسار) |
| 2 | اضغط **+ Create site** (إنشاء موقع) | في الصفحة الرئيسية لـ SharePoint |
| 3 | اختر **Team site** (موقع فريق) | يتضمن صفحة رئيسية + مكتبة مستندات أساسية جاهزة |
| 4 | اختر القالب **Standard team** | أي قالب مناسب — التصميم سنبنيه يدوياً لاحقاً |
| 5 | اسم الموقع: **HSE Management** (أو "إدارة السلامة والصحة المهنية – مشروع KAZ") | يتولّد العنوان تلقائياً مثل: `/sites/HSE-KAZ` — يمكنك تعديله |
| 6 | **Privacy settings: اختر Private** | الخيار: *Private – only members can access this site* — وهذا يحقق شرطك أن الموقع خاص بالكامل ومغلق عن باقي الموظفين |
| 7 | اختر اللغة | العربية أو الإنجليزية (تؤثر على لغة الواجهة فقط، والمحتوى يقبل أي لغة) |
| 8 | **Next ← لا تضف أعضاء الآن ← Finish** | سنضيفهم في الخطوة التالية بالصلاحيات الدقيقة المطلوبة |

### 1-3. تفعيل المشاركة الخارجية للضيوف (Gmail وغيرها)

> نظراً لأن المالك `Ahmuuytt09@gmail.com` والعضو `alnedaweyraed@gmail.com` حسابات Gmail، وبقية الأعضاء على نطاقات شركات أخرى، فجميعهم **ضيوف خارجيون (External Guests)** ويحتاجون هذا الإعداد:

1. افتح **مركز إدارة Microsoft 365** ← admin.microsoft.com ← سجّل الدخول بحساب مسؤول.
2. **Show all ← SharePoint** (مركز إدارة SharePoint).
3. من القائمة: **Policies ← Sharing**، واضبط مستوى SharePoint إلى **"New and existing guests"** على الأقل (لا يُشترط اختيار Anyone/الروابط المجهولة — تجنّبها).
4. ثم **Sites ← Active sites** ← حدد موقع HSE ← علامة التبويب **Policies ← External sharing: Edit** ← اختر **"Only people in your organization" = لا**، بل **"New and existing guests"** ← **Save**.

> بدون هذه الخطوة سيحصل الضيوف على رسالة "You need permission to access this site" مهما حاولت إضافتهم.

### 1-4. توزيع الصلاحيات على الأشخاص المحددين

**البنية الافتراضية لأي موقع خاص:**
- **Site Owners = Full Control** (تحكم كامل)
- **Site Members = Edit** (تحرير)
- **Site Visitors = Read** (قراءة فقط)

**جدول الإسناد المطلوب:**

| # | البريد الإلكتروني | المجموعة | الصلاحية |
|---|-------------------|----------|----------|
| مالك | Ahmuuytt09@gmail.com | Site Owners | **Full Control** |
| 1 | dogukan.arandi@siemens-energy.com | Site Members | Edit |
| 2 | ugur.akbulut@siemens-energy.com | Site Members | Edit |
| 3 | alnedaweyraed@gmail.com | Site Members | Edit |
| 4 | h.farah@al-mialiq.com | Site Members | Edit |
| 5 | hse.dep@al-mialiq.com | Site Members | Edit |
| 6 | planner@al-mialiq.com | Site Members | Edit |
| 7 | onur.ozvatan@siemens-energy.com | Site Members | Edit |

> **ملاحظة Edit أم Read:** كل من تريد له **قراءة فقط** (دون رفع أو تعديل ملفات) ضعه في مجموعة **Site Visitors** بدلاً من Members — باقي الخطوات نفسها. الأكثر شيوعاً لفرق HSE: مهندسو HSE ومدير الموقع بصلاحية Edit، والسيمنر/المخطط Read حسب الحاجة.

**الطريقة A — الواجهة الحديثة (الأسهل):**

1. داخل الموقع: أيقونة الإعدادات **⚙ ← Site permissions** (أذونات الموقع).
2. اضغط **Add members**.
3. أدخل بريد المستخدم ← سيظهر تنبيه أنه *ضيف خارجي* ← اختر الصلاحية من القائمة المنسدلة: **Full Control** للمالك، **Edit** للأعضاء، **Read** للقارئين فقط.
4. أضف رسالة ترحيبية اختيارية ← **Add**.
5. كرر لكل بريد من القائمة.

**الطريقة B — الإعدادات المتقدمة (تحكم أدق، موصى بها عند التعقيد):**

1. **⚙ ← Site permissions ← Advanced permissions settings**.
2. تظهر المجموعات الثلاث: **HSE Owners / HSE Members / HSE Visitors**.
3. افتح المجموعة المطلوبة ← **New ← Add Users** ← ألصق البريد ← فعّل خيار *Send an email invitation* ← **Share**.

**تنبيه مهم بخصوص المالك الخارجي:**
- منح `Ahmuuytt09@gmail.com` صلاحية **Full Control** يعمل بلا مشاكل عبر صفحة Advanced Permissions (مجموعة Site Owners).
- لكن بعض إجراءات "مالك مجموعة Microsoft 365" تتطلب حساباً داخلياً؛ لذا يُفضّل إبقاء **حساب داخلي واحد على الأقل ضمن Owners** كاحتياط إداري.
- بعد قبول الدعوة يصل الضيف إلى الموقع هذا فقط — **لن يرى أي موقع آخر** في المستأجر.

### 1-5. قفل الموقع بالكامل (حجبه عن باقي الموظفين)

1. **⚙ ← Site permissions ← Site sharing settings:**
   - فعّل: **Only site owners can share files, folders, and the site** ← يمنع الأعضاء من إعادة مشاركة الموقع أو محتوياته مع أشخاص غير مصرح لهم.
2. في صفحة **Advanced permissions settings**:
   - تأكد أن مجموعة **Visitors لا تحتوي** على "Everyone" أو "Everyone except external users".
3. **Access requests:** اختر بريلاً لاستقبال طلبات الوصول (يفضل بريد المالك) — أي موظف غير مصرح له سيصل إلى صفحة "Request access" ويرفض تلقائياً ما لم يوافق المالك.

### 1-6. التحقق من النتيجة

1. **Advanced permissions settings ← Check Permissions** ← اكتب أي بريد ← **Check Now** لعرض صلاحياته الفعلية.
2. اختبار سلبي: افتح رابط الموقع في **نافذة تصفح خفي (InPrivate)** بحساب شخصي غير مدرج ← يجب أن تظهر صفحة "You don't have access" — وهذا يؤكد أن الموقع محجوب عن بقية الموظفين.

---

## القسم الثاني: إنشاء مكتبات المستندات والأعمدة المخصصة

### 2-0. خطوات موحّدة لإنشاء أي مكتبة (طبّقها على الأقسام الخمسة)

1. في الموقع: **Site contents** (محتويات الموقع) ← **+ New ← Document library**.
2. الاسم والوصف ← **Create**.
3. داخل المكتبة: زر **+ Add column** في شريط العرض ← اختر نوع العمود.
4. **قاعدة ذهبية:** أنشئ الأعمدة بأسماء **إنجليزية بلا مسافات** مثل `ExpiryDate`، ثم أعد تسميتها لعرضها بالعربية. الاسم الأول يصبح *الاسم الداخلي* المستخدم في الفلاتر والتنسيقات — أي مسافات فيه ستعقّد صيغ Power Automate لاحقاً.
5. **خصّص العمود:** من `+ Add column` تختار النوع (Choice / Date and time / Person / Yes/No / Number / Single line of text / Multiple lines) وتضبط:
   - القيم المسموحة لأعمدة Choice مع ألوانها.
   - **Require that this column contains information = Yes** للأعمدة الإلزامية مثل رقم التصريح وتاريخ الانتهاء.
6. **فهرسة الأعمدة المهمة:** Library settings ← **Indexed columns** ← أضف `ExpiryDate` و`Status` و`PermitType` — يحسّن أداء الفلاتر والتنبيهات على المدى الطويل.
7. **Versioning:** Library settings ← **Versioning settings ← Create major versions** — يحتفظ بتاريخ التعديلات على الملفات (ضروري في التدقيق).

> **نصيحة تنظيمية:** اعتمد **الأعمدة (Metadata) بدلاً من المجلدات المتداخلة**؛ التصنيف بالبيانات الوصفية أسرع في البحث والتصفية والتنبيهات.

### 2-1. مكتبة تصاريح العمل (PTW – Permit to Work)

**الأعمدة المخصصة:**

| العمود (الاسم الداخلي) | النوع | الإعدادات |
|------------------------|-------|-----------|
| PermitNo | سطر نص واحد (Single line of text) | إلزامي — رقم التصريح |
| PermitType | اختيار (Choice) | القيم: Hot Work, Cold Work, Confined Space, Working at Height, Excavation, Electrical, Lifting Operations, Other |
| Contractor | شخص (Person) | الجهة/المقاول المنفذ |
| WorkLocation | سطر نص واحد | موقع العمل داخل المشروع |
| RiskLevel | اختيار (Choice) | Low / Medium / High |
| StartDate | تاريخ (Date and time) | تاريخ بدء العمل |
| ExpiryDate | تاريخ | **إلزامي** — تاريخ انتهاء التصريح |
| Status | اختيار (Choice) | Draft / Pending Approval / Active / Suspended / Expired / Closed (افتراضي: Pending Approval) |
| ApprovedBy | شخص | المعتمد |
| Notes | متعدد الأسطر | ملاحظات وإجراءات رقابية |

**طرق العرض (Views) المقترحة:**
- **Active Permits:** فلتر `Status = Active`.
- **Expiring within 7 days:** `ExpiryDate ≤ [Today]+7` و`Status = Active`.
- **Expired / Closed:** للأرشيف والتدقيق.
- تجميع (Group by) حسب `PermitType`.

### 2-2. مكتبة محادثات السلامة اليومية (TBT – Toolbox Talks)

| العمود | النوع | الإعدادات |
|--------|-------|-----------|
| TalkDate | تاريخ | إلزامي — تاريخ المحاضرة |
| Topic | سطر نص واحد | إلزامي — موضوع المحاضرة |
| Presenter | شخص | المقدم/مهندس السلامة |
| DepartmentArea | سطر نص واحد | القسم/المنطقة |
| AttendeesCount | رقم (Number) | عدد الحضور |
| DurationMin | رقم | المدة بالدقائق |
| Category | اختيار (Choice) | Fall Protection, Electrical Safety, PPE, Fire Safety, LOTO, Heat Stress, Housekeeping, Other |

**Views:** آخر 30 يوماً، تجميع حسب الشهر، تجميع حسب Category.

> كملفات: ارفع نموذج الحضور الموقع (PDF ممسوح ضوئياً) أو صور ورقة التوقيع في نفس السجل.

### 2-3. مكتبة شهادات الفحص من الطرف الثالث (Third Party Certificates)

| العمود | النوع | الإعدادات |
|--------|-------|-----------|
| CertificateNo | سطر نص واحد | إلزامي — رقم الشهادة |
| Category | اختيار (Choice) | Equipment / Personnel / Vehicle / Other |
| ItemName | سطر نص واحد | **إلزامي** — اسم المعدة أو الفرد |
| EquipmentID | سطر نص واحد | الرقم التعريفي للمعدة |
| CertifyingBody | اختيار (Choice) أو نص | TÜV, Bureau Veritas, SGS, DNV, Other |
| IssueDate | تاريخ | تاريخ الإصدار |
| ExpiryDate | تاريخ | **إلزامي** — تاريخ الانتهاء (عمود التنبيهات) |
| RenewalStatus | اختيار (Choice) | Valid / Under Renewal / Expired |

**Views:**
- **Expiring in 30 days:** `ExpiryDate ≤ [Today]+30` (هو الأساس الذي يبنى عليه التنبيه).
- **Expiring in 7 days:** `ExpiryDate ≤ [Today]+7`.
- **Expired:** `ExpiryDate < [Today]`.
- تجميع حسب `Category`.

### 2-4. مكتبة سجلات الموظفين والسلامة (CVs / Personnel Safety Records)

| العمود | النوع | الإعدادات |
|--------|-------|-----------|
| EmployeeName | سطر نص واحد | إلزامي |
| JobTitle | سطر نص واحد | المسمى الوظيفي |
| Company | اختيار (Choice) | Al-Mialiq / Siemens Energy / Subcontractor / Other |
| Department | سطر نص واحد | القسم |
| Qualifications | اختيار متعدد (Choice – Allow multiple) | NEBOSH IGC, IOSH, OSHA, First Aid & CPR, Fire Watcher, Rigger L1–L3, Confined Space Attendant, PTW Holder, Other |
| QualificationExpiry | تاريخ | أقرب تاريخ انتهاء لأهم مؤهل |
| NationalID_No | سطر نص واحد | اختياري — **بيانات حساسة** |

> **تنبيه الخصوصية (PII):** بيانات الموظفين حساسة. إذا أردت حصر هذا القسم على المالك ومدير HSE فقط:
> افتح مكتبة Personnel Records ← **⚙ ← Library settings ← More library settings ← Permissions for this document library ← Stop Inheriting Permissions** ← ثم احذف من لم ترد وأضف المصرح لهم فقط. بقية الموقع يبقى على صلاحياته العامة.

### 2-5. مكتبة ملفات السلامة العامة (General Safety Files)

| العمود | النوع | الإعدادات |
|--------|-------|-----------|
| DocumentNo | سطر نص واحد | رقم الوثيقة |
| DocumentType | اختيار (Choice) | Safety Policy / Incident Report / Near Miss / Emergency Plan / Risk Assessment / Method Statement / Safe Work Procedure / Form-Template / Other |
| Version | سطر نص واحد | رقم الإصدار (مثل V2) |
| EffectiveDate | تاريخ | تاريخ السريان |
| ReviewDate | تاريخ | موعد المراجعة الدورية (أساس تنبيهات مستقلة) |
| DocOwner | شخص | مسؤول الوثيقة |

**Views:** تجميع حسب `DocumentType`، وعرض **Due for Review:** `ReviewDate ≤ [Today]`.

> هذه المكتبة مكانها الطبيعي: سياسات EHS، خطط المشروع (EHS Plan)، تحليل المخاطر (Risk Analysis)، نماذج Near-Miss — مثل الملفات الموجودة لديك حالياً في المشروع.

### 2-6. إبراز الحالة والتواريخ بالألوان (JSON Column Formatting)

**كيفية التطبيق** (نفس الأسلوب لأي عمود): من رأس العمود ← **Column settings ← Format this column ← Advanced mode** ← ألصق الكود ← **Save**.

**1) تلوين حقل `ExpiryDate` (أحمر = منتهٍ، كهرماني = خلال 30 يوماً، أخضر = ساري):**

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "txtContent": "=toLocaleDateString(@currentField)",
  "style": {
    "display": "inline-block",
    "padding": "4px 10px",
    "border-radius": "12px",
    "font-weight": "600",
    "background-color": "=if(@currentField <= @now, '#FDE7E9', if(@currentField <= addDays(@now, 30), '#FFF4CE', '#DFF6DD'))",
    "color": "=if(@currentField <= @now, '#A80000', if(@currentField <= addDays(@now, 30), '#9C6500', '#107C10'))"
  }
}
```

**2) تلوين حقل `Status` كشارة ملونة:**

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/sp/v2/column-formatting.schema.json",
  "elmType": "div",
  "txtContent": "@currentField",
  "style": {
    "display": "inline-block",
    "padding": "4px 12px",
    "border-radius": "12px",
    "font-weight": "bold",
    "color": "#FFFFFF",
    "background-color": "=if(@currentField=='Active' || @currentField=='Valid', '#107C10', if(@currentField=='Pending Approval' || @currentField=='Under Renewal', '#CA5010', if(@currentField=='Suspended', '#986F0B', if(@currentField=='Expired', '#A80000', '#605E5C'))))"
  }
}
```

**تنبيه تقني مهم:** لا تعتمد على عمود **محسوب (Calculated)** فيه `TODAY()` لأن قيمته لا تتحدث تلقائياً مع مرور الأيام (تُحسب عند تعديل العنصر فقط). بدلاً من ذلك استخدم:
- فلاتر الـ **Views بصيغة `[Today]`** (تعمل لحظياً)، و
- تنسيق JSON أعلاه (يستخدم `@now` ويعمل لحظياً)، و
- حساب الفارق داخل **Power Automate** عند الإرسال.

---

## القسم الثالث: التنبيهات التلقائية عبر Power Automate

### 3-1. الخيار السريع المدمج (يستغرق دقيقة)

داخل أي مكتبة فيها عمود تاريخ (مثل Certificates):
**شريط الأدوات ← Automate ← Set a reminder ← Remind me X days in advance** ← حدد عمود `ExpiryDate` وعدد الأيام ← إنشاء.

- ينشئ هذا تدفقاً جاهزاً تلقائياً.
- **الحد:** يرسل التذكير لك أنت فقط (منشئ التدفق) — لا يوزع على فريق كامل. لذلك للفريق المحدد نبني التدفق الاحترافي أدناه.

### 3-2. التدفق الاحترافي: ملخص أسبوعي "شهادات/تصاريح تنتهي خلال 30 يوماً"

**النتيجة:** كل يوم اثنين 08:00 صباحاً، يفحص النظام المكتبة ويرسل بريلاً واحداً منسقاً للأشخاص المحددين بكل ما يقترب انتهاؤه.

**الخطوات بالتفصيل:**

1. افتح **make.powerautomate.com** وسجّل الدخول بنفس حساب Microsoft 365.
2. **Create ← Scheduled cloud flow:**
   - الاسم: `HSE-Expiry-Digest`
   - نمط التشغيل: **Repeat every 1 Week on Monday at 08:00** (تقدر تبدأ Daily ثم تعدله).
3. **+ New step ← "Initialize variable":**
   - Name: `varMaxDate` — Type: **String**
   - Value: انقر **fx** وأدخل: `formatDateTime(addDays(utcNow(), 30), 'yyyy-MM-dd')`
4. **+ New step ← SharePoint ← "Get items":**
   - Site Address: اختر موقع HSE من القائمة.
   - List Name: اختر مكتبة **Third Party Certificates**.
   - افتح **Advanced options ← Filter Query** وألصق (مع علامات الاقتباس المفردة كما هي):
     ```
     ExpiryDate ge '2000-01-01T00:00:00Z' and ExpiryDate le '@{variables('varMaxDate')}'
     ```
     > يستبعد هذا الحقول الفارغة ويلتقط كل ما انتهى أو سينتهي خلال 30 يوماً. لاستثناء المغلق أضف: `and RenewalStatus ne 'Valid'` أو حسب سياستك.
5. **+ New step ← "Select"** (Data Operations):
   - From: اختر `value` من خطوة Get items.
   - Map: أضف الأعمدة: `CertificateNo` ← CertificateNo، `ItemName` ← ItemName، `CertifyingBody` ← CertifyingBody، `ExpiryDate` ← ExpiryDate، `RenewalStatus` ← RenewalStatus.
6. **+ New step ← "Create HTML table":**
   - From: ناتج Select ← Columns: **Automatic**.
7. **+ New step ← "Send an email (V2)"** (Office 365 Outlook):
   - **To:** `hse.dep@al-mialiq.com`
   - **Cc:** `Ahmuuytt09@gmail.com; alnedaweyraed@gmail.com; h.farah@al-mialiq.com; planner@al-mialiq.com`
   - **Subject:** `تقرير أسبوعي HSE – شهادات وتصاريح تنتهي صلاحيتها خلال 30 يوماً`
   - **Body:** ألصق ناتج **Create HTML table** من Dynamic content داخل نص مثل:

     ```html
     <p>تحية طيبة،</p>
     <p>فيما يلي الشهادات التي ستنتهي صلاحيتها خلال الثلاثين يوماً القادمة. يرجى المباشرة بإجراءات التجديد فوراً:</p>
     [هنا أدرِج ناتج HTML table]
     <p>المكتبة المباشرة: [رابط المكتبة]</p>
     <p>مع خالص التحية،<br>فريق السلامة والصحة المهنية – HSE</p>
     ```
   - ضع نفسك سابقاً **Condition** فوق الإرسال إذا أردت: "إذا كان عدد النتائج = 0 لا ترسل" ← Expression: `length(outputs('Get_items')?['body/value']) is greater than 0`.
8. **Save ← Test ← Manually ← Run flow** للتحقق. (أنشئ أولاً عنصر اختبار بتاريخ انتهاء قريب.)
9. **استنساخ التدفق لمكتبة PTW:** من صفحة التدفق ← **… (More) ← Save As ←** غيّر المكتبة إلى **PTW** وعدّل العمود/الموضوع — واحفظ.

### 3-3. نسخة "تنبيه صباحي لكل عنصر" (بديل يومي بسيط)

- نفس الخطوات 1–4 لكن: Recurrence = **Daily 07:30**.
- بدلاً من Select/HTML: **+ New step ← Apply to each (value) ← داخله "Send an email (V2)"**:
  - Subject: `⚠️ تنبيه HSE: الشهادة {CertificateNo} للمعدة {ItemName} ستنتهي بتاريخ {ExpiryDate}`
  - Body: يتضمن رابط العنصر (`Link to item`) وتعليمات التجديد.

### 3-4. تدفقات إضافية مقترحة

| التدفق | المحفز | الإجراء |
|--------|--------|---------|
| إشعار تصريح جديد | When an item is created في مكتبة PTW | بريد فوري لـ hse.dep والمالك بأن تصريحاً جديداً بحاجة للمراجعة والاعتماد |
| تحديث تلقائي للحالة | Recurrence يومي | على عناصر `ExpiryDate < اليوم`: Update item ← Status = Expired |
| تذكير مراجعة السياسات | Recurrence شهري | فلتر `ReviewDate ≤ +30 يوم` في General Safety Files |
| تنبيه عاجل | Recurrence يومي | فلتر `ExpiryDate ≤ +3 أيام` وإرسال لأولوية قصوى |

### 3-5. ملاحظات تشغيل أساسية

- **الترخيص:** موصلات SharePoint وOffice 365 Outlook *قياسية* مشمولة ضمن اشتراك Microsoft 365 — لا حاجة لرخصة Power Automate مميزة.
- البريد يُرسل من **حساب منشئ التدفق** افتراضياً؛ لإظهار مرسل مؤسسي أنشئ بريداً مشتركاً (Shared Mailbox) وامنح نفسك Send As عليه.
- **Run History** في Power Automate يعرض سجل 28 يوماً لتشخيص أي فشل.
- تأكد أن منطقة التوقيت إقليمك (مثلاً Asia/Riyadh) من إعدادات الـ Recurrence.
- في صفحة الموقع: **Site Settings ← Site administration ← Regional settings ← Calendar type = Gregorian** حتى لا تختلط تواريخ الشهادات بالهجري.

---

## القسم الرابع: تصميم الصفحة الرئيسية

### 4-1. الهيكل المقترح

```
┌────────────────────────────────────────────────────────┐
│  HERO: شعار المشروع + عنوان "السلامة أولاً – Safety First"        │
│  + عداد أيام بلا حوادث + زر "بلّغ عن خطر"                      │
├────────────────────────────────────────────────────────┤
│  QUICK LINKS (6 بطاقات ملونة):                                  │
│  📋 تصاريح العمل | 🗣️ محادثات السلامة | 📜 شهادات الطرف الثالث    │
│  👤 سجلات الموظفين | 📁 ملفات السلامة | 🚨 رقم الطوارئ/الشرطة     │
├──────────────────────────┬─────────────────────────────┤
│  تصاريح العمل النشطة      │  شهادات تنتهي خلال 30 يوماً   │
│  (عرض Active Permits)     │  (عرض Expiring in 30 days)    │
├──────────────────────────┴─────────────────────────────┤
│  آخر محادثات السلامة (آخر 5 TBT) + تقويم اجتماعات HSE  │
├──────────────────────────┬─────────────────────────────┤
│  أخبار السلامة (News)     │  جهات اتصال الطوارئ (People)  │
└──────────────────────────┴─────────────────────────────┘
```

### 4-2. خطوات البناء العملية

1. افتح الصفحة الرئيسية للموقع ← **Edit** (أعلى اليمين).
2. **القسم الأول (Full-width):** احذف/عدّل **Hero web part**: عنوان "HSE Management – السلامة والصحة المهنية"، صورة خلفية لموقع المشروع أو أيقونة خوذة، ورابط سريع لمكتبة PTW.
3. **+** أضف قسماً جديداً (Section) بتخطيط **ثلاثة أعمدة (Two columns/Three columns حسب الحاجة):**
4. أضف **Quick links web part** بأسلوب **Grid** وأضف روابط:
   - تصاريح العمل PTW → رابط مكتبة PTW
   - محادثات السلامة TBT → رابط مكتبة TBT
   - شهادات الطرف الثالث → رابط مكتبة Certificates
   - سجلات الموظفين → رابط Personnel Records
   - ملفات السلامة العامة → رابط General Safety Files
   - طوارئ المشروع → رقم/صفحة جهة الاتصال
5. **قسم بعمودين:** في كل عمود أضف **Document library web part** ثم اضبطه:
   - يمين: مكتبة PTW ← من خصائص العنصر اختر العرض **Active Permits** وعدد العناصر 10.
   - يسار: مكتبة Certificates ← العرض **Expiring in 30 days**.
6. **قسم بعمود واحد:** أضف **List/Document web part** لعرض آخر 5 TBT.
7. **قسمان جانبيان:** **Events** لاجتماعات السلامة وتنبيهات المواعيد + **News** لأخبار الحوادث والإنجازات.
8. أضف **People web part** لمسؤولي الطوارئ (فريق HSE والـ warden) وأرقامهم.
9. من **⚙ ← Change the look ← Theme:** اختر ألواناً بحرية/سلامة (أخضر داكن + أصفر تحذيري)، ومن **Site information** أضف شعار الشركة/المشروع.
10. **Republish** ثم من قائمة الصفحة **Make homepage** لتثبيتها كرئيسية.

### 4-3. التنقل الجانبي (Navigation)

عدّل قائمة **التنقل اليسرى** (Edit بجانب القائمة) وأضف بالترتيب:
**الرئيسية ← تصاريح العمل PTW ← محادثات السلامة TBT ← شهادات الطرف الثالث ← سجلات الموظفين ← ملفات السلامة ← Site contents.**

> القاعدة: أي شيء يحتاجه الفريق يومياً يكون على بعد نقرة واحدة من الرئيسية أو القائمة.

---

## ملاحق

### ملحق (أ): قائمة الأعضاء وصلاحياتهم النهائية (Quick Reference)

| البريد | الدور | المجموعة | الصلاحية |
|--------|-------|----------|----------|
| Ahmuuytt09@gmail.com | مسؤول الموقع | Owners | Full Control |
| hse.dep@al-mialiq.com | قسم HSE | Members | Edit |
| h.farah@al-mialiq.com | إدارة المشروع | Members | Edit |
| planner@al-mialiq.com | التخطيط | Members | Edit |
| alnedaweyraed@gmail.com | عضو فريق | Members | Edit |
| dogukan.arandi@siemens-energy.com | Siemens Energy | Members (أو Visitors إذا قراءة فقط) | Edit |
| ugur.akbulut@siemens-energy.com | Siemens Energy | Members (أو Visitors) | Edit |
| onur.ozvatan@siemens-energy.com | Siemens Energy | Members (أو Visitors) | Edit |

### ملحق (ب): تعبيرات جاهزة للنسخ (Power Automate)

```
التاريخ بعد 30 يوماً:
formatDateTime(addDays(utcNow(), 30), 'yyyy-MM-dd')

فلتر "ما انتهى أو سينتهي خلال 30 يوماً":
ExpiryDate ge '2000-01-01T00:00:00Z' and ExpiryDate le '@{variables('varMaxDate')}'

فلتر "ينتهي خلال 7 أيام":
ExpiryDate ge '@{formatDateTime(utcNow(),'yyyy-MM-dd')}' and ExpiryDate le '@{formatDateTime(addDays(utcNow(),7),'yyyy-MM-dd')}'

عدد النتائج (للشرط قبل الإرسال):
length(outputs('Get_items')?['body/value'])

أيام متبقية داخل نص البريد:
div(sub(ticks(item()?['ExpiryDate']), ticks(utcNow())), 864000000000)
```

> عند إدخال التعبيرات داخل خانة Filter Query استخدم **fx (Expression editor)** لضمان تقييم التعبير، مع إبقاء علامات الاقتباس المفردة حول قيمة التاريخ.

### ملحق (ج): قائمة التحقق النهائية (Checklist)

- [ ] إنشاء Team site خاص (Private) باسم HSE.
- [ ] تفعيل External Sharing على مستوى المستأجر ومستوى الموقع.
- [ ] إضافة المالك (Full Control) والأعضاء السبعة (Edit/Read).
- [ ] تقييد مشاركة الموقع على المالكين فقط.
- [ ] إنشاء المكتبات الخمس بالأعمدة والفلاتر والـ Views.
- [ ] تفعيل Versioning + Indexed columns.
- [ ] تطبيق تنسيقات JSON للأعمدة.
- [ ] إنشاء تدفق الملخص الأسبوعي + استنساخه لمكتبة PTW.
- [ ] أتمتة تحديث الحالة إلى Expired عند تجاوز التاريخ.
- [ ] بناء الصفحة الرئيسية وتثبيتها Make homepage.
- [ ] ضبط التوقيت الإقليمي والتقويم الميلادي.
- [ ] اختبار الدخول بحساب غير مصرح (يجب أن يُرفض).
- [ ] اختبار التدفق Test Run بعنصر تجريبي منتهي قريباً.

---

*أُعدّ هذا الدليل خصيصاً لهيكل HSE المطلوب (PTW / TBT / Third Party Certificates / Personnel Records / General Safety Files) ولقائمة الأعضاء المحددة. يمكن توسيعه لاحقاً بسجلات النواقص (Punch List) أو لوحة مؤشرات Power BI HSE.*
