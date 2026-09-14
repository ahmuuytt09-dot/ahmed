# ============================================================================
#  HSE SharePoint Site - Auto Setup Script
#  سكربت الإنشاء الآلي لموقع إدارة السلامة والصحة المهنية
#
#  ما يفعله هذا السكربت آلياً:
#   1) إنشاء موقع Team Site خاص (Private) باسم HSE Management
#   2) تفعيل المشاركة الخارجية على الموقع (لحسابات Gmail وغيرها)
#   3) إنشاء المكتبات الخمس مع كل الأعمدة المخصصة والاختيارات
#   4) إنشاء طرق العرض (Views) بفلاتر التواريخ الذكية [Today]
#   5) تطبيق تنسيقات JSON الملونة على عمود ExpiryDate و Status
#   6) محاولة إضافة المالك والأعضاء سبعة بالصلاحيات المطلوبة
#
#  المتطلبات (مرة واحدة فقط):
#   - PowerShell 7 (أو Windows PowerShell 5.1)
#   - حساب بصلاحية SharePoint Admin للخطوتين 1-2 (أو اطلب من IT)
#   - تثبيت الوحدة:  Install-Module -Name PnP.PowerShell -Scope CurrentUser
#
#  التشغيل:
#   1) عدّل قسم "الإعدادات" أدناه (اسم المستأجر + الإيميلات)
#   2) افتح PowerShell وشغّل:   .\HSE-Site-Auto-Setup.ps1
#   3) سجّل الدخول في المتصفح عند الطلب
# ============================================================================

# ============================ الإعدادات — عدّل هنا ============================
$TenantName   = "YOURTENANT"                       # مثال: ahmedco  → ahmedco.sharepoint.com
$SiteTitle    = "HSE Management"                   # اسم الموقع
$SiteAlias    = "HSE-Management"                   # اختصار الرابط /sites/HSE-Management

$SiteOwner    = "Ahmuuytt09@gmail.com"             # المالك - Full Control
$SiteMembers  = @(                                 # الأعضاء - Edit
  "dogukan.arandi@siemens-energy.com",
  "ugur.akbulut@siemens-energy.com",
  "alnedaweyraed@gmail.com",
  "h.farah@al-mialiq.com",
  "hse.dep@al-mialiq.com",
  "planner@al-mialiq.com",
  "onur.ozvatan@siemens-energy.com"
)
# ==============================================================================

$ErrorActionPreference = "Continue"
$TenantAdminUrl = "https://$TenantName-admin.sharepoint.com"
$SiteUrl        = "https://$TenantName.sharepoint.com/sites/$SiteAlias"

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "  بدء إنشاء موقع HSE — Site URL: $SiteUrl" -ForegroundColor Cyan
Write-Host "=====================================================`n" -ForegroundColor Cyan

# ---------- الخطوة 0: تسجيل التطبيق (مرة واحدة فقط لأول استخدام) ----------
Write-Host "[0] إذا لم تستخدم PnP.PowerShell من قبل على هذا المستأجر، نفّذ مرة واحدة:" -ForegroundColor Yellow
Write-Host '    Register-PnPAzureADApp -ApplicationName "HSESetup" -Tenant "'$TenantName'.onmicrosoft.com" -Interactive' -ForegroundColor Yellow
Write-Host "    (تتطلب موافقة مسؤول/Admin - ثم استخدم -ClientId الذي ينتج)`n" -ForegroundColor Yellow

try {
    # ---------- الخطوة 1: إنشاء الموقع الخاص ----------
    Write-Host "[1] الاتصال بمركز الإدارة وإنشاء الموقع..." -ForegroundColor Green
    Connect-PnPOnline -Url $TenantAdminUrl -Interactive
    New-PnPSite -Type TeamSite -Title $SiteTitle -Alias $SiteAlias -IsPrivate
    Write-Host "    ✔ تم إنشاء الموقع الخاص (Private): $SiteUrl" -ForegroundColor Green

    # ---------- الخطوة 2: تفعيل المشاركة الخارجية ----------
    Write-Host "[2] تفعيل المشاركة الخارجية للضيوف..." -ForegroundColor Green
    Set-PnPTenantSite -Identity $SiteUrl -Sharing ExternalUserSharingOnly
    Write-Host "    ✔ المشاركة الخارجية مفعّلة (New and existing guests)" -ForegroundColor Green

    # ---------- الخطوة 3: الاتصال بالموقع ----------
    Disconnect-PnPOnline
    Connect-PnPOnline -Url $SiteUrl -Interactive

    # ===================== دوال مساعدة =====================
    function New-HseLibrary($name, $desc) {
        New-PnPList -Title $name -Template DocumentLibrary -OnQuickLaunch | Out-Null
        # تفعيل إدارة الإصدارات
        Set-PnPList -Identity $name -MajorVersions 10 | Out-Null
        Write-Host "    ✔ مكتبة: $name" -ForegroundColor Green
    }
    function Add-TextCol($list, $internal, $display, $required=$false) {
        Add-PnPField -List $list -InternalName $internal -DisplayName $display -Type Text -AddToDefaultView -Required:$required | Out-Null
    }
    function Add-NumCol($list, $internal, $display) {
        Add-PnPField -List $list -InternalName $internal -DisplayName $display -Type Number -AddToDefaultView | Out-Null
    }
    function Add-DateCol($list, $internal, $display, $required=$false) {
        Add-PnPField -List $list -InternalName $internal -DisplayName $display -Type DateTime -AddToDefaultView -Required:$required | Out-Null
        # عرض التاريخ فقط بدون وقت
        try { Set-PnPField -List $list -Identity $internal -Values @{ DisplayFormat = 0 } | Out-Null } catch {}
    }
    function Add-PersonCol($list, $internal, $display) {
        Add-PnPField -List $list -InternalName $internal -DisplayName $display -Type User -AddToDefaultView | Out-Null
    }
    function Add-ChoiceCol($list, $internal, $display, $choices, $required=$false, $multi=$false) {
        $type = "Choice"; if ($multi) { $type = "MultiChoice" }
        Add-PnPField -List $list -InternalName $internal -DisplayName $display -Type $type -Choices $choices -AddToDefaultView -Required:$required | Out-Null
    }
    # تنسيق JSON: تلوين تاريخ الانتهاء (أحمر/كهرماني/أخضر)
    $ExpiryJson = @'
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
'@
    # تنسيق JSON: شارة الحالة الملونة
    $StatusJson = @'
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
'@
    function Add-ExpiryViews($list, $fields) {
        Add-PnPView -List $list -Title "Expiring in 30 days" -Fields $fields -Query "<Where><And><Geq><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today/></Value></Geq><Leq><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today OffsetDays='30'/></Value></Leq></And></Where>" | Out-Null
        Add-PnPView -List $list -Title "Expired" -Fields $fields -Query "<Where><Lt><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today/></Value></Lt></Where>" | Out-Null
    }
    # =======================================================

    Write-Host "[3] إنشاء المكتبات والأعمدة..." -ForegroundColor Green

    # ---- مكتبة 1: تصاريح العمل PTW ----
    $ptw = "PTW - Work Permits"
    New-HseLibrary $ptw "Permit to Work - تصاريح العمل"
    Add-TextCol   $ptw "PermitNo"     "Permit No"    $true
    Add-ChoiceCol $ptw "PermitType"   "Permit Type"  @("Hot Work","Cold Work","Confined Space","Working at Height","Excavation","Electrical","Lifting Operations","Other")
    Add-PersonCol $ptw "Contractor"   "Contractor"
    Add-TextCol   $ptw "WorkLocation" "Work Location"
    Add-ChoiceCol $ptw "RiskLevel"    "Risk Level"   @("Low","Medium","High")
    Add-DateCol   $ptw "StartDate"    "Start Date"
    Add-DateCol   $ptw "ExpiryDate"   "Expiry Date"  $true
    Add-ChoiceCol $ptw "Status"       "Status"       @("Draft","Pending Approval","Active","Suspended","Expired","Closed")
    Add-PersonCol $ptw "ApprovedBy"   "Approved By"
    Set-PnPField -List $ptw -Identity "Status"     -Values @{ CustomFormatter = $StatusJson } | Out-Null
    Set-PnPField -List $ptw -Identity "ExpiryDate" -Values @{ CustomFormatter = $ExpiryJson } | Out-Null
    $ptwFields = @("Title","PermitNo","PermitType","Contractor","Status","ExpiryDate")
    Add-PnPView -List $ptw -Title "Active Permits" -Fields $ptwFields -Query "<Where><Eq><FieldRef Name='Status'/><Value Type='Choice'>Active</Value></Eq></Where>" | Out-Null
    Add-PnPView -List $ptw -Title "Expiring in 7 days" -Fields $ptwFields -Query "<Where><And><Geq><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today/></Value></Geq><Leq><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today OffsetDays='7'/></Value></Leq></And></Where>" | Out-Null
    Add-PnPView -List $ptw -Title "Expired" -Fields $ptwFields -Query "<Where><Lt><FieldRef Name='ExpiryDate'/><Value Type='DateTime'><Today/></Value></Lt></Where>" | Out-Null

    # ---- مكتبة 2: محادثات السلامة TBT ----
    $tbt = "TBT - Toolbox Talks"
    New-HseLibrary $tbt "Toolbox Talks - محادثات السلامة اليومية"
    Add-DateCol   $tbt "TalkDate"       "Talk Date"  $true
    Add-TextCol   $tbt "Topic"          "Topic"      $true
    Add-PersonCol $tbt "Presenter"      "Presenter"
    Add-TextCol   $tbt "DepartmentArea" "Department / Area"
    Add-NumCol    $tbt "AttendeesCount" "Attendees Count"
    Add-ChoiceCol $tbt "Category"       "Category"   @("Fall Protection","Electrical Safety","PPE","Fire Safety","LOTO","Heat Stress","Housekeeping","Other")

    # ---- مكتبة 3: شهادات الطرف الثالث ----
    $cert = "Third Party Certificates"
    New-HseLibrary $cert "Third Party Certificates - شهادات الفحص"
    Add-TextCol   $cert "CertificateNo"  "Certificate No"  $true
    Add-ChoiceCol $cert "Category"       "Category"        @("Equipment","Personnel","Vehicle","Other")
    Add-TextCol   $cert "ItemName"       "Item / Person"   $true
    Add-TextCol   $cert "EquipmentID"    "Equipment ID"
    Add-ChoiceCol $cert "CertifyingBody" "Certifying Body" @("TUV","Bureau Veritas","SGS","DNV","Other")
    Add-DateCol   $cert "IssueDate"      "Issue Date"
    Add-DateCol   $cert "ExpiryDate"     "Expiry Date"     $true
    Add-ChoiceCol $cert "RenewalStatus"  "Renewal Status"  @("Valid","Under Renewal","Expired")
    Set-PnPField -List $cert -Identity "RenewalStatus" -Values @{ CustomFormatter = $StatusJson } | Out-Null
    Set-PnPField -List $cert -Identity "ExpiryDate"    -Values @{ CustomFormatter = $ExpiryJson } | Out-Null
    $certFields = @("Title","CertificateNo","Category","ItemName","CertifyingBody","ExpiryDate","RenewalStatus")
    Add-ExpiryViews $cert $certFields

    # ---- مكتبة 4: سجلات الموظفين ----
    $pers = "Personnel Safety Records"
    New-HseLibrary $pers "CVs / Personnel Safety Records - سجلات وتأهيل الأفراد"
    Add-TextCol   $pers "EmployeeName"       "Employee Name"  $true
    Add-TextCol   $pers "JobTitle"           "Job Title"
    Add-ChoiceCol $pers "Company"            "Company"        @("Al-Mialiq","Siemens Energy","Subcontractor","Other")
    Add-TextCol   $pers "Department"         "Department"
    Add-ChoiceCol $pers "Qualifications"     "Qualifications" @("NEBOSH IGC","IOSH","OSHA","First Aid & CPR","Fire Watcher","Rigger L1","Rigger L2","Rigger L3","Confined Space Attendant","PTW Holder","Other") $false $true
    Add-DateCol   $pers "QualificationExpiry" "Qualification Expiry"

    # ---- مكتبة 5: ملفات السلامة العامة ----
    $gen = "General Safety Files"
    New-HseLibrary $gen "General Safety Files - سياسات، حوادث، خطط طوارئ"
    Add-TextCol   $gen "DocumentNo"   "Document No"
    Add-ChoiceCol $gen "DocumentType" "Document Type" @("Safety Policy","Incident Report","Near Miss","Emergency Plan","Risk Assessment","Method Statement","Safe Work Procedure","Form-Template","Other")
    Add-TextCol   $gen "Version"      "Version"
    Add-DateCol   $gen "EffectiveDate" "Effective Date"
    Add-DateCol   $gen "ReviewDate"   "Review Date"
    Add-PersonCol $gen "DocOwner"     "Document Owner"
    Add-PnPView -List $gen -Title "Due for Review" -Fields @("Title","DocumentNo","DocumentType","Version","ReviewDate") -Query "<Where><Leq><FieldRef Name='ReviewDate'/><Value Type='DateTime'><Today/></Value></Leq></Where>" | Out-Null

    Write-Host "    ✔ اكتملت المكتبات الخمس والأعمدة والـ Views والتنسيقات" -ForegroundColor Green

    # ---------- الخطوة 4: الصلاحيات ----------
    Write-Host "[4] إضافة المالك والأعضاء..." -ForegroundColor Green
    $ownersGroup  = "$SiteTitle Owners"
    $membersGroup = "$SiteTitle Members"

    function Try-AddUser($email, $groupName, $roleDesc) {
        $ok = $false
        foreach ($login in @($email, "i:0#.f|membership|$email")) {
            try { Add-PnPUserToGroup -LoginName $login -Identity $groupName -ErrorAction Stop | Out-Null; $ok = $true; break } catch {}
        }
        if ($ok) { Write-Host "    ✔ $email → $groupName ($roleDesc)" -ForegroundColor Green }
        else     { Write-Host "    ⚠ تعذرت إضافة ضيف: $email — أضِفه يدوياً من: ⚙ Site permissions ← Add members" -ForegroundColor Yellow }
    }

    Try-AddUser $SiteOwner $ownersGroup "Full Control"
    foreach ($m in $SiteMembers) { Try-AddUser $m $membersGroup "Edit" }

    # ---------- الخلاصة ----------
    Write-Host "`n=====================================================" -ForegroundColor Cyan
    Write-Host "  ✅ اكتمل الإنشاء الآلي!" -ForegroundColor Cyan
    Write-Host "  رابط الموقع: $SiteUrl" -ForegroundColor Cyan
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "`nالخطوات اليدوية المتبقية (~15 دقيقة - مشروحة في الدليل HSE-SharePoint-Setup-Guide.md):" -ForegroundColor Yellow
    Write-Host "  1) ضيوف لم تُضف أسماؤهم: أضفهم من ⚙ ← Site permissions ← Add members" -ForegroundColor Yellow
    Write-Host "  2) قفل المشاركة: ⚙ ← Site permissions ← Site sharing settings ← Only site owners can share" -ForegroundColor Yellow
    Write-Host "  3) تدفق Power Automate للتنبيهات الأسبوعية (القسم الثالث من الدليل)" -ForegroundColor Yellow
    Write-Host "  4) تصميم الصفحة الرئيسية (القسم الرابع من الدليل)" -ForegroundColor Yellow
    Write-Host "  5) الإعدادات الإقليمية: Site Settings ← Regional settings ← Gregorian`n" -ForegroundColor Yellow
}
catch {
    Write-Host "`n❌ خطأ: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "تأكد من: 1) اسم المستأجر صحيح  2) لديك صلاحية إنشاء المواقع/SharePoint Admin  3) تسجيل تطبيق PnP (الخطوة 0)" -ForegroundColor Red
}
