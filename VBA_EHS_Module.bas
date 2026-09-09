Attribute VB_Name = "modEHS_KAZ_PP"
'===============================================================================
' KAZ PP Expansion Project — EHS Management System
' VBA Module  |  Import: Alt+F11 → File → Import File → VBA_EHS_Module.bas
' Admin password (default): KAZ-EHS-ADMIN-2026   (Cover!C14)
' Requires: Excel for Windows + Microsoft Outlook for email macros
'===============================================================================
Option Explicit

Public Const ADMIN_PWD_DEFAULT As String = "KAZ-EHS-ADMIN-2026"

Private Function AdminPwd() As String
    On Error Resume Next
    AdminPwd = CStr(ThisWorkbook.Worksheets("Cover").Range("C14").Value)
    If Len(AdminPwd) = 0 Then AdminPwd = ADMIN_PWD_DEFAULT
End Function

Private Function AdminVal(addr As String) As Variant
    AdminVal = ThisWorkbook.Worksheets("Admin").Range(addr).Value
End Function

Private Sub UnprotectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=AdminPwd()
End Sub

Private Sub ProtectSheet(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=AdminPwd(), UserInterfaceOnly:=True, _
               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
End Sub

'----- Workbook events should live in ThisWorkbook; call these from there -----
Public Sub EHS_Workbook_Open()
    EHS_RefreshAging
    Application.Calculation = xlCalculationAutomatic
End Sub

'===============================================================================
' PROTECTION
'===============================================================================
Public Sub EHS_ProtectAll()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = AdminPwd()
    ThisWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False
    For Each ws In ThisWorkbook.Worksheets
        UnprotectSheet ws
        ProtectSheet ws
    Next ws
    MsgBox "All sheets locked. Yellow cells remain editable (unlocked)." & vbCrLf & _
           "Admin password: (Cover C14)", vbInformation, "KAZ PP EHS"
End Sub

Public Sub EHS_UnprotectAll()
    Dim ws As Worksheet
    Dim pwd As String
    pwd = InputBox("Enter Admin password to unlock all sheets:", "KAZ PP EHS")
    If pwd <> AdminPwd() Then
        MsgBox "Access denied.", vbCritical
        Exit Sub
    End If
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=pwd
    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect Password:=pwd
    Next ws
    MsgBox "Workbook fully unlocked for Admin.", vbInformation
End Sub

'===============================================================================
' AGING REFRESH (forces overdue-day recalculation)
'===============================================================================
Public Sub EHS_RefreshAging()
    Application.Calculate
    ThisWorkbook.Worksheets("Dashboard").Range("C4").Value = Date
End Sub

'===============================================================================
' FORMAT OBSERVATION + CAPA (local engine; optional OpenAI)
'===============================================================================
Public Sub EHS_FormatObservation()
    Dim ws As Worksheet
    Dim r As Long, raw As String
    Dim apiKey As String
    Set ws = ThisWorkbook.Worksheets("CAPA_AI")
    r = ActiveCell.Row
    If ws.Name <> ActiveSheet.Name Then
        ws.Activate
        r = Application.InputBox("Row number on CAPA_AI to convert:", "CAPA", 10, Type:=1)
        If r < 10 Then Exit Sub
    End If
    If r < 10 Then
        MsgBox "Select a data row (row 10+).", vbExclamation
        Exit Sub
    End If
    raw = CStr(ws.Cells(r, 3).Value)
    If Len(Trim$(raw)) = 0 Then
        MsgBox "Column C (RAW NOTES) is empty.", vbExclamation
        Exit Sub
    End If
    apiKey = CStr(AdminVal("C28"))
    UnprotectSheet ws
    If Len(Trim$(apiKey)) > 0 Then
        On Error GoTo LocalEngine
        ws.Cells(r, 4).Value = OpenAI_Complete(apiKey, _
            "Rewrite as a formal EHS observation for KAZ PP Expansion Project (one paragraph, professional English): " & raw)
        ws.Cells(r, 5).Value = OpenAI_Complete(apiKey, _
            "Provide Corrective and Preventive Actions (CAPA) numbered 1-4 for this EHS finding: " & raw)
        GoTo Done
    End If
LocalEngine:
    ' Formulas already on the sheet compute type/severity/CAPA.
    ' Stamp a timestamp so the user sees VBA ran.
    ws.Cells(r, 8).Value = "Converted " & Format(Now, "yyyy-mm-dd hh:nn")
Done:
    ws.Cells(r, 8).Value = "Converted " & Format(Now, "yyyy-mm-dd hh:nn")
    ProtectSheet ws
    MsgBox "Observation formatted on row " & r & ". Copy D/E into Observations tracker.", vbInformation
End Sub

Private Function OpenAI_Complete(apiKey As String, prompt As String) As String
    ' Late-bound WinHTTP. Optional. If it fails, caller uses local formulas.
    Dim http As Object
    Dim payload As String
    Dim model As String
    Dim resp As String
    model = CStr(AdminVal("C29"))
    If Len(model) = 0 Then model = "gpt-4o-mini"
    payload = "{""model"":""" & model & """,""messages"":[{""role"":""user"",""content"":" & JsonEscape(prompt) & "}],""temperature"":0.2}"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", "https://api.openai.com/v1/chat/completions", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.Send payload
    resp = http.ResponseText
    OpenAI_Complete = ExtractJSONContent(resp)
End Function

Private Function JsonEscape(s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbLf, "\n")
    JsonEscape = """" & t & """"
End Function

Private Function ExtractJSONContent(json As String) As String
    Dim p As Long, p2 As Long
    p = InStr(1, json, """content"":", vbTextCompare)
    If p = 0 Then
        ExtractJSONContent = ""
        Exit Function
    End If
    p = InStr(p, json, """")
    p = InStr(p + 1, json, """")
    p2 = InStr(p + 1, json, """")
    If p2 > p Then
        ExtractJSONContent = Mid$(json, p + 1, p2 - p - 1)
    Else
        ExtractJSONContent = json
    End If
End Function

'===============================================================================
' OUTLOOK — Critical incident
'===============================================================================
Public Sub EHS_SendCriticalAlert()
    Dim ws As Worksheet
    Dim r As Long
    Dim body As String
    If UCase$(CStr(AdminVal("C25"))) <> "YES" Then
        MsgBox "Admin!C25 is not YES — email suppressed.", vbInformation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("Incidents")
    r = Application.InputBox("Incident row number to email (data starts row 5):", "Critical Alert", 5, Type:=1)
    If r < 5 Then Exit Sub
    body = "KAZ PP EXPANSION PROJECT — EHS CRITICAL / INCIDENT ALERT" & vbCrLf & vbCrLf & _
           "ID: " & ws.Cells(r, 2).Value & vbCrLf & _
           "Date: " & ws.Cells(r, 3).Value & vbCrLf & _
           "Classification: " & ws.Cells(r, 5).Value & vbCrLf & _
           "Location: " & ws.Cells(r, 6).Value & vbCrLf & _
           "Subcontractor: " & ws.Cells(r, 7).Value & vbCrLf & _
           "PTW: " & ws.Cells(r, 8).Value & vbCrLf & _
           "LTI?: " & ws.Cells(r, 13).Value & vbCrLf & _
           "Description: " & ws.Cells(r, 14).Value & vbCrLf & vbCrLf & _
           "Dashboard LTIFR/TRIFR will update automatically. Immediate investigation required." & vbCrLf & _
           "This is an automated message from KAZ-EHS-SYS-001."
    SendOutlook CStr(AdminVal("C23")), CStr(AdminVal("C24")), _
                "KAZ PP EHS ALERT — " & ws.Cells(r, 2).Value & " / " & ws.Cells(r, 5).Value, body
End Sub

Public Sub EHS_SendNCRAlert()
    Dim ws As Worksheet
    Dim r As Long
    Dim body As String
    If UCase$(CStr(AdminVal("C26"))) <> "YES" Then
        MsgBox "Admin!C26 is not YES — email suppressed.", vbInformation
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("NCR")
    r = Application.InputBox("NCR row number (data starts row 5):", "NCR Alert", 5, Type:=1)
    If r < 5 Then Exit Sub
    body = "KAZ PP EXPANSION PROJECT — NEW NCR NOTIFICATION" & vbCrLf & vbCrLf & _
           "NCR: " & ws.Cells(r, 2).Value & vbCrLf & _
           "Date: " & ws.Cells(r, 3).Value & vbCrLf & _
           "Subcontractor: " & ws.Cells(r, 4).Value & vbCrLf & _
           "Priority: " & ws.Cells(r, 9).Value & vbCrLf & _
           "Finding: " & ws.Cells(r, 6).Value & vbCrLf & _
           "Required action: " & ws.Cells(r, 7).Value & vbCrLf & _
           "Target: " & ws.Cells(r, 10).Value & vbCrLf & _
           "PTW: " & ws.Cells(r, 15).Value & vbCrLf & vbCrLf & _
           "Scorecard will deduct NCR points automatically. Automated from KAZ-EHS-SYS-001."
    SendOutlook CStr(AdminVal("C23")), CStr(AdminVal("C24")), _
                "KAZ PP NCR — " & ws.Cells(r, 2).Value & " [" & ws.Cells(r, 9).Value & "]", body
End Sub

Private Sub SendOutlook(toAddr As String, ccAddr As String, subject As String, body As String)
    Dim ol As Object, mail As Object
    On Error GoTo Fail
    Set ol = CreateObject("Outlook.Application")
    Set mail = ol.CreateItem(0)
    mail.To = toAddr
    mail.CC = ccAddr
    mail.Subject = subject
    mail.Body = body
    mail.Importance = 2
    mail.Display ' Display first; change to .Send for fully automatic
    Exit Sub
Fail:
    MsgBox "Outlook not available: " & Err.Description & vbCrLf & _
           "Email would have been:" & vbCrLf & subject & vbCrLf & body, vbExclamation
End Sub

'===============================================================================
' WEEKLY PDF
'===============================================================================
Public Sub EHS_ExportWeeklyPDF()
    Dim folder As String
    Dim fname As String
    Dim arr As Variant
    Dim i As Long
    Dim pwd As String
    pwd = AdminPwd()
    folder = CStr(AdminVal("C27"))
    If Len(folder) = 0 Then folder = ThisWorkbook.Path
    On Error Resume Next
    MkDir folder
    On Error GoTo 0
    fname = folder & "\KAZ_PP_EHS_Weekly_" & Format(Date, "yyyymmdd") & ".pdf"
    arr = Array("Dashboard", "Scorecard", "Observations", "Incidents", "NCR", "ManHours")
    ' Unprotect temporarily for export
    For i = LBound(arr) To UBound(arr)
        UnprotectSheet ThisWorkbook.Worksheets(CStr(arr(i)))
    Next i
    ThisWorkbook.Worksheets(arr).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fname, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    ThisWorkbook.Worksheets("Dashboard").Select
    For i = LBound(arr) To UBound(arr)
        ProtectSheet ThisWorkbook.Worksheets(CStr(arr(i)))
    Next i
    MsgBox "Weekly PDF exported:" & vbCrLf & fname, vbInformation, "KAZ PP EHS"
End Sub

'===============================================================================
' NEW OBSERVATION ID
'===============================================================================
Public Sub EHS_NewObservationID()
    Dim ws As Worksheet
    Dim last As Long, r As Long, n As Long
    Set ws = ThisWorkbook.Worksheets("Observations")
    UnprotectSheet ws
    last = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    r = last + 1
    If r < 5 Then r = 5
    n = Application.WorksheetFunction.CountA(ws.Range("B5:B250")) + 1
    ws.Cells(r, 2).Value = "OBS-" & Format(n, "000")
    ws.Cells(r, 3).Value = Date
    ws.Cells(r, 6).Value = "Open"
    ws.Cells(r, r).Select
    ProtectSheet ws
End Sub
)
