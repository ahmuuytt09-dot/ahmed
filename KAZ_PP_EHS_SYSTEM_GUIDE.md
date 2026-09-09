# KAZ PP Expansion Project — Integrated EHS Management System

**Document:** KAZ-EHS-SYS-001 Rev.00  
**Workbook:** `KAZ_PP_EHS_Management_System.xlsx`  
**VBA:** `VBA_EHS_Module.bas`  
**Default Admin password:** `KAZ-EHS-ADMIN-2026`  
**Change it immediately** on sheet **Cover** cell **C14**.

---

## 1. Worksheet structure / هيكل الشيتات

| Sheet | Purpose |
|---|---|
| **Cover** | Identity, password, map, OSHA formula statement |
| **Dashboard** | KPIs: MH without LTI, LTIFR, TRIFR, project score, counters |
| **ManHours** | Daily Seisms / Al-Mayal / Indirect → **one unified total in I5** |
| **Incidents** | Near Miss, UA, UC, First Aid, Recordable, LTI |
| **Observations** | Tracker + Frequency Count + **Overdue Days** (ages daily) + PTW |
| **NCR** | Non-conformance register (Outlook alert) |
| **Scorecard** | Automatic subcontractor EHS scores |
| **CAPA_AI** | Raw notes → official observation + CAPA |
| **Training** | Skilled labour certificates; 30-day colour warning |
| **Equipment** | Heavy plant + 3rd party certs; quarantine if expired |
| **Documents** | HSE/Security/PTW plans + EHS staff CV hyperlinks |
| **Lookups** | Dropdown source lists |
| **Admin** | Targets, SLA, score weights, Outlook addresses, API key |

**Yellow cells = INPUT (unlocked).** **Blue cells = CALCULATED (locked).**

---

## 2. OSHA / KPI formulas

```
LTIFR = (No. of LTI × 1,000,000) / Total Man-Hours
TRIFR = (Recordable cases × 1,000,000) / Total Man-Hours
Man-Hours without LTI = SUM of daily MH with Date > last LTI date
                         (if no LTI → all project MH)
```

- Total Man-Hours: `ManHours!I5`  (also `J4` / named `TotalManHours`)
- LTIFR: `Dashboard!E7`
- TRIFR: `Dashboard!G7`
- Last LTI: `Incidents!W13` or override `Admin!C31`
- Project score uses the same point model as the scorecard (Admin C15–C21).

### Observation aging (Overdue Days) — column O

If status is **Closed** → 0.  
If **Open / In Progress**:

```
MAX(0, TODAY() − Target Close)
```

If Target Close is blank:

```
TODAY() − (Date + SLA days)
SLA = Admin!C10 (default 7)  |  Critical SLA = Admin!C11 (default 1 day)
```

The value **increases by 1 every calendar day** until Closed.

### Frequency Count — column P

```
COUNTIFS(Type, Location, Subcontractor)
```

Repeat flag if count ≥ 2.

---

## 3. Daily Man-Hours (requirement 4)

On **ManHours**, each day:

| Seisms MH | `Headcount × Hrs/person` (G = E×F) |
| Al-Mayal MH | `Headcount × Hrs/person` (J = H×I) |
| Indirect | `Headcount × Admin default hours` |
| Daily total | G+J+L → column M |
| **UNIFIED PROJECT TOTAL** | **`I5 = SUM(M8:M400)`** |

Dashboard reads that single cell.

---

## 4. Subcontractor scorecard

```
Score = MAX(floor, 100
          + OpenOBS × (−2)
          + OverdueOBS × (−5)
          + NCR × (−8)
          + LTI × (−25)
          + Recordable × (−15)
          + NearMiss × (+1) )
```

Weights are editable on **Admin**. Bands A/B/C/D from Admin C36–C39.

---

## 5. CAPA / “AI” hub

Sheet **CAPA_AI**: paste raw Arabic/English notes in column **C**.  
Columns D–G **auto-convert** via Excel formulas (keyword engine: lift, scaffold, excav, confined, height, electrical, PTW).

Optional OpenAI: put API key in `Admin!C28` and run macro `EHS_FormatObservation`.

---

## 6. Training & Equipment colour rules

| Status | Rule |
|---|---|
| VALID | days to expiry > 30 |
| WARNING ≤30d | ≤ Admin!C12 (30) |
| CRITICAL ≤7d | ≤ Admin!C13 (7) |
| EXPIRED / QUARANTINE | days < 0 |

---

## 7. Install VBA (Windows Excel)

1. Open the `.xlsx` → **Save As** `KAZ_PP_EHS_Management_System.xlsm` (macro-enabled).
2. Alt+F11 → File → **Import File** → `VBA_EHS_Module.bas`.
3. Optional ThisWorkbook:

```vb
Private Sub Workbook_Open()
    modEHS_KAZ_PP.EHS_Workbook_Open
End Sub
```

4. File → Options → Trust Center → enable macros for this file.
5. Alt+F8 macros:

| Macro | Function |
|---|---|
| `EHS_ProtectAll` | Lock all sheets with Cover C14 password |
| `EHS_UnprotectAll` | Admin unlock |
| `EHS_SendCriticalAlert` | Outlook mail for an incident row |
| `EHS_SendNCRAlert` | Outlook mail for an NCR row |
| `EHS_ExportWeeklyPDF` | PDF of Dashboard+Scorecard+OBS+INC+NCR+MH |
| `EHS_FormatObservation` | Stamp CAPA conversion |
| `EHS_RefreshAging` | Recalculate overdue days |
| `EHS_NewObservationID` | Next OBS-xxx |

Outlook recipients: **Admin C23 / C24**. Flags C25/C26 = YES.

To send without preview, in `SendOutlook` change `mail.Display` → `mail.Send`.

---

## 8. Protection instructions / حماية الملف

### A. After data structure is ready

1. Review Cover **C14** password.
2. Run **`EHS_ProtectAll`**.
3. Review → Protect Workbook → Structure (same password).
4. File → Info → **Protect Workbook → Encrypt with Password** (file open password). Use a **different** password if you want open vs modify separation; otherwise same.

### B. Manual (no VBA)

- Review tab → Protect Sheet → password `KAZ-EHS-ADMIN-2026`  
  Tick: Select unlocked cells, Use AutoFilter.  
  Yellow input cells already have **Locked = False**.
- Protect Workbook → Structure.
- File → Info → Encrypt with Password.

### C. Users vs Admin

| Role | Access |
|---|---|
| Admin (you) | Full control, VBA, unprotect |
| Users | Yellow cells only; cannot edit formulas, lists, Admin, Lookups if protected |

**Never share C14.** If leaked, change C14, unprotect, re-protect all.

---

## 9. Recommended daily workflow

1. **ManHours** — enter Seisms / Al-Mayal headcount.  
2. **Observations / Incidents / NCR** — log events + PTW number.  
3. **CAPA_AI** — convert raw notes, paste into Observations.  
4. **Dashboard** — check LTIFR/TRIFR vs target.  
5. **Training / Equipment** — act on amber/red.  
6. Friday: **EHS_ExportWeeklyPDF**.  
7. Critical LTI or NCR: run Outlook macros.

---

*Confidential — KAZ PP Expansion Project EHS Department.*
