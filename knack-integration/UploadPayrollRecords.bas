Attribute VB_Name = "UploadPayrollRecords"
Option Explicit

'=========================================================
' UPLOAD PAYROLL - direct Knack API replacement for
' macroUploadPayroll (Button3 "Upload Payroll"), which posted the
' Payroll_Sheet_ACTIVE rows to a Zapier catch webhook.
'
' Field mapping confirmed against a live object_44 record (via
' Zapier's Knack Find Record action) matched by name against every
' column on Payroll_Sheet_ACTIVE. 66 of the sheet's 89 columns have
' a live matching Knack field and are uploaded; the other columns
' listed in PAYROLL_SKIPPED_COLUMNS below have NO matching field on
' object_44 at all (confirmed against the object's live field list)
' and are not sent - same as the prior Zapier mapping must already
' have been doing, since Knack has nowhere to put them. If any of
' these are actually supposed to go somewhere, they need a field
' added to object_44 (or a different target object) before they can
' be uploaded - tell me and I'll wire it in.
'
' This always CREATES a new record per row (matching the original
' webhook's plain POST behavior) - it does not search for or update
' an existing record.
'=========================================================

Private Const PR_OBJECT_KEY As String = "object_44"
Private Const PR_SOURCE_SHEET As String = "Payroll_Sheet_ACTIVE"

Private Const PR_HEADER_ROW As Long = 3
Private Const PR_FIRST_DATA_ROW As Long = 4

' Columns with no matching Knack field on object_44 - not uploaded.
Private Const PAYROLL_SKIPPED_COLUMNS As String = _
    "pt, mgrhrsw, hrswagetype, Not Used, Salary_Pay2, Bonus Paid, " & _
    "Part-Time, DG, DG+, QtrBonusQualify, PriorQtrRev, " & _
    "PriorQtrYearRev, L1, L2, L3, CO OT, Transfers, PR"

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub UploadPayrollReport()

    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(PR_SOURCE_SHEET)

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row

    If lastRow < PR_FIRST_DATA_ROW Then
        MsgBox "No rows found on the '" & PR_SOURCE_SHEET & "' sheet to upload.", _
            vbExclamation, "Upload Payroll"
        Exit Sub
    End If

    Dim confirmResp As VbMsgBoxResult
    confirmResp = MsgBox( _
        "Confirm uploading " & (lastRow - PR_FIRST_DATA_ROW + 1) & _
        " Full Payroll Report record(s) to Knack." & vbCrLf & vbCrLf & _
        "Note: the following columns have no matching Knack field " & _
        "and will NOT be uploaded:" & vbCrLf & PAYROLL_SKIPPED_COLUMNS, _
        vbYesNo, "Upload Payroll")

    If confirmResp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    Dim cache As Object
    Set cache = KnackUploadHelpers.NewLookupCache()

    Dim result As Object
    Set result = KnackUploadHelpers.NewUploadResult()

    Dim rowNumber As Long

    For rowNumber = PR_FIRST_DATA_ROW To lastRow

        UploadOnePayrollRow wsSource, rowNumber, cache, result

    Next rowNumber

    Application.ScreenUpdating = True

    KnackUploadHelpers.ShowUploadSummary "Upload Payroll - Complete", result, _
        "Columns not uploaded (no matching Knack field): " & vbCrLf & PAYROLL_SKIPPED_COLUMNS

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "The Payroll upload could not be completed." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Upload Payroll"

End Sub

'=========================================================
' PER-ROW UPLOAD
'=========================================================

Private Sub UploadOnePayrollRow( _
    ByVal wsSource As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal cache As Object, _
    ByVal result As Object)

    On Error GoTo RowFailed

    Dim managerName As String
    managerName = Trim$(CStr(wsSource.Cells(rowNumber, 1).value))

    If Len(managerName) = 0 Then
        ' Blank row.
        Exit Sub
    End If

    Dim locationName As String
    locationName = Trim$(CStr(wsSource.Cells(rowNumber, 2).value))

    Dim areaNumber As String
    areaNumber = Trim$(CStr(wsSource.Cells(rowNumber, 3).value))

    Dim payPeriodValue As Variant
    payPeriodValue = wsSource.Cells(rowNumber, 71).value ' Pay_Period

    Dim managerId As String
    managerId = KnackUploadHelpers.ResolveManagerId(managerName, cache)

    If Len(managerId) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Could not find a Local Manager record matching '" & managerName & "'."
        Exit Sub
    End If

    Dim locationId As String
    If Len(locationName) > 0 Then
        locationId = KnackUploadHelpers.ResolveLocationId(locationName, cache)
        If Len(locationId) = 0 Then
            KnackUploadHelpers.RecordSkip result, rowNumber, _
                "Could not find a Location record matching '" & locationName & "'."
            Exit Sub
        End If
    End If

    Dim areaId As String
    If Len(areaNumber) > 0 Then
        areaId = KnackUploadHelpers.ResolveAreaId(areaNumber, cache)
        ' Area is optional context - if it doesn't resolve, upload
        ' continues without it rather than skipping the whole row.
    End If

    Dim periodId As String
    If IsDate(payPeriodValue) Then
        periodId = KnackUploadHelpers.ResolvePeriodId(CDate(payPeriodValue), cache)
        If Len(periodId) = 0 Then
            KnackUploadHelpers.RecordSkip result, rowNumber, _
                "Could not find an AFReporting_Period record matching " & _
                Format$(CDate(payPeriodValue), "mm/dd/yyyy") & "."
            Exit Sub
        End If
    Else
        KnackUploadHelpers.RecordSkip result, rowNumber, "Pay_Period is not a valid date."
        Exit Sub
    End If

    Dim f As Collection
    Set f = KnackUploadHelpers.NewJsonFragments()

    ' Note: KnackUploadHelpers is a standard module, not a class -
    ' its members are called fully qualified below (you cannot use
    ' "With KnackUploadHelpers" on a standard module in VBA).

    KnackUploadHelpers.AppendJsonField f, "field_723", KnackUploadHelpers.ConnectionJsonValue(managerId)
    KnackUploadHelpers.AppendJsonField f, "field_724", KnackUploadHelpers.ConnectionJsonValue(locationId)
    KnackUploadHelpers.AppendJsonField f, "field_722", KnackUploadHelpers.ConnectionJsonValue(areaId)
    KnackUploadHelpers.AppendJsonField f, "field_753", KnackUploadHelpers.ConnectionJsonValue(periodId)

    KnackUploadHelpers.AppendJsonField f, "field_652", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 4).value)   ' hourlywage
    KnackUploadHelpers.AppendJsonField f, "field_653", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 5).value)   ' salary
    KnackUploadHelpers.AppendJsonField f, "field_654", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 6).value)   ' nmch
    KnackUploadHelpers.AppendJsonField f, "field_655", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 7).value)   ' nmcc
    KnackUploadHelpers.AppendJsonField f, "field_656", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 8).value)   ' nmop
    KnackUploadHelpers.AppendJsonField f, "field_657", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 9).value)   ' nmca
    KnackUploadHelpers.AppendJsonField f, "field_658", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 10).value)  ' nmti
    KnackUploadHelpers.AppendJsonField f, "field_659", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 11).value)  ' losf
    KnackUploadHelpers.AppendJsonField f, "field_660", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 12).value)  ' lasf
    KnackUploadHelpers.AppendJsonField f, "field_661", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 13).value)  ' coll
    KnackUploadHelpers.AppendJsonField f, "field_662", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 14).value)  ' ppv
    KnackUploadHelpers.AppendJsonField f, "field_663", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 15).value)  ' monthlyNMBonusAmt
    KnackUploadHelpers.AppendJsonField f, "field_664", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 16).value)  ' mdg
    KnackUploadHelpers.AppendJsonField f, "field_665", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 17).value)  ' mnmplus
    KnackUploadHelpers.AppendJsonField f, "field_666", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 18).value)  ' mdg5
    KnackUploadHelpers.AppendJsonField f, "field_667", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 19).value)  ' nmco
    ' col 20 "pt" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_669", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 21).value)  ' pt_percent
    KnackUploadHelpers.AppendJsonField f, "field_974", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 22).value)  ' Cell_Reimb
    ' col 23 "mgrhrsw" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_672", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 24).value)  ' mgrhrsp
    ' col 25 "hrswagetype" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_674", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 26).value)  ' wk1hrsw
    KnackUploadHelpers.AppendJsonField f, "field_675", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 27).value)  ' wk2hrsw
    KnackUploadHelpers.AppendJsonField f, "field_676", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 28).value)  ' wk3hrsw
    KnackUploadHelpers.AppendJsonField f, "field_1050", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 29).value) ' ManualClockCount
    ' col 30 "Not Used" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_679", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 31).value)  ' otweek
    KnackUploadHelpers.AppendJsonField f, "field_680", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 32).value)  ' othrsworked
    KnackUploadHelpers.AppendJsonField f, "field_681", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 33).value)  ' COHrsDL
    KnackUploadHelpers.AppendJsonField f, "field_682", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 34).value)  ' COHrsUL
    KnackUploadHelpers.AppendJsonField f, "field_683", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 35).value)  ' otpay
    KnackUploadHelpers.AppendJsonField f, "field_684", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 36).value)  ' holidayhrs
    KnackUploadHelpers.AppendJsonField f, "field_685", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 37).value)  ' Salary_Pay
    KnackUploadHelpers.AppendJsonField f, "field_686", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 38).value)  ' mgrhrsw_Pay
    KnackUploadHelpers.AppendJsonField f, "field_687", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 39).value)  ' mgrhrsp_Pay
    KnackUploadHelpers.AppendJsonField f, "field_688", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 40).value)  ' mgrhrsOT_Pay
    KnackUploadHelpers.AppendJsonField f, "field_689", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 41).value)  ' mgrhrsh_Pay
    ' col 42 "Salary_Pay2" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_691", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 43).value)  ' PT_Revenue
    KnackUploadHelpers.AppendJsonField f, "field_692", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 44).value)  ' PT Bonus Pmt
    KnackUploadHelpers.AppendJsonField f, "field_693", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 45).value)  ' Miles
    KnackUploadHelpers.AppendJsonField f, "field_694", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 46).value)  ' Miles Reimbursement
    KnackUploadHelpers.AppendJsonField f, "field_695", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 47).value)  ' Reimbursement
    KnackUploadHelpers.AppendJsonField f, "field_696", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 48).value)  ' Other Bonus
    KnackUploadHelpers.AppendJsonField f, "field_751", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 49).value)  ' Salary Pay or Wage
    KnackUploadHelpers.AppendJsonField f, "field_698", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 50).value)  ' Other Pay
    KnackUploadHelpers.AppendJsonField f, "field_699", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 51).value)  ' Total Salary + Other
    KnackUploadHelpers.AppendJsonField f, "field_700", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 52).value)  ' BONUS SECTION
    KnackUploadHelpers.AppendJsonField f, "field_701", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 53).value)  ' Monthly Lead Goal
    KnackUploadHelpers.AppendJsonField f, "field_702", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 54).value)  ' Lead Carryover
    KnackUploadHelpers.AppendJsonField f, "field_703", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 55).value)  ' Pay Period Lead Count
    KnackUploadHelpers.AppendJsonField f, "field_704", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 56).value)  ' Lead Total
    KnackUploadHelpers.AppendJsonField f, "field_705", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 57).value)  ' Pay Period Lead Bonus
    KnackUploadHelpers.AppendJsonField f, "field_706", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 58).value)  ' Lead Bonus Monthly
    KnackUploadHelpers.AppendJsonField f, "field_707", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 59).value)  ' Lead Bonus Excess
    KnackUploadHelpers.AppendJsonField f, "field_708", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 60).value)  ' Monthly Dues Growth
    KnackUploadHelpers.AppendJsonField f, "field_709", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 61).value)  ' Monthly Dues Growth+500
    KnackUploadHelpers.AppendJsonField f, "field_711", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 62).value)  ' Delinquent Bonus
    ' col 63 "Bonus Paid" - no matching field, skipped
    KnackUploadHelpers.AppendJsonField f, "field_712", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 64).value)  ' TOTAL BONUS
    KnackUploadHelpers.AppendJsonField f, "field_752", KnackUploadHelpers.JsonString(wsSource.Cells(rowNumber, 65).value)  ' SUMMARY
    KnackUploadHelpers.AppendJsonField f, "field_714", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 66).value)  ' Wages and Salary
    KnackUploadHelpers.AppendJsonField f, "field_715", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 67).value)  ' Total Reimbursible
    KnackUploadHelpers.AppendJsonField f, "field_716", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 68).value)  ' Total Other
    KnackUploadHelpers.AppendJsonField f, "field_717", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 69).value)  ' Total Bonus4
    KnackUploadHelpers.AppendJsonField f, "field_718", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 70).value)  ' Total Pay
    ' col 71 "Pay_Period" -> field_753, already resolved above
    KnackUploadHelpers.AppendJsonField f, "field_1008", KnackUploadHelpers.JsonDateOnly(wsSource.Cells(rowNumber, 72).value) ' Eff_Wg_Date
    KnackUploadHelpers.AppendJsonField f, "field_1006", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 73).value)  ' Dues Tap Current
    KnackUploadHelpers.AppendJsonField f, "field_1007", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, 74).value)  ' Dues Tap Prior
    ' cols 75-86 (Part-Time, DG, DG+, QtrBonusQualify, PriorQtrRev,
    ' PriorQtrYearRev, L1, L2, L3, CO OT, Transfers, PR) - no
    ' matching field, skipped

    Dim recordJson As String
    recordJson = KnackUploadHelpers.BuildJsonObject(f)

    KnackAPI.KnackCreateRecord PR_OBJECT_KEY, recordJson

    KnackUploadHelpers.RecordSuccess result

    Exit Sub

RowFailed:
    KnackUploadHelpers.RecordFailure result, rowNumber, Err.Description

End Sub
