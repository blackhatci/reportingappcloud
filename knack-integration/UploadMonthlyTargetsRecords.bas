Attribute VB_Name = "UploadMonthlyTargetsRecords"
Option Explicit

'=========================================================
' UPLOAD NEXT MONTH'S TARGETS - direct Knack API replacement for
' UploadMOnthlyTargetRecords (Button5 "Upload Next Months
' Targets"), which mutated the live MonthlyTargets sheet in place
' (overwriting Target Month/Year forward, restoring them
' afterward) and posted the result to a Zapier catch webhook.
'
' This version does NOT mutate the MonthlyTargets sheet at all. For
' every current row, it CREATES ONE NEW object_28 record for next
' month:
'
'   New record field              Source
'   Location                       same as current row (resolved)
'   Local Manager                  same as current row (resolved)
'   AFReporting_Period             VariablesSheet!E2 ("Next Target
'                                   Period", = EOMONTH(current period, 1))
'   Target Year / Target Month     derived from that same next period
'   Monthly New Member Sign Up
'     Target, and other target
'     inputs                       carried through unchanged
'   Prior Month Dues Tap           <- current row's CURRENT Month
'                                   Dues Tap (the rollover this
'                                   button exists to do)
'   Current Month Dues Tap         blank - not yet known for the
'                                   new month
'   Actual New Member Sign Ups,
'   Monthly Member Sign Up
'     Bonus/Excess, New Members
'     Sign Ups                     blank - per-period actuals,
'                                   not yet known for the new month
'
' MBCode and "Manager and MB Code" are Knack concatenation fields
' (computed automatically from Target Month/Year) and are not sent.
' Monthly Revenue Bonus / Monthly Revenue 500 Bonus are Knack
' equation fields (hardcoded to 0 server-side) and are not sent
' either - writing to a formula field has no effect in Knack.
'
' object_28 also has an Area connection (field_940). The
' MonthlyTargets sheet has no Area column to source it from, so it's
' resolved through the Location's own CURRENT Area connection via
' KnackUploadHelpers.ResolveLocationAreaId - the new record always
' gets whatever Area the Location is in right now, regardless of
' what Area the outgoing record was tagged with. If that lookup
' doesn't resolve (e.g. the Location record itself has no Area set),
' field_940 is simply omitted rather than blocking the row - Area is
' not a required field on object_28.
'=========================================================

Private Const MTGT_OBJECT_KEY As String = "object_28"
Private Const MTGT_SOURCE_SHEET As String = "MonthlyTargets"
Private Const MTGT_VARIABLES_SHEET As String = "VariablesSheet"

Private Const MTGT_HEADER_ROW As Long = 1
Private Const MTGT_FIRST_DATA_ROW As Long = 2

' Column positions on the MonthlyTargets sheet (fixed layout).
Private Const COL_LOCATION As Long = 2
Private Const COL_LOCAL_MANAGER As Long = 3
Private Const COL_TARGET As Long = 5
Private Const COL_CURRENT_MONTH_DUES_TAP As Long = 7

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub UploadNextMonthsTargets()

    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Dim wsVariables As Worksheet

    Set wsSource = ThisWorkbook.Worksheets(MTGT_SOURCE_SHEET)
    Set wsVariables = ThisWorkbook.Worksheets(MTGT_VARIABLES_SHEET)

    Dim nextPeriodValue As Variant
    nextPeriodValue = wsVariables.Range("E2").value

    If Not IsDate(nextPeriodValue) Then
        MsgBox "VariablesSheet!E2 (Next Target Period) does not contain " & _
            "a valid date.", vbCritical, "Upload Next Months Targets"
        Exit Sub
    End If

    Dim nextPeriod As Date
    nextPeriod = DateValue(CDate(nextPeriodValue))

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.rows.count, 1).End(xlUp).row

    If lastRow < MTGT_FIRST_DATA_ROW Then
        MsgBox "No rows found on the '" & MTGT_SOURCE_SHEET & "' sheet to roll forward.", _
            vbExclamation, "Upload Next Months Targets"
        Exit Sub
    End If

    Dim confirmResp As VbMsgBoxResult
    confirmResp = MsgBox( _
        "Create " & (lastRow - MTGT_FIRST_DATA_ROW + 1) & _
        " Monthly Target record(s) for " & Format$(nextPeriod, "mmmm yyyy") & _
        "?" & vbCrLf & vbCrLf & _
        "This does not modify the current MonthlyTargets sheet - it only " & _
        "creates new records in Knack for next month.", _
        vbYesNo, "Upload Next Months Targets")

    If confirmResp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    Dim cache As Object
    Set cache = KnackUploadHelpers.NewLookupCache()

    Dim nextPeriodId As String
    nextPeriodId = KnackUploadHelpers.ResolvePeriodId(nextPeriod, cache)

    Dim result As Object
    Set result = KnackUploadHelpers.NewUploadResult()

    If Len(nextPeriodId) = 0 Then

        Application.ScreenUpdating = True

        MsgBox "Could not find an AFReporting_Period record for " & _
            Format$(nextPeriod, "mm/dd/yyyy") & ". Create that period in " & _
            "Knack first, then try again.", _
            vbCritical, "Upload Next Months Targets"

        Exit Sub

    End If

    Dim rowNumber As Long

    For rowNumber = MTGT_FIRST_DATA_ROW To lastRow

        UploadOneTargetRow wsSource, rowNumber, nextPeriod, nextPeriodId, cache, result

    Next rowNumber

    Application.ScreenUpdating = True

    KnackUploadHelpers.ShowUploadSummary _
        "Upload Next Months Targets - Complete", result, _
        "New period: " & Format$(nextPeriod, "mmmm yyyy")

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "The Monthly Targets rollover could not be completed." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Upload Next Months Targets"

End Sub

'=========================================================
' PER-ROW UPLOAD
'=========================================================

Private Sub UploadOneTargetRow( _
    ByVal wsSource As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal nextPeriod As Date, _
    ByVal nextPeriodId As String, _
    ByVal cache As Object, _
    ByVal result As Object)

    On Error GoTo RowFailed

    Dim locationName As String
    Dim managerName As String

    locationName = Trim$(CStr(wsSource.Cells(rowNumber, COL_LOCATION).value))
    managerName = Trim$(CStr(wsSource.Cells(rowNumber, COL_LOCAL_MANAGER).value))

    If Len(locationName) = 0 And Len(managerName) = 0 Then
        ' Blank row - nothing to roll forward.
        Exit Sub
    End If

    If Len(locationName) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, "Location is blank."
        Exit Sub
    End If

    If Len(managerName) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, "Local Manager is blank."
        Exit Sub
    End If

    Dim locationId As String
    locationId = KnackUploadHelpers.ResolveLocationId(locationName, cache)

    If Len(locationId) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Could not find a Location record matching '" & locationName & "'."
        Exit Sub
    End If

    Dim managerId As String
    managerId = KnackUploadHelpers.ResolveManagerId(managerName, cache)

    If Len(managerId) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Could not find a Local Manager record matching '" & managerName & "'."
        Exit Sub
    End If

    ' The rollover this button exists to do: this month's CURRENT
    ' Dues Tap becomes next month's PRIOR Dues Tap.
    Dim priorMonthDuesTap As Variant
    priorMonthDuesTap = wsSource.Cells(rowNumber, COL_CURRENT_MONTH_DUES_TAP).value

    ' Area isn't on the MonthlyTargets sheet, so it's resolved
    ' through the Location's own current Area connection. A miss
    ' here doesn't skip the row - Area is optional on object_28 and
    ' everything else about the record is still valid.
    Dim areaId As String
    areaId = KnackUploadHelpers.ResolveLocationAreaId(locationName, cache)

    ' Built with AppendJsonField/BuildJsonObject (not raw string
    ' concatenation) specifically because field_940 (Area) is
    ' optional and must be omitted entirely when unresolved - a
    ' blank value spliced into a hand-built string would produce
    ' invalid JSON ("field_940":,).
    Dim f As Collection
    Set f = KnackUploadHelpers.NewJsonFragments()

    KnackUploadHelpers.AppendJsonField f, "field_368", KnackUploadHelpers.ConnectionJsonValue(locationId)
    KnackUploadHelpers.AppendJsonField f, "field_364", KnackUploadHelpers.ConnectionJsonValue(managerId)
    KnackUploadHelpers.AppendJsonField f, "field_367", KnackUploadHelpers.ConnectionJsonValue(nextPeriodId)
    KnackUploadHelpers.AppendJsonField f, "field_940", KnackUploadHelpers.ConnectionJsonValue(areaId)
    KnackUploadHelpers.AppendJsonField f, "field_366", KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_TARGET).value)
    KnackUploadHelpers.AppendJsonField f, "field_365", KnackUploadHelpers.JsonNumber(priorMonthDuesTap)
    KnackUploadHelpers.AppendJsonField f, "field_390", "0"
    KnackUploadHelpers.AppendJsonField f, "field_387", KnackUploadHelpers.JsonString(Year(nextPeriod))
    KnackUploadHelpers.AppendJsonField f, "field_385", KnackUploadHelpers.JsonString(Format$(nextPeriod, "mm"))
    KnackUploadHelpers.AppendJsonField f, "field_393", "0"
    KnackUploadHelpers.AppendJsonField f, "field_394", "0"
    KnackUploadHelpers.AppendJsonField f, "field_395", "0"
    KnackUploadHelpers.AppendJsonField f, "field_478", "0"

    Dim recordJson As String
    recordJson = KnackUploadHelpers.BuildJsonObject(f)

    KnackAPI.KnackCreateRecord MTGT_OBJECT_KEY, recordJson

    KnackUploadHelpers.RecordSuccess result

    Exit Sub

RowFailed:
    KnackUploadHelpers.RecordFailure result, rowNumber, Err.Description

End Sub
