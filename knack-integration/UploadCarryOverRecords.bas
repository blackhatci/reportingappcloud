Attribute VB_Name = "UploadCarryOverRecords"
Option Explicit

'=========================================================
' UPLOAD CO TABLE - direct Knack API replacement for
' macroUploadCORecords (Button2 "Upload CO table"), which posted
' the "Upload" sheet's rows to a Zapier catch webhook.
'
' Confirmed against a live object_34 record (via Zapier's Knack
' Find Record action) that the Zap mapped this sheet 1:1 by column
' order into these fields:
'
'   Upload sheet column   -> object_34 field
'   ReportingPeriod        -> field_436 (AFReporting_Period, connection)
'   Local Manager          -> field_437 (connection)
'   OTCarryoverWeekNum     -> field_439 (COWeek)
'   OTCarryoverHours       -> field_438 (COHrsWorkedPrior)
'   NMCOMonth              -> field_479 (NMCO_Month)
'   NMCO_Qty               -> field_480
'   NMCOYear               -> field_748 (NMCOYear)
'
' This always CREATES a new record per row (matching the original
' webhook's plain POST behavior) - it does not search for or update
' an existing record.
'=========================================================

Private Const CO_OBJECT_KEY As String = "object_34"
Private Const CO_SOURCE_SHEET As String = "Upload"

Private Const CO_HEADER_ROW As Long = 1
Private Const CO_FIRST_DATA_ROW As Long = 2

' Column positions on the "Upload" sheet (fixed layout).
Private Const COL_REPORTING_PERIOD As Long = 1
Private Const COL_LOCAL_MANAGER As Long = 2
Private Const COL_OT_WEEK_NUM As Long = 3
Private Const COL_OT_HOURS As Long = 4
Private Const COL_NMCO_MONTH As Long = 5
Private Const COL_NMCO_QTY As Long = 6
Private Const COL_NMCO_YEAR As Long = 7

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub UploadCOTable()

    On Error GoTo ErrorHandler

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Worksheets(CO_SOURCE_SHEET)

    Dim lastRow As Long
    lastRow = wsSource.Cells(wsSource.rows.count, COL_REPORTING_PERIOD).End(xlUp).row

    If lastRow < CO_FIRST_DATA_ROW Then
        MsgBox "No rows found on the '" & CO_SOURCE_SHEET & "' sheet to upload.", _
            vbExclamation, "Upload CO Table"
        Exit Sub
    End If

    Dim confirmResp As VbMsgBoxResult
    confirmResp = MsgBox( _
        "Confirm upload of " & (lastRow - CO_FIRST_DATA_ROW + 1) & _
        " Carryover record(s) to Knack.", _
        vbYesNo, "Upload CO Table")

    If confirmResp <> vbYes Then Exit Sub

    Application.ScreenUpdating = False

    Dim cache As Object
    Set cache = KnackUploadHelpers.NewLookupCache()

    Dim result As Object
    Set result = KnackUploadHelpers.NewUploadResult()

    Dim rowNumber As Long

    For rowNumber = CO_FIRST_DATA_ROW To lastRow

        UploadOneCORow wsSource, rowNumber, cache, result

    Next rowNumber

    Application.ScreenUpdating = True

    KnackUploadHelpers.ShowUploadSummary "Upload CO Table - Complete", result

    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "The CO table upload could not be completed." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, "Upload CO Table"

End Sub

'=========================================================
' PER-ROW UPLOAD
'=========================================================

Private Sub UploadOneCORow( _
    ByVal wsSource As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal cache As Object, _
    ByVal result As Object)

    On Error GoTo RowFailed

    Dim periodValue As Variant
    periodValue = wsSource.Cells(rowNumber, COL_REPORTING_PERIOD).value

    If Not IsDate(periodValue) Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "ReportingPeriod is not a valid date."
        Exit Sub
    End If

    Dim managerName As String
    managerName = Trim$(CStr(wsSource.Cells(rowNumber, COL_LOCAL_MANAGER).value))

    If Len(managerName) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Local Manager is blank."
        Exit Sub
    End If

    Dim periodId As String
    periodId = KnackUploadHelpers.ResolvePeriodId(CDate(periodValue), cache)

    If Len(periodId) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Could not find an AFReporting_Period record matching " & _
            Format$(CDate(periodValue), "mm/dd/yyyy") & "."
        Exit Sub
    End If

    Dim managerId As String
    managerId = KnackUploadHelpers.ResolveManagerId(managerName, cache)

    If Len(managerId) = 0 Then
        KnackUploadHelpers.RecordSkip result, rowNumber, _
            "Could not find a Local Manager record matching '" & managerName & "'."
        Exit Sub
    End If

    Dim recordJson As String
    recordJson = "{" & _
        """field_436"":" & KnackUploadHelpers.ConnectionJsonValue(periodId) & "," & _
        """field_437"":" & KnackUploadHelpers.ConnectionJsonValue(managerId) & "," & _
        """field_439"":" & KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_OT_WEEK_NUM).value) & "," & _
        """field_438"":" & KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_OT_HOURS).value) & "," & _
        """field_479"":" & KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_NMCO_MONTH).value) & "," & _
        """field_480"":" & KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_NMCO_QTY).value) & "," & _
        """field_748"":" & KnackUploadHelpers.JsonNumber(wsSource.Cells(rowNumber, COL_NMCO_YEAR).value) & _
        "}"

    KnackAPI.KnackCreateRecord CO_OBJECT_KEY, recordJson

    KnackUploadHelpers.RecordSuccess result

    Exit Sub

RowFailed:
    KnackUploadHelpers.RecordFailure result, rowNumber, Err.Description

End Sub
