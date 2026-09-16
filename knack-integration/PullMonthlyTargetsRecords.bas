Attribute VB_Name = "PullMonthlyTargetsRecords"
Option Explicit

'=========================================================
' MONTHLY TARGETS CONFIGURATION
'
' Template: PullHoursRecords.bas (same paging / connected-
' period-filter / mapped-write pipeline).
'
' Difference from that template: object_28 has no plain date
' field to filter on Knack's side (only the "AFReporting_Period"
' connection), so every record is paged in and then filtered in
' VBA to the period in VariablesSheet!A2 - the same connected-
' field technique PullHoursRecords/PullLeadsRecords already use
' for their *second* filter pass.
'
' MonthlyTargets (codename Sheet4) is a live sheet that
' macroUploadMonthlyTargets.bas reads and writes directly
' (including the "Next Target Period" column and its own S/T
' staging columns). To avoid clobbering that workflow, this pull
' only ever writes into the columns it recognizes by header name
' (the Knack-backed columns) and never touches anything to their
' right - "Next Target Period" and any staging columns are left
' completely alone.
'=========================================================

Private Const MTGT_OBJECT_KEY As String = "object_28"

' Connected reporting-period field, matched against VariablesSheet!A2.
Private Const MTGT_FIELD_REPORTING_PERIOD As String = "field_367"

Private Const MTGT_VARIABLES_SHEET As String = "VariablesSheet"
Private Const MTGT_OUTPUT_SHEET As String = "MonthlyTargets"

Private Const MTGT_HEADER_ROW As Long = 1
Private Const MTGT_FIRST_DATA_ROW As Long = 2

Private Const MTGT_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullMonthlyTargets()

    On Error GoTo ErrorHandler

    Dim wsVariables As Worksheet
    Dim wsOutput As Worksheet

    Set wsVariables = ThisWorkbook.Worksheets(MTGT_VARIABLES_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(MTGT_OUTPUT_SHEET)

    Dim targetPeriod As Date

    targetPeriod = MtgtRequireDate( _
        wsVariables.Range("A2").value, _
        MTGT_VARIABLES_SHEET & "!A2" _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim records As Collection

    Set records = MtgtPullAllRecords()

    ' Apply the connected reporting-period filter in VBA -
    ' object_28 has no plain date field to filter on Knack's side.
    Set records = MtgtFilterByConnectedPeriod(records, targetPeriod)

    MtgtWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The Monthly Targets pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Monthly Targets Pull"

    Resume CleanExit

End Sub

'=========================================================
' PULL ALL KNACK PAGES (no server-side filter available)
'=========================================================

Private Function MtgtPullAllRecords() As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.KnackGetRecords( _
            MTGT_OBJECT_KEY, _
            page, _
            MTGT_ROWS_PER_PAGE _
        )

        Dim parsed As Object
        Set parsed = JsonConverter.ParseJson(responseText)

        If parsed.Exists("total_pages") Then
            totalPages = CLng(parsed("total_pages"))
        Else
            totalPages = 1
        End If

        If parsed.Exists("records") Then

            Dim rec As Variant

            For Each rec In parsed("records")
                allRecords.Add rec
            Next rec

        End If

        page = page + 1

    Loop

    Set MtgtPullAllRecords = allRecords

End Function

'=========================================================
' VBA-SIDE CONNECTED REPORTING-PERIOD FILTER
'=========================================================

Private Function MtgtFilterByConnectedPeriod( _
    ByVal sourceRecords As Collection, _
    ByVal targetPeriod As Date) As Collection

    Dim filteredRecords As New Collection

    If sourceRecords Is Nothing Then
        Set MtgtFilterByConnectedPeriod = filteredRecords
        Exit Function
    End If

    Dim requiredDate As Date
    requiredDate = DateValue(targetPeriod)

    Dim i As Long

    For i = 1 To sourceRecords.count

        Dim rec As Object
        Set rec = sourceRecords(i)

        Dim connectionText As String

        connectionText = ProcessConnRecords.GetConnValue( _
            rec, _
            MTGT_FIELD_REPORTING_PERIOD _
        )

        Dim connectedDate As Date

        If MtgtTryExtractDate(connectionText, connectedDate) Then

            If DateValue(connectedDate) = requiredDate Then
                filteredRecords.Add rec
            End If

        End If

    Next i

    Set MtgtFilterByConnectedPeriod = filteredRecords

End Function

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'
' Keyed by the exact header text already on the MonthlyTargets
' sheet, so the sheet's current column order drives everything -
' nothing here depends on column position.
'=========================================================

Private Function MtgtFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "ID", "field_363"
    map.Add "Location", "field_368"
    map.Add "Local Manager", "field_364"
    map.Add "AFReporting_Period", "field_367"
    map.Add "Monthly New Member Sign Up Target", "field_366"
    map.Add "Prior Month Dues Tap", "field_365"
    map.Add "Current Month Dues Tap", "field_390"
    map.Add "Monthly Revenue Bonus", "field_389"
    map.Add "Monthly Revenue 500 Bonus", "field_392"
    map.Add "Target Year", "field_387"
    map.Add "Target Month", "field_385"
    map.Add "MBCode", "field_384"
    map.Add "Actual New Member Sign Ups", "field_393"
    map.Add "Monthly Member Sign Up Bonus", "field_394"
    map.Add "Monthly Member Sign Up Excess", "field_395"
    map.Add "Manager and MB Code", "field_397"
    map.Add "New Members Sign Ups", "field_478"

    ' "Next Target Period" is intentionally NOT mapped - it is
    ' computed and written by macroUploadMonthlyTargets.bas and
    ' must survive this pull untouched.

    Set MtgtFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO MonthlyTargets, PRESERVING COLUMN ORDER
'=========================================================

Private Sub MtgtWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        MTGT_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3401, _
            "MtgtWriteMappedRecords", _
            "No header row was found on '" & MTGT_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = MtgtFieldKeyMap()

    ' Resolve each existing column, in its existing left-to-right
    ' order, to a Knack field key. Columns whose header isn't in
    ' fieldMap (e.g. "Next Target Period") are left as 0 and are
    ' never written to or cleared.
    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(MTGT_HEADER_ROW, columnNumber).value _
        ))

        If Len(headerName) > 0 Then
            If fieldMap.Exists(headerName) Then
                columnFieldKeys(columnNumber) = fieldMap(headerName)
                mappedColumnCount = mappedColumnCount + 1
            End If
        End If

    Next columnNumber

    If mappedColumnCount = 0 Then

        Err.Raise _
            vbObjectError + 3402, _
            "MtgtWriteMappedRecords", _
            "None of the headers on '" & MTGT_OUTPUT_SHEET & _
            "' matched a known Monthly Targets field."

    End If

    Dim recordCount As Long

    If records Is Nothing Then
        recordCount = 0
    Else
        recordCount = records.count
    End If

    ' Clear only the mapped columns, down to the larger of the new
    ' record count or whatever data previously occupied the sheet -
    ' "Next Target Period" and any columns past it are never cleared.
    Dim previousLastRow As Long

    previousLastRow = wsOutput.Cells( _
        wsOutput.rows.count, 1 _
    ).End(xlUp).row

    Dim clearThroughRow As Long

    clearThroughRow = previousLastRow

    If MTGT_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = MTGT_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= MTGT_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(MTGT_FIRST_DATA_ROW, columnNumber), _
                    wsOutput.Cells(clearThroughRow, columnNumber) _
                ).ClearContents

            End If

        Next columnNumber

    End If

    If recordCount = 0 Then Exit Sub

    Dim outputValues() As Variant
    ReDim outputValues(1 To recordCount, 1 To lastColumn)

    Dim recordNumber As Long

    For recordNumber = 1 To recordCount

        Dim rec As Object
        Set rec = records(recordNumber)

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                outputValues(recordNumber, columnNumber) = _
                    MtgtGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    ' Write column-by-column so unmapped columns (e.g. "Next
    ' Target Period") are never touched by this bulk write.
    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(MTGT_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

    MtgtFormatOutput wsOutput, recordCount

End Sub

'=========================================================
' FIELD VALUE PROCESSING
'=========================================================

Private Function MtgtGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        MtgtGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            MtgtGetMappedFieldValue = CStr(rec("id"))
        Else
            MtgtGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        MtgtGetMappedFieldValue = vbNullString
        Exit Function
    End If

    MtgtGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function

'=========================================================
' CONNECTED DATE PARSING
'=========================================================

Private Function MtgtTryExtractDate( _
    ByVal connectionText As String, _
    ByRef resultDate As Date) As Boolean

    Dim cleanedText As String

    cleanedText = Trim$(connectionText)

    If Len(cleanedText) = 0 Then Exit Function

    If IsDate(cleanedText) Then

        resultDate = CDate(cleanedText)
        MtgtTryExtractDate = True
        Exit Function

    End If

    cleanedText = Replace(cleanedText, Chr$(160), " ")
    cleanedText = Replace(cleanedText, ",", " ")
    cleanedText = Replace(cleanedText, ";", " ")
    cleanedText = Replace(cleanedText, "|", " ")
    cleanedText = Replace(cleanedText, vbCr, " ")
    cleanedText = Replace(cleanedText, vbLf, " ")

    Dim pieces() As String
    pieces = Split(cleanedText, " ")

    Dim i As Long

    For i = LBound(pieces) To UBound(pieces)

        Dim candidate As String
        candidate = Trim$(pieces(i))

        If Len(candidate) > 0 Then

            If IsDate(candidate) Then

                resultDate = CDate(candidate)
                MtgtTryExtractDate = True
                Exit Function

            End If

        End If

    Next i

End Function

'=========================================================
' OUTPUT FORMATTING (mirrors HoursFormatOutput's intent -
' only reformats the columns this macro actually wrote)
'=========================================================

Private Sub MtgtFormatOutput( _
    ByVal wsOutput As Worksheet, _
    ByVal recordCount As Long)

    Dim firstDataRow As Long
    Dim finalDataRow As Long

    firstDataRow = MTGT_FIRST_DATA_ROW
    finalDataRow = firstDataRow + recordCount - 1

    Dim headerRange As Range
    Set headerRange = wsOutput.Range("A1").CurrentRegion.Rows(1)

    Dim moneyColumns As Variant
    moneyColumns = Array( _
        "Current Month Dues Tap", "Prior Month Dues Tap", _
        "Monthly Revenue Bonus", "Monthly Revenue 500 Bonus" _
    )

    Dim c As Variant

    For Each c In moneyColumns

        Dim matchCell As Range
        Set matchCell = headerRange.Find( _
            what:=c, LookIn:=xlValues, lookat:=xlWhole _
        )

        If Not matchCell Is Nothing Then

            wsOutput.Range( _
                wsOutput.Cells(firstDataRow, matchCell.column), _
                wsOutput.Cells(finalDataRow, matchCell.column) _
            ).NumberFormat = "$#,##0.00"

        End If

    Next c

    Dim dateMatch As Range
    Set dateMatch = headerRange.Find( _
        what:="AFReporting_Period", LookIn:=xlValues, lookat:=xlWhole _
    )

    If Not dateMatch Is Nothing Then

        wsOutput.Range( _
            wsOutput.Cells(firstDataRow, dateMatch.column), _
            wsOutput.Cells(finalDataRow, dateMatch.column) _
        ).NumberFormat = "mm/dd/yyyy"

    End If

End Sub

'=========================================================
' DATE VALIDATION
'=========================================================

Private Function MtgtRequireDate( _
    ByVal cellValue As Variant, _
    ByVal cellDescription As String) As Date

    If IsError(cellValue) Then

        Err.Raise _
            vbObjectError + 3403, _
            "MtgtRequireDate", _
            cellDescription & " contains an Excel error."

    End If

    If Not IsDate(cellValue) Then

        Err.Raise _
            vbObjectError + 3404, _
            "MtgtRequireDate", _
            cellDescription & " does not contain a valid date."

    End If

    MtgtRequireDate = DateValue(CDate(cellValue))

End Function
