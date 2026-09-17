Attribute VB_Name = "PullCarryOverRecords"
Option Explicit

'=========================================================
' OT CARRYOVER CONFIGURATION
'
' Template: PullMonthlyTargetsRecords.bas (itself built from
' PullHoursRecords.bas). Pulls Knack object_34 ("CarryOver")
' into the existing OTCarryOver sheet (codename Sheet14),
' preserving whatever column order is already on that sheet.
'
' object_34 has no plain date field - only a connection to
' AFReporting_Period - so every record is paged in and then
' filtered in VBA to the period in VariablesSheet!B2 (the PRIOR
' period - carryover records are tagged to the period the hours
' carried over FROM, not the current period in A2).
'=========================================================

Private Const CO_OBJECT_KEY As String = "object_34"

' Connected reporting-period field, matched against VariablesSheet!B2
' (the prior period).
Private Const CO_FIELD_REPORTING_PERIOD As String = "field_436"

Private Const CO_VARIABLES_SHEET As String = "VariablesSheet"
Private Const CO_OUTPUT_SHEET As String = "OTCarryOver"

Private Const CO_HEADER_ROW As Long = 1
Private Const CO_FIRST_DATA_ROW As Long = 2

Private Const CO_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullCarryOver()

    On Error GoTo ErrorHandler

    Dim wsVariables As Worksheet
    Dim wsOutput As Worksheet

    Set wsVariables = ThisWorkbook.Worksheets(CO_VARIABLES_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(CO_OUTPUT_SHEET)

    ' OT Carryover records are tagged to the PRIOR period
    ' (VariablesSheet!B2), not the current period (A2) - these
    ' are hours carried over FROM the prior period into the
    ' current one.
    Dim targetPeriod As Date

    targetPeriod = CoRequireDate( _
        wsVariables.Range("B2").value, _
        CO_VARIABLES_SHEET & "!B2" _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim records As Collection

    Set records = CoPullAllRecords()

    Dim rawCount As Long
    rawCount = records.count

    Set records = CoFilterByConnectedPeriod(records, targetPeriod)

    If records.count = 0 Then

        Application.EnableEvents = True
        Application.ScreenUpdating = True

        MsgBox _
            CoDiagnoseEmptyResult(rawCount, targetPeriod), _
            vbExclamation, _
            "OT Carryover Pull - No Matching Records"

        Exit Sub

    End If

    CoWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The OT Carryover pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "OT Carryover Pull"

    Resume CleanExit

End Sub

'=========================================================
' PULL ALL KNACK PAGES (no server-side filter available)
'=========================================================

Private Function CoPullAllRecords() As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.KnackGetRecords( _
            CO_OBJECT_KEY, _
            page, _
            CO_ROWS_PER_PAGE _
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

    Set CoPullAllRecords = allRecords

End Function

'=========================================================
' VBA-SIDE CONNECTED REPORTING-PERIOD FILTER
'=========================================================

Private Function CoFilterByConnectedPeriod( _
    ByVal sourceRecords As Collection, _
    ByVal targetPeriod As Date) As Collection

    Dim filteredRecords As New Collection

    If sourceRecords Is Nothing Then
        Set CoFilterByConnectedPeriod = filteredRecords
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
            CO_FIELD_REPORTING_PERIOD _
        )

        Dim connectedDate As Date

        If CoTryExtractDate(connectionText, connectedDate) Then

            If DateValue(connectedDate) = requiredDate Then
                filteredRecords.Add rec
            End If

        End If

    Next i

    Set CoFilterByConnectedPeriod = filteredRecords

End Function

'=========================================================
' SELF-DIAGNOSIS - builds a message explaining exactly why zero
' records matched, using real data from the actual pull, so this
' doesn't require a separate manual debug pass.
'=========================================================

Private Function CoDiagnoseEmptyResult( _
    ByVal rawCount As Long, _
    ByVal targetPeriod As Date) As String

    Dim msg As String

    msg = "Target period (VariablesSheet!B2, the prior period): " & _
        Format$(targetPeriod, "mm/dd/yyyy") & vbCrLf & _
        "Raw object_34 records pulled: " & rawCount & vbCrLf & vbCrLf

    If rawCount = 0 Then

        msg = msg & _
            "The Knack pull itself returned zero records for " & _
            "object_34 (CarryOver). This is not a filtering issue - " & _
            "either the table is genuinely empty right now, or the " & _
            "API call itself is failing silently. Check the " & _
            "CarryOver table directly in the Knack Builder to " & _
            "confirm it has records."

        CoDiagnoseEmptyResult = msg
        Exit Function

    End If

    msg = msg & _
        "Records were pulled, but none matched the target period " & _
        "after reading field_436. Sample connection text from the " & _
        "first few records:" & vbCrLf & vbCrLf

    Dim records As Collection
    Set records = CoPullAllRecords()

    Dim sampleCount As Long
    sampleCount = 0

    Dim i As Long
    For i = 1 To records.count

        Dim rec As Object
        Set rec = records(i)

        Dim connectionText As String
        connectionText = ProcessConnRecords.GetConnValue( _
            rec, CO_FIELD_REPORTING_PERIOD _
        )

        Dim connectedDate As Date
        Dim parsedOK As Boolean
        parsedOK = CoTryExtractDate(connectionText, connectedDate)

        msg = msg & i & ": '" & connectionText & "'"

        If parsedOK Then
            msg = msg & " -> parsed as " & Format$(connectedDate, "mm/dd/yyyy")
        Else
            msg = msg & " -> COULD NOT PARSE AS A DATE"
        End If

        msg = msg & vbCrLf

        sampleCount = sampleCount + 1
        If sampleCount >= 5 Then Exit For

    Next i

    CoDiagnoseEmptyResult = msg

End Function

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'
' Keyed by the exact header text already on the OTCarryOver
' sheet, so the sheet's current column order drives everything.
'=========================================================

Private Function CoFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "ID", "field_440"
    map.Add "ReportingPeriod", "field_436"
    map.Add "Employee", "field_437"
    map.Add "OTCarryoverWeekNum", "field_439"
    map.Add "OTCarryoverHours", "field_438"
    map.Add "NMCOMonth", "field_479"
    map.Add "NMCO_Qty", "field_480"
    map.Add "Searchable End of Period", "field_441"
    map.Add "Email", "field_735"

    Set CoFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO OTCarryOver, PRESERVING COLUMN ORDER
'=========================================================

Private Sub CoWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        CO_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3501, _
            "CoWriteMappedRecords", _
            "No header row was found on '" & CO_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = CoFieldKeyMap()

    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(CO_HEADER_ROW, columnNumber).value _
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
            vbObjectError + 3502, _
            "CoWriteMappedRecords", _
            "None of the headers on '" & CO_OUTPUT_SHEET & _
            "' matched a known OT Carryover field."

    End If

    Dim recordCount As Long

    If records Is Nothing Then
        recordCount = 0
    Else
        recordCount = records.count
    End If

    Dim previousLastRow As Long

    previousLastRow = wsOutput.Cells( _
        wsOutput.rows.count, 1 _
    ).End(xlUp).row

    Dim clearThroughRow As Long

    clearThroughRow = previousLastRow

    If CO_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = CO_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= CO_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(CO_FIRST_DATA_ROW, columnNumber), _
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
                    CoGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(CO_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

    CoFormatOutput wsOutput, recordCount

End Sub

'=========================================================
' FIELD VALUE PROCESSING
'=========================================================

Private Function CoGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        CoGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            CoGetMappedFieldValue = CStr(rec("id"))
        Else
            CoGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        CoGetMappedFieldValue = vbNullString
        Exit Function
    End If

    CoGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function

'=========================================================
' CONNECTED DATE PARSING
'=========================================================

Private Function CoTryExtractDate( _
    ByVal connectionText As String, _
    ByRef resultDate As Date) As Boolean

    Dim cleanedText As String

    cleanedText = Trim$(connectionText)

    If Len(cleanedText) = 0 Then Exit Function

    If IsDate(cleanedText) Then

        resultDate = CDate(cleanedText)
        CoTryExtractDate = True
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
                CoTryExtractDate = True
                Exit Function

            End If

        End If

    Next i

End Function

'=========================================================
' OUTPUT FORMATTING (only reformats columns this macro wrote)
'=========================================================

Private Sub CoFormatOutput( _
    ByVal wsOutput As Worksheet, _
    ByVal recordCount As Long)

    Dim firstDataRow As Long
    Dim finalDataRow As Long

    firstDataRow = CO_FIRST_DATA_ROW
    finalDataRow = firstDataRow + recordCount - 1

    Dim headerRange As Range
    Set headerRange = wsOutput.Range("A1").CurrentRegion.Rows(1)

    Dim dateColumns As Variant
    dateColumns = Array("ReportingPeriod", "Searchable End of Period")

    Dim c As Variant

    For Each c In dateColumns

        Dim matchCell As Range
        Set matchCell = headerRange.Find( _
            what:=c, LookIn:=xlValues, lookat:=xlWhole _
        )

        If Not matchCell Is Nothing Then

            wsOutput.Range( _
                wsOutput.Cells(firstDataRow, matchCell.column), _
                wsOutput.Cells(finalDataRow, matchCell.column) _
            ).NumberFormat = "mm/dd/yyyy"

        End If

    Next c

End Sub

'=========================================================
' DATE VALIDATION
'=========================================================

Private Function CoRequireDate( _
    ByVal cellValue As Variant, _
    ByVal cellDescription As String) As Date

    If IsError(cellValue) Then

        Err.Raise _
            vbObjectError + 3503, _
            "CoRequireDate", _
            cellDescription & " contains an Excel error."

    End If

    If Not IsDate(cellValue) Then

        Err.Raise _
            vbObjectError + 3504, _
            "CoRequireDate", _
            cellDescription & " does not contain a valid date."

    End If

    CoRequireDate = DateValue(CDate(cellValue))

End Function
