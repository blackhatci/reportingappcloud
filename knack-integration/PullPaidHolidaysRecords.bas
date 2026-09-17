Attribute VB_Name = "PullPaidHolidaysRecords"
Option Explicit

'=========================================================
' PAID HOLIDAYS CONFIGURATION
'
' Template: PullHoursRecords.bas. Pulls Knack object_36
' ("Paid Holidays") into the existing "PaidHolidays" sheet,
' preserving whatever column order is already there.
'
' Filtered to the current reporting period, same as the other
' period-scoped pulls: reads the target date from
' VariablesSheet!A2 and keeps only records whose AFReporting_Period
' connection (field_497) matches it - same connected-field
' technique as PullHoursRecords/PullCarryOverRecords, via
' ProcessConnRecords.GetConnValue, rather than comparing the
' separate Month/Year number fields.
'=========================================================

Private Const HOL_OBJECT_KEY As String = "object_36"

Private Const HOL_FIELD_REPORTING_PERIOD As String = "field_497"

Private Const HOL_VARIABLES_SHEET As String = "VariablesSheet"
Private Const HOL_OUTPUT_SHEET As String = "PaidHolidays"

Private Const HOL_HEADER_ROW As Long = 1
Private Const HOL_FIRST_DATA_ROW As Long = 2

Private Const HOL_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullPaidHolidays()

    On Error GoTo ErrorHandler

    Dim wsVariables As Worksheet
    Dim wsOutput As Worksheet

    Set wsVariables = ThisWorkbook.Worksheets(HOL_VARIABLES_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(HOL_OUTPUT_SHEET)

    Dim targetPeriod As Date

    targetPeriod = HolRequireDate( _
        wsVariables.Range("A2").value, _
        HOL_VARIABLES_SHEET & "!A2" _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim records As Collection
    Set records = HolPullAllRecords()

    Dim rawCount As Long
    rawCount = records.count

    Set records = HolFilterByConnectedPeriod(records, targetPeriod)

    If records.count = 0 Then

        Application.EnableEvents = True
        Application.ScreenUpdating = True

        MsgBox _
            HolDiagnoseEmptyResult(rawCount, targetPeriod), _
            vbExclamation, _
            "Paid Holidays Pull - No Matching Records"

        Exit Sub

    End If

    HolWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The Paid Holidays pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Paid Holidays Pull"

    Resume CleanExit

End Sub

'=========================================================
' SELF-DIAGNOSIS - builds a message explaining exactly why zero
' records matched, using real data from the actual pull.
'=========================================================

Private Function HolDiagnoseEmptyResult( _
    ByVal rawCount As Long, _
    ByVal targetPeriod As Date) As String

    Dim msg As String

    msg = "Target period (from VariablesSheet!A2): " & _
        Format$(targetPeriod, "mm/dd/yyyy") & vbCrLf & _
        "Raw object_36 records pulled: " & rawCount & vbCrLf & vbCrLf

    If rawCount = 0 Then

        msg = msg & _
            "The Knack pull itself returned zero records for " & _
            "object_36 (Paid Holidays). Check the Paid Holidays " & _
            "table directly in the Knack Builder to confirm it has " & _
            "records."

        HolDiagnoseEmptyResult = msg
        Exit Function

    End If

    msg = msg & _
        "Records were pulled, but none matched the target period " & _
        "after reading field_497. Sample connection text from the " & _
        "first few records:" & vbCrLf & vbCrLf

    Dim records As Collection
    Set records = HolPullAllRecords()

    Dim sampleCount As Long
    sampleCount = 0

    Dim i As Long
    For i = 1 To records.count

        Dim rec As Object
        Set rec = records(i)

        Dim connectionText As String
        connectionText = ProcessConnRecords.GetConnValue( _
            rec, HOL_FIELD_REPORTING_PERIOD _
        )

        Dim connectedDate As Date
        Dim parsedOK As Boolean
        parsedOK = HolTryExtractDate(connectionText, connectedDate)

        msg = msg & i & ": '" & connectionText & "'"

        If parsedOK Then
            msg = msg & " -> parsed as " & Format$(connectedDate, "mm/dd/yyyy")
        Else
            msg = msg & " -> COULD NOT PARSE AS A DATE"
        End If

        msg = msg & vbCrLf

        sampleCount = sampleCount + 1
        If sampleCount >= 8 Then Exit For

    Next i

    HolDiagnoseEmptyResult = msg

End Function

'=========================================================
' VBA-SIDE CONNECTED REPORTING-PERIOD FILTER
'=========================================================

Private Function HolFilterByConnectedPeriod( _
    ByVal sourceRecords As Collection, _
    ByVal targetPeriod As Date) As Collection

    Dim filteredRecords As New Collection

    If sourceRecords Is Nothing Then
        Set HolFilterByConnectedPeriod = filteredRecords
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
            rec, HOL_FIELD_REPORTING_PERIOD _
        )

        Dim connectedDate As Date

        If HolTryExtractDate(connectionText, connectedDate) Then

            If DateValue(connectedDate) = requiredDate Then
                filteredRecords.Add rec
            End If

        End If

    Next i

    Set HolFilterByConnectedPeriod = filteredRecords

End Function

'=========================================================
' CONNECTED DATE PARSING
'=========================================================

Private Function HolTryExtractDate( _
    ByVal connectionText As String, _
    ByRef resultDate As Date) As Boolean

    Dim cleanedText As String

    cleanedText = Trim$(connectionText)

    If Len(cleanedText) = 0 Then Exit Function

    If IsDate(cleanedText) Then

        resultDate = CDate(cleanedText)
        HolTryExtractDate = True
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
                HolTryExtractDate = True
                Exit Function

            End If

        End If

    Next i

End Function

'=========================================================
' DATE VALIDATION
'=========================================================

Private Function HolRequireDate( _
    ByVal cellValue As Variant, _
    ByVal cellDescription As String) As Date

    If IsError(cellValue) Then

        Err.Raise _
            vbObjectError + 3803, _
            "HolRequireDate", _
            cellDescription & " contains an Excel error."

    End If

    If Not IsDate(cellValue) Then

        Err.Raise _
            vbObjectError + 3804, _
            "HolRequireDate", _
            cellDescription & " does not contain a valid date."

    End If

    HolRequireDate = DateValue(CDate(cellValue))

End Function

'=========================================================
' PULL ALL KNACK PAGES (no filter - small reference table)
'=========================================================

Private Function HolPullAllRecords() As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.KnackGetRecords( _
            HOL_OBJECT_KEY, _
            page, _
            HOL_ROWS_PER_PAGE _
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

    Set HolPullAllRecords = allRecords

End Function

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'=========================================================

Private Function HolFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "AFReporting_Period", "field_497"
    map.Add "Date", "field_492"
    map.Add "Holiday Name", "field_493"
    map.Add "Holiday Hours", "field_494"
    map.Add "Month", "field_495"
    map.Add "Year", "field_496"
    map.Add "Searchable Reporting Period Date", "field_498"

    Set HolFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO PaidHolidays, PRESERVING COLUMN ORDER
'=========================================================

Private Sub HolWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        HOL_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3801, _
            "HolWriteMappedRecords", _
            "No header row was found on '" & HOL_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = HolFieldKeyMap()

    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(HOL_HEADER_ROW, columnNumber).value _
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
            vbObjectError + 3802, _
            "HolWriteMappedRecords", _
            "None of the headers on '" & HOL_OUTPUT_SHEET & _
            "' matched a known Paid Holidays field."

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

    If HOL_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = HOL_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= HOL_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(HOL_FIRST_DATA_ROW, columnNumber), _
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
                    HolGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(HOL_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

    HolFormatOutput wsOutput, recordCount

End Sub

'=========================================================
' FIELD VALUE PROCESSING
'=========================================================

Private Function HolGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        HolGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            HolGetMappedFieldValue = CStr(rec("id"))
        Else
            HolGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        HolGetMappedFieldValue = vbNullString
        Exit Function
    End If

    HolGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function

'=========================================================
' OUTPUT FORMATTING (only reformats columns this macro wrote)
'=========================================================

Private Sub HolFormatOutput( _
    ByVal wsOutput As Worksheet, _
    ByVal recordCount As Long)

    Dim firstDataRow As Long
    Dim finalDataRow As Long

    firstDataRow = HOL_FIRST_DATA_ROW
    finalDataRow = firstDataRow + recordCount - 1

    Dim headerRange As Range
    Set headerRange = wsOutput.Range("A1").CurrentRegion.Rows(1)

    Dim dateColumns As Variant
    dateColumns = Array( _
        "AFReporting_Period", "Date", _
        "Searchable Reporting Period Date" _
    )

    Dim c As Variant
    Dim matchCell As Range

    For Each c In dateColumns

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
