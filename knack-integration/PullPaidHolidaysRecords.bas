Attribute VB_Name = "PullPaidHolidaysRecords"
Option Explicit

'=========================================================
' PAID HOLIDAYS CONFIGURATION
'
' Template: PullHoursRecords.bas, simplified. Pulls Knack
' object_36 ("Paid Holidays") into the existing "PaidHolidays"
' sheet, preserving whatever column order is already there.
'
' Like Locations, this is a small reference table (holiday
' dates used elsewhere via Month/Year lookups), not tied to a
' single reporting period, so it pulls every record with no
' VariablesSheet date filter.
'=========================================================

Private Const HOL_OBJECT_KEY As String = "object_36"

Private Const HOL_OUTPUT_SHEET As String = "PaidHolidays"

Private Const HOL_HEADER_ROW As Long = 1
Private Const HOL_FIRST_DATA_ROW As Long = 2

Private Const HOL_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullPaidHolidays()

    On Error GoTo ErrorHandler

    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Worksheets(HOL_OUTPUT_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim records As Collection
    Set records = HolPullAllRecords()

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
