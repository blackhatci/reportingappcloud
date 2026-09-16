Attribute VB_Name = "PullAFDelinquencyRecords"
Option Explicit

'=========================================================
' AF DELINQUENCY CONFIGURATION
'
' Template: PullHoursRecords.bas / PullMonthlyTargetsRecords.bas.
' Pulls Knack object_12 ("AFDelinquency") into the existing
' afdelinquency sheet, preserving whatever column order is
' already on that sheet.
'
' object_12 has a real "Mark for Delete" flag, so that part is
' filtered on Knack's side (same as PullHoursRecords /
' PullLeadsRecords do for their delete fields). It has no plain
' date field for the pay period though - only a connection
' ("Pay Period when the Account was Closed") - so the period
' narrowing is done in VBA, same technique as Monthly Targets /
' OT Carryover.
'=========================================================

Private Const DEL_OBJECT_KEY As String = "object_12"

Private Const DEL_FIELD_MARKED_DELETE As String = "field_487"

' Connected pay-period field, matched against VariablesSheet!A2.
Private Const DEL_FIELD_PAY_PERIOD As String = "field_264"

Private Const DEL_VARIABLES_SHEET As String = "VariablesSheet"
Private Const DEL_OUTPUT_SHEET As String = "afdelinquency"

Private Const DEL_HEADER_ROW As Long = 1
Private Const DEL_FIRST_DATA_ROW As Long = 2

Private Const DEL_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullAFDelinquency()

    On Error GoTo ErrorHandler

    Dim wsVariables As Worksheet
    Dim wsOutput As Worksheet

    Set wsVariables = ThisWorkbook.Worksheets(DEL_VARIABLES_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(DEL_OUTPUT_SHEET)

    Dim targetPeriod As Date

    targetPeriod = DelRequireDate( _
        wsVariables.Range("A2").value, _
        DEL_VARIABLES_SHEET & "!A2" _
    )

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim filtersJson As String
    filtersJson = DelBuildFilters()

    Dim records As Collection

    Set records = DelPullAllFilteredRecords(filtersJson)

    ' Apply the connected pay-period check in VBA - object_12 has
    ' no plain date field for the pay period, only a connection.
    Set records = DelFilterByConnectedPayPeriod(records, targetPeriod)

    DelWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The AF Delinquency pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "AF Delinquency Pull"

    Resume CleanExit

End Sub

'=========================================================
' KNACK-SIDE FILTER
'
' Mark for Delete <> Yes
'=========================================================

Private Function DelBuildFilters() As String

    DelBuildFilters = _
        "[{""field"":""" & DEL_FIELD_MARKED_DELETE & """," & _
        """operator"":""is not""," & _
        """value"":""Yes""}]"

End Function

'=========================================================
' PULL ALL MATCHING KNACK PAGES
'=========================================================

Private Function DelPullAllFilteredRecords( _
    ByVal filtersJson As String) As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.GetRecordsByFilters( _
            DEL_OBJECT_KEY, _
            filtersJson, _
            page, _
            DEL_ROWS_PER_PAGE _
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

    Set DelPullAllFilteredRecords = allRecords

End Function

'=========================================================
' VBA-SIDE CONNECTED PAY-PERIOD FILTER
'=========================================================

Private Function DelFilterByConnectedPayPeriod( _
    ByVal sourceRecords As Collection, _
    ByVal targetPeriod As Date) As Collection

    Dim filteredRecords As New Collection

    If sourceRecords Is Nothing Then
        Set DelFilterByConnectedPayPeriod = filteredRecords
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
            DEL_FIELD_PAY_PERIOD _
        )

        Dim connectedDate As Date

        If DelTryExtractDate(connectionText, connectedDate) Then

            If DateValue(connectedDate) = requiredDate Then
                filteredRecords.Add rec
            End If

        End If

    Next i

    Set DelFilterByConnectedPayPeriod = filteredRecords

End Function

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'
' Keyed by the exact header text already on the afdelinquency
' sheet, so the sheet's current column order drives everything.
'
' "Employee Wage Code" is intentionally NOT mapped - object_12
' has only one wage-code field (EmployeeWageCode / field_373),
' which is already mapped to the "1 AFEmployeeWages_Code" column.
' Rather than guess which of the two columns should get it,
' this pull leaves "Employee Wage Code" untouched.
'=========================================================

Private Function DelFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "Location", "field_219"
    map.Add "End of Pay Period", "field_264"
    map.Add "Status", "field_102"
    map.Add "Delinquent Member Full Name", "field_93"
    map.Add "Days_Delinquent", "field_111"
    map.Add "Delinquent Date", "field_101"
    map.Add "Current Delinquent Balance", "field_262"
    map.Add "Reporting Date", "field_95"
    map.Add "Date Resolved", "field_104"
    map.Add "Delinquency Resolution", "field_105"
    map.Add "Number of Phone Calls", "field_106"
    map.Add "Number of Text Messages Sent", "field_107"
    map.Add "Number of Emails Sent", "field_108"
    map.Add "Total Follow Ups", "field_109"
    map.Add "Local Manager", "field_257"
    map.Add "Record ID", "field_260"
    map.Add "BillingCode", "field_110"
    map.Add "DEL over 61 Days Calculation", "field_261"
    map.Add "End of Period Searchable Field", "field_270"
    map.Add "1 AFEmployeeWages_Code", "field_373"
    map.Add "Del Bonus", "field_380"
    map.Add "Area", "field_284"
    map.Add "Marked for Delete", "field_487"
    map.Add "CurrentPeriod", "field_770"

    Set DelFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO afdelinquency, PRESERVING COLUMN ORDER
'=========================================================

Private Sub DelWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        DEL_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3601, _
            "DelWriteMappedRecords", _
            "No header row was found on '" & DEL_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = DelFieldKeyMap()

    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(DEL_HEADER_ROW, columnNumber).value _
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
            vbObjectError + 3602, _
            "DelWriteMappedRecords", _
            "None of the headers on '" & DEL_OUTPUT_SHEET & _
            "' matched a known AF Delinquency field."

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

    If DEL_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = DEL_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= DEL_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(DEL_FIRST_DATA_ROW, columnNumber), _
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
                    DelGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(DEL_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

    DelFormatOutput wsOutput, recordCount

End Sub

'=========================================================
' FIELD VALUE PROCESSING
'=========================================================

Private Function DelGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        DelGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            DelGetMappedFieldValue = CStr(rec("id"))
        Else
            DelGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        DelGetMappedFieldValue = vbNullString
        Exit Function
    End If

    DelGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function

'=========================================================
' CONNECTED DATE PARSING
'=========================================================

Private Function DelTryExtractDate( _
    ByVal connectionText As String, _
    ByRef resultDate As Date) As Boolean

    Dim cleanedText As String

    cleanedText = Trim$(connectionText)

    If Len(cleanedText) = 0 Then Exit Function

    If IsDate(cleanedText) Then

        resultDate = CDate(cleanedText)
        DelTryExtractDate = True
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
                DelTryExtractDate = True
                Exit Function

            End If

        End If

    Next i

End Function

'=========================================================
' OUTPUT FORMATTING (only reformats columns this macro wrote)
'=========================================================

Private Sub DelFormatOutput( _
    ByVal wsOutput As Worksheet, _
    ByVal recordCount As Long)

    Dim firstDataRow As Long
    Dim finalDataRow As Long

    firstDataRow = DEL_FIRST_DATA_ROW
    finalDataRow = firstDataRow + recordCount - 1

    Dim headerRange As Range
    Set headerRange = wsOutput.Range("A1").CurrentRegion.Rows(1)

    Dim moneyColumns As Variant
    moneyColumns = Array("Current Delinquent Balance", "Del Bonus")

    Dim dateColumns As Variant
    dateColumns = Array( _
        "End of Pay Period", "Delinquent Date", _
        "Reporting Date", "Date Resolved", _
        "End of Period Searchable Field" _
    )

    Dim c As Variant
    Dim matchCell As Range

    For Each c In moneyColumns

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

'=========================================================
' DATE VALIDATION
'=========================================================

Private Function DelRequireDate( _
    ByVal cellValue As Variant, _
    ByVal cellDescription As String) As Date

    If IsError(cellValue) Then

        Err.Raise _
            vbObjectError + 3603, _
            "DelRequireDate", _
            cellDescription & " contains an Excel error."

    End If

    If Not IsDate(cellValue) Then

        Err.Raise _
            vbObjectError + 3604, _
            "DelRequireDate", _
            cellDescription & " does not contain a valid date."

    End If

    DelRequireDate = DateValue(CDate(cellValue))

End Function
