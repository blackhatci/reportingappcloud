Attribute VB_Name = "PullLocationsRecords"
Option Explicit

'=========================================================
' LOCATIONS CONFIGURATION
'
' Template: PullHoursRecords.bas, simplified. Pulls Knack
' object_22 ("Locations") into the existing "Area Location"
' sheet, preserving whatever column order is already there.
'
' Unlike the period-scoped pulls (Monthly Targets, CarryOver,
' AFDelinquency), Locations is a reference/lookup table, not
' tied to a reporting period, so there is no VariablesSheet
' date filter here - only active locations are pulled, filtered
' on Knack's side via "Active Location".
'=========================================================

Private Const LOC_OBJECT_KEY As String = "object_22"

Private Const LOC_FIELD_ACTIVE As String = "field_490"

Private Const LOC_OUTPUT_SHEET As String = "Area Location"

Private Const LOC_HEADER_ROW As Long = 1
Private Const LOC_FIRST_DATA_ROW As Long = 2

Private Const LOC_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullLocations()

    On Error GoTo ErrorHandler

    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Worksheets(LOC_OUTPUT_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim filtersJson As String
    filtersJson = LocBuildFilters()

    Dim records As Collection
    Set records = LocPullAllFilteredRecords(filtersJson)

    LocWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The Locations pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Locations Pull"

    Resume CleanExit

End Sub

'=========================================================
' KNACK-SIDE FILTER
'
' Active Location = Yes
'=========================================================

Private Function LocBuildFilters() As String

    LocBuildFilters = _
        "[{""field"":""" & LOC_FIELD_ACTIVE & """," & _
        """operator"":""is""," & _
        """value"":""Yes""}]"

End Function

'=========================================================
' PULL ALL MATCHING KNACK PAGES
'=========================================================

Private Function LocPullAllFilteredRecords( _
    ByVal filtersJson As String) As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.GetRecordsByFilters( _
            LOC_OBJECT_KEY, _
            filtersJson, _
            page, _
            LOC_ROWS_PER_PAGE _
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

    Set LocPullAllFilteredRecords = allRecords

End Function

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'=========================================================

Private Function LocFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "Location", "field_199"
    map.Add "Area", "field_215"

    Set LocFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO Area Location, PRESERVING COLUMN ORDER
'=========================================================

Private Sub LocWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        LOC_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3701, _
            "LocWriteMappedRecords", _
            "No header row was found on '" & LOC_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = LocFieldKeyMap()

    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(LOC_HEADER_ROW, columnNumber).value _
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
            vbObjectError + 3702, _
            "LocWriteMappedRecords", _
            "None of the headers on '" & LOC_OUTPUT_SHEET & _
            "' matched a known Locations field."

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

    If LOC_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = LOC_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= LOC_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(LOC_FIRST_DATA_ROW, columnNumber), _
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
                    LocGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(LOC_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

End Sub

'=========================================================
' FIELD VALUE PROCESSING
'=========================================================

Private Function LocGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        LocGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            LocGetMappedFieldValue = CStr(rec("id"))
        Else
            LocGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        LocGetMappedFieldValue = vbNullString
        Exit Function
    End If

    LocGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function
