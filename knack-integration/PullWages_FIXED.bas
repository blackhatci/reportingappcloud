Attribute VB_Name = "PullWages"
Option Explicit

'=========================================================
' OBJECT 15 - LOCAL MANAGER RATE CONFIGURATION
'=========================================================

Private Const OBJ_LOCAL_MGR_RATE As String = "object_15"
Private Const OBJ_LOCAL_MANAGER As String = "object_6"

' object_15
Private Const FIELD_ACTIVE_RATE As String = "field_140"
Private Const FIELD_LOCAL_MANAGER As String = "field_230"

' object_6
Private Const FIELD_USER_STATUS As String = "field_30"

Private Const ROWS_PER_PAGE As Long = 100

' --- Output sheet (added - this was missing entirely) -----
Private Const WAGE_OUTPUT_SHEET As String = "employeewagecodes"
Private Const WAGE_HEADER_ROW As Long = 1
Private Const WAGE_FIRST_DATA_ROW As Long = 2

'=========================================================
' PUBLIC ENTRY POINT (unchanged) - fetches qualifying
' object_15 records and returns them in memory. Kept as-is
' in case other code calls this expecting just a Collection.
'
' object_15.field_140 = True
'
' AND
'
' connected object_6 record via field_230 has:
' object_6.field_30 = "Active"
'
' Example:
'
'   Dim records As Collection
'   Set records = PullActEmpWages()
'
'=========================================================

Public Function PullActEmpWages() As Collection

    On Error GoTo ErrorHandler

    '-----------------------------------------------------
    ' 1. Pull object_15 records where Active Rate = True
    '-----------------------------------------------------

    Dim rateRecords As Collection

    Set rateRecords = _
        PullObject15ActiveRates()

    '-----------------------------------------------------
    ' 2. Pull object_6 once and build lookup of managers
    '    whose User Status = Active
    '-----------------------------------------------------

    Dim activeManagerLookup As Object

    Set activeManagerLookup = _
        BuildActiveManagerLookup()

    '-----------------------------------------------------
    ' 3. Filter object_15 records by connected manager
    '-----------------------------------------------------

    Dim qualifyingRecords As New Collection

    Dim i As Long

    For i = 1 To rateRecords.count

        Dim rec As Object
        Set rec = rateRecords(i)

        If Object15ManagerIsActive( _
            rec, _
            activeManagerLookup _
        ) Then

            ' Add the ENTIRE object_15 record.
            ' All fields returned by Knack remain available.
            qualifyingRecords.Add rec

        End If

    Next i

    Set PullActEmpWages = _
        qualifyingRecords

    Exit Function

ErrorHandler:

    Err.Raise _
        Err.Number, _
        "PullActEmpWages", _
        Err.Description

End Function

'=========================================================
' NEW PUBLIC ENTRY POINT - this is what should be called
' from RunPayrollFinal instead of the bare PullActEmpWages
' call. Fetches the same qualifying object_15 records AND
' writes them into the employeewagecodes sheet, so the sheet
' mgrWageLoad reads from is actually refreshed from Knack.
'=========================================================

Public Sub PullEmployeeWageCodes()

    On Error GoTo ErrorHandler

    Dim wsOutput As Worksheet
    Set wsOutput = ThisWorkbook.Worksheets(WAGE_OUTPUT_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim records As Collection
    Set records = PullActEmpWages()

    WageWriteMappedRecords records, wsOutput

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox _
        "The Employee Wage Codes pull could not be completed." & _
        vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "Employee Wage Codes Pull"

    Resume CleanExit

End Sub

'=========================================================
' HEADER NAME -> KNACK FIELD KEY
'
' Keyed by the exact header text already on the
' employeewagecodes sheet, so its current column order
' drives everything - nothing here depends on position.
'=========================================================

Private Function WageFieldKeyMap() As Object

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary")

    map.Add "ID", "field_134"
    map.Add "Local Manager", "field_230"
    map.Add "Active Rate", "field_140"
    map.Add "Hourly Wage", "field_233"
    map.Add "Salary", "field_379"
    map.Add "NMCH", "field_235"
    map.Add "NMCC", "field_236"
    map.Add "NMOP", "field_237"
    map.Add "NMCA", "field_371"
    map.Add "NMTI", "field_378"
    map.Add "LOSF", "field_238"
    map.Add "LASF", "field_239"
    map.Add "COLL", "field_240"
    map.Add "PPV", "field_241"
    map.Add "MNMU", "field_242"
    map.Add "MDG", "field_243"
    map.Add "MDG5", "field_375"
    map.Add "MNMPLus", "field_280"
    map.Add "NMPr", "field_376"
    map.Add "PT_Percent", "field_377"
    map.Add "Location", "field_468"
    map.Add "HireDate", "field_512"
    map.Add "HealthCare_Eligible", "field_743"
    map.Add "HealthCare_Eligible_NumberField", "field_744"
    map.Add "HealthcareNotificationDate", "field_745"
    map.Add "Technician Name", "field_835"
    map.Add "EffectiveDate", "field_871"
    map.Add "Cell Reimbursement", "field_973"
    map.Add "Part_Time", "field_1013"
    map.Add "QtrBonusL1", "field_1291"
    map.Add "QtrBonusL2", "field_1292"
    map.Add "QtrBonusL3", "field_1293"

    Set WageFieldKeyMap = map

End Function

'=========================================================
' MAP AND WRITE TO employeewagecodes, PRESERVING COLUMN ORDER
'=========================================================

Private Sub WageWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsOutput As Worksheet)

    Dim lastColumn As Long

    lastColumn = wsOutput.Cells( _
        WAGE_HEADER_ROW, _
        wsOutput.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3901, _
            "WageWriteMappedRecords", _
            "No header row was found on '" & WAGE_OUTPUT_SHEET & "'."

    End If

    Dim fieldMap As Object
    Set fieldMap = WageFieldKeyMap()

    Dim columnFieldKeys() As String
    ReDim columnFieldKeys(1 To lastColumn)

    Dim columnNumber As Long
    Dim mappedColumnCount As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        headerName = Trim$(CStr( _
            wsOutput.Cells(WAGE_HEADER_ROW, columnNumber).value _
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
            vbObjectError + 3902, _
            "WageWriteMappedRecords", _
            "None of the headers on '" & WAGE_OUTPUT_SHEET & _
            "' matched a known Employee Wage Codes field."

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

    If WAGE_FIRST_DATA_ROW + recordCount - 1 > clearThroughRow Then
        clearThroughRow = WAGE_FIRST_DATA_ROW + recordCount - 1
    End If

    If clearThroughRow >= WAGE_FIRST_DATA_ROW Then

        For columnNumber = 1 To lastColumn

            If Len(columnFieldKeys(columnNumber)) > 0 Then

                wsOutput.Range( _
                    wsOutput.Cells(WAGE_FIRST_DATA_ROW, columnNumber), _
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
                    WageGetMappedFieldValue(rec, columnFieldKeys(columnNumber))

            End If

        Next columnNumber

    Next recordNumber

    For columnNumber = 1 To lastColumn

        If Len(columnFieldKeys(columnNumber)) > 0 Then

            wsOutput.Cells(WAGE_FIRST_DATA_ROW, columnNumber).Resize( _
                recordCount, 1 _
            ).value = Application.Index(outputValues, 0, columnNumber)

        End If

    Next columnNumber

End Sub

Private Function WageGetMappedFieldValue( _
    ByVal rec As Object, _
    ByVal fieldKey As String) As Variant

    If rec Is Nothing Then
        WageGetMappedFieldValue = vbNullString
        Exit Function
    End If

    If LCase$(fieldKey) = "id" Then

        If rec.Exists("id") Then
            WageGetMappedFieldValue = CStr(rec("id"))
        Else
            WageGetMappedFieldValue = vbNullString
        End If

        Exit Function

    End If

    If Not rec.Exists(fieldKey) Then
        WageGetMappedFieldValue = vbNullString
        Exit Function
    End If

    WageGetMappedFieldValue = ProcessConnRecords.GetFieldValue(rec, fieldKey)

End Function

'=========================================================
' PULL OBJECT 15 WHERE ACTIVE RATE = TRUE (unchanged)
'=========================================================

Private Function PullObject15ActiveRates() As Collection

    Dim allRecords As New Collection

    Dim filtersJson As String

    filtersJson = _
        "[{""field"":""" & FIELD_ACTIVE_RATE & """," & _
        """operator"":""is""," & _
        """value"":""Yes""}]"

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = _
            KnackAPI.GetRecordsByFilters( _
                OBJ_LOCAL_MGR_RATE, _
                filtersJson, _
                page, _
                ROWS_PER_PAGE _
            )

        Dim parsed As Object

        Set parsed = _
            JsonConverter.ParseJson(responseText)

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

    Set PullObject15ActiveRates = _
        allRecords

End Function

'=========================================================
' BUILD LOOKUP OF ACTIVE OBJECT 6 MANAGERS (unchanged)
'
' Key = object_6 Knack record ID
'=========================================================

Private Function BuildActiveManagerLookup() As Object

    Dim lookup As Object

    Set lookup = _
        CreateObject("Scripting.Dictionary")

    lookup.CompareMode = vbTextCompare

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = _
            KnackAPI.KnackGetRecords( _
                OBJ_LOCAL_MANAGER, _
                page, _
                ROWS_PER_PAGE _
            )

        Dim parsed As Object

        Set parsed = _
            JsonConverter.ParseJson(responseText)

        If parsed.Exists("total_pages") Then
            totalPages = CLng(parsed("total_pages"))
        Else
            totalPages = 1
        End If

        If parsed.Exists("records") Then

            Dim rec As Variant

            For Each rec In parsed("records")

                Dim userStatus As String

                userStatus = _
                    Trim$(CStr( _
                        ProcessConnRecords.GetFieldValue( _
                            rec, _
                            FIELD_USER_STATUS _
                        ) _
                    ))

                If StrComp( _
                    userStatus, _
                    "Active", _
                    vbTextCompare _
                ) = 0 Then

                    If rec.Exists("id") Then

                        Dim managerID As String

                        managerID = _
                            Trim$(CStr(rec("id")))

                        If Len(managerID) > 0 Then

                            If Not lookup.Exists(managerID) Then
                                lookup.Add managerID, True
                            End If

                        End If

                    End If

                End If

            Next rec

        End If

        page = page + 1

    Loop

    Set BuildActiveManagerLookup = lookup

End Function

'=========================================================
' CHECK OBJECT 15 LOCAL MANAGER CONNECTION (unchanged)
'=========================================================

Private Function Object15ManagerIsActive( _
    ByVal rec As Object, _
    ByVal activeManagerLookup As Object) As Boolean

    If rec Is Nothing Then Exit Function

    If Not rec.Exists(FIELD_LOCAL_MANAGER) Then
        Exit Function
    End If

    Dim connectionValue As Variant

    If IsObject(rec(FIELD_LOCAL_MANAGER)) Then

        ' Handle Knack connection collection.
        If TypeName(rec(FIELD_LOCAL_MANAGER)) = "Collection" Then

            Dim connections As Collection
            Set connections = rec(FIELD_LOCAL_MANAGER)

            Dim item As Variant

            For Each item In connections

                If ConnectionIDIsActive( _
                    item, _
                    activeManagerLookup _
                ) Then

                    Object15ManagerIsActive = True
                    Exit Function

                End If

            Next item

        Else

            ' Handle single connected record object.
            If ConnectionIDIsActive( _
                rec(FIELD_LOCAL_MANAGER), _
                activeManagerLookup _
            ) Then

                Object15ManagerIsActive = True

            End If

        End If

    End If

End Function

'=========================================================
' READ CONNECTED RECORD ID (unchanged)
'=========================================================

Private Function ConnectionIDIsActive( _
    ByVal connectionItem As Variant, _
    ByVal activeManagerLookup As Object) As Boolean

    If Not IsObject(connectionItem) Then
        Exit Function
    End If

    On Error GoTo NotActive

    If connectionItem.Exists("id") Then

        Dim connectedID As String

        connectedID = _
            Trim$(CStr(connectionItem("id")))

        If Len(connectedID) > 0 Then

            ConnectionIDIsActive = _
                activeManagerLookup.Exists(connectedID)

        End If

    End If

    Exit Function

NotActive:

    ConnectionIDIsActive = False

End Function
