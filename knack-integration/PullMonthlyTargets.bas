Attribute VB_Name = "PullMonthlyTargets"
Option Explicit

'=========================================================
' MONTHLY TARGETS CONFIGURATION
'
' Pulls Knack object_28 ("Employee Monthly Targets") into a
' dedicated sheet, the same way PullLeadsRecords.bas pulls
' Leads (object_9) into the Leads_Pull sheet.
'
' This is a READ-ONLY pull. It does not touch the existing
' "MonthlyTargets" sheet or the Zapier upload macro
' (UploadMOnthlyTargetRecords in macroUploadMonthlyTargets.bas) -
' those keep working exactly as they do today.
'=========================================================

Private Const MTGT_OBJECT_KEY As String = "object_28"

Private Const MTGT_FIELD_TARGET_MONTH As String = "field_385"
Private Const MTGT_FIELD_TARGET_YEAR As String = "field_387"

Private Const MTGT_MAPPING_SHEET As String = "Knack Field Mappings"
Private Const MTGT_OUTPUT_SHEET As String = "MonthlyTargets_Pull"
Private Const MTGT_OUTPUT_TABLE As String = "MonthlyTargetsTablePull"

' Row pair on the mapping sheet: header names above, Knack field
' keys below - same convention as the Leads block (rows 8-9) and
' the Hours block (rows 1-2).
Private Const MTGT_MAPPING_HEADER_ROW As Long = 10
Private Const MTGT_MAPPING_FIELD_ROW As Long = 11

' Criteria cells live on the output sheet itself (B1 = Target Month,
' B2 = Target Year) so this pull doesn't need to share cells with
' any other macro's inputs. Leave either blank to pull ALL records.
Private Const MTGT_CRITERIA_MONTH_CELL As String = "B1"
Private Const MTGT_CRITERIA_YEAR_CELL As String = "B2"

Private Const MTGT_ROWS_PER_PAGE As Long = 100

'=========================================================
' PUBLIC ENTRY POINT
'=========================================================

Public Sub PullMonthlyTargets()

    On Error GoTo ErrorHandler

    Dim wsMapping As Worksheet
    Dim wsOutput As Worksheet

    Set wsMapping = ThisWorkbook.Worksheets(MTGT_MAPPING_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(MTGT_OUTPUT_SHEET)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim targetMonth As String
    Dim targetYear As String

    targetMonth = Trim$(CStr(wsOutput.Range(MTGT_CRITERIA_MONTH_CELL).value))
    targetYear = Trim$(CStr(wsOutput.Range(MTGT_CRITERIA_YEAR_CELL).value))

    Dim filtersJson As String
    filtersJson = MtgtBuildFilters(targetMonth, targetYear)

    Dim records As Collection
    Set records = MtgtPullAllRecords(filtersJson)

    MtgtWriteMappedRecords records, wsMapping, wsOutput

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
' KNACK-SIDE FILTERS
'
' Target Month = <criteria B1>   (skipped if blank)
' Target Year  = <criteria B2>   (skipped if blank)
'=========================================================

Private Function MtgtBuildFilters( _
    ByVal targetMonth As String, _
    ByVal targetYear As String) As String

    Dim rules As String
    rules = vbNullString

    If Len(targetMonth) > 0 Then
        rules = rules & "{""field"":""" & MTGT_FIELD_TARGET_MONTH & """," & _
                """operator"":""is""," & _
                """value"":""" & Replace(targetMonth, """", "\""") & """}"
    End If

    If Len(targetYear) > 0 Then
        If Len(rules) > 0 Then rules = rules & ","
        rules = rules & "{""field"":""" & MTGT_FIELD_TARGET_YEAR & """," & _
                """operator"":""is""," & _
                """value"":""" & Replace(targetYear, """", "\""") & """}"
    End If

    MtgtBuildFilters = "[" & rules & "]"

End Function

'=========================================================
' PULL ALL MATCHING KNACK PAGES
'=========================================================

Private Function MtgtPullAllRecords( _
    ByVal filtersJson As String) As Collection

    Dim allRecords As New Collection

    Dim page As Long
    Dim totalPages As Long

    page = 1
    totalPages = 1

    Do While page <= totalPages

        Dim responseText As String

        responseText = KnackAPI.GetRecordsByFilters( _
            MTGT_OBJECT_KEY, _
            filtersJson, _
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
' MAP AND WRITE TO MonthlyTargetsTablePull
'=========================================================

Private Sub MtgtWriteMappedRecords( _
    ByVal records As Collection, _
    ByVal wsMapping As Worksheet, _
    ByVal wsOutput As Worksheet)

    Dim outputTable As ListObject

    On Error Resume Next
    Set outputTable = wsOutput.ListObjects(MTGT_OUTPUT_TABLE)
    On Error GoTo 0

    If outputTable Is Nothing Then

        Err.Raise _
            vbObjectError + 3301, _
            "MtgtWriteMappedRecords", _
            "The Excel table '" & MTGT_OUTPUT_TABLE & _
            "' was not found on the '" & MTGT_OUTPUT_SHEET & "' sheet." & _
            vbCrLf & _
            "Create it starting at A4 (rows 1-3 are reserved for the " & _
            "Target Month / Target Year criteria) before running this pull."

    End If

    Dim lastColumn As Long

    lastColumn = wsMapping.Cells( _
        MTGT_MAPPING_FIELD_ROW, _
        wsMapping.columns.count _
    ).End(xlToLeft).column

    If lastColumn < 1 Then

        Err.Raise _
            vbObjectError + 3302, _
            "MtgtWriteMappedRecords", _
            "No Monthly Targets mapping was found in row " & _
            MTGT_MAPPING_FIELD_ROW & " of '" & MTGT_MAPPING_SHEET & "'."

    End If

    MtgtValidateFieldMapping wsMapping, lastColumn

    Dim recordCount As Long

    If records Is Nothing Then
        recordCount = 0
    Else
        recordCount = records.count
    End If

    Dim requiredDataRows As Long

    If recordCount > 0 Then
        requiredDataRows = recordCount
    Else
        requiredDataRows = 1
    End If

    Dim tableTopRow As Long
    tableTopRow = outputTable.HeaderRowRange.row

    Dim newTableRange As Range

    Set newTableRange = wsOutput.Cells(tableTopRow, 1).Resize( _
        requiredDataRows + 1, _
        lastColumn _
    )

    outputTable.Resize newTableRange

    outputTable.HeaderRowRange.value = _
        wsMapping.Cells(MTGT_MAPPING_HEADER_ROW, 1).Resize(1, lastColumn).value

    If Not outputTable.DataBodyRange Is Nothing Then
        outputTable.DataBodyRange.ClearContents
    End If

    If recordCount = 0 Then Exit Sub

    Dim outputValues() As Variant
    ReDim outputValues(1 To recordCount, 1 To lastColumn)

    Dim recordNumber As Long
    Dim columnNumber As Long

    For recordNumber = 1 To recordCount

        Dim rec As Object
        Set rec = records(recordNumber)

        For columnNumber = 1 To lastColumn

            Dim fieldKey As String

            fieldKey = Trim$(CStr( _
                wsMapping.Cells(MTGT_MAPPING_FIELD_ROW, columnNumber).value _
            ))

            If Len(fieldKey) > 0 Then

                outputValues(recordNumber, columnNumber) = _
                    MtgtGetMappedFieldValue(rec, fieldKey)

            Else

                outputValues(recordNumber, columnNumber) = vbNullString

            End If

        Next columnNumber

    Next recordNumber

    outputTable.DataBodyRange.Resize(recordCount, lastColumn).value = outputValues

End Sub

'=========================================================
' FIELD MAPPING VALIDATION
'=========================================================

Private Sub MtgtValidateFieldMapping( _
    ByVal wsMapping As Worksheet, _
    ByVal lastColumn As Long)

    Dim columnNumber As Long

    For columnNumber = 1 To lastColumn

        Dim headerName As String
        Dim fieldKey As String

        headerName = Trim$(CStr( _
            wsMapping.Cells(MTGT_MAPPING_HEADER_ROW, columnNumber).value _
        ))

        fieldKey = Trim$(CStr( _
            wsMapping.Cells(MTGT_MAPPING_FIELD_ROW, columnNumber).value _
        ))

        If Len(headerName) > 0 And Len(fieldKey) = 0 Then

            Err.Raise _
                vbObjectError + 3303, _
                "MtgtValidateFieldMapping", _
                "The Monthly Targets mapping is blank for column " & _
                columnNumber & " (" & headerName & ")."

        End If

    Next columnNumber

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
