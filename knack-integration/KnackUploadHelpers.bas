Attribute VB_Name = "KnackUploadHelpers"
Option Explicit

'=========================================================
' SHARED UPLOAD HELPERS
'
' Used by UploadCarryOverRecords, UploadMonthlyTargetsRecords, and
' UploadPayrollRecords to replace the old macroUploadCORecords /
' macroUploadPayroll / UploadMOnthlyTargetRecords, which posted raw
' sheet rows to Zapier "catch" webhooks and let Zapier resolve
' connections and map fields into Knack. These modules call the
' Knack API directly instead, so this module replicates the one
' thing Zapier was doing invisibly: resolving a plain value (a
' manager's name, a location's name, a period's date, an area
' number) into the Knack record ID a connection field actually
' needs.
'
' Confirmed against live records (via Zapier's Knack "Find Record"
' action, cross-checked against these tables' own field lists):
'   Local Manager (object_6)      identifier = field_37  (Name)
'   Locations (object_22)         identifier = field_199 (Location Load Name)
'   Areas (object_23)             identifier = field_214 (Area Number)
'   AFReporting_Periods (object_20) identifier = field_175 (date)
'=========================================================

Private Const OBJ_LOCAL_MANAGER As String = "object_6"
Private Const OBJ_LOCATIONS As String = "object_22"
Private Const OBJ_AREAS As String = "object_23"
Private Const OBJ_REPORTING_PERIODS As String = "object_20"

Private Const FIELD_MANAGER_NAME As String = "field_37"
Private Const FIELD_LOCATION_NAME As String = "field_199"
Private Const FIELD_AREA_NUMBER As String = "field_214"
Private Const FIELD_PERIOD_DATE As String = "field_175"

'=========================================================
' LOOKUP CACHE
'
' One cache per upload run, passed into every resolver below, so
' the same manager/location/period is only looked up once even if
' it appears on many rows.
'=========================================================

Public Function NewLookupCache() As Object
    Set NewLookupCache = CreateObject("Scripting.Dictionary")
    NewLookupCache.CompareMode = vbTextCompare
End Function

'=========================================================
' GENERIC RESOLVER
'
' Looks up a record ID by exact match on one field. Returns
' vbNullString (not an error) if no match is found or the input is
' blank, so callers can decide how to treat a miss - callers should
' always check for an empty result before using it in a payload.
'=========================================================

Private Function ResolveIdByField( _
    ByVal objectKey As String, _
    ByVal fieldKey As String, _
    ByVal matchValue As String, _
    ByVal cache As Object) As String

    Dim cleanValue As String
    cleanValue = Trim$(matchValue)

    If Len(cleanValue) = 0 Then Exit Function

    Dim cacheKey As String
    cacheKey = objectKey & "|" & fieldKey & "|" & cleanValue

    If cache.Exists(cacheKey) Then
        ResolveIdByField = cache(cacheKey)
        Exit Function
    End If

    Dim filtersJson As String
    filtersJson = "[{""field"":""" & fieldKey & """," & _
        """operator"":""is""," & _
        """value"":""" & Replace(cleanValue, """", "\""") & """}]"

    Dim responseText As String
    responseText = KnackAPI.GetRecordsByFilters(objectKey, filtersJson, 1, 1)

    Dim parsed As Object
    Set parsed = JsonConverter.ParseJson(responseText)

    Dim resolvedId As String

    If parsed.Exists("records") Then
        Dim records As Variant
        Set records = parsed("records")
        If records.count > 0 Then
            Dim rec As Object
            Set rec = records(1)
            If rec.Exists("id") Then
                resolvedId = CStr(rec("id"))
            End If
        End If
    End If

    cache.Add cacheKey, resolvedId

    ResolveIdByField = resolvedId

End Function

Public Function ResolveManagerId( _
    ByVal managerName As String, ByVal cache As Object) As String

    ResolveManagerId = ResolveIdByField( _
        OBJ_LOCAL_MANAGER, FIELD_MANAGER_NAME, managerName, cache)

End Function

Public Function ResolveLocationId( _
    ByVal locationName As String, ByVal cache As Object) As String

    ResolveLocationId = ResolveIdByField( _
        OBJ_LOCATIONS, FIELD_LOCATION_NAME, locationName, cache)

End Function

Public Function ResolveAreaId( _
    ByVal areaNumber As String, ByVal cache As Object) As String

    ResolveAreaId = ResolveIdByField( _
        OBJ_AREAS, FIELD_AREA_NUMBER, areaNumber, cache)

End Function

Public Function ResolvePeriodId( _
    ByVal periodDate As Date, ByVal cache As Object) As String

    ResolvePeriodId = ResolveIdByField( _
        OBJ_REPORTING_PERIODS, FIELD_PERIOD_DATE, _
        Format$(periodDate, "mm/dd/yyyy"), cache)

End Function

'=========================================================
' JSON HELPERS
'=========================================================

' Wraps a resolved Knack record ID as a connection field value:
' ["<id>"]. Returns vbNullString (field omitted) if id is blank -
' never writes an empty/invalid connection.
Public Function ConnectionJsonValue(ByVal recordId As String) As String
    If Len(Trim$(recordId)) = 0 Then Exit Function
    ConnectionJsonValue = "[""" & recordId & """]"
End Function

' Escapes a string for safe embedding inside a JSON string literal.
Public Function JsonEscape(ByVal s As String) As String
    Dim out As String
    out = s
    out = Replace(out, "\", "\\")
    out = Replace(out, """", "\""")
    out = Replace(out, vbCrLf, "\n")
    out = Replace(out, vbCr, "\n")
    out = Replace(out, vbLf, "\n")
    JsonEscape = out
End Function

' Formats a Variant cell value as a raw (unquoted) JSON value:
' numbers as numbers, blanks as 0, everything else as a quoted
' string. Use for plain number fields.
Public Function JsonNumber(ByVal cellValue As Variant) As String
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        JsonNumber = "0"
        Exit Function
    End If
    If Len(Trim$(CStr(cellValue))) = 0 Then
        JsonNumber = "0"
        Exit Function
    End If
    If IsNumeric(cellValue) Then
        JsonNumber = CStr(CDbl(cellValue))
    Else
        JsonNumber = "0"
    End If
End Function

' Formats a Variant cell value as a quoted JSON string.
Public Function JsonString(ByVal cellValue As Variant) As String
    If IsEmpty(cellValue) Or IsNull(cellValue) Then
        JsonString = """"""
        Exit Function
    End If
    JsonString = """" & JsonEscape(CStr(cellValue)) & """"
End Function

' Formats a date-only value the way Knack's API expects a date_time
' field with time ignored: {"date":"mm/dd/yyyy"}. Returns
' vbNullString (field omitted) if the cell isn't a real date -
' never writes an invalid or zero date.
Public Function JsonDateOnly(ByVal cellValue As Variant) As String
    If Not IsDate(cellValue) Then Exit Function
    Dim d As Date
    d = CDate(cellValue)
    If d = 0 Then Exit Function
    JsonDateOnly = "{""date"":""" & Format$(d, "mm/dd/yyyy") & """}"
End Function

'=========================================================
' JSON OBJECT BUILDER
'
' For records with many fields (Payroll has ~65), building one long
' concatenated string literal is error-prone. These let a caller
' append "fieldKey: jsonValue" fragments one at a time - a blank
' jsonValue is silently skipped (field omitted from the payload) -
' then join them into a single JSON object at the end.
'=========================================================

Public Function NewJsonFragments() As Collection
    Set NewJsonFragments = New Collection
End Function

Public Sub AppendJsonField( _
    ByVal fragments As Collection, _
    ByVal fieldKey As String, _
    ByVal jsonValue As String)

    If Len(jsonValue) = 0 Then Exit Sub
    fragments.Add """" & fieldKey & """:" & jsonValue

End Sub

Public Function BuildJsonObject(ByVal fragments As Collection) As String

    If fragments.count = 0 Then
        BuildJsonObject = "{}"
        Exit Function
    End If

    Dim parts() As String
    ReDim parts(1 To fragments.count)

    Dim i As Long
    For i = 1 To fragments.count
        parts(i) = fragments(i)
    Next i

    BuildJsonObject = "{" & Join(parts, ",") & "}"

End Function

'=========================================================
' RUN RESULT TRACKING + SUMMARY POPUP
'
' Every upload macro creates one of these (as a plain Dictionary,
' for simplicity) and calls RecordSuccess/RecordFailure per row,
' then ShowUploadSummary once at the end.
'=========================================================

Public Function NewUploadResult() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.Add "Created", CLng(0)
    result.Add "Skipped", CLng(0)
    result.Add "Errors", New Collection
    Set NewUploadResult = result
End Function

Public Sub RecordSuccess(ByVal result As Object)
    result("Created") = result("Created") + 1
End Sub

Public Sub RecordSkip(ByVal result As Object, ByVal rowNumber As Long, ByVal reason As String)
    result("Skipped") = result("Skipped") + 1
    result("Errors").Add "Row " & rowNumber & " skipped: " & reason
End Sub

Public Sub RecordFailure(ByVal result As Object, ByVal rowNumber As Long, ByVal reason As String)
    result("Errors").Add "Row " & rowNumber & " FAILED: " & reason
End Sub

Public Sub ShowUploadSummary( _
    ByVal title As String, _
    ByVal result As Object, _
    Optional ByVal extraNote As String = vbNullString)

    Dim msg As String

    msg = "Records created: " & result("Created") & vbCrLf & _
        "Rows skipped/failed: " & result("Skipped") & vbCrLf

    If Len(extraNote) > 0 Then
        msg = msg & vbCrLf & extraNote & vbCrLf
    End If

    Dim errorList As Collection
    Set errorList = result("Errors")

    If errorList.count > 0 Then

        msg = msg & vbCrLf & "Details:" & vbCrLf

        Dim shownCount As Long
        Dim i As Long

        For i = 1 To errorList.count

            msg = msg & "- " & errorList(i) & vbCrLf

            shownCount = shownCount + 1
            If shownCount >= 25 Then
                msg = msg & "... and " & (errorList.count - shownCount) & " more."
                Exit For
            End If

        Next i

    End If

    MsgBox msg, vbInformation, title

End Sub
