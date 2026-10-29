Attribute VB_Name = "modHelpers"
Option Explicit

'======================================================================
'  Shared helpers: text normalisation, lookups, sheet and file names.
'  Uses VBA Collections instead of Scripting.Dictionary so the macros
'  also run on Excel for Mac (no Windows scripting runtime needed).
'======================================================================

' Normalise text before comparing: non-breaking spaces, tabs, line breaks,
' double spaces and case. Billing exports are full of these.
Public Function Norm(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    t = CStr(v)
    t = Replace(t, ChrW(160), " ")   ' ChrW, not Chr: Chr(160) is not a non-breaking space on every platform
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Norm = UCase$(Trim$(t))
End Function

' True if the Collection has an item stored under this key.
' Works for both plain values and objects stored as items.
Public Function HasKey(ByVal col As Collection, ByVal key As String) As Boolean
    On Error Resume Next
    Err.Clear
    Dim probe As Boolean
    probe = IsObject(col.Item(key))
    HasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' Column number of a header in row 1 of a 2-D array (0 if not found).
' Compares normalised text, so "Amount EUR " still matches "Amount EUR".
Public Function FindHeaderInArray(ByRef data As Variant, ByVal headerText As String) As Long
    Dim c As Long, target As String
    target = Norm(headerText)
    For c = LBound(data, 2) To UBound(data, 2)
        If Norm(data(1, c)) = target Then
            FindHeaderInArray = c
            Exit Function
        End If
    Next c
End Function

' Value of a setting in CONFIG_GENERAL (column A = setting, column B = value).
Public Function GetSetting(ByVal wb As Workbook, ByVal settingName As String, _
                           Optional ByVal required As Boolean = True) As String
    Dim ws As Worksheet, r As Long, lastRow As Long
    If Not SheetExists(wb, "CONFIG_GENERAL") Then
        Err.Raise vbObjectError + 100, , "Missing sheet CONFIG_GENERAL."
    End If
    Set ws = wb.Worksheets("CONFIG_GENERAL")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Norm(ws.Cells(r, 1).Value) = Norm(settingName) Then
            GetSetting = Trim$(CStr(ws.Cells(r, 2).Value))
            Exit Function
        End If
    Next r
    If required Then
        Err.Raise vbObjectError + 101, , "Setting '" & settingName & "' not found in CONFIG_GENERAL."
    End If
End Function

' Comma-separated list -> Collection keyed by each normalised item.
Public Function ListToSet(ByVal csv As String) As Collection
    Dim result As New Collection, parts As Variant, i As Long, code As String
    If Len(Trim$(csv)) > 0 Then
        parts = Split(csv, ",")
        For i = LBound(parts) To UBound(parts)
            code = Norm(parts(i))
            If Len(code) > 0 Then
                If Not HasKey(result, code) Then result.Add code, code
            End If
        Next i
    End If
    Set ListToSet = result
End Function

Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Excel sheet names: no / \ ? * [ ] : ' and at most 31 characters.
Public Function CleanSheetName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("/", "\", "?", "*", "[", "]", ":", "'")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), " ")
    Next i
    s = Norm(s)
    If Len(s) > 31 Then s = Trim$(Left$(s, 31))
    CleanSheetName = s
End Function

' File names: remove characters Windows and macOS do not allow.
Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), " ")
    Next i
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanFileName = Trim$(s)
End Function

' Column number -> letters (1 -> A, 28 -> AB).
Public Function ColLetter(ByVal colNum As Long) As String
    Dim s As String, m As Long
    Do While colNum > 0
        m = (colNum - 1) Mod 26
        s = Chr(65 + m) & s
        colNum = (colNum - m - 1) \ 26
    Loop
    ColLetter = s
End Function

' Create a folder and any missing parent folders.
Public Sub EnsureFolder(ByVal folderPath As String)
    Dim sep As String, parts As Variant, i As Long, current As String
    sep = Application.PathSeparator
    parts = Split(folderPath, sep)
    For i = LBound(parts) To UBound(parts)
        If i = LBound(parts) Then
            current = parts(i)
        Else
            current = current & sep & parts(i)
        End If
        If Len(current) > 0 And Right$(current, 1) <> ":" Then
            If Len(Dir(current, vbDirectory)) = 0 Then
                On Error Resume Next
                MkDir current
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

' Simple in-place sort for a small 1-D array of strings.
Public Sub SortStrings(ByRef arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

Public Sub StyleHeader(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(12, 24, 43)
    End With
End Sub
