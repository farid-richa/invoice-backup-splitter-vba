Attribute VB_Name = "modEntityTabs"
Option Explicit

'======================================================================
'  GenerateEntityTabs
'  Splits the ALL sheet into one tab per billing entity and rule, then
'  builds SUMMARY (amounts per tab, intra-EU vs extra-EU) and CHECK
'  (live formulas that prove nothing was lost or double counted).
'
'  Routing priority for each line:
'    1. special rule in CONFIG_RULES (entity + origin + destination)
'    2. in-house group carrier (keyword in CONFIG_GENERAL)
'    3. carrier payment terms in CONFIG_TERMS
'    4. the entity's base tab
'  Lines with no entity go to UNASSIGNED instead of being skipped.
'======================================================================

Private Const SRC_SHEET As String = "ALL"
Private Const SUMMARY_SHEET As String = "SUMMARY"
Private Const CHECK_SHEET As String = "CHECK"
Private Const REGISTRY_SHEET As String = "_GENERATED"
Private Const UNASSIGNED_TAB As String = "UNASSIGNED"
' Tolerance as TEXT: concatenating a Double into a formula uses the local decimal
' separator (0,005 in Spanish Excel), which breaks the English .Formula syntax.
Private Const TOLERANCE_TXT As String = "0.005"

' Entity mapping loaded from CONFIG_ENTITIES
Private mapMatch() As String, mapType() As String, mapTab() As String, mapDisplay() As String
Private mapCount As Long
' Special rules loaded from CONFIG_RULES
Private ruleEntity() As String, ruleSuffix() As String
Private ruleOrigins() As Variant, ruleDests() As Variant   ' each item is a Collection
Private ruleCount As Long

Public Sub GenerateEntityTabs()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim calcMode As Long
    calcMode = Application.Calculation

    On Error GoTo Fail
    If Not SheetExists(wb, SRC_SHEET) Then
        MsgBox "Sheet '" & SRC_SHEET & "' not found.", vbCritical, "Invoice Backup Splitter"
        Exit Sub
    End If

    Dim wsSrc As Worksheet: Set wsSrc = wb.Worksheets(SRC_SHEET)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then
        MsgBox "No data in sheet '" & SRC_SHEET & "'.", vbExclamation, "Invoice Backup Splitter"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' ---- read everything into memory once (fast, no cell-by-cell access)
    Dim data As Variant
    data = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value

    Dim colEntity As Long, colCarrier As Long, colOrig As Long, colDest As Long, colAmount As Long
    colEntity = RequireColumn(data, GetSetting(wb, "Entity Column"))
    colCarrier = RequireColumn(data, GetSetting(wb, "Carrier Column"))
    colOrig = RequireColumn(data, GetSetting(wb, "Origin Column"))
    colDest = RequireColumn(data, GetSetting(wb, "Destination Column"))
    colAmount = RequireColumn(data, GetSetting(wb, "Amount Column"))

    Dim inhouseKey As String, inhouseSuffix As String
    inhouseKey = Norm(GetSetting(wb, "In-house Carrier Keyword", False))
    inhouseSuffix = Norm(GetSetting(wb, "In-house Tab Suffix", False))
    If Len(inhouseSuffix) = 0 Then inhouseSuffix = "GROUP"

    Dim euSet As Collection
    Set euSet = ListToSet(GetSetting(wb, "EU Countries", False))

    LoadEntityMap wb
    LoadRules wb
    Dim terms As Collection: Set terms = LoadTerms(wb)

    ' ---- route every line to a tab
    Dim n As Long: n = UBound(data, 1)
    Dim rowTab() As String, rowIntra() As String, rowAmount() As Double, rowBad() As Boolean
    ReDim rowTab(2 To n): ReDim rowIntra(2 To n): ReDim rowAmount(2 To n): ReDim rowBad(2 To n)

    Dim groups As New Collection     ' tab name -> Collection of row numbers
    Dim tabNames As New Collection   ' tab names in first-seen order
    Dim tabBase As New Collection    ' tab name -> base entity tab
    Dim tabSuffix As New Collection  ' tab name -> suffix ("" for base)

    Dim r As Long, baseTab As String, suffix As String, tabName As String
    Dim carrier As String, orig As String, dest As String
    Dim badAmounts As Long, newGroup As Collection

    For r = 2 To n
        baseTab = MapEntity(Norm(data(r, colEntity)))
        carrier = Norm(data(r, colCarrier))
        orig = Norm(data(r, colOrig))
        dest = Norm(data(r, colDest))

        suffix = ""
        If baseTab <> UNASSIGNED_TAB Then
            suffix = MatchRule(baseTab, orig, dest)
            If Len(suffix) = 0 Then
                If Len(inhouseKey) > 0 And InStr(1, carrier, inhouseKey, vbBinaryCompare) > 0 Then
                    suffix = inhouseSuffix
                ElseIf HasKey(terms, carrier) Then
                    suffix = CStr(terms(carrier)) & " DAYS"
                End If
            End If
        End If

        If Len(suffix) > 0 Then
            tabName = CleanSheetName(baseTab & " - " & suffix)
        Else
            tabName = CleanSheetName(baseTab)
        End If

        rowTab(r) = tabName
        rowIntra(r) = IIf(HasKey(euSet, orig) And HasKey(euSet, dest), "Y", "N")
        If IsNumeric(data(r, colAmount)) And Not IsEmpty(data(r, colAmount)) Then
            rowAmount(r) = CDbl(data(r, colAmount))
        Else
            rowBad(r) = True
            badAmounts = badAmounts + 1
        End If

        If Not HasKey(groups, tabName) Then
            Set newGroup = New Collection
            groups.Add newGroup, tabName
            tabNames.Add tabName
            tabBase.Add baseTab, tabName
            tabSuffix.Add suffix, tabName
        End If
        groups(tabName).Add r
    Next r

    ' ---- sorted list of tabs
    Dim tabs() As String, i As Long
    ReDim tabs(1 To tabNames.Count)
    For i = 1 To tabNames.Count
        tabs(i) = tabNames(i)
    Next i
    SortStrings tabs

    ' ---- remove tabs created by the previous run, so nothing stale is left
    DeletePreviousTabs wb

    ' ---- write one tab per group (header + 2 extra columns)
    Dim outCols As Long: outCols = lastCol + 2
    Dim ws As Worksheet, rowsIn As Collection, outArr As Variant
    Dim k As Long, c As Long, srcRow As Long
    Dim prevSheet As Worksheet: Set prevSheet = wsSrc

    For i = 1 To UBound(tabs)
        Set rowsIn = groups(tabs(i))
        ReDim outArr(1 To rowsIn.Count + 1, 1 To outCols)
        For c = 1 To lastCol
            outArr(1, c) = data(1, c)
        Next c
        outArr(1, lastCol + 1) = "Target Tab"
        outArr(1, lastCol + 2) = "Intra-EU"
        For k = 1 To rowsIn.Count
            srcRow = rowsIn(k)
            For c = 1 To lastCol
                outArr(k + 1, c) = data(srcRow, c)
            Next c
            outArr(k + 1, lastCol + 1) = tabs(i)
            outArr(k + 1, lastCol + 2) = rowIntra(srcRow)
        Next k

        Set ws = GetOrCreateSheet(wb, tabs(i), prevSheet)
        ws.Cells.Clear
        ws.Range(ws.Cells(1, 1), ws.Cells(rowsIn.Count + 1, outCols)).Value = outArr
        FormatDataSheet ws, outCols, colAmount
        Set prevSheet = ws
    Next i
    WriteRegistry wb, tabs

    ' ---- SUMMARY and CHECK
    BuildSummary wb, tabs, groups, tabBase, tabSuffix, rowAmount, rowIntra, rowBad
    BuildCheck wb, tabs, colAmount, lastRow - 1, badAmounts

    ' ---- order: ALL, SUMMARY, CHECK, generated tabs
    wb.Worksheets(SUMMARY_SHEET).Move After:=wsSrc
    wb.Worksheets(CHECK_SHEET).Move After:=wb.Worksheets(SUMMARY_SHEET)
    Set prevSheet = wb.Worksheets(CHECK_SHEET)
    For i = 1 To UBound(tabs)
        wb.Worksheets(tabs(i)).Move After:=prevSheet
        Set prevSheet = wb.Worksheets(tabs(i))
    Next i

    Application.Calculation = calcMode
    Application.Calculate
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wb.Worksheets(CHECK_SHEET).Activate

    MsgBox "Done: " & UBound(tabs) & " tabs from " & (lastRow - 1) & " lines." & vbCrLf & _
           "CHECK status: " & wb.Worksheets(CHECK_SHEET).Range("B3").Value, _
           vbInformation, "Invoice Backup Splitter"
    Exit Sub

Fail:
    Application.Calculation = calcMode
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Stopped: " & Err.Description, vbCritical, "Invoice Backup Splitter"
End Sub

'----------------------------------------------------------------------
'  Configuration
'----------------------------------------------------------------------
Private Function RequireColumn(ByRef data As Variant, ByVal header As String) As Long
    RequireColumn = FindHeaderInArray(data, header)
    If RequireColumn = 0 Then
        Err.Raise vbObjectError + 200, , "Column '" & header & "' not found in sheet ALL."
    End If
End Function

Private Sub LoadEntityMap(ByVal wb As Workbook)
    Dim ws As Worksheet, lastRow As Long, r As Long
    mapCount = 0
    If Not SheetExists(wb, "CONFIG_ENTITIES") Then Exit Sub
    Set ws = wb.Worksheets("CONFIG_ENTITIES")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ReDim mapMatch(1 To lastRow - 1): ReDim mapType(1 To lastRow - 1)
    ReDim mapTab(1 To lastRow - 1): ReDim mapDisplay(1 To lastRow - 1)
    For r = 2 To lastRow
        If Len(Norm(ws.Cells(r, 1).Value)) > 0 Then
            mapCount = mapCount + 1
            mapMatch(mapCount) = Norm(ws.Cells(r, 1).Value)
            mapType(mapCount) = Norm(ws.Cells(r, 2).Value)
            mapTab(mapCount) = Norm(ws.Cells(r, 3).Value)
            mapDisplay(mapCount) = Trim$(CStr(ws.Cells(r, 4).Value))
        End If
    Next r
End Sub

' Raw entity name -> base tab. Unknown names keep their own (normalised) name.
Private Function MapEntity(ByVal entity As String) As String
    Dim i As Long
    If Len(entity) = 0 Then
        MapEntity = UNASSIGNED_TAB
        Exit Function
    End If
    For i = 1 To mapCount
        If mapType(i) = "CONTAINS" Then
            If InStr(1, entity, mapMatch(i), vbBinaryCompare) > 0 Then MapEntity = mapTab(i): Exit Function
        Else
            If entity = mapMatch(i) Then MapEntity = mapTab(i): Exit Function
        End If
    Next i
    MapEntity = entity
End Function

Private Function DisplayNameFor(ByVal baseTab As String) As String
    Dim i As Long
    For i = 1 To mapCount
        If mapTab(i) = baseTab And Len(mapDisplay(i)) > 0 Then
            DisplayNameFor = mapDisplay(i)
            Exit Function
        End If
    Next i
    DisplayNameFor = baseTab
End Function

Private Sub LoadRules(ByVal wb As Workbook)
    Dim ws As Worksheet, lastRow As Long, r As Long
    ruleCount = 0
    If Not SheetExists(wb, "CONFIG_RULES") Then Exit Sub
    Set ws = wb.Worksheets("CONFIG_RULES")
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    ReDim ruleEntity(1 To lastRow - 1): ReDim ruleSuffix(1 To lastRow - 1)
    ReDim ruleOrigins(1 To lastRow - 1): ReDim ruleDests(1 To lastRow - 1)
    For r = 2 To lastRow
        If Len(Norm(ws.Cells(r, 2).Value)) > 0 And Len(Norm(ws.Cells(r, 5).Value)) > 0 Then
            ruleCount = ruleCount + 1
            ruleEntity(ruleCount) = Norm(ws.Cells(r, 2).Value)
            Set ruleOrigins(ruleCount) = ListToSet(CStr(ws.Cells(r, 3).Value))
            Set ruleDests(ruleCount) = ListToSet(CStr(ws.Cells(r, 4).Value))
            ruleSuffix(ruleCount) = Norm(ws.Cells(r, 5).Value)
        End If
    Next r
End Sub

' First matching rule wins. An empty country list means "any country".
Private Function MatchRule(ByVal baseTab As String, ByVal orig As String, ByVal dest As String) As String
    Dim i As Long
    For i = 1 To ruleCount
        If ruleEntity(i) = baseTab Then
            If (ruleOrigins(i).Count = 0 Or HasKey(ruleOrigins(i), orig)) And _
               (ruleDests(i).Count = 0 Or HasKey(ruleDests(i), dest)) Then
                MatchRule = ruleSuffix(i)
                Exit Function
            End If
        End If
    Next i
End Function

' Carrier name -> pay days, from CONFIG_TERMS. Any number of days works.
Private Function LoadTerms(ByVal wb As Workbook) As Collection
    Dim result As New Collection, ws As Worksheet, lastRow As Long, r As Long
    Dim carrier As String, days As Long
    If SheetExists(wb, "CONFIG_TERMS") Then
        Set ws = wb.Worksheets("CONFIG_TERMS")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastRow
            carrier = Norm(ws.Cells(r, 1).Value)
            days = CLng(Val(CStr(ws.Cells(r, 2).Value)))
            If Len(carrier) > 0 And days > 0 Then
                If Not HasKey(result, carrier) Then result.Add days, carrier
            End If
        Next r
    End If
    Set LoadTerms = result
End Function

'----------------------------------------------------------------------
'  Sheets
'----------------------------------------------------------------------
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String, _
                                  ByVal afterSheet As Worksheet) As Worksheet
    If SheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=afterSheet)
        GetOrCreateSheet.Name = sheetName
    End If
    GetOrCreateSheet.Visible = xlSheetVisible
End Function

' The registry is a hidden sheet listing the tabs this macro created,
' written in one go so the next run knows exactly what to remove.
Private Sub WriteRegistry(ByVal wb As Workbook, ByRef tabs() As String)
    Dim reg As Worksheet, outArr As Variant, i As Long
    If Not SheetExists(wb, REGISTRY_SHEET) Then
        Set reg = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        reg.Name = REGISTRY_SHEET
    End If
    Set reg = wb.Worksheets(REGISTRY_SHEET)
    reg.Cells.ClearContents
    ReDim outArr(1 To UBound(tabs) + 1, 1 To 1)
    outArr(1, 1) = "Tabs created by GenerateEntityTabs (do not edit)"
    For i = 1 To UBound(tabs)
        outArr(i + 1, 1) = tabs(i)
    Next i
    reg.Range(reg.Cells(1, 1), reg.Cells(UBound(tabs) + 1, 1)).Value = outArr
    reg.Visible = xlSheetHidden
End Sub

Private Sub DeletePreviousTabs(ByVal wb As Workbook)
    Dim reg As Worksheet, lastRow As Long, r As Long, nm As String
    If Not SheetExists(wb, REGISTRY_SHEET) Then Exit Sub
    Set reg = wb.Worksheets(REGISTRY_SHEET)
    lastRow = Application.WorksheetFunction.CountA(reg.Columns(1))
    For r = 2 To lastRow
        nm = CStr(reg.Cells(r, 1).Value)
        If Len(nm) > 0 And SheetExists(wb, nm) Then
            If Not IsProtectedSheet(nm) Then wb.Worksheets(nm).Delete
        End If
    Next r
    If lastRow >= 2 Then reg.Range(reg.Cells(2, 1), reg.Cells(lastRow, 1)).ClearContents
End Sub

Private Function IsProtectedSheet(ByVal nm As String) As Boolean
    Select Case UCase$(nm)
        Case "README", SRC_SHEET, SUMMARY_SHEET, CHECK_SHEET, REGISTRY_SHEET, "EXPORT_LOG"
            IsProtectedSheet = True
        Case Else
            IsProtectedSheet = (Left$(UCase$(nm), 7) = "CONFIG_")
    End Select
End Function

Private Sub FormatDataSheet(ByVal ws As Worksheet, ByVal lastCol As Long, ByVal colAmount As Long)
    StyleHeader ws, lastCol
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
    ws.Columns(colAmount).NumberFormat = "#,##0.00"
    ws.Columns.AutoFit
End Sub

'----------------------------------------------------------------------
'  SUMMARY: one row per tab, totals split intra-EU / extra-EU
'----------------------------------------------------------------------
Private Sub BuildSummary(ByVal wb As Workbook, ByRef tabs() As String, ByVal groups As Collection, _
                         ByVal tabBase As Collection, ByVal tabSuffix As Collection, _
                         ByRef rowAmount() As Double, ByRef rowIntra() As String, ByRef rowBad() As Boolean)
    Dim ws As Worksheet, i As Long, k As Long, srcRow As Long, outRow As Long
    Dim lines As Long, amt As Double, intraAmt As Double, extraAmt As Double
    Dim totLines As Long, totAmt As Double, totIntra As Double, totExtra As Double
    Dim label As String

    If SheetExists(wb, SUMMARY_SHEET) Then
        Set ws = wb.Worksheets(SUMMARY_SHEET)
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(SRC_SHEET))
        ws.Name = SUMMARY_SHEET
    End If

    ws.Range("A1:G1").Value = Array("Tab", "Billing Entity", "Routing", "Lines", _
                                    "Amount EUR", "Intra-EU EUR", "Extra-EU EUR")
    outRow = 2
    For i = 1 To UBound(tabs)
        lines = 0: amt = 0: intraAmt = 0: extraAmt = 0
        For k = 1 To groups(tabs(i)).Count
            srcRow = groups(tabs(i))(k)
            lines = lines + 1
            If Not rowBad(srcRow) Then
                amt = amt + rowAmount(srcRow)
                If rowIntra(srcRow) = "Y" Then
                    intraAmt = intraAmt + rowAmount(srcRow)
                Else
                    extraAmt = extraAmt + rowAmount(srcRow)
                End If
            End If
        Next k

        If tabBase(tabs(i)) = UNASSIGNED_TAB Then
            label = "Needs review: no billing entity"
        Else
            label = DisplayNameFor(tabBase(tabs(i)))
        End If

        ws.Cells(outRow, 1).Value = tabs(i)
        ws.Cells(outRow, 2).Value = label
        ws.Cells(outRow, 3).Value = IIf(Len(tabSuffix(tabs(i))) = 0, "Base", tabSuffix(tabs(i)))
        ws.Cells(outRow, 4).Value = lines
        ws.Cells(outRow, 5).Value = Round(amt, 2)
        ws.Cells(outRow, 6).Value = Round(intraAmt, 2)
        ws.Cells(outRow, 7).Value = Round(extraAmt, 2)
        totLines = totLines + lines: totAmt = totAmt + amt
        totIntra = totIntra + intraAmt: totExtra = totExtra + extraAmt
        outRow = outRow + 1
    Next i

    ws.Cells(outRow, 1).Value = "TOTAL"
    ws.Cells(outRow, 4).Value = totLines
    ws.Cells(outRow, 5).Value = Round(totAmt, 2)
    ws.Cells(outRow, 6).Value = Round(totIntra, 2)
    ws.Cells(outRow, 7).Value = Round(totExtra, 2)
    ws.Rows(outRow).Font.Bold = True

    StyleHeader ws, 7
    ws.Range("D:D").NumberFormat = "#,##0"
    ws.Range("E:G").NumberFormat = "#,##0.00"
    ws.Columns("A:G").AutoFit
End Sub

'----------------------------------------------------------------------
'  CHECK: live formulas, so the reconciliation stays true even if
'  someone edits a tab by hand after the macro ran.
'----------------------------------------------------------------------
Private Sub BuildCheck(ByVal wb As Workbook, ByRef tabs() As String, ByVal colAmount As Long, _
                       ByVal sourceLines As Long, ByVal badAmounts As Long)
    Dim ws As Worksheet, i As Long, r As Long, amtCol As String, firstTabRow As Long, lastTabRow As Long
    amtCol = ColLetter(colAmount)

    If SheetExists(wb, CHECK_SHEET) Then
        Set ws = wb.Worksheets(CHECK_SHEET)
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(SUMMARY_SHEET))
        ws.Name = CHECK_SHEET
    End If

    ws.Range("A1").Value = "Reconciliation"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "Overall status"
    ws.Range("A3").Font.Bold = True

    ' per-tab block
    ws.Range("A12:E12").Value = Array("Tab", "SUMMARY amount", "Live sum of tab", "Difference", "Status")
    StyleHeaderRange ws, "A12:E12"
    firstTabRow = 13
    r = firstTabRow
    For i = 1 To UBound(tabs)
        ws.Cells(r, 1).Value = tabs(i)
        ws.Cells(r, 2).Formula = "=SUMIF(" & SUMMARY_SHEET & "!A:A,A" & r & "," & SUMMARY_SHEET & "!E:E)"
        ws.Cells(r, 3).Formula = "=SUM('" & tabs(i) & "'!" & amtCol & ":" & amtCol & ")"
        ws.Cells(r, 4).Formula = "=ROUND(C" & r & "-B" & r & ",2)"
        ws.Cells(r, 5).Formula = "=IF(ABS(D" & r & ")<" & TOLERANCE_TXT & ",""OK"",""REVIEW"")"
        r = r + 1
    Next i
    lastTabRow = r - 1

    ' totals block
    ws.Range("A5:E5").Value = Array("Control", "Expected", "Actual", "Difference", "Status")
    StyleHeaderRange ws, "A5:E5"

    ws.Range("A6").Value = "Amount: ALL sheet vs sum of tabs"
    ws.Range("B6").Formula = "=SUM(" & SRC_SHEET & "!" & amtCol & ":" & amtCol & ")"
    ws.Range("C6").Formula = "=SUM(C" & firstTabRow & ":C" & lastTabRow & ")"

    ws.Range("A7").Value = "Amount: SUMMARY total vs sum of tabs"
    ws.Range("B7").Formula = "=SUMIF(" & SUMMARY_SHEET & "!A:A,""TOTAL""," & SUMMARY_SHEET & "!E:E)"
    ws.Range("C7").Formula = "=C6"

    ws.Range("A8").Value = "Lines: ALL sheet vs sum of tabs"
    ws.Range("B8").Value = sourceLines
    ws.Range("C8").Formula = "=SUMIF(" & SUMMARY_SHEET & "!A:A,""<>TOTAL""," & SUMMARY_SHEET & "!D:D)"

    ws.Range("A9").Value = "Lines with no billing entity (UNASSIGNED)"
    ws.Range("B9").Value = 0
    ws.Range("C9").Formula = "=SUMIF(" & SUMMARY_SHEET & "!A:A,""" & UNASSIGNED_TAB & """," & SUMMARY_SHEET & "!D:D)"

    ws.Range("A10").Value = "Lines with a non-numeric amount"
    ws.Range("B10").Value = 0
    ws.Range("C10").Value = badAmounts

    ' rows 6-8 must reconcile (REVIEW if not); rows 9-10 are data-quality warnings
    For r = 6 To 10
        ws.Cells(r, 4).Formula = "=ROUND(C" & r & "-B" & r & ",2)"
        If r <= 8 Then
            ws.Cells(r, 5).Formula = "=IF(ABS(D" & r & ")<" & TOLERANCE_TXT & ",""OK"",""REVIEW"")"
        Else
            ws.Cells(r, 5).Formula = "=IF(C" & r & "=0,""OK"",""WARNING"")"
        End If
    Next r

    ws.Range("B3").Formula = "=IF(COUNTIF(E6:E8,""REVIEW"")+COUNTIF(E" & firstTabRow & ":E" & lastTabRow & _
                             ",""REVIEW"")>0,""REVIEW"",IF(COUNTIF(E9:E10,""WARNING"")>0," & _
                             """OK - see warnings"",""OK""))"
    ws.Range("B3").Font.Bold = True

    ws.Range("B6:D7").NumberFormat = "#,##0.00"
    ws.Range("B" & firstTabRow & ":D" & lastTabRow).NumberFormat = "#,##0.00"
    ws.Columns("A:E").AutoFit
End Sub

Private Sub StyleHeaderRange(ByVal ws As Worksheet, ByVal addr As String)
    With ws.Range(addr)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(12, 24, 43)
    End With
End Sub
