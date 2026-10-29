Attribute VB_Name = "modCarrierExport"
Option Explicit

'======================================================================
'  Export one file per carrier with its pending or rejected invoices,
'  ready to send. One macro with a parameter replaces a copy per status.
'
'  Output:  <workbook folder>\exports\YYYYMMDD\
'           YYYYMMDD <CarrierID> <Carrier Name> PENDING_INVOICES.xlsx
'  Each file is written once, with all its lines (no reopening per row),
'  and re-running the same day overwrites instead of duplicating lines.
'  The ALL sheet is not modified. Every file is listed in EXPORT_LOG.
'======================================================================

Private Const SRC_SHEET As String = "ALL"
Private Const LOG_SHEET As String = "EXPORT_LOG"
Private Const FORMAT_XLSX As Long = 51   ' same value as Excel's xlOpenXMLWorkbook

Public Sub ExportPendingInvoices()
    ExportByStatus "PENDING"
End Sub

Public Sub ExportRejectedInvoices()
    ExportByStatus "REJECTED"
End Sub

Public Sub ExportByStatus(ByVal statusWanted As String)
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sep As String: sep = Application.PathSeparator
    statusWanted = Norm(statusWanted)

    On Error GoTo Fail
    If Len(wb.Path) = 0 Then
        MsgBox "Save this workbook first: files are exported next to it.", vbExclamation, "Carrier export"
        Exit Sub
    End If
    If Not SheetExists(wb, SRC_SHEET) Then
        MsgBox "Sheet '" & SRC_SHEET & "' not found.", vbCritical, "Carrier export"
        Exit Sub
    End If

    Dim wsSrc As Worksheet: Set wsSrc = wb.Worksheets(SRC_SHEET)
    Dim lastRow As Long, lastCol As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    If lastRow < 2 Then Exit Sub

    Dim data As Variant
    data = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value

    Dim colStatus As Long, colCarrier As Long, colCarrierId As Long, colAmount As Long
    colStatus = FindHeaderInArray(data, GetSetting(wb, "Status Column"))
    colCarrier = FindHeaderInArray(data, GetSetting(wb, "Carrier Column"))
    colCarrierId = FindHeaderInArray(data, GetSetting(wb, "Carrier ID Column", False))
    colAmount = FindHeaderInArray(data, GetSetting(wb, "Amount Column"))
    If colStatus = 0 Or colCarrier = 0 Or colAmount = 0 Then
        Err.Raise vbObjectError + 300, , "Status, carrier or amount column not found in sheet ALL."
    End If

    ' ---- group matching lines by carrier
    Dim groups As New Collection, keys As New Collection, labels As New Collection
    Dim r As Long, key As String, carrierName As String, carrierId As String
    Dim newGroup As Collection
    For r = 2 To lastRow
        If Norm(data(r, colStatus)) = statusWanted Then
            carrierName = Norm(data(r, colCarrier))
            carrierId = ""
            If colCarrierId > 0 Then carrierId = Norm(data(r, colCarrierId))
            key = carrierId & "|" & carrierName
            If Not HasKey(groups, key) Then
                Set newGroup = New Collection
                groups.Add newGroup, key
                keys.Add key
                labels.Add Trim$(carrierId & " " & carrierName), key
            End If
            groups(key).Add r
        End If
    Next r

    If keys.Count = 0 Then
        MsgBox "No " & statusWanted & " invoices found.", vbInformation, "Carrier export"
        Exit Sub
    End If

    ' ---- output folder
    Dim stamp As String, folder As String, exportRoot As String
    stamp = Format(Date, "yyyymmdd")
    exportRoot = GetSetting(wb, "Export Folder", False)
    If Len(exportRoot) = 0 Then exportRoot = "exports"
    folder = wb.Path & sep & exportRoot & sep & stamp
    EnsureFolder folder

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- one workbook per carrier, written in one go
    Dim i As Long, k As Long, c As Long, rowsIn As Collection, outArr As Variant
    Dim newWb As Workbook, ws As Worksheet, filePath As String, amt As Double
    Dim filesWritten As Long, linesWritten As Long

    For i = 1 To keys.Count
        key = keys(i)
        Set rowsIn = groups(key)
        ReDim outArr(1 To rowsIn.Count + 1, 1 To lastCol)
        For c = 1 To lastCol
            outArr(1, c) = data(1, c)
        Next c
        amt = 0
        For k = 1 To rowsIn.Count
            For c = 1 To lastCol
                outArr(k + 1, c) = data(rowsIn(k), c)
            Next c
            If IsNumeric(data(rowsIn(k), colAmount)) Then amt = amt + CDbl(data(rowsIn(k), colAmount))
        Next k

        Set newWb = Workbooks.Add(xlWBATWorksheet)
        Set ws = newWb.Worksheets(1)
        ws.Name = statusWanted & " INVOICES"
        ws.Range(ws.Cells(1, 1), ws.Cells(rowsIn.Count + 1, lastCol)).Value = outArr
        StyleHeader ws, lastCol
        ws.Columns(colAmount).NumberFormat = "#,##0.00"
        ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).AutoFilter
        ws.Columns.AutoFit

        filePath = folder & sep & CleanFileName(stamp & " " & labels(key) & " " & statusWanted & "_INVOICES") & ".xlsx"
        newWb.SaveAs Filename:=filePath, FileFormat:=FORMAT_XLSX
        newWb.Close SaveChanges:=False

        LogExport wb, statusWanted, Split(key, "|")(0), Split(key, "|")(1), rowsIn.Count, amt, filePath
        filesWritten = filesWritten + 1
        linesWritten = linesWritten + rowsIn.Count
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox filesWritten & " files written (" & linesWritten & " " & LCase$(statusWanted) & _
           " lines)" & vbCrLf & "Folder: " & folder, vbInformation, "Carrier export"
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Stopped: " & Err.Description, vbCritical, "Carrier export"
End Sub

Private Sub LogExport(ByVal wb As Workbook, ByVal status As String, ByVal carrierId As String, _
                      ByVal carrierName As String, ByVal lines As Long, ByVal amount As Double, _
                      ByVal filePath As String)
    Dim ws As Worksheet, nextRow As Long
    If Not SheetExists(wb, LOG_SHEET) Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(SRC_SHEET))
        ws.Name = LOG_SHEET
        ws.Range("A1:G1").Value = Array("Run Date", "Status", "Carrier ID", "Carrier Name", _
                                        "Lines", "Amount EUR", "File")
        StyleHeader ws, 7
    End If
    Set ws = wb.Worksheets(LOG_SHEET)
    ' CountA instead of End(xlUp): reliable even when another workbook is active
    nextRow = Application.WorksheetFunction.CountA(ws.Columns(1)) + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(nextRow, 2).Value = status
    ws.Cells(nextRow, 3).Value = carrierId
    ws.Cells(nextRow, 4).Value = carrierName
    ws.Cells(nextRow, 5).Value = lines
    ws.Cells(nextRow, 6).Value = Round(amount, 2)
    ws.Cells(nextRow, 6).NumberFormat = "#,##0.00"
    ws.Cells(nextRow, 7).Value = filePath
End Sub
