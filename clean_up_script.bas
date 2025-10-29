Attribute VB_Name = "Module1"
Sub CleanHRData()
    Dim wb As Workbook
    Dim ws As Worksheet, wsClean As Worksheet, wsIssues As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long, issRow As Long
    Dim nameVal As String, emailVal As String, deptVal As String
    Dim seen As Object
    
    Set wb = ActiveWorkbook               ' the workbook you can see
    Set ws = wb.Worksheets("Employees")   ' <-- change if your sheet name differs

    ' Guard: sheet must have data
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        MsgBox "The 'Employees' sheet is empty.", vbExclamation: Exit Sub
    End If

    ' Find the last used row safely (column A must have the Name header)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows under the header.", vbExclamation: Exit Sub
    End If

    ' Recreate output sheets
    On Error Resume Next
    Application.DisplayAlerts = False
    wb.Worksheets("Cleaned").Delete
    wb.Worksheets("Issues").Delete
    wb.Worksheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsClean = wb.Worksheets.Add(After:=ws): wsClean.Name = "Cleaned"
    Set wsIssues = wb.Worksheets.Add(After:=wsClean): wsIssues.Name = "Issues"
    Set wsSummary = wb.Worksheets.Add(After:=wsIssues): wsSummary.Name = "Summary"

    ' Headers
    wsClean.Range("A1:F1").Value = Array("Name", "Email", "Department", "Phone", "StartDate", "Skills")
    wsIssues.Range("A1:H1").Value = Array("Row", "Name", "Email", "Department", "Phone", "StartDate", "Skills", "Issue")

    Set seen = CreateObject("Scripting.Dictionary")
    outRow = 2: issRow = 2

    For r = 2 To lastRow
        nameVal = StrConv(Trim(ws.Cells(r, 1).Value), vbProperCase)
        emailVal = LCase(Trim(ws.Cells(r, 2).Value))
        deptVal = ws.Cells(r, 3).Value

        ' simple email check
        If InStr(emailVal, "@") = 0 Or InStr(emailVal, ".") = 0 Then
            wsIssues.Cells(issRow, 1).Resize(1, 8).Value = _
                Array(r, nameVal, emailVal, deptVal, ws.Cells(r, 4).Value, ws.Cells(r, 5).Value, ws.Cells(r, 6).Value, "Invalid email")
            issRow = issRow + 1
            GoTo NextRow
        End If

        If seen.Exists(emailVal) Then GoTo NextRow
        seen(emailVal) = True

        wsClean.Cells(outRow, 1).Resize(1, 6).Value = _
            Array(nameVal, emailVal, deptVal, ws.Cells(r, 4).Value, ws.Cells(r, 5).Value, ws.Cells(r, 6).Value)
        outRow = outRow + 1
NextRow:
    Next r

    ' Summary
    wsSummary.Range("A1").Value = "Metric": wsSummary.Range("B1").Value = "Value"
    wsSummary.Range("A2:B4").Value = Array( _
        Array("Total Cleaned", outRow - 2), _
        Array("Invalid Emails", issRow - 2), _
        Array("Run Date", Now))

    ' Pretty tables
    FormatAsTable wsClean, "CleanedTable", "TableStyleMedium9"
    FormatAsTable wsIssues, "IssuesTable", "TableStyleMedium3"
    FormatAsTable wsSummary, "SummaryTable", "TableStyleMedium7"

    MsgBox "Done. See Cleaned / Issues / Summary.", vbInformation
End Sub

Private Sub FormatAsTable(ws As Worksheet, tblName As String, styleName As String)
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    Dim lo As ListObject
    ' remove previous table if exists with same name
    On Error Resume Next
    ws.ListObjects(tblName).Unlist
    On Error GoTo 0

    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    lo.Name = tblName
    lo.TableStyle = styleName
End Sub


