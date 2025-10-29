Attribute VB_Name = "Module1"
Sub CleanHRData()
    Dim ws As Worksheet, wsClean As Worksheet, wsIssues As Worksheet, wsSummary As Worksheet
    Dim lastRow As Long, r As Long, outRow As Long, issRow As Long
    Dim nameVal As String, emailVal As String, deptVal As String
    
    Set ws = ThisWorkbook.Sheets(1) ' Use the first sheet with raw data
    
    ' Create or reset sheets
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("Cleaned").Delete
    Sheets("Issues").Delete
    Sheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsClean = Sheets.Add: wsClean.Name = "Cleaned"
    Set wsIssues = Sheets.Add: wsIssues.Name = "Issues"
    Set wsSummary = Sheets.Add: wsSummary.Name = "Summary"
    
    ' Headers
    wsClean.Range("A1:F1").Value = Array("Name", "Email", "Department", "Phone", "StartDate", "Skills")
    wsIssues.Range("A1:H1").Value = Array("Row", "Name", "Email", "Department", "Phone", "StartDate", "Skills", "Issue")
    
    outRow = 2
    issRow = 2
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    
    For r = 2 To lastRow
        nameVal = StrConv(Trim(ws.Cells(r, 1).Value), vbProperCase)
        emailVal = LCase(Trim(ws.Cells(r, 2).Value))
        deptVal = ws.Cells(r, 3).Value
        
        ' Validate email
        If InStr(emailVal, "@") = 0 Or InStr(emailVal, ".") = 0 Then
            wsIssues.Cells(issRow, 1).Resize(1, 8).Value = _
                Array(r, nameVal, emailVal, deptVal, ws.Cells(r, 4).Value, ws.Cells(r, 5).Value, ws.Cells(r, 6).Value, "Invalid email")
            issRow = issRow + 1
            GoTo NextRow
        End If
        
        ' Skip duplicates (by email)
        If seen.Exists(emailVal) Then GoTo NextRow
        seen(emailVal) = True
        
        ' Write to Cleaned sheet
        wsClean.Cells(outRow, 1).Resize(1, 6).Value = _
            Array(nameVal, emailVal, deptVal, ws.Cells(r, 4).Value, ws.Cells(r, 5).Value, ws.Cells(r, 6).Value)
        outRow = outRow + 1
        
NextRow:
    Next r
    
    ' Build Summary
    wsSummary.Range("A1").Value = "Total Cleaned"
    wsSummary.Range("B1").Value = outRow - 2
    wsSummary.Range("A2").Value = "Invalid Emails"
    wsSummary.Range("B2").Value = issRow - 2
    wsSummary.Range("A3").Value = "Run Date"
    wsSummary.Range("B3").Value = Now
    
    MsgBox "Data cleaned. Check 'Cleaned', 'Issues', and 'Summary' sheets."
End Sub


