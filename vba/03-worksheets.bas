Attribute VB_Name = "WorksheetTools"
Option Explicit

Public Function GetOrCreateWorksheet(ByVal worksheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(worksheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = worksheetName
    End If

    Set GetOrCreateWorksheet = ws
End Function

Public Sub CreateReportSheet()
    Dim ws As Worksheet
    Set ws = GetOrCreateWorksheet("Report")

    ws.Cells.Clear
    ws.Range("A1").Value = "Sales Report"
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit
End Sub

Public Sub CountWorksheets()
    MsgBox "Počet hárkov: " & ThisWorkbook.Worksheets.Count, vbInformation, "excel-lab"
End Sub
