Attribute VB_Name = "ConditionsLoops"
Option Explicit

Public Sub CheckRevenue()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim revenue As Double

    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Cells(1, 8).Value = "Status"

    For i = 2 To lastRow
        revenue = CDbl(ws.Cells(i, 5).Value) * CDbl(ws.Cells(i, 6).Value)

        If revenue >= 1000 Then
            ws.Cells(i, 8).Value = "Vysoká tržba"
        ElseIf revenue >= 500 Then
            ws.Cells(i, 8).Value = "Stredná tržba"
        Else
            ws.Cells(i, 8).Value = "Nízka tržba"
        End If
    Next i
End Sub

Public Sub NumberRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Cells(1, 9).Value = "RowNo"

    For i = 2 To lastRow
        ws.Cells(i, 9).Value = i - 1
    Next i
End Sub

Public Sub HighlightNegativeValues()
    Dim cell As Range

    For Each cell In ThisWorkbook.Worksheets("Report").Range("F2:F500")
        If IsNumeric(cell.Value) And cell.Value < 0 Then
            cell.Font.Bold = True
            cell.Font.Color = vbRed
        End If
    Next cell
End Sub

Public Sub ProcessUntilBlank()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Data")
    i = 2

    Do While ws.Cells(i, 1).Value <> ""
        ws.Cells(i, 10).Value = UCase$(CStr(ws.Cells(i, 2).Value))
        i = i + 1
    Loop

    ws.Cells(1, 10).Value = "RegionUpper"
End Sub
