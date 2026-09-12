Attribute VB_Name = "RangeCells"
Option Explicit

Public Sub WriteToRange()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    ws.Range("H1").Value = "Poznámka"
    ws.Range("H2").Value = "Zapísané cez Range"
End Sub

Public Sub FillWithCells()
    Dim ws As Worksheet
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets("Data")

    For i = 2 To 11
        ws.Cells(i, 9).Value = i - 1
    Next i

    ws.Cells(1, 9).Value = "Poradie"
End Sub

Public Function LastUsedRow(ByVal ws As Worksheet, Optional ByVal columnNumber As Long = 1) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
End Function
