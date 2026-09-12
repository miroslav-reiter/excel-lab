Attribute VB_Name = "PodmienkyCykly"
Option Explicit

Public Sub SkontrolujTrzbu()
    Dim harok As Worksheet
    Dim poslednyRiadok As Long
    Dim i As Long
    Dim trzba As Double

    Set harok = ThisWorkbook.Worksheets("Dáta")
    poslednyRiadok = harok.Cells(harok.Rows.Count, 1).End(xlUp).Row

    harok.Cells(1, 8).Value = "Hodnotenie"

    For i = 2 To poslednyRiadok
        trzba = CDbl(harok.Cells(i, 5).Value) * CDbl(harok.Cells(i, 6).Value)

        If trzba >= 1000 Then
            harok.Cells(i, 8).Value = "Vysoká tržba"
        ElseIf trzba >= 500 Then
            harok.Cells(i, 8).Value = "Stredná tržba"
        Else
            harok.Cells(i, 8).Value = "Nízka tržba"
        End If
    Next i
End Sub

Public Sub OcislovanieRiadkov()
    Dim harok As Worksheet
    Dim poslednyRiadok As Long
    Dim i As Long

    Set harok = ThisWorkbook.Worksheets("Dáta")
    poslednyRiadok = harok.Cells(harok.Rows.Count, 1).End(xlUp).Row
    harok.Cells(1, 9).Value = "Číslo riadka"

    For i = 2 To poslednyRiadok
        harok.Cells(i, 9).Value = i - 1
    Next i
End Sub

Public Sub ZvyrazniZaporneHodnoty()
    Dim bunka As Range

    For Each bunka In ThisWorkbook.Worksheets("Prehľad").Range("F2:F500")
        If IsNumeric(bunka.Value) And bunka.Value < 0 Then
            bunka.Font.Bold = True
            bunka.Font.Color = vbRed
        End If
    Next bunka
End Sub

Public Sub SpracujPoPrazdnyRiadok()
    Dim harok As Worksheet
    Dim i As Long

    Set harok = ThisWorkbook.Worksheets("Dáta")
    i = 2

    Do While harok.Cells(i, 1).Value <> ""
        harok.Cells(i, 10).Value = UCase$(CStr(harok.Cells(i, 2).Value))
        i = i + 1
    Loop

    harok.Cells(1, 10).Value = "Región veľkými písmenami"
End Sub
