Attribute VB_Name = "RozsahyBunky"
Option Explicit

Public Sub ZapisDoRozsahu()
    Dim harok As Worksheet
    Set harok = ThisWorkbook.Worksheets("Dáta")

    harok.Range("H1").Value = "Poznámka"
    harok.Range("H2").Value = "Zapísané pomocou objektu Range"
End Sub

Public Sub VyplnPomocouBuniek()
    Dim harok As Worksheet
    Dim i As Long

    Set harok = ThisWorkbook.Worksheets("Dáta")

    For i = 2 To 11
        harok.Cells(i, 9).Value = i - 1
    Next i

    harok.Cells(1, 9).Value = "Poradie"
End Sub

Public Function PoslednyPouzityRiadok(ByVal harok As Worksheet, Optional ByVal cisloStlpca As Long = 1) As Long
    PoslednyPouzityRiadok = harok.Cells(harok.Rows.Count, cisloStlpca).End(xlUp).Row
End Function
