Attribute VB_Name = "NastrojeHarkov"
Option Explicit

Public Function ZiskajAleboVytvorHarok(ByVal nazovHarka As String) As Worksheet
    Dim harok As Worksheet

    On Error Resume Next
    Set harok = ThisWorkbook.Worksheets(nazovHarka)
    On Error GoTo 0

    If harok Is Nothing Then
        Set harok = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        harok.Name = nazovHarka
    End If

    Set ZiskajAleboVytvorHarok = harok
End Function

Public Sub VytvorHarokPrehlad()
    Dim harok As Worksheet
    Set harok = ZiskajAleboVytvorHarok("Prehľad")

    harok.Cells.Clear
    harok.Range("A1").Value = "Prehľad predaja"
    harok.Range("A1").Font.Bold = True
    harok.Columns.AutoFit
End Sub

Public Sub SpocitajHarky()
    MsgBox "Počet hárkov: " & ThisWorkbook.Worksheets.Count, vbInformation, "excel-lab"
End Sub
