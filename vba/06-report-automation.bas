Attribute VB_Name = "AutomatizaciaPrehladu"
Option Explicit

Public Sub VytvorPrehladPredaja()
    Dim harokData As Worksheet
    Dim harokPrehlad As Worksheet
    Dim poslednyRiadok As Long
    Dim povodnyVypocet As XlCalculation
    Dim povodnePrekreslovanie As Boolean
    Dim povodneUdalosti As Boolean

    On Error GoTo Chyba

    povodnyVypocet = Application.Calculation
    povodnePrekreslovanie = Application.ScreenUpdating
    povodneUdalosti = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set harokData = ThisWorkbook.Worksheets("Dáta")
    Set harokPrehlad = ZiskajAleboVytvorHarok("Prehľad")

    harokPrehlad.Cells.Clear
    poslednyRiadok = harokData.Cells(harokData.Rows.Count, 1).End(xlUp).Row

    harokData.Range("A1:F" & poslednyRiadok).Copy Destination:=harokPrehlad.Range("A1")

    harokPrehlad.Range("G1").Value = "Tržba"
    harokPrehlad.Range("G2:G" & poslednyRiadok).FormulaR1C1 = "=RC[-2]*RC[-1]"
    harokPrehlad.Range("G2:G" & poslednyRiadok).NumberFormat = "#,##0.00 [$€-sk-SK]"

    harokPrehlad.Range("A1:G1").Font.Bold = True
    harokPrehlad.Range("A1:G1").AutoFilter
    harokPrehlad.Columns("A:G").AutoFit

    harokPrehlad.Range("I1").Value = "Celková tržba"
    harokPrehlad.Range("I2").Formula = "=SUM(G2:G" & poslednyRiadok & ")"
    harokPrehlad.Range("I2").NumberFormat = "#,##0.00 [$€-sk-SK]"

    MsgBox "Prehľad bol vytvorený.", vbInformation, "excel-lab"

Koniec:
    Application.Calculation = povodnyVypocet
    Application.EnableEvents = povodneUdalosti
    Application.ScreenUpdating = povodnePrekreslovanie
    Exit Sub

Chyba:
    MsgBox "Chyba " & Err.Number & ": " & Err.Description, vbExclamation, "excel-lab"
    Resume Koniec
End Sub
