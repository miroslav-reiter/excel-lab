Attribute VB_Name = "ReportAutomation"
Option Explicit

Public Sub CreateSalesReport()
    Dim wsData As Worksheet
    Dim wsReport As Worksheet
    Dim lastRow As Long
    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    On Error GoTo CleanFail

    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsData = ThisWorkbook.Worksheets("Data")
    Set wsReport = GetOrCreateWorksheet("Report")

    wsReport.Cells.Clear
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    wsData.Range("A1:F" & lastRow).Copy Destination:=wsReport.Range("A1")

    wsReport.Range("G1").Value = "Revenue"
    wsReport.Range("G2:G" & lastRow).FormulaR1C1 = "=RC[-2]*RC[-1]"
    wsReport.Range("G2:G" & lastRow).NumberFormat = "#,##0.00 [$€-sk-SK]"

    wsReport.Range("A1:G1").Font.Bold = True
    wsReport.Range("A1:G1").AutoFilter
    wsReport.Columns("A:G").AutoFit

    wsReport.Range("I1").Value = "Total Revenue"
    wsReport.Range("I2").Formula = "=SUM(G2:G" & lastRow & ")"
    wsReport.Range("I2").NumberFormat = "#,##0.00 [$€-sk-SK]"

    MsgBox "Report bol vytvorený.", vbInformation, "excel-lab"

CleanExit:
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

CleanFail:
    MsgBox "Chyba " & Err.Number & ": " & Err.Description, vbExclamation, "excel-lab"
    Resume CleanExit
End Sub
