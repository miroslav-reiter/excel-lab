Attribute VB_Name = "FirstMacro"
Option Explicit

Public Sub HelloExcel()
    MsgBox "Ahoj Excel!", vbInformation, "excel-lab"
End Sub

Public Sub WriteCourseName()
    ThisWorkbook.Worksheets("VBA Guide").Range("B2").Value = "Microsoft Excel IV. Expert - VBA"
End Sub
