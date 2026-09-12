Attribute VB_Name = "PrveMakro"
Option Explicit

Public Sub AhojExcel()
    MsgBox "Ahoj, Excel!", vbInformation, "excel-lab"
End Sub

Public Sub ZapisNazovKurzu()
    ThisWorkbook.Worksheets("Sprievodca VBA").Range("B2").Value = "Microsoft Excel IV. Expert - Programovanie vo VBA"
End Sub
