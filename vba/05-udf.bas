Attribute VB_Name = "UserFunctions"
Option Explicit

Public Function PriceWithVAT(ByVal priceWithoutVAT As Double, ByVal vatRate As Double) As Double
    PriceWithVAT = priceWithoutVAT * (1 + vatRate)
End Function

Public Function SafePercent(ByVal numerator As Double, ByVal denominator As Double) As Variant
    If denominator = 0 Then
        SafePercent = CVErr(xlErrDiv0)
    Else
        SafePercent = numerator / denominator
    End If
End Function

Public Function Revenue(ByVal unitPrice As Double, ByVal quantity As Double) As Double
    Revenue = unitPrice * quantity
End Function
