Attribute VB_Name = "PouzivatelskeFunkcie"
Option Explicit

Public Function CenaSDPH(ByVal cenaBezDPH As Double, ByVal sadzbaDPH As Double) As Double
    CenaSDPH = cenaBezDPH * (1 + sadzbaDPH)
End Function

Public Function BezpecnePercento(ByVal citatel As Double, ByVal menovatel As Double) As Variant
    If menovatel = 0 Then
        BezpecnePercento = CVErr(xlErrDiv0)
    Else
        BezpecnePercento = citatel / menovatel
    End If
End Function

Public Function Trzba(ByVal jednotkovaCena As Double, ByVal mnozstvo As Double) As Double
    Trzba = jednotkovaCena * mnozstvo
End Function
