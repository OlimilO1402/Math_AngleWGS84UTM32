Attribute VB_Name = "MMath"
Option Explicit
Public Pi

Public Sub Init()
    Pi = CDec(CDec(4) * CDec(Atn(1)))
End Sub

'##########'               Zusätzliche Funktionen               '##########'
Public Function SinusCardinalis(ByVal x As Double) As Double ' aka sinc
    If x = 0 Then
        SinusCardinalis = 1
    Else
        SinusCardinalis = VBA.Math.Sin(x) / x
    End If
End Function

Public Function Log10(ByVal d As Double) As Double
    Log10 = VBA.Math.Log(d) / VBA.Math.Log(10)
End Function

Public Function LN(ByVal d As Double) As Double
  LN = VBA.Math.Log(d)
End Function

Public Function LogN(ByVal x As Double, Optional ByVal N As Double = 10#) As Double
                     'n darf nicht eins und nicht 0 sein
    LogN = VBA.Math.Log(x) / VBA.Math.Log(N)
End Function

Public Function BigMul(ByVal A As Long, ByVal b As Long) As Variant
    BigMul = CDec(A) * CDec(b)
End Function

Public Function Floor(ByVal A As Double) As Double
    Floor = CDbl(Int(A))
End Function

Public Function Ceiling(ByVal A As Double) As Double
    Ceiling = CDbl(Int(A))
    If A <> 0 Then If Abs(Ceiling / A) <> 1 Then Ceiling = Ceiling + 1
End Function


