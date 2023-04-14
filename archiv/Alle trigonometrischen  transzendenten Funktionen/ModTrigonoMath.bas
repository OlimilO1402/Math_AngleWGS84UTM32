Attribute VB_Name = "ModTrigonoMath"
Option Explicit
Public Const PI As Double = 3.14159265358979

'##########'            Trigonometrische Funktionen             '##########'
Public Function Sinus(ByVal A As Double) As Double   ' aka sin
    Sinus = VBA.Math.Sin(A)
End Function
Public Function Cosinus(ByVal A As Double) As Double ' aka cos
    Cosinus = VBA.Math.Cos(A)
End Function
Public Function Tangens(ByVal A As Double) As Double ' aka tan
    Tangens = VBA.Math.Tan(A)
End Function

Public Function Cosecans(ByVal A As Double) As Double  ' aka csc
    Cosecans = 1 / VBA.Math.Sin(A)
End Function
Public Function Secans(ByVal A As Double) As Double    ' aka sec
    Secans = 1 / VBA.Math.Cos(A)
End Function
Public Function Cotangens(ByVal A As Double) As Double ' aka cot
    Cotangens = 1 / VBA.Math.Tan(A)
End Function

'##########'         Trigonometrische Umkehrfunktionen          '##########'
Public Function ArcusSinus(ByVal y As Double) As Double   ' aka arcsin
    Select Case y
        Case 1
            ArcusSinus = 0.5 * PI
        Case -1
            ArcusSinus = -0.5 * PI
        Case Else
            ArcusSinus = VBA.Math.Atn(y / Sqr(1 - y * y))
    End Select
End Function
Public Function ArcusCosinus(ByVal x As Double) As Double ' aka arccos
    ArcusCosinus = 0.5 * PI - ArcusSinus(x)
End Function
Public Function ArcusTangens(ByVal t As Double) As Double ' aka arctan
    ArcusTangens = VBA.Math.Atn(t)
End Function

'ArcusTangensXY: also known as ATan2
Public Function ArcusTangensXY(ByVal x As Double, _
                               ByVal y As Double) As Double
    If y > 0 Then
        If x > 0 Then       ' 1. Quadrant
            ArcusTangensXY = Atn(Abs(y) / Abs(x)) '+ PI * 0#
        ElseIf x < 0 Then   ' 2. Quadrant
            ArcusTangensXY = -Atn(Abs(y) / Abs(x)) + PI '* 1#
        Else 'If x = 0 Then ' pos. Y-Achse
            ArcusTangensXY = 0.5 * PI
        End If
    ElseIf y < 0 Then
        If x < 0 Then       ' 3. Quadrant
            ArcusTangensXY = Atn(Abs(y) / Abs(x)) + PI '* 1#
        ElseIf x > 0 Then   ' 4. Quadrant
            ArcusTangensXY = -Atn(Abs(y) / Abs(x)) + PI * 2
        Else 'If x = 0 Then ' neg. Y-Achse
            ArcusTangensXY = 1.5 * PI
        End If
    Else 'If y = 0 Then
        If x > 0 Then       ' pos. X-Achse
            ArcusTangensXY = 0
        ElseIf x < 0 Then   ' neg. X-Achse
            ArcusTangensXY = PI
        Else 'If x = 0 Then ' Nullpunkt
            ArcusTangensXY = 0
        End If
    End If
End Function

Public Function ArcusCosecans(ByVal y As Double) As Double  ' aka arccsc
    ArcusCosecans = ArcusSinus(1 / y)
End Function
Public Function ArcusSecans(ByVal x As Double) As Double    ' aka arcsec
    ArcusSecans = ArcusCosinus(1 / x)
End Function
Public Function ArcusCotangens(ByVal t As Double) As Double ' aka arccot
    ArcusCotangens = PI * 0.5 - VBA.Math.Atn(t)
End Function

'######################'  Hyperbolische Funktionen   '#####################'
Public Function SinusHyperbolicus(ByVal A As Double) As Double   ' aka sinh
    SinusHyperbolicus = (Exp(A) - Exp(-A)) / 2
End Function
Public Function CosinusHyperbolicus(ByVal A As Double) As Double ' aka sinh
    CosinusHyperbolicus = (Exp(A) + Exp(-A)) / 2
End Function
Public Function TangensHyperbolicus(ByVal A As Double) As Double ' aka tanh
    TangensHyperbolicus = (Exp(A) - Exp(-A)) / (Exp(A) + Exp(-A))
End Function

Public Function CosecansHyperbolicus(ByVal y As Double) As Double  ' aka csch
    CosecansHyperbolicus = 2 / (Exp(y) - Exp(-y))
End Function
Public Function SecansHyperbolicus(ByVal x As Double) As Double    ' aka sech
    SecansHyperbolicus = 2 / (Exp(x) + Exp(-x))
End Function
Public Function CotangensHyperbolicus(ByVal t As Double) As Double ' aka coth
    CotangensHyperbolicus = (Exp(t) + Exp(-t)) / (Exp(t) - Exp(-t))
End Function

'##########'           Hyperbolische Umkehrfunktionen           '##########'
Public Function AreaSinusHyperbolicus(ByVal y As Double) As Double   ' aka arsinh
    AreaSinusHyperbolicus = VBA.Math.Log(y + Sqr(y * y + 1))
End Function
Public Function AreaCosinusHyperbolicus(ByVal x As Double) As Double ' aka arcosh
    AreaCosinusHyperbolicus = VBA.Math.Log(x + Sqr(x * x - 1))
End Function
Public Function AreaTangensHyperbolicus(ByVal t As Double) As Double ' aka artanh
    AreaTangensHyperbolicus = VBA.Math.Log((1 + t) / (1 - t)) / 2
End Function

Public Function AreaCosecansHyperbolicus(ByVal x As Double) As Double  ' aka arcsch
    AreaCosecansHyperbolicus = VBA.Math.Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function
Public Function AreaSecansHyperbolicus(ByVal x As Double) As Double    ' aka arsech
    AreaSecansHyperbolicus = VBA.Math.Log((Sqr(-x * x + 1) + 1) / x)
End Function
Public Function AreaCotangensHyperbolicus(ByVal x As Double) As Double ' aka arcoth
    AreaCotangensHyperbolicus = VBA.Math.Log((x + 1) / (x - 1)) / 2
End Function


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
Public Function LogN(ByVal x As Double, _
                     Optional ByVal N As Double = 10#) As Double
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


'##########'    Winkelumrechnung    '##########'
' Grad(=Deg), Neugrad(=Gon) und Radians(=Rad)
Public Function DegToRad(ByVal A As Double) As Double
    DegToRad = A * PI / 180
End Function
Public Function DegToGon(ByVal A As Double) As Double
    DegToGon = A / 0.9
End Function
Public Function RadToDeg(ByVal A As Double) As Double
    RadToDeg = A * 180 / PI
End Function
Public Function RadToGon(ByVal A As Double) As Double
    RadToGon = A * 200 / PI
End Function
Public Function GonToDeg(ByVal A As Double) As Double
    GonToDeg = A * 0.9
End Function
Public Function GonToRad(ByVal A As Double) As Double
    GonToRad = A * PI / 200
End Function

