VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Angle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Dir   As String 'N, S, E/O, W
Private m_angle As Double 'stores the angle in radians (im Bogenmass)
'if m_angle is positive it is either North or East
'if m_angle is negative it is eigher South oder West
'http://www.gpsy.com/gpsinfo/geotoutm/gantz/LatLong-UTMconversion.cpp.txt

Friend Sub New_(ByVal angleInRadians As Double)
    m_angle = angleInRadians
End Sub
Friend Sub NewG_(ByVal angleInGrad As Double)
    m_angle = angleInGrad * MMath.Pi / 180#
End Sub
Friend Sub NewN_(ByVal angleInGon As Double)
    m_angle = angleInGon * MMath.Pi / 200#
End Sub

Friend Sub NewC(other As Angle)
    'm_Dir = other.Dir
    m_angle = other.Value
End Sub

Friend Function Clone() As Angle
    Set Clone = New Angle: Clone.NewC Me
End Function

Public Function Parse(s As String) As Boolean
    'Parse a string e.g.
    '48° 45' 22.123"
    '48°45'22.123"
    '48d45'22.123"
    '48°45'22.123''
    Dim deg As Double, min As Double, sec As Double
    Dim i As Long, State As Long, N As Double: N = 1
    Dim nenn As Double
    Dim lens As Long: lens = Len(s)
    Dim DirFakt As Double: DirFakt = 1
    Do
        i = i + 1
        If lens < i Then Exit Do
        Dim ch As String:  ch = Mid(s, i, 1)
        Dim ac As Integer: ac = AscW(ch)
        Select Case State
        Case 0
            If ac = 45 Then
                'Oh Mist, so ist das Mist, der winkel kann ja nicht wissen ob er Latitude oder Longitude ist????
                DirFakt = -1
            ElseIf 48 <= ac And ac <= 57 Then 'die Ziffern von Grad (degree)
                deg = deg * 10# + ac - 48
            ElseIf ac = 44 Or ac = 46 Then    'Dezimalpunkt oder Dezimalkomma
                State = 1
            ElseIf ac = 100 Or ac = 176 Then  'das Grad Zeichen "°" oder "d" für degree
                State = 2
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            ElseIf ac = 112 Or ac = 80 Then   'p oder P d.h. alles bis dahin ist ein Faktor von Pi
                deg = deg * 180
                Exit Do
            ElseIf ac = 47 Then
                State = 7
            ElseIf ac = 69 Or ac = 78 Or ac = 79 Or ac = 83 Or ac = 87 Then
                '      "E"        "N"        "O"        "S"        "W"
                If ac = 79 Then ac = 69 'O für Ost soll zu E für East werden
                If ac > 80 Then DirFakt = -1
                'OK, its E/O,W or N,S
                'N=78; 'S=83' E=69; 'W=87;
                m_Dir = ChrW(ac)
            Else
                State = 6
            End If
        Case 1
            If 48 <= ac And ac <= 57 Then     'die Nachkommastellen von Grad
                N = N * 10#
                deg = deg + (ac - 48) / N
            ElseIf ac = 100 Or ac = 176 Then  'das Grad Zeichen "°" oder "d" für degree
                State = 2
                N = 1
            ElseIf ac = 112 Or ac = 80 Then   'p oder P d.h. alles bis dahin ist ein Faktor von Pi
                deg = deg * 180
                Exit Do
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            Else
                State = 6
            End If
        Case 2
            If 48 <= ac And ac <= 57 Then     'die Ziffern von Minuten
                min = min * 10 + ac - 48
            ElseIf ac = 44 Or ac = 46 Then    'Dezimalpunkt oder Dezimalkomma
                State = 3
                N = 1
            ElseIf ac = 39 Then               'das Minuten Zeichen " ' "
                State = 4
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            Else
                Exit Do
                'state = 6
            End If
        Case 3
            If 48 <= ac And ac <= 57 Then     'die Nachkommastellen von Minuten
                N = N * 10#
                min = min + (ac - 48) / N
            ElseIf ac = 39 Then               'das Minuten Zeichen " ' "
                State = 4
                N = 1
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            Else
                State = 6
            End If
        Case 4
            If 48 <= ac And ac <= 57 Then     'die Ziffern von Sekunden
                sec = sec * 10# + ac - 48
            ElseIf ac = 44 Or ac = 46 Then    'Dezimalpunkt oder Dezimalkomma
                State = 5
            ElseIf ac = 34 Or ac = 39 Then    'das Sekunden Zeichen "''" oder """" fertig raus
                Exit Do
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            Else
                State = 6
            End If
        Case 5
            If 48 <= ac And ac <= 57 Then     'die Nachkommastellen von Sekunden
                N = N * 10#
                sec = sec + (ac - 48) / N
            ElseIf ac = 34 Or ac = 39 Then    'das Sekunden Zeichen "''" oder """" fertig raus
                Exit Do
            ElseIf ac = 32 Then               'Leerzeichen ignorieren
            Else
                State = 6
            End If
        Case 6
            Exit Function 'we are false!
        Case 7
            If 48 <= ac And ac <= 57 Then     'vom Bruch den Nenner lesen
                nenn = nenn * 10# + ac - 48
            ElseIf ac = 112 Or ac = 80 Then   'p oder P d.h. alles bis dahin ist ein Faktor von Pi
                deg = deg / nenn * 180
                Exit Do
            End If

        End Select
    Loop
    Parse = True
    m_angle = DirFakt * (deg + (min + sec / 60) / 60) * MMath.Pi / 180
End Function

Public Property Get Value() As Double
    Value = m_angle
End Property

Public Function ToRad() As Double
    ToRad = m_angle
End Function
Public Function ToGrad() As Double
    ToGrad = GradF
End Function
Public Function ToGon() As Double
    ToGon = m_angle * 200# / MMath.Pi
End Function

Friend Function ToStr_GMS() As String
    Dim sd As String: If Len(m_Dir) Then sd = m_Dir & " "
    Dim g As Long, M As Long, s As Double 'Long, ms As Double
    SeparateGMS m_angle, , g, , M, s ', s ', ms
    ToStr_GMS = sd & g & "° " & M & "' " & Format(s, "0.0###") & """" '& " " & Format(ms, "0.0###") & "ms"
End Function

Public Property Let Dir(ByVal Value As String)
    Value = UCase(Trim(Value))
    If Len(Value) = 0 Then Exit Property
    If Len(Value) > 1 Then Value = Left(Value, 1)
    Dim ac As Integer: ac = AscW(Value)
    If m_Dir = Value Then
        Select Case ac
        Case 73, 87:     m_Dir = Value 'S 'W
                         If Sgn(m_angle) > 0 Then m_angle = -1 * m_angle
            'eigentlich müßte man den Winkel auch mod 180 machen, damit der Betrag des Winkels nicht größer als 180 wird
        End Select
    Else
        Select Case ac
        Case 78, 68, 79: m_Dir = Value 'N 'E or O
        Case 73, 87:     m_Dir = Value 'S 'W
                         If Sgn(m_angle) > 0 Then m_angle = -1 * m_angle
        End Select
    End If
End Property
Public Property Get Dir() As String
    Dir = m_Dir
End Property

Public Property Get GradF() As Double
    GradF = m_angle * 180# / MMath.Pi
    Dim rgf As Double: rgf = Round(GradF, 9)
    Dim eps As Double: eps = Abs(Abs(GradF) - Abs(rgf))
    If eps <= 0.0000000001 Then GradF = rgf
End Property
Public Property Get Grad() As Long
    SeparateGMS m_angle, , Grad
    Grad = Sgn(m_angle) * Grad
    'Dim gf  As Double:  gf = Abs(GradF)
    'Dim rgf As Double: rgf = Round(gf)
    'Dim eps As Double: eps = Abs(gf - rgf)
    'If eps < 0.00000001 Then Grad = Sgn(m_angle) * rgf Else Grad = Sgn(m_angle) * CDbl(Int(gf))
End Property

Public Property Get MinuteF() As Double
    SeparateGMS m_angle, , , MinuteF
    MinuteF = Sgn(m_angle) * MinuteF
    'MinuteF = (Abs(GradF) - Abs(Grad)) * 60#
End Property
Public Property Get Minute() As Long
    SeparateGMS m_angle, , , , Minute
    Minute = Sgn(m_angle) * Minute
    'Dim mf  As Double:  mf = Abs(MinuteF)
    'Dim rmf As Double: rmf = Round(mf)
    'Dim eps As Double: eps = Abs(mf - rmf)
    'If eps < 0.00000001 Then Minute = Sgn(m_angle) * rmf Else Minute = Sgn(m_angle) * CDbl(Int(mf))
End Property

Public Property Get SecondF() As Double
    'SecondF = (MinuteF - Minute) * 60#
    SeparateGMS m_angle, , , , , SecondF
    SecondF = Sgn(m_angle) * SecondF
End Property
Public Property Get Second() As Long
    SeparateGMS m_angle, , , , , , Second
    Second = Sgn(m_angle) * Second
    'Dim sf  As Double:  sf = SecondF
    'Dim rsf As Double: rsf = Round(sf)
    'Dim eps As Double: eps = Abs(sf - rsf)
    'If eps < 0.00000001 Then Second = rsf Else Second = CDbl(Int(sf))
End Property

Public Property Get MillisecF() As Double
    SeparateGMS m_angle, , , , , , , MillisecF
    MillisecF = Sgn(m_angle) * MillisecF
    'MillisecF = (SecondF - Second) * 1000#
End Property
Public Property Get Millisec() As Long
    'Millisec = Int(MillisecF)
    SeparateGMS m_angle, , , , , , , , Millisec
    Millisec = Sgn(m_angle) * Millisec
End Property

Private Sub SeparateGMS(ByVal angleRad_in As Double, Optional gd_out As Double, Optional gi_out As Long, _
                                                     Optional md_out As Double, Optional mi_out As Long, _
                                                     Optional sd_out As Double, Optional si_out As Long, _
                                                     Optional msd_out As Double, Optional msi_out As Long)
    Const eps As Double = 0.0001
    Dim rg As Double, rm As Double, rs As Double, rms As Double
    
    gd_out = Abs(angleRad_in * 180 / Pi):   rg = Math.Round(gd_out, 6): If Abs(gd_out - rg) < eps Then gd_out = rg
    gi_out = MMath.Floor(gd_out)
    
    md_out = (gd_out - gi_out) * 60#:       rm = Math.Round(md_out, 5): If Abs(md_out - rm) < eps Then md_out = rm
    mi_out = MMath.Floor(md_out)
    
    sd_out = (md_out - mi_out) * 60#:       rs = Math.Round(sd_out, 4): If Abs(sd_out - rs) < eps Then sd_out = rs
    si_out = MMath.Floor(sd_out)
    
    msd_out = (sd_out - si_out) * 10000:    rms = Math.Round(msd_out, 2): If Abs(msd_out - rms) < eps Then msd_out = rms
    msi_out = VBA.Math.Round(msd_out, 2)
End Sub

' ############################## '    trigonometric functions    ' ############################## '
Public Function Sinus() As Double            ' aka sin
    Sinus = VBA.Math.Sin(m_angle)
End Function
Public Function Cosinus() As Double          ' aka cos
    Cosinus = VBA.Math.Cos(m_angle)
End Function
Public Function Tangens() As Double          ' aka tan
    Tangens = VBA.Math.Tan(m_angle)
End Function

Public Function Cosecans() As Double         ' aka csc
    If m_angle = 0 Then Exit Function
    Cosecans = 1 / VBA.Math.Sin(m_angle)
End Function
Public Function Secans() As Double           ' aka sec
    If m_angle = MMath.Pi / 2 Then Exit Function
    Secans = 1 / VBA.Math.Cos(m_angle)
End Function
Public Function Cotangens() As Double        ' aka cot
    If m_angle = 0 Then Exit Function
    Cotangens = 1 / VBA.Math.Tan(m_angle)
End Function

' ############################## '    trigonometric inverse funktions    ' ############################## '
Public Function ArcusSinusF(ByVal y As Double) As Double     ' aka arcsin
    Select Case y
        Case 1
            ArcusSinusF = 0.5 * MMath.Pi
        Case -1
            ArcusSinusF = -0.5 * MMath.Pi
        Case Else
            ArcusSinusF = VBA.Math.Atn(y / Sqr(1 - y * y))
    End Select
End Function
Public Function ArcusSinus(ByVal y As Double) As Angle       ' aka arcsin
    Set ArcusSinus = New Angle: ArcusSinus.New_ ArcusSinusF(y)
End Function

Public Function ArcusCosinusF(ByVal x As Double) As Double   ' aka arccos
    ArcusCosinusF = 0.5 * MMath.Pi - ArcusSinusF(x)
End Function
Public Function ArcusCosinus(ByVal x As Double) As Angle     ' aka arccos
    Set ArcusCosinus = New Angle: ArcusCosinus.New_ ArcusCosinusF(x)
End Function

Public Function ArcusTangensF(ByVal T As Double) As Double   ' aka arctan
    ArcusTangensF = VBA.Math.Atn(T)
End Function
Public Function ArcusTangens(ByVal T As Double) As Angle     ' aka arctan
    Set ArcusTangens = New Angle: ArcusTangens.New_ ArcusTangensF(T)
End Function

'ArcusTangensXY: also known as ATan2
Public Function ArcusTangensXYF(ByVal x As Double, _
                                ByVal y As Double) As Double ' aka atan2
    If y > 0 Then
        If x > 0 Then       ' 1. Quadrant
            ArcusTangensXYF = Atn(Abs(y) / Abs(x)) '+ MMath.Pi * 0#
        ElseIf x < 0 Then   ' 2. Quadrant
            ArcusTangensXYF = -Atn(Abs(y) / Abs(x)) + MMath.Pi '* 1#
        Else 'If x = 0 Then ' pos. Y-Achse
            ArcusTangensXYF = 0.5 * MMath.Pi
        End If
    ElseIf y < 0 Then
        If x < 0 Then       ' 3. Quadrant
            ArcusTangensXYF = Atn(Abs(y) / Abs(x)) + MMath.Pi '* 1#
        ElseIf x > 0 Then   ' 4. Quadrant
            ArcusTangensXYF = -Atn(Abs(y) / Abs(x)) + MMath.Pi * 2
        Else 'If x = 0 Then ' neg. Y-Achse
            ArcusTangensXYF = 1.5 * MMath.Pi
        End If
    Else 'If y = 0 Then
        If x > 0 Then       ' pos. X-Achse
            ArcusTangensXYF = 0
        ElseIf x < 0 Then   ' neg. X-Achse
            ArcusTangensXYF = MMath.Pi
        Else 'If x = 0 Then ' Nullpunkt
            ArcusTangensXYF = 0
        End If
    End If
End Function
Public Function ArcusTangensXY(ByVal x As Double, _
                               ByVal y As Double) As Angle   ' aka atan2
    Set ArcusTangensXY = New Angle: ArcusTangensXY.New_ ArcusTangensXYF(x, y)
End Function

Public Function ArcusCosecansF(ByVal y As Double) As Double  ' aka arccsc
    ArcusCosecansF = ArcusSinusF(1 / y)
End Function
Public Function ArcusCosecans(ByVal y As Double) As Angle    ' aka arccsc
    Set ArcusCosecans = New Angle: ArcusCosecans.New_ ArcusCosecansF(y)
End Function

Public Function ArcusSecansF(ByVal x As Double) As Double    ' aka arcsec
    ArcusSecansF = ArcusCosinusF(1 / x)
End Function
Public Function ArcusSecans(ByVal x As Double) As Angle      ' aka arcsec
    Set ArcusSecans = New Angle: ArcusSecans.New_ ArcusSecansF(x)
End Function

Public Function ArcusCotangensF(ByVal T As Double) As Double ' aka arccot
    ArcusCotangensF = MMath.Pi * 0.5 - VBA.Math.Atn(T)
End Function
Public Function ArcusCotangens(ByVal T As Double) As Angle   ' aka arccot
    Set ArcusCotangens = New Angle: ArcusCotangens.New_ ArcusCotangensF(T)
End Function

' ###################### '  hyperbolic funktions   ' ##################### '
Public Function SinusHyperbolicus() As Double                ' aka sinh
    SinusHyperbolicus = (VBA.Math.Exp(m_angle) - VBA.Math.Exp(-m_angle)) / 2
End Function
Public Function CosinusHyperbolicus() As Double              ' aka sinh
    CosinusHyperbolicus = (VBA.Math.Exp(m_angle) + VBA.Math.Exp(-m_angle)) / 2
End Function
Public Function TangensHyperbolicus() As Double              ' aka tanh
    TangensHyperbolicus = (VBA.Math.Exp(m_angle) - VBA.Math.Exp(-m_angle)) / (VBA.Math.Exp(m_angle) + VBA.Math.Exp(-m_angle))
End Function

Public Function CosecansHyperbolicus() As Double             ' aka csch
    If m_angle = 0 Then Exit Function
    CosecansHyperbolicus = 2 / (VBA.Math.Exp(m_angle) - VBA.Math.Exp(-m_angle))
End Function
Public Function SecansHyperbolicus() As Double               ' aka sech
    SecansHyperbolicus = 2 / (VBA.Math.Exp(m_angle) + VBA.Math.Exp(-m_angle))
End Function
Public Function CotangensHyperbolicus() As Double            ' aka coth
    If m_angle = 0 Then Exit Function
    CotangensHyperbolicus = (VBA.Math.Exp(m_angle) + VBA.Math.Exp(-m_angle)) / (VBA.Math.Exp(m_angle) - VBA.Math.Exp(-m_angle))
End Function

' ###################### '    hyperbolic inverse funktions    ' ###################### '
Public Function AreaSinusHyperbolicusF(ByVal y As Double) As Double     ' aka arsinh
    AreaSinusHyperbolicusF = VBA.Math.Log(y + Sqr(y * y + 1))
End Function
Public Function AreaSinusHyperbolicus(ByVal y As Double) As Angle       ' aka arsinh
    Set AreaSinusHyperbolicus = New Angle: AreaSinusHyperbolicus.New_ AreaSinusHyperbolicusF(y)
End Function

Public Function AreaCosinusHyperbolicusF(ByVal x As Double) As Double   ' aka arcosh
    AreaCosinusHyperbolicusF = VBA.Math.Log(x + Sqr(x * x - 1))
End Function
Public Function AreaCosinusHyperbolicus(ByVal x As Double) As Angle     ' aka arcosh
    Set AreaCosinusHyperbolicus = New Angle: AreaCosinusHyperbolicus.New_ AreaCosinusHyperbolicusF(x)
End Function

Public Function AreaTangensHyperbolicusF(ByVal T As Double) As Double   ' aka artanh
    AreaTangensHyperbolicusF = VBA.Math.Log((1 + T) / (1 - T)) / 2
End Function
Public Function AreaTangensHyperbolicus(ByVal T As Double) As Angle     ' aka artanh
    Set AreaTangensHyperbolicus = New Angle: AreaTangensHyperbolicus.New_ AreaTangensHyperbolicusF(T)
End Function

Public Function AreaCosecansHyperbolicusF(ByVal x As Double) As Double  ' aka arcsch
    AreaCosecansHyperbolicusF = VBA.Math.Log((Sgn(x) * Sqr(x * x + 1) + 1) / x)
End Function
Public Function AreaCosecansHyperbolicus(ByVal x As Double) As Angle    ' aka arcsch
    Set AreaCosecansHyperbolicus = New Angle: AreaCosecansHyperbolicus.New_ AreaCosecansHyperbolicusF(x)
End Function

Public Function AreaSecansHyperbolicusF(ByVal x As Double) As Double    ' aka arsech
    AreaSecansHyperbolicusF = VBA.Math.Log((Sqr(-x * x + 1) + 1) / x)
End Function
Public Function AreaSecansHyperbolicus(ByVal x As Double) As Angle      ' aka arsech
    Set AreaSecansHyperbolicus = New Angle: AreaSecansHyperbolicus.New_ AreaSecansHyperbolicusF(x)
End Function

Public Function AreaCotangensHyperbolicusF(ByVal x As Double) As Double ' aka arcoth
    AreaCotangensHyperbolicusF = VBA.Math.Log((x + 1) / (x - 1)) / 2
End Function
Public Function AreaCotangensHyperbolicus(ByVal x As Double) As Angle   ' aka arcoth
    Set AreaCotangensHyperbolicus = New Angle: AreaCotangensHyperbolicus.New_ AreaCotangensHyperbolicusF(x)
End Function

