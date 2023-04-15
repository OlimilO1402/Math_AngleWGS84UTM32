VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AngleDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Angle
Private m_Dir   As String  'N, S, E/O, W
Private m_Angle As Variant 'stores the angle in radians (im Bogenmass) as Decimal
'if m_angle is positive it is either North or East
'if m_angle is negative it is eigher South oder West
'http://www.gpsy.com/gpsinfo/geotoutm/gantz/LatLong-UTMconversion.cpp.txt

Friend Sub New_(angleInRadians)
    m_Angle = CDec(angleInRadians)
End Sub
Friend Sub NewD_(angleInDegrees)
    m_Angle = CDec(CDec(angleInDegrees) * MMath.Pi / CDec(180))
End Sub
Friend Sub NewG_(angleInGon)
    m_Angle = CDec(CDec(angleInGon) * MMath.Pi / CDec(200))
End Sub
Friend Sub NewDMS_(aDeg, aMin, aSec)
    m_Angle = CDec((CDec(aDeg) + CDec(aMin) / CDec(60) + CDec(aSec) / CDec(3600)) * CDec(MMath.Pi) / CDec(180#))
End Sub
'copy to MNew:
'Public Function AngleDec(ByVal angleInRadians) As AngleDec
'    Set AngleDec = New AngleDec: AngleDec.New_ angleInRadians
'End Function
'Public Function AngleDecD(ByVal angleInDegrees) As AngleDec
'    Set AngleDecD = New AngleDec: AngleDecD.NewD_ angleInDegrees
'End Function
'Public Function AngleDecG(ByVal angleInGon As Double) As AngleDec
'    Set AngleDecG = New AngleDec: AngleDecG.NewG_ angleInGon
'End Function
'Public Function AngleDecDMS(ByVal deg As Long, ByVal min As Long, ByVal sec As Double) As AngleDec
'    Set AngleDecDMS = New AngleDec: AngleDecDMS.NewDMS_ deg, min, sec
'End Function
'Public Function AngleDecS(ByVal angleInDMS As String) As AngleDec
'    Set AngleDecS = New AngleDec: AngleDecS.Parse angleInDMS
'End Function

Friend Sub NewC(other As AngleDec)
    'm_Dir = other.Dir 'Fuck why not????
    m_Dir = other.Dir
    m_Angle = other.Value
End Sub

Friend Function Clone() As AngleDec
    Set Clone = New AngleDec: Clone.NewC Me
End Function

Public Function Parse(s As String) As Boolean
    'Parse a string e.g.
    '48° 45' 22.123"
    '48°45'22.123"
    '48d45'22.123"
    '48°45'22.123''
    Dim deg, Min, sec
    Dim i As Long, State As Long, N: N = 1
    Dim nenn
    Dim lens As Long: lens = Len(s)
    Dim DirFakt: DirFakt = 1
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
                Min = Min * 10 + ac - 48
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
                Min = Min + (ac - 48) / N
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
    m_Angle = CDec(CDec(DirFakt) * (CDec(deg) + (CDec(Min) + CDec(sec) / CDec(60)) / CDec(60)) * MMath.Pi / CDec(180))
End Function

Public Property Get Value()
    Value = CDec(m_Angle)
End Property

Public Function ToRad()
    ToRad = CDec(m_Angle)
End Function
Public Function ToGrad()
    ToGrad = CDec(GradF)
End Function
Public Function ToGon()
    ToGon = CDec(m_Angle * CDec(200) / MMath.Pi)
End Function

Friend Function ToStr_DMS() As String
    Dim sd As String: If Len(m_Dir) Then sd = m_Dir & " "
    Dim g As Long, M As Long, s As Double 'Long, ms As Double
    SeparateDMS m_Angle, , g, , M, s ', s ', ms
    ToStr_DMS = sd & g & "° " & M & "' " & Format(s, "0.0###") & """" '& " " & Format(ms, "0.0###") & "ms"
End Function

Public Property Let Dir(ByVal Value As String)
    Value = UCase(Trim(Value))
    If Len(Value) = 0 Then Exit Property
    If Len(Value) > 1 Then Value = Left(Value, 1)
    Dim ac As Integer: ac = AscW(Value)
    If m_Dir = Value Then
        Select Case ac
        Case 83, 87:     m_Dir = Value 'S 'W
                         If Sgn(m_Angle) > 0 Then m_Angle = -1 * m_Angle
            'eigentlich müßte man den Winkel auch mod 180 machen, damit der Betrag des Winkels nicht größer als 180 wird
        End Select
    Else
        Select Case ac
        Case 78, 69, 79: 'N 'E or O
            If Not (m_Dir = "O" And ac = 69) Then
                'leave "O" as it is ;-)
                m_Dir = Value
            End If
        Case 83, 87:     m_Dir = Value 'S 'W
                         If Sgn(m_Angle) > 0 Then m_Angle = -1 * m_Angle
        End Select
    End If
End Property
Public Property Get Dir() As String
    Dir = m_Dir
End Property

Public Property Get GradF()
    GradF = CDec(m_Angle * CDec(180) / MMath.Pi)
    Dim rgf: rgf = CDec(Round(GradF, 9))
    Dim eps: eps = CDec(Abs(Abs(GradF) - Abs(rgf)))
    If eps <= CDec(0.0000000001) Then GradF = rgf
End Property
'Public Property Let GradF(ByVal Value)
'    Dim gd ', md, sd
'    SeparateDMS m_Angle, gd, , md, , sd
'    NewDMS_ Value, gd, sd
'End Property

Public Property Get Grad() As Long
    SeparateDMS m_Angle, , Grad
    Grad = Sgn(m_Angle) * Grad
    'Dim gf  As Double:  gf = Abs(GradF)
    'Dim rgf As Double: rgf = Round(gf)
    'Dim eps As Double: eps = Abs(gf - rgf)
    'If eps < 0.00000001 Then Grad = Sgn(m_angle) * rgf Else Grad = Sgn(m_angle) * CDbl(Int(gf))
End Property
Public Property Let Grad(ByVal Value As Long)
    Dim gi As Long, mi As Long, sd
    SeparateDMS m_Angle, , gi, , mi, sd
    NewDMS_ Value, mi, sd
End Property

Public Property Get MinuteF()
    SeparateDMS m_Angle, , , MinuteF
    MinuteF = Sgn(m_Angle) * MinuteF
End Property
'Public Property Let MinuteF(ByVal Value)
'    Dim gi As Long, md ', sd
'    SeparateDMS m_Angle, , gi, md, , sd
'    NewDMS_ gi, Value, sd
'End Property

Public Property Get Minute() As Long
    SeparateDMS m_Angle, , , , Minute
    Minute = Sgn(m_Angle) * Minute
End Property
Public Property Let Minute(ByVal Value As Long)
    Dim gi As Long, mi As Long, sd
    SeparateDMS m_Angle, , gi, , mi, sd
    NewDMS_ gi, Value, sd
End Property

Public Property Get SecondF()
    'SecondF = (MinuteF - Minute) * 60#
    SeparateDMS m_Angle, , , , , SecondF
    SecondF = Sgn(m_Angle) * SecondF
End Property
Public Property Let SecondF(ByVal Value)
    Dim gi As Long, mi As Long, sd
    SeparateDMS m_Angle, , gi, , mi, sd
    NewDMS_ gi, mi, Value
End Property

Public Property Get Second() As Long
    SeparateDMS m_Angle, , , , , , Second
    Second = Sgn(m_Angle) * Second
    'Dim sf  As Double:  sf = SecondF
    'Dim rsf As Double: rsf = Round(sf)
    'Dim eps As Double: eps = Abs(sf - rsf)
    'If eps < 0.00000001 Then Second = rsf Else Second = CDbl(Int(sf))
End Property
Public Property Let Second(ByVal Value As Long)
    Dim gi As Long, mi As Long, si As Long, msd
    SeparateDMS m_Angle, , gi, , mi, , si, msd
    NewDMS_ gi, mi, CDec(Value) + CDec(msd) / CDec(1000)
End Property

Public Property Get MillisecF()
    SeparateDMS m_Angle, , , , , , , MillisecF
    MillisecF = Sgn(m_Angle) * MillisecF
    'MillisecF = (SecondF - Second) * 1000#
End Property
Public Property Let MillisecF(ByVal Value)
    Dim gi As Long, mi As Long, si As Long
    SeparateDMS m_Angle, , gi, , mi, , si
    NewDMS_ gi, mi, CDec(CDec(si) + CDec(Value) / CDec(1000))
End Property

Public Property Get Millisec() As Long
    'Millisec = Int(MillisecF)
    SeparateDMS m_Angle, , , , , , , , Millisec
    Millisec = Sgn(m_Angle) * Millisec
End Property
Public Property Let Millisec(ByVal Value As Long)
    Dim gi As Long, mi As Long, si As Long
    SeparateDMS m_Angle, , gi, , mi, , si
    NewDMS_ gi, mi, CDec(CDec(si) + CDec(Value) / CDec(1000))
End Property

Private Sub SeparateDMS(ByVal angleRad_in, Optional gd_out, Optional gi_out As Long, _
                                           Optional md_out, Optional mi_out As Long, _
                                           Optional sd_out, Optional si_out As Long, _
                                           Optional msd_out, Optional msi_out As Long)
    Const eps = 0.00000001
    Dim rg, rm, rs, rms
    
    gd_out = CDec(Abs(angleRad_in * CDec(180) / MMath.Pi)):   rg = CDec(Math.Round(gd_out, 6)):  If Abs(gd_out - rg) < eps Then gd_out = rg
    gi_out = CDec(MMath.Floor(gd_out))
    
    md_out = CDec((gd_out - gi_out) * CDec(60)):              rm = CDec(Math.Round(md_out, 5)):  If Abs(md_out - rm) < eps Then md_out = rm
    mi_out = CDec(MMath.Floor(md_out))
    
    sd_out = CDec((md_out - mi_out) * CDec(60)):              rs = CDec(Math.Round(sd_out, 4)):  If Abs(sd_out - rs) < eps Then sd_out = rs
    si_out = CDec(MMath.Floor(sd_out))
    
    msd_out = CDec((sd_out - si_out) * CDec(10000)):         rms = CDec(Math.Round(msd_out, 2)): If Abs(msd_out - rms) < eps Then msd_out = rms
    msi_out = CDec(VBA.Math.Round(msd_out, 2))
End Sub

Public Function Add(a As Angle) As AngleDec
    Set Add = MNew.AngleDec(m_Angle + CDec(a.Value))
End Function

Public Sub AddDeg(ByVal d)
    m_Angle = m_Angle + d * MMath.Pi / 180#
End Sub

Public Function Subt(a As Angle) As AngleDec
    Set Subt = MNew.AngleDec(m_Angle - CDec(a.Value))
End Function

Public Function Mul(ByVal Value) As AngleDec
    Set Mul = MNew.AngleDec(m_Angle * CDec(Value))
End Function

Public Function Div(ByVal Value) As AngleDec
    Set Div = MNew.AngleDec(m_Angle / CDec(Value))
End Function

Public Function Abss(a As Angle) As AngleDec
    Set Abss = MNew.AngleDec(CDec(Math.Abs(m_Angle)))
End Function

' ############################## '    trigonometric functions    ' ############################## '
Public Function Sinus()              ' aka sin
    Sinus = CDec(VBA.Math.Sin(m_Angle))
End Function
Public Function Cosinus()            ' aka cos
    Cosinus = CDec(VBA.Math.Cos(m_Angle))
End Function
Public Function Tangens()            ' aka tan
    Tangens = CDec(VBA.Math.Tan(m_Angle))
End Function

Public Function Cosecans()           ' aka csc
    If m_Angle = 0 Then Exit Function
    Cosecans = CDec(CDec(1) / CDec(VBA.Math.Sin(m_Angle)))
End Function
Public Function Secans()             ' aka sec
    If m_Angle = CDec(MMath.Pi / CDec(2)) Then Exit Function
    Secans = CDec(CDec(1) / CDec(VBA.Math.Cos(m_Angle)))
End Function
Public Function Cotangens()          ' aka cot
    If m_Angle = 0 Then Exit Function
    Cotangens = CDec(CDec(1) / CDec(VBA.Math.Tan(m_Angle)))
End Function

' ############################## '    trigonometric inverse funktions    ' ############################## '
Public Function ArcusSinusF(ByVal y)         ' aka arcsin
    Select Case y
        Case 1
            ArcusSinusF = CDec(0.5) * MMath.Pi
        Case -1
            ArcusSinusF = CDec(-0.5) * MMath.Pi
        Case Else
            ArcusSinusF = CDec(VBA.Math.Atn(CDec(y) / Sqr(CDec(1) - CDec(y) * CDec(y))))
    End Select
End Function
Public Function ArcusSinus(ByVal y) As AngleDec         ' aka arcsin
    Set ArcusSinus = New Angle: ArcusSinus.New_ ArcusSinusF(CDec(y))
End Function

Public Function ArcusCosinusF(ByVal x)       ' aka arccos
    If x < -1 Or 1 < x Then Exit Function
    ArcusCosinusF = 0.5 * MMath.Pi - ArcusSinusF(CDec(x))
End Function
Public Function ArcusCosinus(ByVal x) As AngleDec       ' aka arccos
    Set ArcusCosinus = New AngleDec: ArcusCosinus.New_ ArcusCosinusF(CDec(x))
End Function

Public Function ArcusTangensF(ByVal T)    ' aka arctan
    ArcusTangensF = VBA.Math.Atn(CDec(T))
End Function
Public Function ArcusTangens(ByVal T) As AngleDec      ' aka arctan
    Set ArcusTangens = New AngleDec: ArcusTangens.New_ ArcusTangensF(CDec(T))
End Function

'ArcusTangensXY: also known as ATan2
Public Function ArcusTangensXYF(ByVal x, _
                                ByVal y)   ' aka atan2
    If y > 0 Then
        If x > 0 Then       ' 1. Quadrant
            ArcusTangensXYF = Atn(Abs(y) / Abs(CDec(x))) '+ MMath.Pi * 0#
        ElseIf x < 0 Then   ' 2. Quadrant
            ArcusTangensXYF = -Atn(Abs(y) / Abs(CDec(x))) + MMath.Pi '* 1#
        Else 'If x = 0 Then ' pos. Y-Achse
            ArcusTangensXYF = 0.5 * MMath.Pi
        End If
    ElseIf y < 0 Then
        If x < 0 Then       ' 3. Quadrant
            ArcusTangensXYF = CDec(Atn(CDec(Abs(CDec(y))) / CDec(Abs(CDec(x)))) + MMath.Pi) '* 1#
        ElseIf x > 0 Then   ' 4. Quadrant
            ArcusTangensXYF = CDec(-Atn(CDec(Abs(CDec(y))) / CDec(Abs(CDec(x)))) + MMath.Pi * CDec(2))
        Else 'If x = 0 Then ' neg. Y-Achse
            ArcusTangensXYF = CDec(1.5) * MMath.Pi
        End If
    Else 'If y = 0 Then
        If x > 0 Then       ' pos. X-Achse
            ArcusTangensXYF = CDec(0)
        ElseIf x < 0 Then   ' neg. X-Achse
            ArcusTangensXYF = MMath.Pi
        Else 'If x = 0 Then ' Nullpunkt
            ArcusTangensXYF = CDec(0)
        End If
    End If
End Function
Public Function ArcusTangensXY(ByVal x, _
                               ByVal y) As AngleDec    ' aka atan2
    Set ArcusTangensXY = New AngleDec: ArcusTangensXY.New_ ArcusTangensXYF(CDec(x), CDec(y))
End Function

Public Function ArcusCosecansF(ByVal y)    ' aka arccsc
    If y = 0 Then Exit Function
    ArcusCosecansF = ArcusSinusF(1 / y)
End Function
Public Function ArcusCosecans(ByVal y) As AngleDec     ' aka arccsc
    Set ArcusCosecans = New AngleDec: ArcusCosecans.New_ ArcusCosecansF(CDec(y))
End Function

Public Function ArcusSecansF(ByVal x)     ' aka arcsec
    If x = 0 Then Exit Function
    ArcusSecansF = ArcusCosinusF(1 / x)
End Function
Public Function ArcusSecans(ByVal x) As AngleDec       ' aka arcsec
    Set ArcusSecans = New AngleDec: ArcusSecans.New_ ArcusSecansF(CDec(x))
End Function

Public Function ArcusCotangensF(ByVal T)   ' aka arccot
    ArcusCotangensF = MMath.Pi * CDec(0.5) - CDec(VBA.Math.Atn(CDec(T)))
End Function
Public Function ArcusCotangens(ByVal T) As AngleDec    ' aka arccot
    Set ArcusCotangens = New AngleDec: ArcusCotangens.New_ ArcusCotangensF(CDec(T))
End Function

' ###################### '  hyperbolic funktions   ' ##################### '
Public Function SinusHyperbolicus()                 ' aka sinh
    SinusHyperbolicus = CDec((CDec(VBA.Math.Exp(m_Angle)) - CDec(VBA.Math.Exp(-m_Angle))) / CDec(2))
End Function
Public Function CosinusHyperbolicus()               ' aka sinh
    CosinusHyperbolicus = CDec((CDec(VBA.Math.Exp(m_Angle)) + CDec(VBA.Math.Exp(-m_Angle))) / CDec(2))
End Function
Public Function TangensHyperbolicus()               ' aka tanh
    TangensHyperbolicus = CDec((CDec(VBA.Math.Exp(m_Angle)) - CDec(VBA.Math.Exp(-m_Angle))) / CDec((CDec(VBA.Math.Exp(m_Angle)) + CDec(VBA.Math.Exp(-m_Angle)))))
End Function

Public Function CosecansHyperbolicus()              ' aka csch
    If m_Angle = 0 Then Exit Function
    CosecansHyperbolicus = CDec(CDec(2) / (CDec(VBA.Math.Exp(m_Angle)) - CDec(VBA.Math.Exp(-m_Angle))))
End Function
Public Function SecansHyperbolicus()                ' aka sech
    SecansHyperbolicus = CDec(CDec(2) / (CDec(VBA.Math.Exp(m_Angle)) + CDec(VBA.Math.Exp(-m_Angle))))
End Function
Public Function CotangensHyperbolicus()             ' aka coth
    If m_Angle = 0 Then Exit Function
    CotangensHyperbolicus = CDec((CDec(VBA.Math.Exp(m_Angle)) + CDec(VBA.Math.Exp(-m_Angle))) / (CDec(VBA.Math.Exp(m_Angle)) - CDec(VBA.Math.Exp(-m_Angle))))
End Function

' ###################### '    hyperbolic inverse funktions    ' ###################### '
Public Function AreaSinusHyperbolicusF(ByVal y)      ' aka arsinh
    AreaSinusHyperbolicusF = CDec(VBA.Math.Log(CDec(y) + CDec(Sqr(CDec(y) * CDec(y) + CDec(1)))))
End Function
Public Function AreaSinusHyperbolicus(ByVal y) As AngleDec        ' aka arsinh
    Set AreaSinusHyperbolicus = New AngleDec: AreaSinusHyperbolicus.New_ AreaSinusHyperbolicusF(y)
End Function

Public Function AreaCosinusHyperbolicusF(ByVal x)   ' aka arcosh
    If x < 1 Then Exit Function
    AreaCosinusHyperbolicusF = CDec(VBA.Math.Log(CDec(x) + CDec(Sqr(CDec(x) * CDec(x) - CDec(1)))))
End Function
Public Function AreaCosinusHyperbolicus(ByVal x) As AngleDec      ' aka arcosh
    Set AreaCosinusHyperbolicus = New AngleDec: AreaCosinusHyperbolicus.New_ AreaCosinusHyperbolicusF(CDec(x))
End Function

Public Function AreaTangensHyperbolicusF(ByVal T)   ' aka artanh
    If T < -1 Or 1 < T Then Exit Function
    AreaTangensHyperbolicusF = VBA.Math.Log((1 + T) / (1 - T)) / 2
End Function
Public Function AreaTangensHyperbolicus(ByVal T) As AngleDec      ' aka artanh
    Set AreaTangensHyperbolicus = New AngleDec: AreaTangensHyperbolicus.New_ AreaTangensHyperbolicusF(CDec(T))
End Function

Public Function AreaCosecansHyperbolicusF(ByVal x)   ' aka arcsch
    If x = 0 Then Exit Function
    AreaCosecansHyperbolicusF = CDec(VBA.Math.Log((CDec(Sgn(CDec(x))) * CDec(Sqr(CDec(x) * CDec(x) + CDec(1))) + CDec(1)) / CDec(x)))
End Function
Public Function AreaCosecansHyperbolicus(ByVal x) As AngleDec     ' aka arcsch
    Set AreaCosecansHyperbolicus = New AngleDec: AreaCosecansHyperbolicus.New_ AreaCosecansHyperbolicusF(CDec(x))
End Function

Public Function AreaSecansHyperbolicusF(ByVal x)      ' aka arsech
    If x <= 0 Or 1 <= x Then Exit Function
    AreaSecansHyperbolicusF = CDec(VBA.Math.Log((CDec(Sqr(CDec(-x) * CDec(x) + CDec(1))) + CDec(1)) / CDec(x)))
End Function
Public Function AreaSecansHyperbolicus(ByVal x) As AngleDec       ' aka arsech
    Set AreaSecansHyperbolicus = New AngleDec: AreaSecansHyperbolicus.New_ AreaSecansHyperbolicusF(CDec(x))
End Function

Public Function AreaCotangensHyperbolicusF(ByVal x)   ' aka arcoth
    If x < -1 Or 1 < x Then AreaCotangensHyperbolicusF = CDec(VBA.Math.Log((CDec(x) + CDec(1)) / (CDec(x) - CDec(1))) / CDec(2))
End Function
Public Function AreaCotangensHyperbolicus(ByVal x) As AngleDec    ' aka arcoth
    Set AreaCotangensHyperbolicus = New AngleDec: AreaCotangensHyperbolicus.New_ AreaCotangensHyperbolicusF(CDec(x))
End Function


' v ############################## v '      Implements Angle      ' v ############################## v '

Private Function Angle_Abss(a As Angle) As Angle
    Set Angle_Abss = Me.Abss(a)
End Function

Private Function Angle_Add(a As Angle) As Angle
    Set Angle_Add = Me.Add(a)
End Function

Private Sub Angle_AddDeg(ByVal d As Double)
    m_Angle = m_Angle + d * MMath.Pi / 180#
End Sub

Private Function Angle_ArcusCosecans(ByVal y As Double) As Angle
    Set Angle_ArcusCosecans = Me.ArcusCosecans(CDec(y))
End Function

Private Function Angle_ArcusCosecansF(ByVal y As Double) As Double
    Angle_ArcusCosecansF = Me.ArcusCosecansF(CDec(y))
End Function

Private Function Angle_ArcusCosinus(ByVal x As Double) As Angle
    Set Angle_ArcusCosinus = Me.ArcusCosinus(CDec(x))
End Function

Private Function Angle_ArcusCosinusF(ByVal x As Double) As Double
    Angle_ArcusCosinusF = Me.ArcusCosinusF(CDec(x))
End Function

Private Function Angle_ArcusCotangens(ByVal T As Double) As Angle
    Set Angle_ArcusCotangens = Me.ArcusCotangens(CDec(T))
End Function

Private Function Angle_ArcusCotangensF(ByVal T As Double) As Double
    Angle_ArcusCotangensF = Me.ArcusCotangensF(CDec(T))
End Function

Private Function Angle_ArcusSecans(ByVal x As Double) As Angle
    Set Angle_ArcusSecans = Me.ArcusSecans(CDec(x))
End Function

Private Function Angle_ArcusSecansF(ByVal x As Double) As Double
    Angle_ArcusSecansF = Me.ArcusSecansF(CDec(x))
End Function

Private Function Angle_ArcusSinus(ByVal y As Double) As Angle
    Set Angle_ArcusSinus = Me.ArcusSinus(CDec(y))
End Function

Private Function Angle_ArcusSinusF(ByVal y As Double) As Double
    Angle_ArcusSinusF = Me.ArcusSinusF(CDec(y))
End Function

Private Function Angle_ArcusTangens(ByVal T As Double) As Angle
    Set Angle_ArcusTangens = Me.ArcusTangens(CDec(T))
End Function

Private Function Angle_ArcusTangensF(ByVal T As Double) As Double
    Angle_ArcusTangensF = Me.ArcusTangensF(CDec(T))
End Function

Private Function Angle_ArcusTangensXY(ByVal x As Double, ByVal y As Double) As Angle
    Angle_ArcusTangensXY = Me.ArcusTangensXY(CDec(x), CDec(y))
End Function

Private Function Angle_ArcusTangensXYF(ByVal x As Double, ByVal y As Double) As Double
    Angle_ArcusTangensXYF = Me.ArcusTangensXYF(CDec(x), CDec(y))
End Function

Private Function Angle_AreaCosecansHyperbolicus(ByVal x As Double) As Angle
    Angle_AreaCosecansHyperbolicus = Me.AreaCosecansHyperbolicus(CDec(x))
End Function

Private Function Angle_AreaCosecansHyperbolicusF(ByVal x As Double) As Double
    Angle_AreaCosecansHyperbolicusF = Me.AreaCosecansHyperbolicusF(CDec(x))
End Function

Private Function Angle_AreaCosinusHyperbolicus(ByVal x As Double) As Angle
    Angle_AreaCosinusHyperbolicus = Me.AreaCosinusHyperbolicus(CDec(x))
End Function

Private Function Angle_AreaCosinusHyperbolicusF(ByVal x As Double) As Double
    Angle_AreaCosinusHyperbolicusF = Me.AreaCosinusHyperbolicusF(CDec(x))
End Function

Private Function Angle_AreaCotangensHyperbolicus(ByVal x As Double) As Angle
    Angle_AreaCotangensHyperbolicus = Me.AreaCotangensHyperbolicus(CDec(x))
End Function

Private Function Angle_AreaCotangensHyperbolicusF(ByVal x As Double) As Double
    Angle_AreaCotangensHyperbolicusF = Me.AreaCotangensHyperbolicusF(CDec(x))
End Function

Private Function Angle_AreaSecansHyperbolicus(ByVal x As Double) As Angle
    Set Angle_AreaSecansHyperbolicus = Me.AreaSecansHyperbolicus(CDec(x))
End Function

Private Function Angle_AreaSecansHyperbolicusF(ByVal x As Double) As Double
    Angle_AreaSecansHyperbolicusF = Me.AreaSecansHyperbolicusF(CDec(x))
End Function

Private Function Angle_AreaSinusHyperbolicus(ByVal y As Double) As Angle
    Set Angle_AreaSinusHyperbolicus = Me.AreaSinusHyperbolicus(CDec(y))
End Function

Private Function Angle_AreaSinusHyperbolicusF(ByVal y As Double) As Double
    Angle_AreaSinusHyperbolicusF = Me.AreaSinusHyperbolicusF(CDec(y))
End Function

Private Function Angle_AreaTangensHyperbolicus(ByVal T As Double) As Angle
    Set Angle_AreaTangensHyperbolicus = Me.AreaTangensHyperbolicus(CDec(T))
End Function

Private Function Angle_AreaTangensHyperbolicusF(ByVal T As Double) As Double
    Angle_AreaTangensHyperbolicusF = Me.AreaTangensHyperbolicusF(CDec(T))
End Function

Private Function Angle_Cosecans() As Double
    Angle_Cosecans = Me.Cosecans
End Function

Private Function Angle_CosecansHyperbolicus() As Double
    Angle_CosecansHyperbolicus = Me.CosecansHyperbolicus
End Function

Private Function Angle_Cosinus() As Double
    Angle_Cosinus = Me.Cosinus
End Function

Private Function Angle_CosinusHyperbolicus() As Double
    Angle_CosinusHyperbolicus = Me.CosinusHyperbolicus
End Function

Private Function Angle_Cotangens() As Double
    Angle_Cotangens = Me.Cotangens
End Function

Private Function Angle_CotangensHyperbolicus() As Double
    Angle_CotangensHyperbolicus = Me.CotangensHyperbolicus
End Function

Private Property Let Angle_Dir(ByVal RHS As String)
    Me.Dir = RHS
End Property

Private Property Get Angle_Dir() As String
    Angle_Dir = Me.Dir
End Property

Private Function Angle_Div(ByVal Value As Double) As Angle
    Set Angle_Div = Me.Div(CDec(Value))
End Function

Private Property Let Angle_Grad(ByVal RHS As Long)
    Me.Grad = RHS
End Property

Private Property Get Angle_Grad() As Long
    Angle_Grad = Me.Grad
End Property

Private Property Get Angle_GradF() As Double
    Angle_GradF = Me.GradF
End Property

Private Property Let Angle_Millisec(ByVal RHS As Long)
    Me.Millisec = RHS
End Property

Private Property Get Angle_Millisec() As Long
    Angle_Millisec = Me.Millisec
End Property

Private Property Let Angle_MillisecF(ByVal RHS As Double)
    Me.MillisecF = RHS
End Property

Private Property Get Angle_MillisecF() As Double
    Angle_MillisecF = Me.MillisecF
End Property

Private Property Let Angle_Minute(ByVal RHS As Long)
    Me.Minute = RHS
End Property

Private Property Get Angle_Minute() As Long
    Angle_Minute = Me.Minute
End Property

Private Property Get Angle_MinuteF() As Double
    Angle_MinuteF = Me.MinuteF
End Property

Private Function Angle_Mul(ByVal Value As Double) As Angle
    Set Angle_Mul = Me.Mul(CDec(Value))
End Function

Private Function Angle_Parse(s As String) As Boolean
    Angle_Parse = Me.Parse(s)
End Function

Private Function Angle_Secans() As Double
    Angle_Secans = Me.Secans
End Function

Private Function Angle_SecansHyperbolicus() As Double
    Angle_SecansHyperbolicus = Me.SecansHyperbolicus
End Function

Private Property Let Angle_Second(ByVal RHS As Long)
    Me.Second = RHS
End Property

Private Property Get Angle_Second() As Long
    Angle_Second = Me.Second
End Property

Private Property Let Angle_SecondF(ByVal RHS As Double)
    Me.SecondF = RHS
End Property

Private Property Get Angle_SecondF() As Double
    Angle_SecondF = Me.SecondF
End Property

Private Function Angle_Sinus() As Double
    Angle_Sinus = Me.Sinus
End Function

Private Function Angle_SinusHyperbolicus() As Double
    Angle_SinusHyperbolicus = Me.SinusHyperbolicus
End Function

Private Function Angle_Subt(a As Angle) As Angle
    Set Angle_Subt = Me.Subt(a)
End Function

Private Function Angle_Tangens() As Double
    Angle_Tangens = Me.Tangens
End Function

Private Function Angle_TangensHyperbolicus() As Double
    Angle_TangensHyperbolicus = Me.TangensHyperbolicus
End Function

Private Function Angle_ToGon() As Double
    Angle_ToGon = Me.ToGon
End Function

Private Function Angle_ToGrad() As Double
    Angle_ToGrad
End Function

Private Function Angle_ToRad() As Double
    Angle_ToRad = Me.ToRad
End Function

Private Property Get Angle_Value()
    Angle_Value = m_Angle
End Property
