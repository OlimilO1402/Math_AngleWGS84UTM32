VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbGrafikTest 
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox PBTrigo 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   480
      Width           =   5895
   End
   Begin VB.CommandButton BtnTextTest 
      Caption         =   "TextTest"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TxtTrigo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label LblGrafikTest 
      Caption         =   "Grafik Test:"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call FillCombo(CmbGrafikTest)
    Call BtnTextTest_Click
    PBTrigo.ScaleMode = 7 'cm
End Sub
Private Sub FillCombo(aCMB As ComboBox)
    Dim A: A = Array( _
    "Sinus", "Cosinus", "Tangens", _
    "Cosecans", "Secans", "Cotangens", _
    "ArcusSinus", "ArcusCosinus", "ArcusTangens", _
    "ArcusCosecans", "ArcusSecans", "ArcusCotangens", _
    "SinusHyperbolicus", "CosinusHyperbolicus", "TangensHyperbolicus", _
    "CosecansHyperbolicus", "SecansHyperbolicus", "CotangensHyperbolicus", _
    "AreaSinusHyperbolicus", "AreaCosinusHyperbolicus", "AreaTangensHyperbolicus", _
    "AreaCosecansHyperbolicus", "AreaSecansHyperbolicus", "AreaCotangensHyperbolicus", _
    "SinusCardinalis")
    Dim i As Long
    With aCMB
        .Clear
        For i = 0 To UBound(A)
            Call .AddItem(A(i))
        Next
    End With
    aCMB.Text = "Sinus"
End Sub

Private Sub CmbGrafikTest_Click()
    PBTrigo.ZOrder 0
    PBTrigo.Refresh
End Sub
'Private Sub CmbGrafikTest_Change()
'    PBTrigo.ZOrder 0
'    PBTrigo.Refresh
'End Sub

Private Sub BtnTextTest_Click()
    TxtTrigo.ZOrder 0
    Call TextTestTrigo
End Sub

Private Sub TextTestTrigo()
    Dim t As String
    t = t & TestTrigonoMath
    t = t & TestATAN
    t = t & TestAngleConverter
    t = t & TestLogarithm
    t = t & TestFloorCeilingBigMul
    TxtTrigo.Text = t
End Sub
Private Function TestTrigonoMath() As String
    Dim s As String
    Dim A As Double, r As Double
    
    s = s & "Test Trigono Math" & vbCrLf
    
    A = PI / 3
    
    r = ModTrigonoMath.Sinus(A)
    s = s & " Sinus(" & Format(RadToDeg(A), "0.0°") & ") = " & Format(r, "0.000") & vbCrLf
    
    r = ModTrigonoMath.Cosinus(A)
    s = s & " Cosinus(" & Format(RadToDeg(A), "0.0°") & ") = " & Format(r, "0.000") & vbCrLf
    
    r = ModTrigonoMath.Tangens(A)
    s = s & " Tangens(" & Format(RadToDeg(A), "0.0°") & ") = " & Format(r, "0.000") & vbCrLf
    
    r = ModTrigonoMath.Cosecans(A)
    s = s & " Cosecans(" & Format(RadToDeg(A), "0.0°") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.Secans(A)
    s = s & " Secans(" & Format(RadToDeg(A), "0.0°") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.Cotangens(A)
    s = s & " Cotangens(" & Format(RadToDeg(A), "0.0°") & ") = " & CStr(r) & vbCrLf
    
    A = 0.5
    r = ModTrigonoMath.ArcusSinus(A)
    s = s & " ArcusSinus(" & Format(A, "0.000") & ") = " & Format(RadToDeg(r), "0.0°") & vbCrLf

    r = ModTrigonoMath.ArcusCosinus(A)
    s = s & " ArcusCosinus(" & Format(A, "0.000") & ") = " & Format(RadToDeg(r), "0.0°") & vbCrLf
    
    A = Sqr(3)
    r = ModTrigonoMath.ArcusTangens(A)
    s = s & " ArcusTangens(" & Format(A, "0.000") & ") = " & Format(RadToDeg(r), "0.0°") & vbCrLf

    A = 1.5
    r = ModTrigonoMath.ArcusCosecans(A)
    s = s & " ArcusCosecans(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.ArcusSecans(A)
    s = s & " ArcusSecans(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.ArcusCotangens(A)
    s = s & " ArcusCotangens(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    
    r = ModTrigonoMath.SinusHyperbolicus(A)
    s = s & " SinusHyperbolicus(" & Format(RadToDeg(A), "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.CosinusHyperbolicus(A)
    s = s & " CosinusHyperbolicus(" & Format(RadToDeg(A), "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.TangensHyperbolicus(A)
    s = s & " TangensHyperbolicus(" & Format(RadToDeg(A), "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.CosecansHyperbolicus(A)
    s = s & " CosecansHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.SecansHyperbolicus(A)
    s = s & " SecansHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.CotangensHyperbolicus(A)
    s = s & " CotangensHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    A = 0.5
    r = ModTrigonoMath.AreaSinusHyperbolicus(A)
    s = s & " AreaSinusHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    A = 1.5
    r = ModTrigonoMath.AreaCosinusHyperbolicus(A)
    s = s & " AreaCosinusHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    A = 0.5
    r = ModTrigonoMath.AreaTangensHyperbolicus(A)
    s = s & " AreaTangensHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.AreaCosecansHyperbolicus(A)
    s = s & " AreaCosecansHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    r = ModTrigonoMath.AreaSecansHyperbolicus(A)
    s = s & " AreaSecansHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    A = 1.5
    r = ModTrigonoMath.AreaCotangensHyperbolicus(A)
    s = s & " AreaCotangensHyperbolicus(" & Format(A, "0.000") & ") = " & CStr(r) & vbCrLf
    
    TestTrigonoMath = s & vbCrLf
End Function

Private Function TestATAN() As String
    Dim s As String
    s = s & "Test ArcusTangensXY (aka Atan2)" & vbCrLf
    s = s & TestATANToString(0, 0)
    s = s & TestATANToString(1.5, 0)
    s = s & TestATANToString(0, 1.5)
    s = s & TestATANToString(1.2, 1.5)
    s = s & TestATANToString(-1.5, 0)
    s = s & TestATANToString(0, -1.5)
    s = s & TestATANToString(-1.2, -1.5)
    s = s & TestATANToString(1.2, -1.5)
    s = s & TestATANToString(-1.2, 1.5)
    TestATAN = s & vbCrLf
End Function
Private Function TestATANToString(ByVal x As Double, _
                                  ByVal y As Double) As String
    Dim A As Double
    Dim s As String
    A = ArcusTangensXY(x, y)
    s = s & " ArcusTangensXY(x := " & _
       Format$(x, "0.0") & "; y := " & Format$(x, "0.0") & ") = " & _
       Format$(A, "0.000") & vbCrLf
    TestATANToString = s
End Function
Public Function TestAngleConverter() As String

    Dim angleD As Double ' Winkel in Grad
    Dim angleR As Double ' Winkel in Radians
    Dim angleG As Double ' Winkel in Gon

    angleD = 180#

    angleR = DegToRad(angleD)
    angleG = DegToGon(angleD)

    angleD = RadToDeg(angleR)
    angleG = RadToGon(angleR)

    angleD = GonToDeg(angleG)
    angleR = GonToRad(angleG)

    ' 180 3,14159265358979 200
    TestAngleConverter = "Test Winkelkonvertierung: " & vbCrLf & _
        " Angle [deg] = " & Format$(angleD, "0.0") & vbCrLf & _
        " Angle [rad] = " & Format$(angleR, "0.00000") & vbCrLf & _
        " Angle [gon] = " & Format$(angleG, "0.0") & vbCrLf

End Function

Public Function TestLogarithm() As String

    Dim s As String
    Dim x As Double
    Dim b As Double
    Dim N As Double
    Dim L As Double

    s = s & vbCrLf & "Test Logarithmus: " & vbCrLf

    x = 10000
    b = 10
    L = LogN(x, b)
    s = s & " LogN(" & CStr(x) & ", " & CStr(b) & ") = " & CStr(L) & vbCrLf

    b = 5
    L = LogN(x, b)
    s = s & " LogN(" & CStr(x) & ", " & CStr(b) & ") = " & CStr(L) & vbCrLf

    b = 4
    L = LogN(x, b)
    s = s & " LogN(" & CStr(x) & ", " & CStr(b) & ") = " & CStr(L) & vbCrLf

    N = 2
    L = LogN(x, N)
    s = s & " LogN(" & CStr(x) & ", " & CStr(N) & ") = " & CStr(L) & vbCrLf

    ' N = 10
    L = LogN(x)
    s = s & " LogN(" & CStr(x) & ") = " & CStr(L) & vbCrLf

    ' N = 10
    L = Log10(x) ' , N)
    s = s & " Log10(" & CStr(x) & ") = " & CStr(L) & vbCrLf
    
    x = 2
    L = LN(x)
    s = s & " Ln(" & CStr(x) & ") = " & CStr(L) & vbCrLf
    
    ' N = 1
    L = LogN(x, 2)
    s = s & " LogN(" & CStr(x) & ") = " & CStr(L) & vbCrLf
    TestLogarithm = s

End Function
Private Function TestFloorCeilingBigMul() As String
    Dim s As String
    Dim d As Double
    s = s & vbCrLf & "Test Floor, Ceiling, BigMul" & vbCrLf
    
    d = 2147483649.12345
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = 2147483649.56789
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = 1#
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = 0#
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = -1#
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = -2147483649.12345
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    d = -2147483649.56789
    s = s & MessDFC(d, Floor(d), Ceiling(d))
    
    Dim dec
    dec = BigMul(999999999#, 999999999#)
    s = s & " BigMul(999999999, 999999999) = " & CStr(dec)
    TestFloorCeilingBigMul = s
End Function
Private Function MessDFC(ByVal d As Double, _
                         ByVal f As Double, _
                         ByVal c As Double) As String
    MessDFC = "   Floor(" & CStr(d) & ") = " & CStr(f) & vbCrLf & _
              " Ceiling(" & CStr(d) & ") = " & CStr(c) & vbCrLf
End Function
Private Sub PBTrigo_Paint()
'is rather quick'n'dirty
'soll nur zum Testen der Funktionen dienen
    On Error Resume Next
    Dim X1 As Double, Y1 As Double
    Dim X2 As Double, Y2 As Double
    Dim XN As Double, YN As Double
    
    'die Koordinaten des Nullpunkts (zur Verschiebung)
    XN = PBTrigo.ScaleWidth / 2
    YN = PBTrigo.ScaleHeight / 2
    'Koordinatensystem zeichnen
    'die X-Achse zeichnen
    X1 = 0:                  Y1 = YN
    X2 = PBTrigo.ScaleWidth: Y2 = Y1
    PBTrigo.Line (X1, Y1)-(X2, Y2)
    'die Y-Achse zeichnen
    X1 = XN: Y1 = 0
    X2 = X1: Y2 = PBTrigo.ScaleHeight
    PBTrigo.Line (X1, Y1)-(X2, Y2)
    
    'Kurve zeichnen
    Dim i As Long
    Dim DrawItem As String
    DrawItem = CmbGrafikTest.Text
    ReDim Pts(0 To 720) As Double 'zwei Perioden zeichnen -pi...0...+pi
    Dim p As Double
    'Array füllen
    For i = 0 To UBound(Pts)
        p = CDbl(DegToRad(i - 360))
        Select Case DrawItem 'DrawItem
        'Trigonometrische Funktionen
        Case "Sinus"
            Pts(i) = ModTrigonoMath.Sinus(p)
        Case "Cosinus"
            Pts(i) = ModTrigonoMath.Cosinus(p)
        Case "Tangens"
            Pts(i) = ModTrigonoMath.Tangens(p)
        Case "Cosecans"
            Pts(i) = ModTrigonoMath.Cosecans(p)
        Case "Secans"
            Pts(i) = ModTrigonoMath.Secans(p)
        Case "Cotangens"
            Pts(i) = ModTrigonoMath.Cotangens(p)
        'Trigonometrische Umkehrfunktionen
        Case "ArcusSinus"
            Pts(i) = ModTrigonoMath.ArcusSinus(p)
        Case "ArcusCosinus"
            Pts(i) = ModTrigonoMath.ArcusCosinus(p)
        Case "ArcusTangens"
            Pts(i) = ModTrigonoMath.ArcusTangens(p)
        Case "ArcusCosecans"
            Pts(i) = ModTrigonoMath.ArcusCosecans(p)
        Case "ArcusSecans"
            Pts(i) = ModTrigonoMath.ArcusSecans(p)
        Case "ArcusCotangens"
            Pts(i) = ModTrigonoMath.ArcusCotangens(p)
        'Hyperbolische Funktionen
        Case "SinusHyperbolicus"
            Pts(i) = ModTrigonoMath.SinusHyperbolicus(p)
        Case "CosinusHyperbolicus"
            Pts(i) = ModTrigonoMath.CosinusHyperbolicus(p)
        Case "TangensHyperbolicus"
            Pts(i) = ModTrigonoMath.TangensHyperbolicus(p)
        Case "CosecansHyperbolicus"
            Pts(i) = ModTrigonoMath.CosecansHyperbolicus(p)
        Case "SecansHyperbolicus"
            Pts(i) = ModTrigonoMath.SecansHyperbolicus(p)
        Case "CotangensHyperbolicus"
            Pts(i) = ModTrigonoMath.CotangensHyperbolicus(p)
        'Hyperbolische Umkehrfunktionen
        Case "AreaSinusHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaSinusHyperbolicus(p)
        Case "AreaCosinusHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaCosinusHyperbolicus(p)
        Case "AreaTangensHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaTangensHyperbolicus(p)
        Case "AreaCosecansHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaCosecansHyperbolicus(p)
        Case "AreaSecansHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaSecansHyperbolicus(p)
        Case "AreaCotangensHyperbolicus"
            Pts(i) = ModTrigonoMath.AreaCotangensHyperbolicus(p)
        'Spezielle Funktionen
        Case "SinusCardinalis"
            Pts(i) = ModTrigonoMath.SinusCardinalis(p * 2 * PI) * 4
        End Select
    Next
    'Kurve in Array zeichnen
    For i = LBound(Pts) To UBound(Pts) - 1
        X1 = (-2 * PI + i * PI / 180) + XN:       Y1 = -(Pts(i)) + YN
        X2 = (-2 * PI + (i + 1) * PI / 180) + XN: Y2 = -(Pts(i + 1)) + YN
        PBTrigo.Line (X1, Y1)-(X2, Y2)
    Next
  
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, w As Single, H As Single
    Dim brdr As Single
    brdr = 8 * 15
    L = brdr: t = TxtTrigo.Top
    w = Me.ScaleWidth - L - brdr
    H = Me.ScaleHeight - t - brdr
    If w > 0 And H > 0 Then
        Call TxtTrigo.Move(L, t, w, H)
        Call PBTrigo.Move(L, t, w, H)
        PBTrigo.Cls
        Call PBTrigo_Paint
    End If
End Sub

