VERSION 5.00
Begin VB.Form FTestAngle 
   Caption         =   "Test Angle"
   ClientHeight    =   7215
   ClientLeft      =   15795
   ClientTop       =   3000
   ClientWidth     =   9015
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FTestAngle.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7215
   ScaleWidth      =   9015
   Begin VB.CommandButton BtnTestGraphic 
      Caption         =   "Test Graphic"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton BtnTestNumeric 
      Caption         =   "Test Numeric"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.PictureBox PnlTestNumeric 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   8295
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3765
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2415
      End
      Begin VB.CommandButton BtnParseAngle 
         Caption         =   "Parse Angle"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "45,55555"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   4800
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   $"FTestAngle.frx":1782
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   7050
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "        "
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   1
         Top             =   1320
         Width           =   360
      End
   End
   Begin VB.PictureBox PnlTestGraphic 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4875
      ScaleWidth      =   8235
      TabIndex        =   8
      Top             =   480
      Width           =   8295
      Begin VB.ComboBox CmbGrafikTest 
         Height          =   345
         Left            =   2640
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   120
         Width           =   2895
      End
   End
End
Attribute VB_Name = "FTestAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_alpha As Angle

Private Type Point2D
    X As Double
    Y As Double
End Type
Private Type Points
    Arr() As Point2D
End Type
Private m_Graphs() As Points

Private Sub Form_Load()
    With List1
        .AddItem "123d46'12.3456''"
        .AddItem " 47.37816667"
        .AddItem "-8.23250000"
        .AddItem "N 47.38195°"
        .AddItem " E 8.54879° "
        .AddItem "S 47°12.625'"
        .AddItem " W 7° 27.103' "
        .AddItem "N 47°12.625'"
        .AddItem "N 47°22.690'"
        .AddItem " E 8° 13.950'"
        .AddItem "E7d26'22.500"""
        .AddItem "-1/2p"
        .AddItem "1/3p"
        .AddItem "1/4p"
        .AddItem "1/5p"
        .AddItem "2/3p"
        .AddItem "3/4p"
        .AddItem "3/2p"
    End With
    FillCombo CmbGrafikTest
    FillGraphs
    PnlTestGraphic.ScaleMode = 7 'cm
    BtnTestNumeric.Value = True '_Click
    BtnParseAngle.Value = True
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = PnlTestNumeric.Left:   T = PnlTestNumeric.Top
    W = Me.ScaleWidth:  H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then PnlTestNumeric.Move L, T, W, H
    If W > 0 And H > 0 Then PnlTestGraphic.Move L, T, W, H
    L = (PnlTestGraphic.ScaleWidth - CmbGrafikTest.Width) / 2: T = CmbGrafikTest.Top
    CmbGrafikTest.Move L, T
    L = List1.Left:  T = List1.Top
    W = List1.Width: H = PnlTestNumeric.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    L = Text2.Left: T = Text2.Top
    W = PnlTestNumeric.ScaleWidth - L: H = PnlTestNumeric.ScaleHeight - T
    If W > 0 And H > 0 Then Text2.Move L, T, W, H
End Sub

Private Sub FillCombo(aCMB As ComboBox)
    Dim a: a = Array( _
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
        For i = 0 To UBound(a)
            Call .AddItem(a(i))
        Next
    End With
    aCMB.Text = "Sinus"
End Sub

Private Sub BtnTestNumeric_Click()
    PnlTestNumeric.ZOrder 0
    BtnTestNumeric.ZOrder 0
    BtnTestGraphic.ZOrder 0
    BtnTestNumeric.Default = True
End Sub

Private Sub BtnTestGraphic_Click()
    PnlTestGraphic.ZOrder 0
    BtnTestNumeric.ZOrder 0
    BtnTestGraphic.ZOrder 0
    BtnTestGraphic.Default = True
    CmbGrafikTest.ListIndex = 0
End Sub

Private Sub PnlTestGraphic_Resize()
    PnlTestGraphic.Refresh
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then BtnParseAngle_Click
End Sub

Private Sub BtnParseAngle_Click()
    ParseAngle Text1.Text
    List1.AddItem Text1.Text
End Sub

Private Sub List1_Click()
    ParseAngle List1.Text
End Sub

Sub ParseAngle(s As String)
    'Shows parsing the angle is correct, done right and leaves many useful options to the user
    Set m_alpha = MNew.AngleS(s)
    UpdateView
End Sub

Public Sub UpdateView()
    With m_alpha
        Label1.Caption = .Value & " (rad)" & vbCrLf & _
                         .ToGrad & " °" & vbCrLf & _
                         .ToStr_DMS & vbCrLf & _
                         .GradF & " °" & vbCrLf & _
                         .Grad & " °" & vbCrLf & _
                         .MinuteF & " '" & vbCrLf & _
                         .Minute & " '" & vbCrLf & _
                         .SecondF & """" & vbCrLf & _
                         .Second & """" & vbCrLf & _
                         .MillisecF & vbCrLf & _
                         .Millisec
        'shows all trigonometric functions
        Dim s As String: s = ""
        s = s & "Sin(alpha) = " & .Sinus & vbCrLf
        s = s & "Cos(alpha) = " & .Cosinus & vbCrLf
        s = s & "Tan(alpha) = " & .Tangens & vbCrLf
        s = s & "Sec(alpha) = " & .Secans & vbCrLf
        s = s & "Csc(alpha) = " & .Cosecans & vbCrLf
        s = s & "Cot(alpha) = " & .Cotangens & vbCrLf
        s = s & "Sinh(alpha) = " & .SinusHyperbolicus & vbCrLf
        s = s & "Cosh(alpha) = " & .CosinusHyperbolicus & vbCrLf
        s = s & "Tanh(alpha) = " & .TangensHyperbolicus & vbCrLf
        s = s & "Sech(alpha) = " & .SecansHyperbolicus & vbCrLf
        s = s & "Csch(alpha) = " & .CosecansHyperbolicus & vbCrLf
        s = s & "Coth(alpha) = " & .CotangensHyperbolicus & vbCrLf
        Text2.Text = s
    End With
End Sub

Private Sub CmbGrafikTest_Click()
    PnlTestGraphic.Refresh
End Sub

Sub FillGraphs()
    
    ReDim m_Graphs(0 To 25)
    Dim MinPoints As Long: MinPoints = -360
    Dim MaxPoints As Long: MaxPoints = 360
    ReDim AngleArr(MinPoints To MaxPoints) As Angle
    ReDim Points(MinPoints To MaxPoints) As Point2D
    Dim i As Long
    For i = MinPoints To MaxPoints
        Set AngleArr(i) = MNew.AngleD(i)
        Points(i).X = AngleArr(i).Value
    Next
    
    Dim j As Long
    Dim trigono As New Angle
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Sinus:             Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Cosinus:           Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Tangens:           Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Cosecans:          Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Secans:            Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).Cotangens:         Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusSinusF(Points(i).X):        Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusCosinusF(Points(i).X):      Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusTangensF(Points(i).X):      Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusCosecansF(Points(i).X):     Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusSecansF(Points(i).X):       Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.ArcusCotangensF(Points(i).X):    Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).SinusHyperbolicus:     Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).CosinusHyperbolicus:   Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).TangensHyperbolicus:   Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).CosecansHyperbolicus:  Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).SecansHyperbolicus:    Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = AngleArr(i).CotangensHyperbolicus: Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaSinusHyperbolicusF(Points(i).X):       Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaCosinusHyperbolicusF(Points(i).X):     Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaTangensHyperbolicusF(Points(i).X):     Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaCosecansHyperbolicusF(Points(i).X):    Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaSecansHyperbolicusF(Points(i).X):      Next
    m_Graphs(j).Arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).Y = trigono.AreaCotangensHyperbolicusF(Points(i).X):   Next
    m_Graphs(j).Arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).Y = MMath.SinusCardinalis(Points(i).X * 2 * MMath.Pi) * 4: Next
    m_Graphs(j).Arr = Points: j = j + 1
    
End Sub

Private Sub PnlTestGraphic_Paint()
    On Error Resume Next
    Dim XN As Double, YN As Double
    'the coordinates of the center point (for translation)
    XN = PnlTestGraphic.ScaleWidth / 2
    YN = PnlTestGraphic.ScaleHeight / 2
    
    'draw the coord-system
    'draw the X-axis
    Dim X1 As Double, Y1 As Double
    Dim X2 As Double, Y2 As Double
    X1 = 0:                         Y1 = YN
    X2 = PnlTestGraphic.ScaleWidth: Y2 = Y1
    PnlTestGraphic.Line (X1, Y1)-(X2, Y2)
    
    'draw the Y-axis
    X1 = XN: Y1 = 0
    X2 = X1: Y2 = PnlTestGraphic.ScaleHeight
    PnlTestGraphic.Line (X1, Y1)-(X2, Y2)
    
    Dim j As Long: j = CmbGrafikTest.ListIndex
    Dim Pts() As Point2D: Pts = m_Graphs(j).Arr
    'draw the curve
    Dim i As Long
    For i = LBound(Pts) To UBound(Pts) - 1
        X1 = Pts(i).X + XN:        Y1 = -Pts(i).Y + YN
        X2 = Pts(i + 1).X + XN:    Y2 = -Pts(i + 1).Y + YN
        PnlTestGraphic.Line (X1, Y1)-(X2, Y2)
    Next
    
End Sub
