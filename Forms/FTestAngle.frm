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
    x As Double
    y As Double
End Type
Private Type Points
    arr() As Point2D
End Type
Private m_Graphs() As Points

Private Sub Form_Load()
    FillGraphs
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
    PnlTestGraphic.ScaleMode = 7 'cm
End Sub

Private Sub Form_Resize()
    Dim L As Single, t As Single, W As Single, H As Single
    L = PnlTestNumeric.Left:   t = PnlTestNumeric.Top
    W = Me.ScaleWidth:  H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then PnlTestNumeric.Move L, t, W, H
    If W > 0 And H > 0 Then PnlTestGraphic.Move L, t, W, H
    L = (PnlTestGraphic.ScaleWidth - CmbGrafikTest.Width) / 2: t = CmbGrafikTest.Top
    CmbGrafikTest.Move L, t
    L = List1.Left:  t = List1.Top
    W = List1.Width: H = PnlTestNumeric.ScaleHeight - t
    If W > 0 And H > 0 Then List1.Move L, t, W, H
    L = Text2.Left: t = Text2.Top
    W = PnlTestNumeric.ScaleWidth - L: H = PnlTestNumeric.ScaleHeight - t
    If W > 0 And H > 0 Then Text2.Move L, t, W, H
    Dim i As Long: i = CmbGrafikTest.ListIndex
    If i >= 0 Then DrawGraph i
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
    'Show parsing the angle is correct, done right and leaves many useful options to the user
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
        'show all trigonometric functions
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
    DrawGraph CmbGrafikTest.ListIndex
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
        Points(i).x = AngleArr(i).Value
    Next
    
    Dim j As Long
    Dim trigono As New Angle
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Sinus:             Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Cosinus:           Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Tangens:           Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Cosecans:          Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Secans:            Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).Cotangens:         Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusSinusF(Points(i).x):        Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusCosinusF(Points(i).x):      Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusTangensF(Points(i).x):      Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusCosecansF(Points(i).x):     Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusSecansF(Points(i).x):       Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.ArcusCotangensF(Points(i).x):    Next
    m_Graphs(j).arr = Points: j = j + 1
    
    
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).SinusHyperbolicus:     Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).CosinusHyperbolicus:   Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).TangensHyperbolicus:   Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).CosecansHyperbolicus:  Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).SecansHyperbolicus:    Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = AngleArr(i).CotangensHyperbolicus: Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaSinusHyperbolicusF(Points(i).x):       Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaCosinusHyperbolicusF(Points(i).x):     Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaTangensHyperbolicusF(Points(i).x):     Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaCosecansHyperbolicusF(Points(i).x):    Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaSecansHyperbolicusF(Points(i).x):      Next
    m_Graphs(j).arr = Points: j = j + 1
    For i = MinPoints To MaxPoints: Points(i).y = trigono.AreaCotangensHyperbolicusF(Points(i).x):   Next
    m_Graphs(j).arr = Points: j = j + 1
    
    For i = MinPoints To MaxPoints: Points(i).y = MMath.SinusCardinalis(Points(i).x * 2 * MMath.Pi) * 4: Next
    m_Graphs(j).arr = Points: j = j + 1
    
End Sub

'Private Sub PnlTestGraphic_Paint()
'    'DrawGraph
'    'nope do not do this here, otherwise it will crash
'End Sub

Sub DrawGraph(ByVal Index As Long)
Try: On Error GoTo Catch
    Dim XN As Double, YN As Double
    'the coordinates of the center point (for translation)
    XN = PnlTestGraphic.ScaleWidth / 2
    YN = PnlTestGraphic.ScaleHeight / 2
    
    PnlTestGraphic.Cls
    'draw the coord-system
    'draw the X-axis
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, Y2 As Double
    x1 = 0:                         y1 = YN
    x2 = PnlTestGraphic.ScaleWidth: Y2 = y1
    PnlTestGraphic.Line (x1, y1)-(x2, Y2)
    
    'draw the Y-axis
    x1 = XN: y1 = 0
    x2 = x1: Y2 = PnlTestGraphic.ScaleHeight
    PnlTestGraphic.Line (x1, y1)-(x2, Y2)
    
    Dim j As Long: j = Index
    Dim Pts() As Point2D: Pts = m_Graphs(j).arr
    'draw the curve
    Dim i As Long
    For i = LBound(Pts) To UBound(Pts) - 1
        x1 = Pts(i).x + XN:        y1 = -Pts(i).y + YN
        x2 = Pts(i + 1).x + XN:    Y2 = -Pts(i + 1).y + YN
        If x1 < 0 Then x1 = -1
        If x2 < 0 Then x2 = -1
        If y1 < 0 Then y1 = -1
        If Y2 < 0 Then Y2 = -1
        If x1 > PnlTestGraphic.ScaleWidth Then x1 = PnlTestGraphic.ScaleWidth
        If x2 > PnlTestGraphic.ScaleWidth Then x2 = PnlTestGraphic.ScaleWidth
        If y1 > PnlTestGraphic.ScaleHeight Then y1 = PnlTestGraphic.ScaleHeight
        If Y2 > PnlTestGraphic.ScaleHeight Then Y2 = PnlTestGraphic.ScaleHeight
        PnlTestGraphic.Line (x1, y1)-(x2, Y2)
    Next
    Exit Sub
Catch:
    MsgBox "Error in :" & TypeName(Me) & "::DrawGraph"
End Sub
