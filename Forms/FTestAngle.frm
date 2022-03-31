VERSION 5.00
Begin VB.Form FTestAngle 
   Caption         =   "Test Angle"
   ClientHeight    =   3975
   ClientLeft      =   15795
   ClientTop       =   3000
   ClientWidth     =   8175
   Icon            =   "FTestAngle.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3975
   ScaleWidth      =   8175
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
      Height          =   3255
      Left            =   4800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "45,55555"
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton BtnParseAngle 
      Caption         =   "Parse Angle"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   $"FTestAngle.frx":1782
      Height          =   555
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "'        "
      Height          =   195
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   390
   End
End
Attribute VB_Name = "FTestAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim alpha As Angle

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
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then BtnParseAngle_Click
End Sub

Private Sub BtnParseAngle_Click()
    ParseAngle Text1.Text
End Sub

Private Sub List1_Click()
    ParseAngle List1.Text
End Sub

Sub ParseAngle(s As String)
    'Shows parsing the angle is correct, done right and leaves many useful options to the user
    Set alpha = MNew.AngleS(s)
    UpdateView
End Sub

Public Sub UpdateView()
    With alpha
        Label1.Caption = .Value & " (rad)" & vbCrLf & _
                         .ToGrad & " °" & vbCrLf & _
                         .ToStr_GMS & vbCrLf & _
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
