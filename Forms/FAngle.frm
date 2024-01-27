VERSION 5.00
Begin VB.Form FAngle 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dialog Angle"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame1 
      Caption         =   "Angle [ ° ][ ' ][ '' ] (=degree minute second)"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5295
      Begin VB.TextBox TxtDMSAngleSec 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox TxtDMSAngleMin 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtDMSAngleDeg 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Second (Float) [ ' ]:"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Minute (Int) [ ' ]:"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Degree (Int) [ ° ]:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.TextBox TxtAngleDeg 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox TxtAngleRad 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   2880
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
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
      Left            =   1200
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Angle (Float) [ ° ]:"
      Height          =   225
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Angle (Float) [rad]:"
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1485
   End
End
Attribute VB_Name = "FAngle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modal Dialog
Private m_Angle  As AngleDec
Private m_Result As VbMsgBoxResult
Private m_isUpdatingView As Boolean
Private m_LastTB As TextBox
Private m_PropA  As Func1 'PropLet

Public Function ShowDialog(aAngle As AngleDec, FOwner As Form) As VbMsgBoxResult
    Set m_Angle = aAngle.Clone
    UpdateView
    Me.Show vbModal, FOwner
    ShowDialog = m_Result
    If m_Result = vbCancel Then Exit Function
    aAngle.NewC m_Angle
End Function

Private Sub UpdateView()
    m_isUpdatingView = True
    TxtAngleRad.Text = Format(m_Angle.ToRad, "0.###########")
    TxtAngleDeg.Text = Format(m_Angle.ToGrad, "0.##########")
    TxtDMSAngleDeg.Text = Format(m_Angle.Grad, "0")
    TxtDMSAngleMin.Text = Format(m_Angle.Minute, "0")
    TxtDMSAngleSec.Text = Format(m_Angle.SecondF, "0.######")
    m_isUpdatingView = False
End Sub

Private Sub UpdateData()
    If m_LastTB Is Nothing Then Exit Sub
    TB_OnLostFocus
End Sub

Private Sub BtnOK_Click()
    UpdateData
    m_Result = vbOK:     Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel: Unload Me
End Sub

Private Sub TB_OnLostFocus()
    If m_isUpdatingView Then Exit Sub
    Dim s As String: s = m_LastTB.Text
    Dim alp
    If MString.Decimal_TryParse(s, alp) Then
        'm_PropA.Invoke = alp
        m_PropA.Invoke alp
    Else
        MsgBox "Failed to parse a numeric value from: " & s
        Exit Sub
    End If
    UpdateView
End Sub

Private Sub TxtAngleRad_GotFocus()
    Set m_LastTB = TxtAngleRad:       Set m_PropA = MNew.PropLet(m_Angle, "Value")
End Sub
Private Sub TxtAngleRad_LostFocus()
    TB_OnLostFocus
End Sub

Private Sub TxtAngleDeg_GotFocus()
    Set m_LastTB = TxtAngleDeg:       Set m_PropA = MNew.PropLet(m_Angle, "GradF")
End Sub
Private Sub TxtAngleDeg_LostFocus()
    TB_OnLostFocus
End Sub

Private Sub TxtDMSAngleDeg_GotFocus()
    Set m_LastTB = TxtDMSAngleDeg:    Set m_PropA = MNew.PropLet(m_Angle, "Grad")
End Sub
Private Sub TxtDMSAngleDeg_LostFocus()
    TB_OnLostFocus
End Sub

Private Sub TxtDMSAngleMin_GotFocus()
    Set m_LastTB = TxtDMSAngleMin:    Set m_PropA = MNew.PropLet(m_Angle, "Minute")
End Sub
Private Sub TxtDMSAngleMin_LostFocus()
    TB_OnLostFocus
End Sub

Private Sub TxtDMSAngleSec_GotFocus()
    Set m_LastTB = TxtDMSAngleSec:    Set m_PropA = MNew.PropLet(m_Angle, "SecondF")
End Sub
Private Sub TxtDMSAngleSec_LostFocus()
    TB_OnLostFocus
End Sub
