VERSION 5.00
Begin VB.Form FGeoPos 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dialog Geo Position"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
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
   ScaleHeight     =   2295
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtDescription 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.ComboBox CmbEW 
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.ComboBox CmbNS 
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
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
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
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
      Left            =   2400
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton BtnEditLongitude 
      Caption         =   "Edit Longitude"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton BtnEditLatitude 
      Caption         =   "Edit Latitude"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Angle (Deg Min Sec)"
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Angle (Deg Min Sec)"
      Height          =   225
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "FGeoPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modal Dialog
Private m_GeoPos As GeoPos
Private m_Result As VbMsgBoxResult
Private m_FOwner As Form

Private Sub BtnEditLatitude_Click()
    If FAngle.ShowDialog(m_GeoPos.Latitude, m_FOwner) Then
        UpdateView
    End If
End Sub

Private Sub BtnEditLongitude_Click()
    If FAngle.ShowDialog(m_GeoPos.Longitude, m_FOwner) Then
        UpdateView
    End If
End Sub

Private Sub CmbNS_Click()
    m_GeoPos.Latitude.Dir = Left(CmbNS.Text, 1)
    UpdateView
End Sub

Private Sub CmbEW_Click()
    m_GeoPos.Longitude.Dir = Left(CmbEW.Text, 1)
    UpdateView
End Sub

Private Sub Form_Load()
    With Me.CmbEW
        .Clear
        .AddItem "East"
        .AddItem "West"
    End With
    With Me.CmbNS
        .Clear
        .AddItem "North"
        .AddItem "South"
    End With
End Sub

Public Function ShowDialog(aGeoPos As GeoPos, FOwner As Form) As VbMsgBoxResult
    Set m_FOwner = FOwner
    Set m_GeoPos = aGeoPos.Clone
    UpdateView
    Me.Show vbModal, FOwner
    ShowDialog = m_Result
    If m_Result = vbCancel Then Exit Function
    aGeoPos.NewC m_GeoPos
End Function

Private Sub BtnOK_Click()
    m_Result = vbOK:     Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel: Unload Me
End Sub

Private Sub UpdateView()
    CmbNS.Text = m_GeoPos.Latitude.Dir
    CmbEW.Text = m_GeoPos.Longitude.Dir
    Label1.Caption = m_GeoPos.Latitude.ToStr_DMS
    Label2.Caption = m_GeoPos.Longitude.ToStr_DMS
    TxtDescription.Text = m_GeoPos.Name
End Sub

Private Sub TxtDescription_LostFocus()
    m_GeoPos.Name = TxtDescription.Text
End Sub
