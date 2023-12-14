VERSION 5.00
Begin VB.Form FGeoPos 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dialog Geo Position"
   ClientHeight    =   3615
   ClientLeft      =   150
   ClientTop       =   795
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
   ScaleHeight     =   3615
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.OptionButton OptUTM32 
      Caption         =   "UTM32"
      Height          =   255
      Left            =   1440
      TabIndex        =   17
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OptWGS84 
      Caption         =   "WGS84"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
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
      TabIndex        =   0
      Top             =   3120
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
      Left            =   2880
      TabIndex        =   1
      Top             =   3120
      Width           =   1455
   End
   Begin VB.PictureBox PnlWGS84 
      BorderStyle     =   0  'Kein
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   2
      Top             =   600
      Width           =   5295
      Begin VB.CommandButton BtnEditLatitude 
         Caption         =   "Edit"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton BtnEditLongitude 
         Caption         =   "Edit"
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
         Left            =   4320
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox CmbNS 
         Height          =   345
         Left            =   1080
         TabIndex        =   8
         Top             =   0
         Width           =   975
      End
      Begin VB.ComboBox CmbEW 
         Height          =   345
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtDescription 
         Height          =   975
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox TxtLatitude 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox TxtLongitude 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtNHN 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   225
         Left            =   0
         TabIndex        =   15
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Latitude:"
         Height          =   225
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Longitude:"
         Height          =   225
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height above sea level:"
         Height          =   345
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[m+NHN]"
         Height          =   225
         Left            =   4320
         TabIndex        =   11
         Top             =   960
         Width           =   810
      End
   End
   Begin VB.PictureBox PnlUTM32 
      BorderStyle     =   0  'Kein
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   18
      Top             =   600
      Width           =   5295
      Begin VB.TextBox TxtUTMZone 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtNHNUTM32 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TxtEasting 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   0
         Width           =   3135
      End
      Begin VB.TextBox TxtNorthing 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox TxtDescriptionUTM32 
         Height          =   975
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label LblUTMZone 
         AutoSize        =   -1  'True
         Caption         =   "UTM Zone:"
         Height          =   225
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "[m+NHN]"
         Height          =   225
         Left            =   4320
         TabIndex        =   27
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Height above sea level:"
         Height          =   345
         Left            =   0
         TabIndex        =   26
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Easting:"
         Height          =   225
         Left            =   1320
         TabIndex        =   24
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Northing:"
         Height          =   225
         Left            =   1320
         TabIndex        =   23
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   225
         Left            =   0
         TabIndex        =   22
         Top             =   1440
         Width           =   945
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Begin VB.Menu mnuOptConvert 
         Caption         =   "Convert"
         Checked         =   -1  'True
      End
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
Private m_ShowWGS84 As Boolean
Private m_UTMGps As UTM32

Private Sub Form_Load()
    m_ShowWGS84 = True
    AddEW Me.CmbEW
    AddNS Me.CmbNS
    'ToggleView
End Sub

Private Sub AddEW(Cmb As ComboBox)
     With Cmb: .Clear: .AddItem "East": .AddItem "West": End With
End Sub
Private Sub AddNS(Cmb As ComboBox)
     With Cmb: .Clear: .AddItem "North": .AddItem "South": End With
End Sub

Public Function ShowDialog(aGeoPos As GeoPos, FOwner As Form) As VbMsgBoxResult
    Set m_FOwner = FOwner
    Set m_GeoPos = aGeoPos.Clone
    Set m_UTMGps = m_GeoPos.ToUTM32(MUTM.EllipsoWGS84)
    UpdateView
    Me.Show vbModal, m_FOwner
    ShowDialog = m_Result
    If m_Result = vbCancel Then Exit Function
    aGeoPos.NewC m_GeoPos
End Function


Private Sub OptWGS84_Click()
    ToggleView
End Sub
Private Sub OptUTM32_Click()
    ToggleView
End Sub
Private Sub ToggleView()
    If mnuOptConvert.Checked Then UpdateData
    m_ShowWGS84 = Not m_ShowWGS84
    If m_ShowWGS84 Then PnlWGS84.ZOrder 0 Else PnlUTM32.ZOrder 0
    If Not m_GeoPos Is Nothing Then
        UpdateView
    End If
End Sub

Private Sub UpdateView()
    If m_ShowWGS84 Then
        CmbNS.Text = m_GeoPos.Latitude.Dir
        CmbEW.Text = m_GeoPos.Longitude.Dir
        TxtLatitude.Text = m_GeoPos.Latitude.ToStr_DMS
        TxtLongitude.Text = m_GeoPos.Longitude.ToStr_DMS
        TxtNHN.Text = m_GeoPos.Height
        TxtDescription.Text = m_GeoPos.Name
    Else
        'If m_UTMGps Is Nothing Then
        '    Set m_UTMGps = m_GeoPos.ToUTM32(MUTM.EllipsoWGS84)
        'End If
        TxtNorthing.Text = m_UTMGps.Northing
        TxtEasting.Text = m_UTMGps.Easting
        TxtUTMZone.Text = m_UTMGps.Zone 'Str
        
        TxtNHNUTM32.Text = m_UTMGps.Height
        TxtDescriptionUTM32.Text = m_UTMGps.Name
    End If
End Sub

Private Sub UpdateData()
    If m_ShowWGS84 Then
        m_GeoPos.Latitude.Dir = Left(CmbNS.Text, 1)
        m_GeoPos.Longitude.Dir = Left(CmbEW.Text, 1)
        m_GeoPos.Name = TxtDescription.Text
        m_GeoPos.Latitude.Parse TxtLatitude.Text
        m_GeoPos.Longitude.Parse TxtLongitude.Text
        Dim H As Double
        If Not Double_TryParse(TxtNHN.Text, H) Then Exit Sub
        m_GeoPos.Height = H
        If mnuOptConvert.Checked Then
            'Dim u As UTM32: Set u = m_GeoPos.ToUTM32
            'm_UTMGps.NewC u
            Set m_UTMGps = m_GeoPos.ToUTM32(MUTM.EllipsoWGS84)
        End If
    Else
        'Dim n As Double
        'If Not CheckParse(TxtNorthing.Text, n) Then Exit Sub
        'Dim e As Double
        'If Not CheckParse(TxtEasting.Text, e) Then Exit Sub
        'Set m_UTMGps = MNew.UTM32(n, e, TxtUTMZone.Text)
        If mnuOptConvert.Checked Then
            Set m_GeoPos = m_UTMGps.ToWGS84(MUTM.EllipsoWGS84)
        End If
    End If
End Sub

Private Sub BtnOK_Click()
    UpdateData
    m_Result = vbOK:     Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel: Unload Me
End Sub

Private Sub BtnEditLatitude_Click()
    FAngle.Move Me.Left + BtnEditLatitude.Left, Me.Top + BtnEditLatitude.Top
    If FAngle.ShowDialog(m_GeoPos.Latitude, m_FOwner) Then
        UpdateView
    End If
End Sub

Private Sub BtnEditLongitude_Click()
    FAngle.Move Me.Left + BtnEditLongitude.Left, Me.Top + BtnEditLongitude.Top
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

Private Sub mnuOptConvert_Click()
    mnuOptConvert.Checked = Not mnuOptConvert.Checked
End Sub
Private Sub TxtLatitude_LostFocus()
    m_GeoPos.Latitude.Parse TxtLatitude.Text
    UpdateView
End Sub
Private Sub TxtLongitude_LostFocus()
    m_GeoPos.Longitude.Parse TxtLongitude.Text
    UpdateView
End Sub
Private Sub TxtNHN_LostFocus()
    Dim H As Double
    If Not Double_TryParse(TxtNHN.Text, H) Then Exit Sub
    m_GeoPos.Height = H
    UpdateView
End Sub
Private Sub TxtDescription_LostFocus()
    m_GeoPos.Name = TxtDescription.Text
    UpdateView
End Sub


Private Sub TxtNorthing_LostFocus()
    Dim d As Double
    If Not CheckParse(TxtNorthing.Text, d) Then Exit Sub
    m_UTMGps.Northing = d
    UpdateView
End Sub
Private Sub TxtEasting_LostFocus()
    Dim d As Double
    If Not CheckParse(TxtEasting.Text, d) Then Exit Sub
    m_UTMGps.Easting = d
    UpdateView
End Sub
Private Sub TxtUTMZone_LostFocus()
    m_UTMGps.Zone = TxtUTMZone.Text
    UpdateView
End Sub
Private Sub TxtNHNUTM32_LostFocus()
    Dim d As Double
    If Not CheckParse(TxtNHNUTM32.Text, d) Then Exit Sub
    m_UTMGps.Height = d
    UpdateView
End Sub
Private Sub TxtDescriptionUTM32_LostFocus()
    m_UTMGps.Name = TxtDescriptionUTM32.Text
End Sub

Private Function CheckParse(s As String, d_out As Double) As Boolean
    CheckParse = Double_TryParse(s, d_out)
    If Not CheckParse Then
        MsgBox "Could not parse the value: " & s & vbCrLf & "Please give a valid number"
    End If
End Function

