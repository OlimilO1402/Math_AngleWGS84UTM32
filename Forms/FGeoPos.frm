VERSION 5.00
Begin VB.Form FGeoPos 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Dialog Geo Position"
   ClientHeight    =   3615
   ClientLeft      =   150
   ClientTop       =   495
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
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.OptionButton OptWGS84 
      Caption         =   "WGS84"
      Height          =   255
      Left            =   120
      TabIndex        =   0
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
      TabIndex        =   26
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
      TabIndex        =   27
      Top             =   3120
      Width           =   1455
   End
   Begin VB.PictureBox PnlWGS84 
      BorderStyle     =   0  'Kein
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   5295
      TabIndex        =   28
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
         TabIndex        =   5
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
         TabIndex        =   3
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
         TabIndex        =   14
         Top             =   1440
         Width           =   4215
      End
      Begin VB.TextBox TxtLatitude 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox TxtLongitude 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TxtNHN 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   225
         Left            =   0
         TabIndex        =   13
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Latitude:"
         Height          =   225
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Longitude:"
         Height          =   225
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Height above sea level:"
         Height          =   345
         Left            =   0
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "[m+NHN]"
         Height          =   225
         Left            =   4320
         TabIndex        =   12
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
      TabIndex        =   29
      Top             =   600
      Width           =   5295
      Begin VB.TextBox TxtUTMZone 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtNHNUTM32 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox TxtEasting 
         Alignment       =   2  'Zentriert
         Height          =   375
         Left            =   2160
         TabIndex        =   18
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
         TabIndex        =   25
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label LblUTMZone 
         AutoSize        =   -1  'True
         Caption         =   "UTM Zone:"
         Height          =   225
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "[m+NHN]"
         Height          =   225
         Left            =   4320
         TabIndex        =   23
         Top             =   960
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Height above sea level:"
         Height          =   345
         Left            =   0
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Easting:"
         Height          =   225
         Left            =   1320
         TabIndex        =   17
         Top             =   0
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Northing:"
         Height          =   225
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   225
         Left            =   0
         TabIndex        =   24
         Top             =   1440
         Width           =   945
      End
   End
   Begin VB.Menu mnuOption 
      Caption         =   "Option"
      Visible         =   0   'False
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
Private m_FOwner As Form
Private m_GeoPos As GeoPos
Private m_UTMGps As UTM32
Private m_Result As VbMsgBoxResult

Private m_isUpdatingView As Boolean
Private m_LastTB    As TextBox
Private mPropA      As Func1
Private mUpdateData As Func1

'the following problem arises:
'in TextBox_LostFocus we take the given value, we have to parse and validate it
'when user hits Return the Button ok is activated before the last LostFocus
'so the last value never gets taken
'the solution now is: we have to take all values again in OK-click
'but this is not very satisfying
'we could do better
'we have to store the last textbox in gotfocus
'and in ok-click we parse only the last value

'OK was haben wir
'
'auf der Seite WGS84 haben wir:
' * Latitude:
'   - ComboBox: muss ma nix weiter machen
'   - TextBox:  die Function GeoPos.Latitude.Parse mit Parameter TextBox.Text aufrufen        -> a
' * Longitude:
'   - ComboBox: muss ma nix weiter machen
'   - TextBox:  die Function GeoPos.Longitude.Parse mit Parameter TextBox.Text aufrufen       -> a
' * TextBox: die Höhe erst als Double parsen dann als PropLet übernehmen                      -> b
' * TextBox: die Beschreibung als PropLet übernehmen                                          -> c
'
'auf der Seite UTM32 haben wir:
' * TextBox: UTMZone als PropLet As String an UTM32.UTMZone übergeben                         -> c
' * TextBox: Easting  As Double parsen dann als PropLet As String an Prop Easting übergeben   -> b
' * TextBox: Northing As double parsen dann als PropLet As String an Prop Northing übergeben  -> b
' * TextBox: die Höhe erst als Double parsen dann als PropLet übernehmen                      -> b
' * TextBox: die Beschreibung als PropLet übernehmen                                          -> c

'-> wir haben drei verschiedene Arten Daten zu übernehmen
' a) String von TextBox an obj-Function übergeben z.b. Parse
' b) String von TextBox nach double parsen Fehler ausgeben und an PropLet übergeben
' c) String von TextBox direkt an PropLet übergeben

Private Sub Form_Load()
    AddEW Me.CmbEW
    AddNS Me.CmbNS
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
    If mnuOptConvert.Checked Then
        Convert
    End If
    PnlWGS84.ZOrder IIf(OptWGS84.Value, 0, 1)
    If Not m_GeoPos Is Nothing Then
        UpdateView
    End If
End Sub

Private Sub Convert()
    If OptWGS84.Value Then
        Set m_GeoPos = m_UTMGps.ToWGS84(MUTM.EllipsoWGS84)
    Else
        Set m_UTMGps = m_GeoPos.ToUTM32(MUTM.EllipsoWGS84)
    End If
End Sub

Private Sub UpdateView()
    m_isUpdatingView = True
    If OptWGS84.Value Then
        CmbNS.Text = m_GeoPos.Latitude.Dir
        CmbEW.Text = m_GeoPos.Longitude.Dir
        TxtLatitude.Text = m_GeoPos.Latitude.ToStr_DMS
        TxtLongitude.Text = m_GeoPos.Longitude.ToStr_DMS
        TxtNHN.Text = m_GeoPos.Height
        TxtDescription.Text = m_GeoPos.Name
    Else
        TxtUTMZone.Text = m_UTMGps.Zone
        TxtEasting.Text = m_UTMGps.Easting
        TxtNorthing.Text = m_UTMGps.Northing
        TxtNHNUTM32.Text = m_UTMGps.Height
        TxtDescriptionUTM32.Text = m_UTMGps.Name
    End If
    m_isUpdatingView = False
End Sub

Private Function UpdateData() As Boolean
    If m_isUpdatingView Then Exit Function
    Dim d
    If OptWGS84.Value Then
        
        m_GeoPos.Latitude.Dir = CmbNS.Text
        m_GeoPos.Longitude.Dir = CmbEW.Text
        
        UpdateData = AngleParse(m_GeoPos.Latitude, TxtLatitude.Text)
        If Not UpdateData Then Exit Function
        
        UpdateData = AngleParse(m_GeoPos.Longitude, TxtLongitude.Text)
        If Not UpdateData Then Exit Function
        
        UpdateData = FloatParse(TxtNHN.Text, d)
        If Not UpdateData Then Exit Function
        m_GeoPos.Height = d
        
        m_GeoPos.Name = TxtDescription.Text
    Else
        m_UTMGps.Zone = TxtUTMZone.Text
        
        UpdateData = FloatParse(TxtEasting.Text, d)
        If Not UpdateData Then Exit Function
        m_UTMGps.Easting = d
        
        UpdateData = FloatParse(TxtNorthing.Text, d)
        If Not UpdateData Then Exit Function
        m_UTMGps.Northing = d
        
        UpdateData = FloatParse(TxtNHNUTM32.Text, d)
        If Not UpdateData Then Exit Function
        m_UTMGps.Height = d
        
        m_UTMGps.Name = TxtDescriptionUTM32.Text
    End If
'    If mnuOptConvert.Checked Then
'        If OptWGS84.Value Then
'            Set m_UTMGps = m_GeoPos.ToUTM32(MUTM.EllipsoWGS84)
'        Else
'            Set m_GeoPos = m_UTMGps.ToWGS84(MUTM.EllipsoWGS84)
'        End If
'    End If
End Function

Private Function AngleParse(a As AngleDec, s As String) As Boolean
    AngleParse = a.Parse(s)
    If Not AngleParse Then MsgBox "Could not parse: " & vbCrLf & s
End Function

Private Function FloatParse(s As String, d_out) As Boolean
    FloatParse = MString.Decimal_TryParse(s, d_out)
    If Not FloatParse Then MsgBox "Could not parse convert the value to a float: " & s
End Function

Private Sub BtnOK_Click()
    If Not UpdateData Then
        UpdateView
        Exit Sub
    End If
    m_Result = vbOK:     Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_Result = vbCancel: Unload Me
End Sub

Private Sub mnuOptConvert_Click()
    mnuOptConvert.Checked = Not mnuOptConvert.Checked
End Sub


'WGS84
Private Sub CmbNS_Click()
    m_GeoPos.Latitude.Dir = Left(CmbNS.Text, 1)
    UpdateView
End Sub
Private Sub CmbEW_Click()
    m_GeoPos.Longitude.Dir = Left(CmbEW.Text, 1)
    UpdateView
End Sub

Private Sub TxtLatitude_GotFocus()
    Set m_LastTB = TxtLatitude:       Set mPropA = MNew.Func1(m_GeoPos.Latitude, "Parse")
End Sub
Private Sub TxtLatitude_LostFocus()
    Call AngleParse(m_GeoPos.Latitude, TxtLatitude.Text)
    UpdateView
End Sub
Private Sub BtnEditLatitude_Click()
    FAngle.Move Me.Left + BtnEditLatitude.Left, Me.Top + BtnEditLatitude.Top
    If FAngle.ShowDialog(m_GeoPos.Latitude, m_FOwner) Then
        UpdateView
    End If
End Sub

Private Sub TxtLongitude_GotFocus()
    Set m_LastTB = TxtLongitude:       Set mPropA = MNew.PropLet(m_GeoPos.Longitude, "Parse")
End Sub
Private Sub TxtLongitude_LostFocus()
    Call AngleParse(m_GeoPos.Longitude, TxtLongitude.Text)
    UpdateView
End Sub
Private Sub BtnEditLongitude_Click()
    FAngle.Move Me.Left + BtnEditLongitude.Left, Me.Top + BtnEditLongitude.Top
    If FAngle.ShowDialog(m_GeoPos.Longitude, m_FOwner) Then
        UpdateView
    End If
End Sub
Private Sub TxtNHN_LostFocus()
    Dim d As Double
    If Not CheckParse(TxtNHN.Text, d) Then Exit Sub
    m_GeoPos.Height = d
    UpdateView
End Sub
Private Sub TxtDescription_LostFocus()
    m_GeoPos.Name = TxtDescription.Text
    UpdateView
End Sub

'UTM32
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

Private Sub TxtDescriptionUTM32_GotFocus()
    Set m_LastTB = TxtDescriptionUTM32:   Set mPropA = MNew.PropLet(m_UTMGps, "Name")
End Sub
Private Sub TxtDescriptionUTM32_LostFocus()
    m_UTMGps.Name = TxtDescriptionUTM32.Text
    UpdateView
End Sub

Private Function CheckParse(s As String, d_out As Double) As Boolean
    If IsNumeric(s) Then
        CheckParse = Double_TryParse(s, d_out)
    End If
    If Not CheckParse Then
        MsgBox "Could not parse the value: " & s & vbCrLf & "Please give a valid number"
    End If
End Function


'Private Sub TxtAngleRad_GotFocus()
'    Set m_LastTB = TxtAngleRad:       Set m_PropA = MNew.PropLet(m_Angle, "Value")
'End Sub
'Private Sub TxtAngleRad_LostFocus()
'    TB_OnLostFocus
'End Sub

'Private Sub TB_OnLostFocus()
'    If m_isUpdatingView Then Exit Sub
'    Dim s As String: s = m_LastTB.Text
'    'Dim v
'    'If MString.Decimal_TryParse(s, v) Then
'        Call m_PropA.Invoke(s)
'    'Else
'    '    MsgBox "Failed to parse a numeric value from: " & s
'    '    Exit Sub
'    'End If
'    UpdateView
'End Sub


' a) String von TextBox an obj-Function übergeben z.b. Parse
' b) String von TextBox nach double parsen Fehler ausgeben und an PropLet übergeben
' c) String von TextBox direkt an PropLet übergeben

Public Function OnAngleParse() As Boolean
    If m_isUpdatingView Then Exit Function
    Dim s As String: s = m_LastTB.Text
    OnAngleParse = mPropA.Invoke(s)
    If Not OnAngleParse Then
        MsgBox "Could not parse angle: " & s
    End If
    'UpdateView
End Function

Public Function OnFloatParse() As Boolean
    If m_isUpdatingView Then Exit Function
    Dim s As String: s = m_LastTB.Text
    Dim v
    OnFloatParse = MString.Decimal_TryParse(s, v)
    If OnFloatParse Then
        Call mPropA.Invoke(s)
    Else
        MsgBox "Failed to parse the value: " & s
        Exit Function
    End If
    'UpdateView
End Function

Public Function OnStringPropLet() As Boolean
    If m_isUpdatingView Then Exit Function
    Dim s As String: s = m_LastTB.Text
    Call mPropA.Invoke(s)
    OnStringPropLet = True
    'UpdateView
End Function
