VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   6615
   ClientLeft      =   3240
   ClientTop       =   3330
   ClientWidth     =   12750
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   6615
   ScaleWidth      =   12750
   Begin VB.ListBox LBTrip 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   6360
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   2
      Top             =   240
      Width           =   6375
   End
   Begin VB.TextBox TxtResults 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   3240
      Width           =   12735
   End
   Begin VB.ListBox LBFamousPlaces 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   0
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Geo-positions of your famous places:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   2955
   End
   Begin VB.Label LblTripResults 
      AutoSize        =   -1  'True
      Caption         =   "Trip Length:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6360
      TabIndex        =   3
      Top             =   0
      Width           =   945
   End
   Begin VB.Menu mnuPopGPS 
      Caption         =   "mnuPopGPS"
      Begin VB.Menu mnuGeoPosAdd 
         Caption         =   "Add New Geo Position"
      End
      Begin VB.Menu mnuGeoPosEdit 
         Caption         =   "Edit Geo Position"
      End
      Begin VB.Menu mnuStartKoUmre 
         Caption         =   "Start Koordinaten-Umrechner.de"
      End
      Begin VB.Menu mnuStartGEarth 
         Caption         =   "Show Position in Google Earth"
      End
      Begin VB.Menu mnuAddToTrip 
         Caption         =   "Add To Trip"
      End
      Begin VB.Menu mnuGeoPosDelete 
         Caption         =   "Delete Item"
      End
   End
   Begin VB.Menu mnuPopTrip 
      Caption         =   "mnuPopTrip"
      Begin VB.Menu mnuTripStartGEarth 
         Caption         =   "Show Trip in Google Earth"
      End
      Begin VB.Menu mnuTripShowRoute 
         Caption         =   "Show Route in Google Earth (from:to:)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTripMoveUp 
         Caption         =   "Move ^_up_^"
      End
      Begin VB.Menu mnuTripMoveDown 
         Caption         =   "Move v_down_v"
      End
      Begin VB.Menu mnuTripRemovePlace 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu mnuTripClear 
         Caption         =   "Clear Trip"
      End
   End
   Begin VB.Menu mnuPopOptions 
      Caption         =   "mnuPopOptions"
      Begin VB.Menu mnuOptFolder 
         Caption         =   "Write kml-file to Folder"
         Begin VB.Menu mnuOptFolderTemp 
            Caption         =   "Temp"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOptFolderDocs 
            Caption         =   "Documents"
         End
      End
      Begin VB.Menu mnuOptFolderOpen 
         Caption         =   "Open Folder: Temp"
      End
      Begin VB.Menu mnuOptShowAngles 
         Caption         =   "Show Angle-Dialog"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://www.koordinaten-umrechner.de/decimal/51.000000,10.000000?karte=OpenStreetMap&zoom=8
Private m_FamousPlaces As Collection 'Of GeoPos
Private m_Trip         As Collection 'Of GeoPos
Private m_pfnKml       As String 'pathfilename of kml-file
'https://earth.google.com/web

'Wishes:
' * do file operations with class PathFileName get the
'   chance to call the default program for html-files
' * check how to send kml-file to Google-Earth Web
'   https://www.youtube.com/watch?v=-wXcH5Uzsos
' * if desktop version Google-Earth-Pro exists
'        offer option for usigng old (pro) or new (web)
'   else use google earth web

Private Sub Form_Load()

    Me.Caption = "Angle, GeoPos(gps)WGS84,UTM32 v" & App.Major & "." & App.Minor & "." & App.Revision
    AddPlaces
    Set m_Trip = New Collection
    
    m_pfnKml = pathTemp & "\" & fnKml
    
    mnuPopGPS.Visible = False
    mnuPopTrip.Visible = False
    mnuPopOptions.Visible = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopOptions
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FileExists(m_pfnKml) Then Kill m_pfnKml
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = LBFamousPlaces.Top
    Dim W As Single: W = Me.ScaleWidth / 2
    Dim H As Single: H = Me.ScaleHeight - LBFamousPlaces.Top - TxtResults.Height
    If W > 0 And H > 0 Then LBFamousPlaces.Move L, T, W, H
    L = LBFamousPlaces.Left + LBFamousPlaces.Width
    If W > 0 And H > 0 Then LBTrip.Move L, T, W, H: LblTripResults.Left = L
    L = 0
    H = TxtResults.Height 'Me.ScaleHeight - T - LBFamousPlaces.Height
    T = Me.ScaleHeight - H
    W = Me.ScaleWidth
    If W > 0 And H > 0 Then TxtResults.Move L, T, W, H
End Sub

Private Sub AddPlaces()
    Set m_FamousPlaces = New Collection
    With m_FamousPlaces
        .Add MNew.GeoPos(MNew.AngleDecS("N 52° 30' 35,31"""), MNew.AngleDecS("O 13° 22' 32,41"""), 34, "Berlin, Potsdamer Platz")
        .Add MNew.GeoPos(MNew.AngleDecS("N 48° 51' 29,69"""), MNew.AngleDecS("O 02° 17' 40,38"""), 216, "Paris, Eiffelturm")
        .Add MNew.GeoPos(MNew.AngleDecS("N 51° 30' 02,48"""), MNew.AngleDecS("W 00° 07' 28,53"""), 85, "London, Big Ben")
        .Add MNew.GeoPos(MNew.AngleDecS("N 55° 45' 14,77"""), MNew.AngleDecS("O 37° 37' 13,46"""), 147, "Moskau, Roter Platz")
        .Add MNew.GeoPos(MNew.AngleDecS("N 48° 08' 13,91"""), MNew.AngleDecS("O 11° 34' 31,75"""), 515, "München, Marienplatz")
        .Add MNew.GeoPos(MNew.AngleDecS("N 40° 44' 54,39"""), MNew.AngleDecS("W 73° 59' 08,39"""), 35, "New York, Empire State Building")
        .Add MNew.GeoPos(MNew.AngleDecS("S 22° 57' 06,95"""), MNew.AngleDecS("W 43° 12' 37,66"""), 704, "Rio de Janeiro, Cristo Redentor")
        .Add MNew.GeoPos(MNew.AngleDecS("N 51° 03' 03,96"""), MNew.AngleDecS("O 05° 51' 58,81"""), 34, "Deutschland, westlichster Punkt")
        .Add MNew.GeoPos(MNew.AngleDecS("N 51° 16' 22,54"""), MNew.AngleDecS("O 15° 02' 30,91"""), 163, "Deutschland, östlichster Punkt")
        .Add MNew.GeoPos(MNew.AngleDecS("N 55° 03' 30,95"""), MNew.AngleDecS("O 08° 25' 03,98"""), 0, "Deutschland, nördlichster Punkt")
        .Add MNew.GeoPos(MNew.AngleDecS("N 47° 16' 12,40"""), MNew.AngleDecS("O 10° 10' 42,04"""), 0, "Deutschland, südlichster Punkt")
    End With
    UpdateView
End Sub

Sub UpdateView()
    With LBFamousPlaces
        .Clear
        Dim gps As GeoPos
        For Each gps In m_FamousPlaces
            .AddItem gps.ToStr
        Next
    End With
End Sub

Private Sub LBFamousPlaces_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then Exit Sub
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(LBFamousPlaces.Text)
    Dim utm As UTM32: Set utm = gps.ToUTM32(MUTM.Ellipsoids(22))   'the ellipsoid WGS-84
    TxtResults.Text = utm.ToStr & vbCrLf & gps.ToStr & vbCrLf & gps.ToStrKml & vbCrLf & gps.ToKoUmrLink
End Sub

Private Sub LBFamousPlaces_DblClick()
    mnuGeoPosEdit_Click
End Sub

Private Sub LBFamousPlaces_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopGPS
    End If
End Sub

Private Sub LBTrip_Click()
    If LBTrip.ListCount = 1 Then
        mnuTripMoveUp.Enabled = False
        mnuTripMoveDown.Enabled = False
        Exit Sub
    End If
    mnuTripMoveUp.Enabled = True
    mnuTripMoveDown.Enabled = True
    Select Case LBTrip.ListIndex
    Case 0:                    mnuTripMoveUp.Enabled = False
    Case LBTrip.ListCount - 1: mnuTripMoveDown.Enabled = False
    End Select
End Sub

Private Sub LBTrip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopTrip
    End If
End Sub

Private Sub LBTrip_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuAddToTrip_Click
End Sub

Sub UpdateTripLengthView()
    Dim tl As Double: tl = CalcTripLength
    Dim s As String
    If tl < 1000 Then
        s = Round(tl, 2) & " m"
    Else
        tl = tl / 1000
        s = Round(tl, 2) & " km"
    End If
    LblTripResults.Caption = "Trip Length: " & s
End Sub

Private Function CalcTripLength() As Double
    If m_Trip.Count < 2 Then Exit Function
    Dim gps1 As GeoPos: Set gps1 = m_Trip.Item(1)
    Dim gps2 As GeoPos: Set gps2 = m_Trip.Item(2)
    Dim tl As Double: tl = gps1.HaverSineDistanceTo(gps2)
    If m_Trip.Count < 3 Then
        CalcTripLength = tl
        Exit Function
    End If
    Set gps1 = gps2
    Dim i As Long
    For i = 3 To m_Trip.Count
        Set gps2 = m_Trip.Item(i)
        tl = tl + gps1.HaverSineDistanceTo(gps2)
        Set gps1 = gps2
    Next
    CalcTripLength = tl
End Function

Private Function GetTripForKml() As String
    If m_Trip.Count < 2 Then Exit Function
    Dim i As Long: i = 1
    Dim gps As GeoPos: Set gps = m_Trip.Item(1)
    Dim s   As String: s = gps.Coords_ToKml
    For i = 2 To m_Trip.Count
        Set gps = m_Trip.Item(i)
        s = s & " " & gps.Coords_ToKml
    Next
    GetTripForKml = s
End Function

Private Function GetRouteForGE() As String
    If m_Trip.Count < 2 Then Exit Function
    Dim i As Long: i = 1
    Dim gps As GeoPos
    
    Set gps = m_Trip.Item(1)
    GetRouteForGE = "from:" & Trim(Str(gps.Latitude.ToGrad)) & "," & Trim(Str(gps.Longitude.ToGrad))
    
    Set gps = m_Trip.Item(2)
    GetRouteForGE = GetRouteForGE & " to:" & Trim(Str(gps.Latitude.ToGrad)) & "," & Trim(Str(gps.Longitude.ToGrad))
End Function

Private Function GetGeoPos(s As String) As GeoPos
    Dim gps As GeoPos
    For Each gps In m_FamousPlaces
        If gps.ToStr = s Then
            Set GetGeoPos = gps
            Exit Function
        End If
    Next
End Function

' ----------~~~~~~~~~~==========########## '     Menu handler    ' ##########==========~~~~~~~~~~---------- '
' ----------~~~~~~~~~~==========########## '      mnuPopGPS      ' ##########==========~~~~~~~~~~---------- '

Private Sub mnuGeoPosAdd_Click()
    'Dim s As String: s = InputBox("Add a new Place, values separated by semicolon.", "Add New Place: Lat; Lon; Height; Name", "45°; 10°; 150; My new place")
    'If s = vbNullString Then Exit Sub 'Cancel
    
    Dim s As String: s = "45°; 10°; 150; My new place"
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    
    FGeoPos.Move Me.Left + (Me.Width - FGeoPos.Width) / 2, Me.Top + (Me.Height - FGeoPos.Height) / 2
    If FGeoPos.ShowDialog(gps, Me) = vbCancel Then Exit Sub
    
    m_FamousPlaces.Add gps
    UpdateView
End Sub

Private Sub mnuGeoPosEdit_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = GetGeoPos(s)
    If gps Is Nothing Then Exit Sub
    
    FGeoPos.Move Me.Left + (Me.Width - FGeoPos.Width) / 2, Me.Top + (Me.Height - FGeoPos.Height) / 2
    If FGeoPos.ShowDialog(gps, Me) = vbCancel Then Exit Sub
    
    'Dim ns As String: ns = InputBox("Edit Position, values separated by semicolon.", "Edit Position: Lat; Lon; Height; Name", s)
    'If ns = vbNullString Then Exit Sub 'Cancel
    'gps.Parse ns
    LBFamousPlaces.List(LBFamousPlaces.ListIndex) = gps.ToStr
End Sub

Private Sub mnuGeoPosDelete_Click()
    Dim i As Long: i = LBFamousPlaces.ListIndex
    If i < 0 Then Exit Sub
    m_FamousPlaces.Remove i + 1
    LBFamousPlaces.RemoveItem i
    'UpdateView
End Sub

Private Sub mnuStartKoUmre_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    Dim cmd As String: cmd = """" & pfnFF & """" & " " & """" & gps.ToKoUmrLink & """"
    Shell cmd, vbNormalFocus
End Sub

Private Sub mnuStartGEarth_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    If FileExists(m_pfnKml) Then Kill m_pfnKml
    If SaveFile(m_pfnKml, gps.ToStrKml) Then
        'maybe here edit the path to your Google Earth installation
        Dim cmd As String: cmd = """" & pfnGE & """" & " " & """" & m_pfnKml & """"
        Shell cmd, vbNormalFocus
    End If
End Sub

Private Sub mnuAddToTrip_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = GetGeoPos(s)
    If gps Is Nothing Then Exit Sub
    m_Trip.Add gps
    LBTrip.AddItem s
    UpdateTripLengthView
End Sub

' ----------~~~~~~~~~~==========########## '      mnuPopTrip      ' ##########==========~~~~~~~~~~---------- '

Private Sub mnuTripShowRoute_Click()
    If m_Trip.Count < 2 Then MsgBox "Minimum 2 Places in a route!": Exit Sub
    Dim ft As String: ft = GetRouteForGE 'FromTo
    If Len(ft) = 0 Then Exit Sub
    If FileExists(pfnGE) Then
        'OM: 2022-05-08
        'NOPE DOES NOT WORK; NOT WITH GOOGLE EARTH AND NOT WITH GOOGLE MAPS
        'THE LINKS MUST BE SOMEWHAT DIFFERENT, DONT KNOW HOW
        'Dim cmd As String: cmd = """" & pfnGE & """" & " " & """" & ft & """"
        Dim cmd As String: cmd = """" & pfnFF & """" & " " & """" & "https://www.google.com/maps/dir/" & ft
        Shell cmd, vbNormalFocus
        Debug.Print cmd
    Else
        'trying to load the kml-file to Google-Earth-Web
        MsgBox "Please install desktop-version Google Earth Pro"
    End If

End Sub

Private Sub mnuTripStartGEarth_Click()
    If m_Trip.Count < 2 Then MsgBox "Minimum 2 Places in a trip!": Exit Sub
    Dim st As String: st = GetTripForKml
    If Len(st) = 0 Then Exit Sub
    Dim s As String: s = ""
    s = s & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
            "<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">" & vbCrLf & _
            "<Document>" & vbCrLf & _
            "    <name>MyTrip</name>" & vbCrLf & _
            "    <StyleMap id=""inline"">" & vbCrLf & _
            "        <Pair>" & vbCrLf & _
            "            <key>normal</key>" & vbCrLf & _
            "            <styleUrl>#inline0</styleUrl>" & vbCrLf & _
            "        </Pair>" & vbCrLf & _
            "        <Pair>" & vbCrLf & _
            "            <key>highlight</key>" & vbCrLf & _
            "            <styleUrl>#inline1</styleUrl>" & vbCrLf & _
            "        </Pair>" & vbCrLf & _
            "    </StyleMap>" & vbCrLf & _
            "    <Style id=""inline0"">" & vbCrLf & _
            "        <LineStyle>" & vbCrLf & _
            "            <color>ff0000ff</color>" & vbCrLf & _
            "            <width>2</width>" & vbCrLf & _
            "        </LineStyle>" & vbCrLf & _
            "        <PolyStyle>" & vbCrLf & _
            "            <fill>0</fill>" & vbCrLf & _
            "        </PolyStyle>" & vbCrLf & _
            "    </Style>" & vbCrLf & _
            "    <Style id=""inline1"">" & vbCrLf & _
            "        <LineStyle>"
    s = s & "            <color>ff0000ff</color>" & vbCrLf & _
            "            <width>2</width>" & vbCrLf & _
            "        </LineStyle>" & vbCrLf & _
            "        <PolyStyle>" & vbCrLf & _
            "            <fill>0</fill>" & vbCrLf & _
            "        </PolyStyle>" & vbCrLf & _
            "    </Style>" & vbCrLf & _
            "    <Placemark>" & vbCrLf & _
            "        <name>Pfadmesswert</name>" & vbCrLf & _
            "        <styleUrl>#inline</styleUrl>" & vbCrLf & _
            "        <LineString>" & vbCrLf & _
            "            <tessellate>1</tessellate>" & vbCrLf & _
            "            <coordinates>"
    s = s & "                " & st & vbCrLf
    s = s & "            </coordinates>" & vbCrLf & _
            "        </LineString>" & vbCrLf & _
            "    </Placemark>" & vbCrLf & _
            "</Document>" & vbCrLf & _
            "</kml>"
    If FileExists(m_pfnKml) Then Kill m_pfnKml
    If SaveFile(m_pfnKml, s) Then
        'maybe here edit the path to your Google Earth installation
        If FileExists(pfnGE) Then
            Dim cmd As String: cmd = """" & pfnGE & """" & " " & """" & m_pfnKml & """"
            Shell cmd, vbNormalFocus
        Else
            'trying to load the kml-file to Google-Earth-Web
            MsgBox "Please install desktop-version Google Earth Pro"
        End If
    End If
End Sub

Private Sub mnuTripMoveUp_Click()
    Dim i As Long: i = LBTrip.ListIndex
    If i < 0 Then MsgBox "Select item first": Exit Sub
    'View aktualisieren
    Dim tmp As String
    tmp = LBTrip.List(i)
    LBTrip.List(i) = LBTrip.List(i - 1)
    LBTrip.List(i - 1) = tmp
    LBTrip.ListIndex = i - 1
    i = i + 1  'collection is 1-based
    Dim gps As GeoPos: Set gps = m_Trip.Item(i)
    m_Trip.Remove i
    m_Trip.Add gps, , i - 1
End Sub

Private Sub mnuTripMoveDown_Click()
    Dim i As Long: i = LBTrip.ListIndex
    If i < 0 Then MsgBox "Select item first": Exit Sub
    Dim tmp As String
    tmp = LBTrip.List(i)
    LBTrip.List(i) = LBTrip.List(i + 1)
    LBTrip.List(i + 1) = tmp
    LBTrip.ListIndex = i + 1
    i = i + 1 'collection is 1-based
    Dim gps As GeoPos: Set gps = m_Trip.Item(i)
    m_Trip.Remove i
    m_Trip.Add gps, , , i '- 1
End Sub

Private Sub mnuTripRemovePlace_Click()
    Dim i As Long: i = LBTrip.ListIndex
    If i < 0 Then MsgBox "Select item first": Exit Sub
    m_Trip.Remove i + 1
    LBTrip.RemoveItem i
    UpdateTripLengthView
End Sub

Private Sub mnuTripClear_Click()
    LBTrip.Clear
    Set m_Trip = New Collection
End Sub

' ----------~~~~~~~~~~==========########## '      mnuPopOpt      ' ##########==========~~~~~~~~~~---------- '
Private Sub mnuOptFolderDocs_Click()
    mnuOptFolderDocs.Checked = True
    mnuOptFolderTemp.Checked = False
    mnuOptFolderOpen.Caption = "Open Folder: Documents"
    m_pfnKml = pathDocs & "\" & fnKml
End Sub

Private Sub mnuOptFolderTemp_Click()
    mnuOptFolderDocs.Checked = False
    mnuOptFolderTemp.Checked = True
    mnuOptFolderOpen.Caption = "Open Folder: Temp"
    m_pfnKml = pathTemp & "\" & fnKml
End Sub

Private Sub mnuOptFolderOpen_Click()
    Dim cmd As String
    cmd = "explorer.exe " & IIf(mnuOptFolderTemp.Checked, pathTemp, pathDocs)
    Shell cmd, vbNormalFocus
End Sub

Private Sub mnuOptShowAngles_Click()
    FTestAngle.Show
End Sub

