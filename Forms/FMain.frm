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
      MultiSelect     =   1  '1 -Einfach
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
      Begin VB.Menu mnuGeoPosSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGeoPosCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuGeoPosCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuGeoPosPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
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
      Begin VB.Menu mnuOptStartGEWeb 
         Caption         =   "Start Google-Earth-Web (not Pro)"
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
    Dim h As Single: h = Me.ScaleHeight - LBFamousPlaces.Top - TxtResults.Height
    If W > 0 And h > 0 Then LBFamousPlaces.Move L, T, W, h
    L = LBFamousPlaces.Left + LBFamousPlaces.Width
    If W > 0 And h > 0 Then LBTrip.Move L, T, W, h: LblTripResults.Left = L
    L = 0
    h = TxtResults.Height 'Me.ScaleHeight - T - LBFamousPlaces.Height
    T = Me.ScaleHeight - h
    W = Me.ScaleWidth
    If W > 0 And h > 0 Then TxtResults.Move L, T, W, h
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
        
        .Add MNew.GeoPos(MNew.AngleDecS("N 31° 19' 28,34''"), MNew.AngleDecS("O 120° 42' 43,40''"), 0, "China Suzhou Suzhou Supertower, H=450m, 98etag.  y(dFst)=2019 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 32°  3' 46,17''"), MNew.AngleDecS("O 118° 46' 42,83''"), 0, "China Nanjing Zifeng Tower, H=450m, 89etag.  y(dFst)=2010 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N  3°  9' 29,04''"), MNew.AngleDecS("O 101° 42' 43,41''"), 0, "Malaysien Kuala Lumpur Petronas Towers, H=452m, y(dFst)=2004 Zwillingstürme auf 172m durch Skybridge verbunden")
        .Add MNew.GeoPos(MNew.AngleDecS("N 28° 11' 42,84''"), MNew.AngleDecS("O 112° 58' 26,00''"), 0, "China Changsa Changsa IFS Tower, H=452m, 94etag. ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 10° 46' 17,45''"), MNew.AngleDecS("O 106° 42' 15,85''"), 0, "Vietnam Ho-Chi-Minh-Stadt (=Saigon) Landmark 81 (Bitexco financial Tower), H=461,2m, 81etag. ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 59° 59' 12,47''"), MNew.AngleDecS("O  30° 10' 38,99''"), 0, "Russland Sankt Petersburg Lakhta Center, H=462m, Höchstes Gebäude Europas, Hauptquartier Gazprom")
        .Add MNew.GeoPos(MNew.AngleDecS("N 40° 45' 59,36''"), MNew.AngleDecS("W  73° 58' 50,39''"), 0, "USA New York Central Park Tower, H=472m, 98etag.  y(dFst)=2020 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 22° 18' 12,56''"), MNew.AngleDecS("O 114°  9' 37,18''"), 0, "China Hongkong International Commerce Center, H=484m, 108etag.  y(dFst)=2010 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 31° 14'  7,62''"), MNew.AngleDecS("O 121° 30'  4,94''"), 0, "China Shanghai Shanghai World Financial Center, H=492m, 101etag.  y(dFst)=2008 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 25°  1' 59,94''"), MNew.AngleDecS("O 121° 33' 54,34''"), 0, "Taiwan Taipeh Taipei 101, H=508m, 101etag.  y(dFst)=2007 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 39° 54' 41,28''"), MNew.AngleDecS("O 116° 27' 38,81''"), 0, "China Peking CITIC Tower Zhongguo Zun, H=528m, 108etag.  y(dFst)=2018 steht immer noch leer")
        .Add MNew.GeoPos(MNew.AngleDecS("N 39°  7' 43,98''"), MNew.AngleDecS("O 117° 11' 49,06''"), 0, "China Tianjin Tianjin CTF Finance Centre, H=530m, 97etag.  y(dFst)=2019 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 23°  6' 29,84''"), MNew.AngleDecS("O 113° 19' 11,10''"), 0, "China Guangzhou Chow Tai Fook Centre, H=530m, 111etag.  y(dFst)=2016 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 40° 42' 45,88''"), MNew.AngleDecS("W  74°  0' 48,17''"), 0, "USA New York One World Trade Centre, H=541,3m, 105etag.  y(dFst)=2014 steht am Ort des ehem WTC das 2001 zerstört wurde")
        .Add MNew.GeoPos(MNew.AngleDecS("N 37° 30' 45,21''"), MNew.AngleDecS("O 127°  6'  9,13''"), 0, "Südkorea Seoul Lotte World Tower, H=555m, 123etag. ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 22° 32' 11,26''"), MNew.AngleDecS("O 114°  3'  1,80''"), 0, "China Shenzhen Ping an International Finance Center, H=599m, 115etag.  y(dFst)=2017 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 21° 25'  7,09''"), MNew.AngleDecS("O  39° 49' 34,27''"), 0, "Saudi Arabien Mekka Makkah Royal Clock Tower, H=601m, 120etag.  y(dFst)=2011 Hochausgruppe Abraj Al Bait Größte Turmuhr der Welt (sieht aus wie bingben)")
        .Add MNew.GeoPos(MNew.AngleDecS("N 31° 14'  7,22''"), MNew.AngleDecS("O 121° 30'  6,88''"), 0, "China Shanghai Shanghai Tower, H=632m, 132etag. ")
        .Add MNew.GeoPos(MNew.AngleDecS("N 30° 35'  9,51''"), MNew.AngleDecS("O 114° 19'  3,25''"), 0, "China Wuhan Wuhan Greenland Center, H=636m, 125etag.  y(dFst)=2019 ")
        .Add MNew.GeoPos(MNew.AngleDecS("N  3°  8' 29,78''"), MNew.AngleDecS("O 101° 42'  4,03''"), 45, "Malaysien Kuala Lumpur Merdeka 118, H=678,9m, 118etag.  y(dFst)=(vorauss. End) 2023 Turmspitze symbol die ausgestreckte Hand des Premierministers Merdeka=Unabhängigkeit")
        .Add MNew.GeoPos(MNew.AngleDecS("N 25° 11' 49,91''"), MNew.AngleDecS("O  55° 16' 27,76''"), 0, "Vereinigte Arabische Emirate Dubai Burj Khalifa, H=828m, 163etag.  y(dFst)=2010 ")
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
    LBFamousPlaces_Click
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
    
    Dim s As String: s = "45°; 10°; 40; My new place"
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    
    FGeoPos.Move Me.Left + (Me.Width - FGeoPos.Width) / 2, Me.Top + (Me.Height - FGeoPos.Height) / 2
    If FGeoPos.ShowDialog(gps, Me) = vbCancel Then Exit Sub
    
    m_FamousPlaces.Add gps
    UpdateView
End Sub

Private Sub mnuGeoPosCut_Click()
    Dim i As Long: i = LBFamousPlaces.ListIndex + 1
    If i = 0 Then Exit Sub
    'Dim s As String: s = LBFamousPlaces.List(i)
    Dim gps As GeoPos: Set gps = m_FamousPlaces.Item(i)  ' MNew.GeoPosS(s)
    Clipboard.SetText gps.ToStrClipBoard
    m_FamousPlaces.Remove i
    UpdateView
    'LBFamousPlaces.RemoveItem i
End Sub

Private Sub mnuGeoPosCopy_Click()
    Dim i As Long: i = LBFamousPlaces.ListIndex + 1
    If i = 0 Then Exit Sub
    'Dim s As String: s = LBFamousPlaces.List(i)
    Dim gps As GeoPos: Set gps = m_FamousPlaces.Item(i) ' MNew.GeoPosS(s)
    Clipboard.SetText gps.ToStrClipBoard
End Sub

Private Sub mnuGeoPosPaste_Click()
    Dim s As String: s = Clipboard.GetText
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    Dim i As Long: i = LBFamousPlaces.ListIndex + 1
    If i = 0 Then
        m_FamousPlaces.Add gps
    Else
        m_FamousPlaces.Add gps, , i
    End If
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

Private Function LBFamousPlaces_GetSelection() As Long()
    Dim i As Long, c As Long
    ReDim v(0 To LBFamousPlaces.SelCount - 1)
    For i = 0 To LBFamousPlaces.ListCount - 1
        If LBFamousPlaces.Selected(i) Then
            v(c) = i
            c = c + 1
        End If
    Next
End Function

Private Sub mnuGeoPosDelete_Click()
'    Dim i As Long: i = LBFamousPlaces.ListIndex
'    If i < 0 Then Exit Sub
'    m_FamousPlaces.Remove i + 1
'    LBFamousPlaces.RemoveItem i
    'remove the items in the collection m_FamousPlaces and in the ListBox LBFamousPlaces
    Dim selectedindices() As Long: v = LBFamousPlaces_GetSelection
    Dim i As Long
    For i = 0 To UBound(selectedindices)
        
    Next
    'UpdateView
End Sub

Private Sub mnuOptStartGEWeb_Click()
    mnuOptStartGEWeb.Checked = Not mnuOptStartGEWeb.Checked
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
    Dim cmd As String
    If Not mnuOptStartGEWeb.Checked Then
        If FileExists(pfnGE) Then
            If FileExists(m_pfnKml) Then Kill m_pfnKml
            If SaveFile(m_pfnKml, gps.ToStrKml) Then
                'maybe here edit the path to your Google Earth installation
                cmd = """" & pfnGE & """" & " " & """" & m_pfnKml & """"
                Shell cmd, vbNormalFocus
            Else
                MsgBox "Could not write kmlfile: " & vbCrLf & m_pfnKml
            End If
            Exit Sub
        End If
    End If
Try: On Error GoTo Catch
    'https://earth.google.com/web/@48.01091401,10.61795265,624.60371552a,100.07091926d,35y,0h,0t,0r
    cmd = """" & pfnFF & """" & " " & """" & MMain.GEWeb & gps.ToGEWeb & """"
    Shell cmd, vbNormalFocus
    Exit Sub
Catch:
    MsgBox "Could not start google earth, maybe googleearth.exe or firefox.exe not found"
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
    FTestAngle.Show vbModal, Me
End Sub

