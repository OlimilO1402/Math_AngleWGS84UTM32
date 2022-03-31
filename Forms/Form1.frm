VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   3240
   ClientTop       =   3030
   ClientWidth     =   12255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   12255
   Begin VB.ListBox LBTrip 
      Height          =   2010
      Left            =   6120
      TabIndex        =   2
      Top             =   240
      Width           =   6135
   End
   Begin VB.TextBox TxtResults 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   2280
      Width           =   12255
   End
   Begin VB.ListBox LBFamousPlaces 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Geo-Positions of Your Famous Places:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   2700
   End
   Begin VB.Label LblTripResults 
      AutoSize        =   -1  'True
      Caption         =   "Trip Length:"
      Height          =   195
      Left            =   6120
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "mnuPopup1"
      Visible         =   0   'False
      Begin VB.Menu mnuStartKoUmre 
         Caption         =   "Start Koordinaten-Umrechner.de"
      End
      Begin VB.Menu mnuStartGEarth 
         Caption         =   "Show Position in Google Earth"
      End
      Begin VB.Menu mnuAddGeoPos 
         Caption         =   "Add New Position"
      End
      Begin VB.Menu mnuEditGeoPos 
         Caption         =   "Edit Position"
      End
      Begin VB.Menu mnuAddToTrip 
         Caption         =   "Add To Trip"
      End
      Begin VB.Menu mnuOpenTempDir 
         Caption         =   "Open Temp-Folder"
      End
   End
   Begin VB.Menu mnuPopup2 
      Caption         =   "mnuPopup2"
      Visible         =   0   'False
      Begin VB.Menu mnuTripRemovePlace 
         Caption         =   "Remove From Trip"
      End
      Begin VB.Menu mnuTripStartGEarth 
         Caption         =   "Show Trip in Google Earth"
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
Private m_Trip         As Collection
Private m_pfn          As String

Private Sub Form_Load()
    Me.Caption = "Angle,WGS84,UTM32 v" & App.Major & "." & App.Minor & "." & App.Revision
    AddPlaces
    Set m_Trip = New Collection
    m_pfn = Environ("Temp") & "\" & "AngleWGS84UTM32GoogleEarth.kml"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If FileExists(m_pfn) Then Kill m_pfn
End Sub

Private Sub Form_Resize()
    Dim l As Single, T As Single: T = LBFamousPlaces.Top
    Dim W As Single: W = Me.ScaleWidth / 2
    Dim H As Single: H = Me.ScaleHeight - LBFamousPlaces.Top - TxtResults.Height
    If W > 0 And H > 0 Then LBFamousPlaces.Move l, T, W, H
    l = LBFamousPlaces.Left + LBFamousPlaces.Width
    If W > 0 And H > 0 Then LBTrip.Move l, T, W, H: LblTripResults.Left = l
    l = 0
    H = TxtResults.Height 'Me.ScaleHeight - T - LBFamousPlaces.Height
    T = Me.ScaleHeight - H
    W = Me.ScaleWidth
    If W > 0 And H > 0 Then TxtResults.Move l, T, W, H
End Sub

Private Sub AddPlaces()
    Set m_FamousPlaces = New Collection
    With m_FamousPlaces
        .Add MNew.GeoPos(MNew.AngleS("N 52° 30' 35,31"""), MNew.AngleS("O 13° 22' 32,41"""), 34, "Berlin, Potsdamer Platz")
        .Add MNew.GeoPos(MNew.AngleS("N 48° 51' 29,69"""), MNew.AngleS("O 02° 17' 40,38"""), 216, "Paris, Eiffelturm")
        .Add MNew.GeoPos(MNew.AngleS("N 51° 30' 02,48"""), MNew.AngleS("W 00° 07' 28,53"""), 85, "London, Big Ben")
        .Add MNew.GeoPos(MNew.AngleS("N 55° 45' 14,77"""), MNew.AngleS("O 37° 37' 13,46"""), 147, "Moskau, Roter Platz")
        .Add MNew.GeoPos(MNew.AngleS("N 48° 08' 13,91"""), MNew.AngleS("O 11° 34' 31,75"""), 515, "München, Marienplatz")
        .Add MNew.GeoPos(MNew.AngleS("N 40° 44' 54,39"""), MNew.AngleS("W 73° 59' 08,39"""), 35, "New York, Empire State Building")
        .Add MNew.GeoPos(MNew.AngleS("S 22° 57' 06,95"""), MNew.AngleS("W 43° 12' 37,66"""), 704, "Rio de Janeiro, Cristo Redentor")
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

Private Sub LBFamousPlaces_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopup1
    End If
End Sub

Private Sub LBTrip_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu mnuPopup2
    End If
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

Private Function GetGeoPos(s As String) As GeoPos
    Dim gps As GeoPos
    For Each gps In m_FamousPlaces
        If gps.ToStr = s Then
            Set GetGeoPos = gps
            Exit Function
        End If
    Next
End Function
' ----------~~~~~~~~~~==========########## '     Menu handler     ' ##########==========~~~~~~~~~~---------- '
Private Sub mnuStartKoUmre_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    'maybe here edit the path to your preferred internet browser
    Dim cmd As String: cmd = """" & "C:\Program Files\Mozilla Firefox\firefox.exe" & """" & " " & """" & gps.ToKoUmrLink & """"
    Shell cmd, vbNormalFocus
End Sub

Private Sub mnuStartGEarth_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    If FileExists(m_pfn) Then Kill m_pfn
    If SaveFile(m_pfn, gps.ToStrKml) Then
        'maybe here edit the path to your Google Earth installation
        Dim cmd As String: cmd = """" & "C:\Program Files\Google\Google Earth Pro\client\googleearth.exe" & """" & " " & """" & m_pfn & """"
        Shell cmd, vbNormalFocus
    End If
End Sub

Private Sub mnuAddGeoPos_Click()
    Dim s As String: s = InputBox("Add a new Place, values separated by semicolon.", "Add New Place: Lat; Lon; Height; Name", "45°; 10°; 150; My new place")
    If s = vbNullString Then Exit Sub 'Cancel
    Dim gps As GeoPos: Set gps = MNew.GeoPosS(s)
    m_FamousPlaces.Add gps
    UpdateView
End Sub

Private Sub mnuEditGeoPos_Click()
    Dim s As String: s = LBFamousPlaces.Text
    If Len(s) = 0 Then MsgBox "Select item first": Exit Sub
    Dim gps As GeoPos: Set gps = GetGeoPos(s)
    If gps Is Nothing Then Exit Sub
    Dim ns As String: ns = InputBox("Edit Position, values separated by semicolon.", "Edit Position: Lat; Lon; Height; Name", s)
    If ns = vbNullString Then Exit Sub 'Cancel
    gps.Parse ns
    LBFamousPlaces.List(LBFamousPlaces.ListIndex) = gps.ToStr
    'UpdateView
    'UpdateTripLengthView
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

Private Sub mnuOpenTempDir_Click()
    Dim cmd As String: cmd = "explorer.exe " & Environ("Temp") & "\"
    Shell cmd, vbNormalFocus
End Sub

Private Sub mnuTripRemovePlace_Click()
    Dim i As Long: i = LBTrip.ListIndex
    If i < 0 Then MsgBox "Select item first": Exit Sub
    m_Trip.Remove i + 1
    LBTrip.RemoveItem i
    UpdateTripLengthView
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
    If FileExists(m_pfn) Then Kill m_pfn
    If SaveFile(m_pfn, s) Then
        'maybe here edit the path to your Google Earth installation
        Dim cmd As String: cmd = """" & "C:\Program Files\Google\Google Earth Pro\client\googleearth.exe" & """" & " " & """" & m_pfn & """"
        Shell cmd, vbNormalFocus
    End If
End Sub
