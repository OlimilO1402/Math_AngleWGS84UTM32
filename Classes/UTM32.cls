VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UTM32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Universal Transverse Mercator
'https://de.wikipedia.org/wiki/UTM-Koordinatensystem
'https://en.wikipedia.org/wiki/Universal_Transverse_Mercator_coordinate_system
'For example, the CN Tower is at 43°38'33.24"N 79°23'13.7"W, which is in UTM zone 17, and the grid position is 630084 m east, 4833438 m north.
Private m_Name As String
Private m_Northing As Double
Private m_Easting  As Double
Private m_ZoneNum  As Integer
Private m_ZoneStr  As String
Private m_Height As Double

Friend Sub New_(ByVal Northing As Double, ByVal Easting As Double, ByVal UTMZone As String)
    m_Northing = Northing: m_Easting = Easting: Me.Zone = UTMZone
End Sub

Friend Sub NewG(aGeoPos As GeoPos, elli As Ellipsoid)
    LLtoUTM elli.EquatorialRadius, elli.eccentricitySquared, aGeoPos.Latitude.ToGrad, aGeoPos.Longitude.ToGrad  '* aGeoPos.Longitude.Direction
    m_Name = aGeoPos.Name: m_Height = aGeoPos.Height
End Sub

Friend Sub NewC(other As UTM32)
    With other
        m_Name = .Name
        m_Northing = .Northing
        m_Easting = .Easting
        m_ZoneNum = .ZoneNum
        m_ZoneStr = .ZoneStr
    End With
End Sub

Friend Function Clone() As UTM32
    Set Clone = New UTM32: Clone.NewC Me
End Function


Public Property Let Northing(ByVal Value As Double)
    m_Northing = Value
End Property
Public Property Get Northing() As Double
    Northing = m_Northing
End Property


Public Property Let Easting(ByVal Value As Double)
    m_Easting = Value
End Property
Public Property Get Easting() As Double
    Easting = m_Easting
End Property


Public Property Get ZoneNum() As Integer
    ZoneNum = m_ZoneNum
End Property
Public Property Get ZoneStr() As String
    ZoneStr = m_ZoneStr
End Property

Public Property Let Zone(ByVal Value As String)
    If Len(Value) Then
        m_ZoneNum = StrToUL(Trim(Value), m_ZoneStr)
    Else
        m_ZoneStr = "C"
        m_ZoneNum = 1
    End If
End Property
Public Property Get Zone() As String
    Zone = CStr(m_ZoneNum) & m_ZoneStr 'm_UTMZone
End Property

Private Function StrToUL(ByVal Value As String, sRest_out As String) As Long
    Dim l As Long: l = Len(Value)
    If l = 0 Then
        Value = "01C"
        l = 3
    End If
    Dim State As Integer, i As Long, ch As String, ac As Integer
    Dim zs As String, zi As Integer ': zi = -1
    Do
        i = i + 1
        ch = Mid(Value, i, 1)
        ac = AscW(ch)
        Select Case State
        Case 0
             Select Case ac
             Case 48 To 57 '0 to 9
                'If zi < 0 Then zi = 1
                zi = zi * 10 + ac - 48
                'zi = zi * 10 '+ ac - 57
             Case Else
                'If zi < 0 Then State = 2 Else State = 1
                State = 1
                i = i - 1
             End Select
        Case 1
             Select Case ac
             Case 65 To 90:  zs = zs & ch
             Case 97 To 122: zs = zs & ChrW(ac - 32)
             Case Else: Exit Do
             End Select
        Case 2 'Error
            Exit Do
        End Select
    Loop While i < l
    StrToUL = zi
    sRest_out = zs
End Function

Public Property Let Height(ByVal Value As Double)
    m_Height = Value
End Property
Public Property Get Height() As Double
    Height = m_Height
End Property

Public Property Let Name(ByVal Value As String)
    m_Name = Value
End Property
Public Property Get Name() As String
    Name = m_Name
End Property

Friend Function ToStr() As String
    ToStr = Me.Zone & " E: " & m_Easting & " N: " & m_Northing
End Function

Friend Function ToWGS84(elli As Ellipsoid) As GeoPos
    Dim lat As Double, lon As Double
    'LLtoUTM elli.EquatorialRadius, elli.eccentricitySquared, lat, lon
    UTMtoLL elli.EquatorialRadius, elli.eccentricitySquared, m_Northing, m_Easting, Me.Zone, lat, lon
    'Set ToWGS84 = New GeoPos: ToWGS84.New_ MNew.AngleD(lat), MNew.AngleD(lon), m_Height, m_Name
    'Dim ALat As AngleDec: Set ALat = MNew.AngleDecD(lat)
    'Dim ALon As AngleDec: Set ALon = MNew.AngleDecD(lon)
    Set ToWGS84 = MNew.GeoPos(MNew.AngleDecD(lat), MNew.AngleDecD(lon), m_Height, m_Name)
End Function

Private Sub LLtoUTM(ByVal ERadius As Double, ByVal eccs As Double, ByVal LatGrad As Double, ByVal LonGrad As Double)
'//converts lat/long to UTM coords.  Equations from USGS Bulletin 1532
'//East Longitudes are positive, West longitudes are negative.
'//North latitudes are positive, South latitudes are negative
'//Lat and Long are in decimal degrees
'//Written by Chuck Gantz- chuck.gantz@globalstar.com
    Dim lat        As Double: lat = LatGrad
    Dim lon        As Double: lon = LonGrad
    Dim a          As Double: a = ERadius
    Dim eccSquared As Double: eccSquared = eccs
    Const k0       As Double = 0.9996
        
'//Make sure the longitude is between -180.00 .. 179.9
    Dim LonTmp As Double: LonTmp = (lon + 180) - Int((lon + 180) / 360) * 360 - 180
    Dim LatRad As Double: LatRad = lat * MMath.Pi / 180
    Dim LonRad As Double: LonRad = LonTmp * MMath.Pi / 180
    
    
    Dim ZoneNumber As Long: ZoneNumber = Int((LonTmp + 180) / 6) + 1
    
    If lat >= 56 And lat < 64 And LonTmp >= 3 And LonTmp < 12 Then
        ZoneNumber = 32
    End If
    
  'Special zones for Svalbard
    If 72# <= lat And lat < 84# Then
        If 0# <= LonTmp And LonTmp < 9# Then
            ZoneNumber = 31
        ElseIf 9# <= LonTmp And LonTmp < 21# Then
            ZoneNumber = 33
        ElseIf 21# <= LonTmp And LonTmp < 33# Then
            ZoneNumber = 35
        ElseIf 33# <= LonTmp And LonTmp < 42# Then
            ZoneNumber = 37
        End If
    End If
    Dim LonOrigin    As Double: LonOrigin = (ZoneNumber - 1) * 6 - 180 + 3 ' +3 puts origin in middle of zone
    Dim LonOriginRad As Double: LonOriginRad = LonOrigin * MMath.Pi / 180 'deg2rad;

    '//compute the UTM Zone from the latitude and longitude
    'sprintf(UTMZone, "%d%c", ZoneNumber, UTMLetterDesignator(Lat));
    'Me.Zone = ZoneNumber & UTMLetterDesignator(Lat)
    m_ZoneNum = ZoneNumber
    m_ZoneStr = UTMLetterDesignator(lat)
    'Debug.Print "UTMZone: " & m_UTMZone & " ZoneNumber: " & ZoneNumber & " UTMLetter: " & UTMLetterDesignator(Lat)
    
    Dim eccPrimeSquared As Double: eccPrimeSquared = (eccSquared) / (1 - eccSquared)

    Dim n  As Double: n = a / VBA.Math.Sqr(1 - eccSquared * Sin(LatRad) * Sin(LatRad))
    Dim T  As Double: T = VBA.Math.Tan(LatRad) * VBA.Math.Tan(LatRad)
    Dim c  As Double: c = eccPrimeSquared * VBA.Math.Cos(LatRad) * VBA.Math.Cos(LatRad)
    Dim AA As Double: AA = Cos(LatRad) * (LonRad - LonOriginRad)
    Dim M  As Double
    M = a * ((1 - eccSquared / 4 - 3 * eccSquared * eccSquared / 64 - 5 * eccSquared * eccSquared * eccSquared / 256) * LatRad _
                - (3 * eccSquared / 8 + 3 * eccSquared * eccSquared / 32 + 45 * eccSquared * eccSquared * eccSquared / 1024) * Sin(2 * LatRad) _
                                    + (15 * eccSquared * eccSquared / 256 + 45 * eccSquared * eccSquared * eccSquared / 1024) * Sin(4 * LatRad) _
                                    - (35 * eccSquared * eccSquared * eccSquared / 3072) * Sin(6 * LatRad))
    
    m_Easting = CDbl((k0 * n * (AA + (1 - T + c) * AA * AA * AA / 6 _
                    + (5 - 18 * T + T * T + 72 * c - 58 * eccPrimeSquared) * AA * AA * AA * AA * AA / 120) _
                    + 500000#))

    m_Northing = CDbl((k0 * (M + n * Tan(LatRad) * (AA * AA / 2 + (5 - T + 9 * c + 4 * c * c) * AA * AA * AA * AA / 24 _
                 + (61 - 58 * T + T * T + 600 * c - 330 * eccPrimeSquared) * AA * AA * AA * AA * AA * AA / 720))))
    If (lat < 0) Then m_Northing = m_Northing + 10000000#  '; //10000000 meter offset for southern hemisphere
    
End Sub

Private Sub UTMtoLL(ByVal ERadius As Double, ByVal eccs As Double, ByVal UTMNorthing As Double, ByVal UTMEasting As Double, ByVal UTMZone As String, Lat_out As Double, Lon_out As Double)
'//converts UTM coords to lat/long.  Equations from USGS Bulletin 1532
'//East Longitudes are positive, West longitudes are negative.
'//North latitudes are positive, South latitudes are negative
'//Lat and Long are in decimal degrees.
'//Written by Chuck Gantz- chuck.gantz@globalstar.com

    Const k0 As Double = 0.9996
    Dim a As Double: a = ERadius 'ellipsoid[ReferenceEllipsoid].EquatorialRadius;
    Dim eccSquared As Double: eccSquared = eccs 'ellipsoid[ReferenceEllipsoid].eccentricitySquared;
    
    Dim e1 As Double: e1 = (1 - VBA.Sqr(1 - eccSquared)) / (1 + VBA.Sqr(1 - eccSquared))
    Dim x  As Double:  x = UTMEasting - 500000#  '//remove 500,000 meter offset for longitude
    Dim y  As Double:  y = UTMNorthing
    
    Dim ZoneLetter As String
    Dim ZoneNumber As Long:  ZoneNumber = StrToUL(UTMZone, ZoneLetter)
    m_ZoneNum = ZoneNumber
    m_ZoneStr = ZoneLetter
    If Len(ZoneLetter) = 0 Then Exit Sub
    
    Dim NorthernHemisphere As Long '//1 for northern hemispher, 0 for southern
    If ((Asc(ZoneLetter) - Asc("N")) >= 0) Then
        NorthernHemisphere = 1 '//point is in northern hemisphere
    Else
        NorthernHemisphere = 0 '//point is in southern hemisphere
        y = y - 10000000#      '//remove 10,000,000 meter offset used for southern hemisphere
    End If
    Dim LonOrigin       As Double: LonOrigin = (ZoneNumber - 1) * 6 - 180 + 3 '//+3 puts origin in middle of zone
    Dim eccPrimeSquared As Double: eccPrimeSquared = (eccSquared) / (1 - eccSquared)
    
    Dim M  As Double:  M = y / k0
    Dim mu As Double: mu = M / (a * (1 - eccSquared / 4 - 3 * eccSquared * eccSquared / 64 - 5 * eccSquared * eccSquared * eccSquared / 256))
    
    Dim phi1Rad As Double: phi1Rad = mu + (3 * e1 / 2 - 27 * e1 * e1 * e1 / 32) * Sin(2 * mu) _
                                        + (21 * e1 * e1 / 16 - 55 * e1 * e1 * e1 * e1 / 32) * Sin(4 * mu) _
                                        + (151 * e1 * e1 * e1 / 96) * Sin(6 * mu)
    Dim phi1    As Double:    phi1 = phi1Rad * 180 / MMath.Pi 'rad2deg
    
    Dim N1      As Double:      N1 = a / VBA.Math.Sqr(1 - eccSquared * VBA.Math.Sin(phi1Rad) * VBA.Math.Sin(phi1Rad))
    Dim T1      As Double:      T1 = VBA.Math.Tan(phi1Rad) * VBA.Math.Tan(phi1Rad)
    Dim c1      As Double:      c1 = eccPrimeSquared * VBA.Math.Cos(phi1Rad) * VBA.Math.Cos(phi1Rad)
    Dim R1      As Double:      R1 = a * (1 - eccSquared) / ((1 - eccSquared * VBA.Math.Sin(phi1Rad) * VBA.Math.Sin(phi1Rad)) ^ 1.5)
    Dim d       As Double:       d = x / (N1 * k0)
    Dim lat     As Double
    lat = phi1Rad - (N1 * VBA.Math.Tan(phi1Rad) / R1) * (d * d / 2 - (5 + 3 * T1 + 10 * c1 - 4 * c1 * c1 - 9 * eccPrimeSquared) * d * d * d * d / 24 _
                    + (61 + 90 * T1 + 298 * c1 + 45 * T1 * T1 - 252 * eccPrimeSquared - 3 * c1 * c1) * d * d * d * d * d * d / 720)
    lat = lat * 180 / MMath.Pi 'rad2deg
    
    Dim lon As Double
    lon = (d - (1 + 2 * T1 + c1) * d * d * d / 6 + (5 - 2 * c1 + 28 * T1 - 3 * c1 * c1 + 8 * eccPrimeSquared + 24 * T1 * T1) _
                    * d * d * d * d * d / 120) / Cos(phi1Rad)
    lon = LonOrigin + lon * 180 / MMath.Pi 'rad2deg
    
    Lat_out = lat
    Lon_out = lon
End Sub
