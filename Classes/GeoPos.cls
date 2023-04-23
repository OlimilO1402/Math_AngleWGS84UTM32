VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeoPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Name   As String
Private m_Lat    As AngleDec
Private m_Lon    As AngleDec
Private m_Height As Double

Friend Sub New_(Latitude As AngleDec, Longitude As AngleDec, Optional ByVal Height As Double, Optional ByVal aName As String)
    Set m_Lat = Latitude: Set m_Lon = Longitude: m_Height = Height: m_Name = aName
End Sub

Friend Sub NewC(other As GeoPos)
    m_Name = other.Name:     Set m_Lat = other.Latitude.Clone
    m_Height = other.Height: Set m_Lon = other.Longitude.Clone
End Sub

Friend Function Clone() As GeoPos
    Set Clone = New GeoPos: Clone.NewC Me
End Function

Public Property Get Latitude() As AngleDec
    Set Latitude = m_Lat
End Property
Public Property Get Longitude() As AngleDec
    Set Longitude = m_Lon
End Property

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

Public Sub Parse(s As String)
    Dim sa() As String: sa = GetSA(s)
    Dim u As Long: u = UBound(sa)
    Dim i As Long
    If i <= u Then Set m_Lat = MNew.AngleDecS(sa(i)): i = i + 1
    If i <= u Then Set m_Lon = MNew.AngleDecS(sa(i)): i = i + 1
    If i <= u Then m_Height = CDbl(sa(i)): i = i + 1
    If i <= u Then m_Name = sa(i): i = i + 1
    If m_Lat.Value < 0 Then m_Lat.Dir = "S" Else m_Lat.Dir = "N"
    If m_Lon.Value < 0 Then m_Lon.Dir = "W" Else m_Lon.Dir = "E"
End Sub

Private Function GetSA(s As String) As String()
    Dim sa1() As String: sa1 = Split(s, "; ")
    Dim u1 As Long: u1 = UBound(sa1)
    Dim sa2() As String: sa2 = Split(s, vbTab)
    Dim u2 As Long: u2 = UBound(sa2)
    If u2 > u1 Then GetSA = sa2 Else GetSA = sa1
End Function

Public Function HaverSineDistanceTo(other As GeoPos) As Double
   HaverSineDistanceTo = HaverSineDistance(Me, other)
   'HaverSineDistanceTo = PythagorasDistance(Me, other)
End Function

'1 Flugmeile = 1 nm = 1852m

Private Function HaverSineDistance(p1 As GeoPos, p2 As GeoPos) As Double
    
    'http://rosettacode.org/wiki/Haversine_formula
    'https://www.movable-type.co.uk/scripts/latlong.html
    
    'const R = 6371000 '; // metres
    'const phi1 = lat1 * Math.PI/180 '; // f, ? in radians
    'const phi2 = lat2 * Math.PI/180 ';
    'const Delta_phi = (lat2-lat1) * Math.PI/180 ';
    'const Delta_lam = (lon2-lon1) * Math.PI/180 ';
    '
    'const a = Math.sin(Delta_phi / 2) * Math.sin(Delta_phi / 2) +
    '          Math.cos(phi_1) * Math.cos(phi_2) *
    '          Math.sin(Delta_lam / 2) * Math.sin(Delta_lam / 2) ';
    'const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)) ';
    '
    'const d = R * c; // in metres
    
    Dim dlat As Double: dlat = Abs(p1.Latitude.ToRad - p2.Latitude.ToRad)    'longitude and latitude are angles
    Dim dlon As Double: dlon = Abs(p1.Longitude.ToRad - p2.Longitude.ToRad)  'longitude and latitude are angles
    Dim sindlat2 As Double: sindlat2 = VBA.Math.Sin(dlat / 2)
    Dim sindlon2 As Double: sindlon2 = VBA.Math.Sin(dlon / 2)
    'haversin(theta) = sin^2(theta/2) = (1-cos(theta))/2
    Dim a  As Double: a = (sindlat2 * sindlat2) + p1.Latitude.Cosinus * p2.Latitude.Cosinus * (sindlon2 * sindlon2)
    'Dim sq As Double: sq = Math.Sqr(a)
    'Dim mn As Double: mn = minD(1#, Math.Sqr(a))
    Dim c  As Double: c = 2 * m_Lat.ArcusSinusF(minD(1#, Math.Sqr(a)))
    Dim r  As Double: 'R = 6371008.767 'earth radius in m
                      'R = 6378137#
                      'R = 6356752.3142
                      'R = 6367444.66 / 0.9996
                      r = 6369993
                      'R = 6373000
                      'R = 6369900 * 0.9996
    HaverSineDistance = r * c
End Function

'Private Function HaverSineDistance(p1 As GeoPos, p2 As GeoPos) As Double
'
'    'http://rosettacode.org/wiki/Haversine_formula
'    'https://www.movable-type.co.uk/scripts/latlong.html
'
'    'const R = 6371000 '; // metres
'    'const phi1 = lat1 * Math.PI/180 '; // f, ? in radians
'    'const phi2 = lat2 * Math.PI/180 ';
'    'const Delta_phi = (lat2-lat1) * Math.PI/180 ';
'    'const Delta_lam = (lon2-lon1) * Math.PI/180 ';
'    '
'    'const a = Math.sin(Delta_phi / 2) * Math.sin(Delta_phi / 2) +
'    '          Math.cos(phi_1) * Math.cos(phi_2) *
'    '          Math.sin(Delta_lam / 2) * Math.sin(Delta_lam / 2) ';
'    'const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)) ';
'    '
'    'const d = R * c; // in metres
'
'    Dim dlat As Double: dlat = Abs(p1.Latitude.ToRad - p2.Latitude.ToRad)   'longitude and latitude are angles
'    Dim dlon As Double: dlon = Abs(p1.Longitude.ToRad - p2.Longitude.ToRad) 'longitude and latitude are angles
'    Dim sindlat2 As Double: sindlat2 = VBA.Math.Sin(dlat / 2)
'    Dim sindlon2 As Double: sindlon2 = VBA.Math.Sin(dlon / 2)
'    'haversin(theta) = sin^2(theta/2) = (1-cos(theta))/2
'    Dim a  As Double: a = (sindlat2 * sindlat2) + p1.Latitude.Cosinus * p2.Latitude.Cosinus * (sindlon2 * sindlon2)
'    Dim sq As Double: sq = Math.Sqr(a)
'    Dim mn As Double: mn = minD(1#, sq)
'    Dim c  As Double: c = 2 * m_Lat.ArcusSinusF(mn)
'    Dim R  As Double: ' R = 6371008.767 'earth radius in m
'                      R = 6378137#
'                      'R = 6356752.3142
'                      'R = 6367444.66
'                      'R = 6373000
'    HaverSineDistance = R * c
'End Function

Private Function minD(V1 As Double, V2 As Double) As Double
    If V1 < V2 Then minD = V1 Else minD = V2
End Function

Friend Function ToUTM32(elli As Ellipsoid) As UTM32
    Set ToUTM32 = New UTM32: ToUTM32.NewG Me, elli
End Function

Friend Function ToStr() As String
    'Dim sfm As String: sfm = "0.000"
    'ToString = "{lat:" & Format(m_Lat.Value, sfm) & "; lon: " & Format(m_Lon.Value, sfm) & "}"
    'ToStr = "{lat:" & m_Lat.ToStr_GMS & "; lon: " & m_Lon.ToStr_GMS & "}"
    ToStr = m_Lat.ToStr_DMS & "; " & m_Lon.ToStr_DMS & "; " & m_Height & "; " & m_Name
End Function

Public Function ToStrClipBoard() As String
    ToStrClipBoard = m_Lat.ToStr_DMS & vbTab & m_Lon.ToStr_DMS & vbTab & m_Height & vbTab & m_Name
End Function

Private Function GetStr(ByVal v As Double) As String
    'Converts a Double to String by using the function Str for ensuring "." as a decimalseparator
    'we could also use CDbl and eventually replace comma (",") with period (".")
    GetStr = Trim(Str(v))
    Dim c As Integer: c = AscW(Left(GetStr, 1))
    Select Case c
    'Asc("0") = 48; Asc("9") = 57;
    Case 48 To 57: Exit Function
    End Select
    'Asc(".") = 46
    If c = 46 Then GetStr = "0" & GetStr: Exit Function
    'Asc("-") = 45
    If c = 45 Then
        c = AscW(Mid(GetStr, 2, 1))
        If c = 46 Then GetStr = "-0" & Mid(GetStr, 2)
    End If
End Function

Public Function Coords_ToKml() As String
    'for kml-file, here first Longitude then Latitude
    Coords_ToKml = GetStr(m_Lon.ToGrad) & "," & GetStr(m_Lat.ToGrad) & "," & m_Height
End Function

Public Function ToStrKml(Optional ByVal aName As String) As String
    Dim coords As String: coords = Coords_ToKml 'GetStr(m_Lon.ToGrad * m_Lon.Direction) & "," & GetStr(m_Lat.ToGrad * m_Lat.Direction) & "," & m_Height
    If Len(aName) = 0 Then aName = m_Name
    If Len(aName) = 0 Then aName = coords
    Dim s As String: s = ""
    s = s & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
            "<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">" & vbCrLf & _
            "<Placemark>" & vbCrLf & _
            "    <name>" & aName & "</name>" & vbCrLf & _
            "    <Point>" & vbCrLf & _
            "        <coordinates>" & coords & "</coordinates>" & vbCrLf & _
            "    </Point>" & vbCrLf & _
            "</Placemark>" & vbCrLf & _
            "</kml>"
    ToStrKml = s
End Function

'https://www.koordinaten-umrechner.de/decimal/48.137394,11.575447?karte=OpenStreetMap&zoom=8
Public Function ToKoUmrLink() As String
    'here first Latitude then Longitude, the other way around for kml-file
    Dim coords As String: coords = GetStr(m_Lat.ToGrad) & "," & GetStr(m_Lon.ToGrad)
    ToKoUmrLink = "https://www.koordinaten-umrechner.de/decimal/" & coords & "?karte=OpenStreetMap&zoom=8"
End Function

Public Function ToGEWeb() As String
    ToGEWeb = Trim(Str(Me.Latitude.GradF)) & "," & Trim(Str(Me.Longitude.GradF)) & "," & Trim(Str(Me.Height) & "a") & "," & Trim(Str(1000) & "d") & ", 35y, 0h, 0t, 0r"
End Function
