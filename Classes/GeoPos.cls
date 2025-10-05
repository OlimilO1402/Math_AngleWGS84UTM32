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
    Set m_Lat = Latitude: Set m_Lon = Longitude:                 m_Height = Height:                 m_Name = aName
    'Breitengrade
    'Wendekreis bei 23,5° N/S in between is the tropic
    'latitudes have a range
    'from 0° at the equator to 90° north at the northpole and
    'from 0° at the equator to 90° south at the southpole
    CheckLat m_Lat, m_Lon
    
    'Längengrade oder Meridiane
    'longitude has a range
    'from 0° at the Greenwich-line to 179°59'59.99'' east at the international date line (Datumsgrenze) and
    'from 179°59'59.99'' west to 0° back to the Greenwich-line
    CheckLon m_Lon
    
End Sub
'copy to MNew:
'Public Function GeoPos(Latitude As AngleDec, Longitude As AngleDec, Optional ByVal Height As Double, Optional ByVal Name As String) As GeoPos
'    Set GeoPos = New GeoPos: GeoPos.New_ Latitude, Longitude, Height, Name
'End Function
'Public Function GeoPosS(s As String) As GeoPos
'    Set GeoPosS = New GeoPos: GeoPosS.Parse s
'End Function

Private Sub CheckLat(Latitude As AngleDec, Longitude As AngleDec)
    Dim latVal: latVal = Latitude.GradF
    Dim lonVal: lonVal = Longitude.GradF
    Dim latDir As String: latDir = IIf(latVal < 0, "S", "N")
    If latVal < 0 Then latVal = CDec(-1) * latVal
    If latVal > 360 Then latVal = latVal Mod 360
    If latVal < 90 Then
        'OK
    ElseIf latVal < 180 Then
        latVal = CDec(180) - latVal
        lonVal = lonVal + 179.999
    ElseIf latVal < 270 Then
        latVal = CDec(270) - latVal
        lonVal = lonVal + 179.999 '180
    ElseIf latVal < 360 Then
        latVal = CDec(360) - latVal
    End If
    
    Latitude.GradF = latVal
    Latitude.Dir = latDir
    Longitude.GradF = lonVal
End Sub
Private Sub CheckLon(Longitude As AngleDec)
    Dim lonVal: lonVal = Longitude.GradF
    Dim lonDir As String: lonDir = IIf(lonVal < 0, "W", "E")
    If lonVal < 0 Then lonVal = CDec(-1) * lonVal
    If lonVal > 360 Then lonVal = lonVal Mod 360
    If lonVal < 180 Then
        'OK
    ElseIf lonVal < 360 Then
        lonVal = CDec(360) - lonVal
    End If
    Longitude.GradF = lonVal
    Longitude.Dir = lonDir
End Sub

Friend Sub NewC(other As GeoPos)
    m_Name = other.Name:     Set m_Lat = other.Latitude.Clone
    m_Height = other.Height: Set m_Lon = other.Longitude.Clone
End Sub

Friend Function Clone() As GeoPos
    Set Clone = New GeoPos: Clone.NewC Me
End Function

Public Function IsEqual(other As GeoPos) As Boolean
    If Not Me.Latitude.IsEqual(other.Latitude) Then Exit Function
    If Not Me.Longitude.IsEqual(other.Longitude) Then Exit Function
    If Me.Name <> other.Name Then Exit Function
    If Me.Height <> other.Height Then Exit Function
    IsEqual = True
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

'MB-Baustatik
'[Geodat]
'  GB = 48.18885
'  GL = 10.98411
Public Function Parse(ByVal s As String) As Boolean
Try: On Error GoTo Catch
    s = Trim(s)
    If Len(s) = 0 Then Exit Function
    If IsMBBaustatikFormat(s) Then
        s = Trim(Mid(s, 9))
        Dim Lat As Double: Lat = Parse_MB_Value1(s)
        Dim Lon As Double: Lon = Parse_MB_Value2(s)
        Set m_Lat = MNew.AngleDecD(Lat)
        If m_Lat.Value < 0 Then m_Lat.Dir = "S" Else m_Lat.Dir = "N"
        Set m_Lon = MNew.AngleDecD(Lon)
        If m_Lon.Value < 0 Then m_Lon.Dir = "W" Else m_Lon.Dir = "E"
        m_Name = "MB-Baustatik-Geodat"
        Parse = True
        Exit Function
    End If
    'Else
    Dim sa() As String: sa = GetSA(s)
    Dim u As Long: u = UBound(sa)
    If u < 1 Then Exit Function
    Dim i As Long
    Dim tmp As New Angle
    If i <= u Then
        Parse = tmp.Parse(sa(i))
        If Not Parse Then Exit Function
        Set m_Lat = MNew.AngleDecS(sa(i))
        i = i + 1
        If m_Lat.Value < 0 Then m_Lat.Dir = "S" Else m_Lat.Dir = "N"
    End If
    If i <= u Then
        Parse = tmp.Parse(sa(i))
        If Not Parse Then Exit Function
        Set m_Lon = MNew.AngleDecS(sa(i))
        i = i + 1
        If m_Lon.Value < 0 Then m_Lon.Dir = "W" Else m_Lon.Dir = "E"
    End If
    If i <= u Then
        m_Height = CDbl(sa(i))
        i = i + 1
    End If
    If i <= u Then
        m_Name = sa(i)
        'i = i + 1
    End If
    Parse = True
Catch:
End Function

Private Function IsMBBaustatikFormat(ByVal s As String) As Boolean
    s = UCase(s)
    IsMBBaustatikFormat = InStr(s, "[GEODAT]") > 0
    If Not IsMBBaustatikFormat Then Exit Function
    IsMBBaustatikFormat = Len(s) > 20
End Function

Private Function Parse_MB_Value1(ByVal s As String) As Double
    'das erste istgleich-Zeichen
    Dim p1 As Long: p1 = InStr(1, s, "=")
    Dim z As String, ch As String, ds As Boolean
    If p1 > 0 Then
        s = Trim(Mid(s, p1 + 1))
        Dim i As Long
        For i = 1 To Len(s)
            ch = Mid(s, i, 1)
            Select Case Asc(ch)
            Case 44, 46:    If ds Then Exit For
                            z = z & ".": ds = True
            Case 48 To 57  ', . 0 - 9
                            z = z & ch
            Case Else: Exit For
            End Select
        Next
        If IsNumeric(z) Then Parse_MB_Value1 = Val(z)
        'das zweite istgleich-Zeichen
        'Dim p2 As Long: p2 = InStr(p1 + 2, s, "=")
    End If
End Function
Private Function Parse_MB_Value2(ByVal s As String) As Double
    'das erste istgleich-Zeichen
    Dim p1 As Long: p1 = InStr(1, s, "=")
    Dim z As String, ch As String, ds As Boolean
    If p1 > 0 Then
        'das zweite istgleich-Zeichen
        Dim p2 As Long: p2 = InStr(p1 + 2, s, "=")
        If p2 > 0 Then
            s = Trim(Mid(s, p2 + 1))
            Dim i As Long
            For i = 1 To Len(s)
                ch = Mid(s, i, 1)
                Select Case Asc(ch)
                Case 44, 46:    If ds Then Exit For
                                z = z & ".": ds = True
                Case 48 To 57  ', . 0 - 9
                                z = z & ch
                Case Else: Exit For
                End Select
            Next
            If IsNumeric(z) Then Parse_MB_Value2 = Val(z)
        End If
    End If
End Function

Private Function GetSA(s As String) As String()
    Dim sa1() As String: sa1 = Split(s, "; ")
    Dim u1 As Long: u1 = UBound(sa1)
    Dim sa2() As String: sa2 = Split(s, vbTab)
    Dim u2 As Long: u2 = UBound(sa2)
    If u2 > u1 Then
        GetSA = sa2
    Else
        sa2 = Split(s, " ")
        u2 = UBound(sa2)
        If u2 > u1 Then
            GetSA = sa2
        Else
            GetSA = sa1
        End If
    End If
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
    GetStr = Trim(str(v))
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

'Public Function ToKmlPlacemark(Optional ByVal aName As String) As String
'    Dim Coords As String: Coords = Coords_ToKml 'GetStr(m_Lon.ToGrad * m_Lon.Direction) & "," & GetStr(m_Lat.ToGrad * m_Lat.Direction) & "," & m_Height
'    If Len(aName) = 0 Then aName = m_Name
'    If Len(aName) = 0 Then aName = Coords
'    Dim s As String: s = ""
'    s = s & "        <Placemark>" & vbCrLf & _
'            "            <name>" & aName & "</name>" & vbCrLf & _
'            "            <Point>" & vbCrLf & _
'            "                <coordinates>" & Coords & "</coordinates>" & vbCrLf & _
'            "            </Point>" & vbCrLf & _
'            "        </Placemark>"
'    ToKmlPlacemark = s
'End Function

'https://www.koordinaten-umrechner.de/decimal/48.137394,11.575447?karte=OpenStreetMap&zoom=8
Public Function ToKoUmrLink() As String
    'here first Latitude then Longitude, the other way around for kml-file
    Dim Coords As String: Coords = GetStr(m_Lat.ToGrad) & "," & GetStr(m_Lon.ToGrad)
    ToKoUmrLink = "https://www.koordinaten-umrechner.de/decimal/" & Coords & "?karte=OpenStreetMap&zoom=8"
End Function

Public Function ToGEWeb() As String
    ToGEWeb = Trim(str(Me.Latitude.GradF)) & "," & Trim(str(Me.Longitude.GradF)) & "," & Trim(str(Me.Height) & "a") & "," & Trim(str(1000) & "d") & ",35y,0h,0t,0r"
End Function
