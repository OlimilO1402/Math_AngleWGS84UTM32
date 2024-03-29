Attribute VB_Name = "MMain"
Option Explicit
Public pathTemp  As String 'path to tmp directory
Public pathDocs  As String 'path to documents directory
Public pathProgs As String 'path to program files directory
Public pfnGE     As String 'pathfilename to google earth pro (or link to google earth web)
Public GEWeb     As String 'base link to google earth web
Public pfnFF     As String 'pathfilename to firefox
Public fnKml     As String 'default filename of kml file

Sub Main()
    
    pathTemp = Environ("Temp")
    pathDocs = Environ("Homedrive") & Environ("Homepath") & "\Documents\"
    pathProgs = Environ("ProgramW6432")
    
    'maybe here edit the path to your default or preferred internet browser
    pfnFF = pathProgs & "\Mozilla Firefox\firefox.exe"
    pfnGE = pathProgs & "\Google\Google Earth Pro\client\googleearth.exe"
    GEWeb = "https://earth.google.com/web/@" 'LatitudeInDegree, LongitudeInDegree, height,
    fnKml = "AngleWGS84UTM32GoogleEarth.kml"
    
    MMath.Init
    MUTM.Init
    FMain.Show
    'FTestAngle.Show
End Sub

Public Function GetStr(ByVal v As Double) As String
    'Converts a Double to String by using the function Str for ensuring "." as a decimalseparator
    'we could also use cdbl and eventually replace a comma (",") with a period (".")
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

Public Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(FileName) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Public Function SaveFile(PFN As String, FCont As String) As Boolean
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , FCont
    SaveFile = True
    GoTo Finally
Catch:
    MsgBox "Error during writing file occored:" & vbCrLf & PFN
Finally:
    Close FNr
End Function

Public Sub test()
    Dim Pi: Pi = 4 * Atn(1)
    Dim hilfswert As String
    Dim lat1 As Double
    Dim lon1 As Double
    Dim lat2 As Double
    Dim lon2 As Double
    
    'lat1 = 33.942501: lon1 = -118.407997
    'lat2 = 40.639801: lon2 = -73.7789
    
    'lat1 = 40.748442: lon1 = -73.985664:
    'lat2 = 55.754103: lon2 = 37.620406
    
    Dim grad_lat1     As Long:     grad_lat1 = MMath.Floor(lat1)
    Dim minuten_lat1  As Long:  minuten_lat1 = MMath.Floor((lat1 - grad_lat1) * 60)
    Dim sekunden_lat1 As Long: sekunden_lat1 = MMath.Floor((lat1 - grad_lat1) * 60)
    
    Dim grad_lon1     As Long:     grad_lon1 = MMath.Floor(lon1)
    Dim minuten_lon1  As Long:  minuten_lon1 = MMath.Floor((lon1 - grad_lon1) * 60)
    Dim sekunden_lon1 As Long: sekunden_lon1 = MMath.Floor((lon1 - grad_lon1) * 60)
    
    Dim grad_lat2     As Long:     grad_lat2 = MMath.Floor(lat2)
    Dim minuten_lat2  As Long:  minuten_lat2 = MMath.Floor((lat2 - grad_lat2) * 60)
    Dim sekunden_lat2 As Long: sekunden_lat2 = MMath.Floor((lat2 - grad_lat2) * 60)
    
    Dim grad_lon2     As Long:     grad_lon2 = MMath.Floor(lon2)
    Dim minuten_lon2  As Long:  minuten_lon2 = MMath.Floor((lon2 - grad_lon2) * 60)
    Dim sekunden_lon2 As Long: sekunden_lon2 = MMath.Floor((lon2 - grad_lon2) * 60)
    
    hilfswert = minuten_lat1 & "." & sekunden_lat1
    lat1 = grad_lat1 + CDbl(Val(hilfswert)) / 60
    lat1 = lat1 * Pi / 180
    
    hilfswert = minuten_lon1 & "." & sekunden_lon1
    lon1 = grad_lon1 + CDbl(Val(hilfswert)) / 60
    lon1 = lon1 * Pi / 180
    
    hilfswert = minuten_lat2 & "." & sekunden_lat2
    lat2 = grad_lat2 + CDbl(Val(hilfswert)) / 60
    lat2 = lat2 * Pi / 180
    
    hilfswert = minuten_lon2 & "." & sekunden_lon2
    lon2 = grad_lon2 + CDbl(Val(hilfswert)) / 60
    lon2 = lon2 * Pi / 180
    
    Dim test As Double: test = 2 * ArcusSinus(VBA.Math.Sqr((Sin((lat1 - lat2) / 2)) ^ 2 + Cos(lat1) * Cos(lat2) * (Sin((lon1 - lon2) / 2)) ^ 2))
    test = (test * 180 * 60) / Pi
    
    MsgBox MMath.Floor(test * 1.852)
    
End Sub

Public Function ArcusCosinus(ByVal x As Double) As Double
    ArcusCosinus = 0.5 * Pi - ArcusSinus(x)
End Function

Public Function ArcusSinus(ByVal y As Double) As Double
    Select Case y
        Case 1
            ArcusSinus = 0.5 * Pi
        Case -1
            ArcusSinus = -0.5 * Pi
        Case Else
            ArcusSinus = VBA.Math.Atn(y / Sqr(1 - y * y))
    End Select
End Function


