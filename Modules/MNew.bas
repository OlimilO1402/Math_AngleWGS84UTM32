Attribute VB_Name = "MNew"
Option Explicit

Public Function Angle(ByVal angleInRadians As Double) As Angle
    Set Angle = New Angle: Angle.New_ angleInRadians
End Function

Public Function AngleD(ByVal angleInDegrees As Double) As Angle
    Set AngleD = New Angle: AngleD.NewD_ angleInDegrees
End Function

Public Function AngleG(ByVal angleInGon As Double) As Angle
    Set AngleG = New Angle: AngleG.NewG_ angleInGon
End Function

Public Function AngleDMS(ByVal aDeg As Long, ByVal aMin As Double, ByVal aSec As Double) As Angle
    Set AngleDMS = New Angle: AngleDMS.NewDMS_ aDeg, aMin, aSec
End Function

Public Function AngleS(ByVal angleInDMS As String) As Angle
    Set AngleS = New Angle: AngleS.Parse angleInDMS
End Function

Public Function AngleDec(ByVal angleInRadians) As AngleDec
    Set AngleDec = New AngleDec: AngleDec.New_ angleInRadians
End Function

Public Function AngleDecD(ByVal angleInDegrees) As AngleDec
    Set AngleDecD = New AngleDec: AngleDecD.NewD_ angleInDegrees
End Function

Public Function AngleDecG(ByVal angleInGon As Double) As AngleDec
    Set AngleDecG = New AngleDec: AngleDecG.NewG_ angleInGon
End Function

Public Function AngleDecDMS(ByVal deg As Long, ByVal min As Long, ByVal sec As Double) As AngleDec
    Set AngleDecDMS = New AngleDec: AngleDecDMS.NewDMS_ deg, min, sec
End Function

Public Function AngleDecS(ByVal angleInDMS As String) As AngleDec
    Set AngleDecS = New AngleDec: AngleDecS.Parse angleInDMS
End Function


Public Function GeoPos(Latitude As AngleDec, Longitude As AngleDec, Optional ByVal Height As Double, Optional ByVal Name As String) As GeoPos
    Set GeoPos = New GeoPos: GeoPos.New_ Latitude, Longitude, Height, Name
End Function
Public Function GeoPosS(s As String) As GeoPos
    Set GeoPosS = New GeoPos: GeoPosS.Parse s
End Function

Public Function UTM32(ByVal Northing As Double, ByVal Easting As Double, ByVal UTMZone As String) As UTM32
    Set UTM32 = New UTM32: UTM32.New_ Northing, Easting, UTMZone
End Function
Public Function UTM32G(aGeoPos As GeoPos, elli As Ellipsoid) As UTM32
    Set UTM32G = New UTM32: UTM32G.NewG aGeoPos, elli
End Function
