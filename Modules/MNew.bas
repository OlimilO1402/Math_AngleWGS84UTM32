Attribute VB_Name = "MNew"
Option Explicit

Public Function Angle(ByVal angleInRadians As Double) As Angle
    Set Angle = New Angle: Angle.New_ angleInRadians
End Function
Public Function AngleG(ByVal angleInGrad As Double) As Angle
    Set AngleG = New Angle: AngleG.NewG_ angleInGrad
End Function
Public Function AngleS(ByVal angleInGMS As String) As Angle
    Set AngleS = New Angle: AngleS.Parse angleInGMS
End Function

Public Function GeoPos(Latitude As Angle, Longitude As Angle, Optional ByVal Height As Double, Optional ByVal Name As String) As GeoPos
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
