Attribute VB_Name = "MUTM"
Option Explicit

Public Type Ellipsoid
    id                  As Long
    Name                As String
    EquatorialRadius    As Double
    eccentricitySquared As Double
End Type

'Public Type Ellipsoids
'    Arr() As Ellipsoid
'    Count As Long
'End Type

'Public Ellipsoids As Ellipsoids
Public Ellipsoids() As Ellipsoid

Public Sub Init()
    ReDim Ellipsoids(-1 To 22)
    Dim i As Long: i = -1
    Ellipsoids(i) = New_Ellipsoid(i, "Placeholder           ", 0, 0):                 i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Airy                  ", 6377563, 0.00667054):  i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Australian National   ", 6378160, 0.006694542): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Bessel 1841           ", 6377397, 0.006674372): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Bessel 1841 (Nambia)  ", 6377484, 0.006674372): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Clarke 1866           ", 6378206, 0.006768658): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Clarke 1880           ", 6378249, 0.006803511): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Everest               ", 6377276, 0.006637847): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Fischer 1960 (Mercury)", 6378166, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Fischer 1968          ", 6378150, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "GRS 1967              ", 6378160, 0.006694605): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "GRS 1980              ", 6378137, 0.00669438):  i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Helmert 1906          ", 6378200, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Hough                 ", 6378270, 0.00672267):  i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "International         ", 6378388, 0.00672267):  i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Krassovsky            ", 6378245, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Modified Airy         ", 6377340, 0.00667054):  i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Modified Everest      ", 6377304, 0.006637847): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "Modified Fischer 1960 ", 6378155, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "South American 1969   ", 6378160, 0.006694542): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "WGS 60                ", 6378165, 0.006693422): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "WGS 66                ", 6378145, 0.006694542): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "WGS-72                ", 6378135, 0.006694318): i = i + 1
    Ellipsoids(i) = New_Ellipsoid(i, "WGS-84                ", 6378137, 0.00669438)
    
End Sub
Public Function New_Ellipsoid(aID As Long, aName As String, aEquRadius As Double, aEccSqu As Double) As Ellipsoid
    With New_Ellipsoid
        .id = aID: .Name = Trim(aName)
        .EquatorialRadius = aEquRadius
        .eccentricitySquared = aEccSqu
    End With
End Function

Public Sub TestIt()
    Dim i As Long, c1 As String, c2 As String
    For i = -80 To 84 Step 1#
        c1 = UTMLetterDesignator(i)
        c2 = Lat2UTMLetter(i)
        If c1 <> c2 Then
            Debug.Print "c1: " & c1 & " <> " & "c2: " & c2 & " lat: " & i
        End If
    Next
End Sub
Public Function UTMLetterDesignator(ByVal aLatitude As Double) As String
'Asc("C") = 67
'Asc("X") = 88
    Dim s As String
    Dim Lat As Double: Lat = aLatitude
    Select Case True

    Case (-80 <= Lat) And (Lat < -72): s = "C"
    Case (-72 <= Lat) And (Lat < -64): s = "D"
    Case (-64 <= Lat) And (Lat < -56): s = "E"
    Case (-56 <= Lat) And (Lat < -48): s = "F"
    Case (-48 <= Lat) And (Lat < -40): s = "G"
    Case (-40 <= Lat) And (Lat < -32): s = "H"

    Case (-32 <= Lat) And (Lat < -24): s = "J"
    Case (-24 <= Lat) And (Lat < -16): s = "K"
    Case (-16 <= Lat) And (Lat < -8):  s = "L"
    Case (-8 <= Lat) And (Lat < 0):    s = "M"
    Case (0 <= Lat) And (Lat < 8):     s = "N"

    Case (8 <= Lat) And (Lat < 16):    s = "P"
    Case (16 <= Lat) And (Lat < 24):   s = "Q"
    Case (24 <= Lat) And (Lat < 32):   s = "R"
    Case (32 <= Lat) And (Lat < 40):   s = "S"
    Case (40 <= Lat) And (Lat < 48):   s = "T"
    Case (48 <= Lat) And (Lat < 56):   s = "U"
    Case (56 <= Lat) And (Lat < 64):   s = "V"
    Case (64 <= Lat) And (Lat < 72):   s = "W"
    Case (72 <= Lat) And (Lat <= 84):  s = "X"

    End Select
    UTMLetterDesignator = s
End Function

Public Function Lat2UTMLetter(ByVal aLatitude As Double) As String
    'does the same as function "UTMLetterDesignator" but simpler
    Dim ch As Integer: ch = 67
    If -32 <= aLatitude Then ch = ch + 1
    If 8 <= aLatitude Then ch = ch + 1
    If 80 <= aLatitude Then ch = ch - 1
    Lat2UTMLetter = ChrW(Int((aLatitude + 80) / 8) + ch)
End Function
'Public Sub AddEllipsoid(aElli As Ellipsoid)
'    With Ellipsoids
'        If .Count = 0 Then ReDim .Arr(-1 To 3)
'    End With
'End Sub
