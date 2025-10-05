VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DocumentKml"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_PFNkml    As String
Private m_GPSList   As Collection 'Of GeoPos
Private m_hasPin    As Boolean
Private m_hasLine   As Boolean
Private m_LineColor As Long
Private Const cind  As Byte = 2
Private m_iis       As Byte
Private m_ind       As String

Private Sub Class_Initialize()
    Set m_GPSList = New Collection
End Sub

Friend Sub New_(aPFN As String, Optional ByVal HasPin As Boolean = True, Optional ByVal HasLine As Boolean = False, Optional ByVal LineColor As Long = &HFF)
    m_PFNkml = aPFN: m_hasLine = HasLine: m_LineColor = LineColor: m_hasPin = HasPin
End Sub

Public Function Add(aGPS As GeoPos) As GeoPos
    m_GPSList.Add aGPS
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get HasPin() As Boolean
    HasPin = m_hasPin
End Property
Public Property Let HasPin(ByVal Value As Boolean)
    m_hasPin = Value
End Property

Public Property Get HasLine() As Boolean
    HasLine = m_hasLine
End Property
Public Property Let HasLine(ByVal Value As Boolean)
    m_hasLine = Value
End Property

Public Property Get LineColor() As Long
    LineColor = m_LineColor
End Property
Public Property Let LineColor(ByVal Value As Long)
    m_LineColor = Value
End Property

'Indent Stack
Private Function PeekInd() As String
    PeekInd = m_ind
End Function
Private Function PushInd() As String
    m_iis = m_iis + 1
    m_ind = Space(cind * m_iis)
    PushInd = m_ind
End Function
Private Function PopInd() As String
    m_iis = m_iis - 1
    m_ind = Space(cind * m_iis)
    PopInd = m_ind
End Function

Private Function WriteKmlHeader() As String
    WriteKmlHeader = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                     "<kml xmlns=""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"" xmlns:kml=""http://www.opengis.net/kml/2.2"" xmlns:atom=""http://www.w3.org/2005/Atom"">"
End Function

Public Function WritePlacemark(gps As GeoPos) As String
    With gps
        Dim Coords As String: Coords = .Coords_ToKml
        Dim aName  As String:  aName = .Name
    End With
    If Len(aName) = 0 Then aName = Coords
    Dim s As String: s = PeekInd & "<Placemark>" & vbCrLf & _
                         PushInd & "<name>" & aName & "</name>" & vbCrLf & _
                         PeekInd & "<Point>" & vbCrLf & _
                         PushInd & "<coordinates>" & Coords & "</coordinates>" & vbCrLf & _
                          PopInd & "</Point>" & vbCrLf & _
                          PopInd & "</Placemark>"
    WritePlacemark = s
End Function

Public Function ToStr() As String
    Dim s As String: s = WriteKmlHeader & vbCrLf
    s = s & PushInd & "<Document>" & vbCrLf
        PushInd
        Dim gps As GeoPos
        For Each gps In m_GPSList
            s = s & WritePlacemark(gps) & vbCrLf
        Next
        PopInd
    s = s & PeekInd & "</Document>" & vbCrLf
    PopInd
    s = s & PeekInd & "</kml>" ' & vbCrLf
    ToStr = s
End Function

Public Function Save() As Boolean
Try: On Error GoTo Catch
'    If Not FileExists(pfnGE) Then
'        MsgBox "Path to Google Earth Pro not found" & vbCrLf & pfnGE
'        Exit Sub
'    End If
    If FileExists(m_PFNkml) Then
        Kill m_PFNkml
    End If
    Dim filecont As String: filecont = Me.ToStr
    If Not MMain.SaveFile(m_PFNkml, filecont) Then
        MsgBox "Could not write kmlfile: " & vbCrLf & m_PFNkml & vbCrLf & "----------" & vbCrLf & filecont
        Exit Function
    End If
    'maybe here edit the path to your Google Earth installation
'    Dim cmd As String: cmd = """" & pfnGE & """" & " " & """" & m_PFNkml & """"
'    Shell cmd, vbNormalFocus
    Save = True
    Exit Function
Catch:
    MsgBox "Errors during writing of kml file"
End Function

'Public Function MMain.SaveFile(PFN As String, FCont As String) As Boolean
'Try: On Error GoTo Catch
'    Dim FNr As Integer: FNr = FreeFile
'    Open PFN For Binary Access Write As FNr
'    Put FNr, , FCont
'    SaveFile = True
'    GoTo Finally
'Catch:
'    MsgBox "Error during writing file occored:" & vbCrLf & PFN
'Finally:
'    Close FNr
'End Function
