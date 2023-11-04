Attribute VB_Name = "MOpenLC"
Option Explicit
'Open Location Code aka Plus-Code
'https://de.wikipedia.org/wiki/Open_Location_Code
Private Enum EPlusCode
    olc_2 = 0     '50
    olc_3 ' = 1   '51
    olc_4 ' = 2   '52
    olc_5 ' = 3   '53
    olc_6 ' = 4   '54
    olc_7 ' = 5   '55
    olc_8 ' = 6   '56
    olc_9 ' = 7   '57
    
    'A            '65
    'B            '66
    olc_C ' = 8   '67
    'D            '68
    'E            '69
    olc_F ' = 9   '70
    olc_G ' = 10  '71
    olc_H ' = 11  '72
    'I            '73
    olc_J ' = 12  '74
    'K            '75
    'L            '76
    olc_M ' = 13  '77
    'N            '78
    'O            '79
    olc_P ' = 14  '80
    olc_Q ' = 15  '81
    olc_R ' = 16  '82
    'S            '83
    'T            '84
    'U            '85
    olc_V ' = 17  '86
    olc_W ' = 18  '87
    olc_X ' = 19  '88
End Enum

Public Function PlusCode_ToAngle(ByVal aPlusCode As String) As Angle
    Dim i As Long, n As Long: n = Len(aPlusCode)
    If n = 0 Then Exit Function
    aPlusCode = UCase(aPlusCode)
    'Dim a 'As Variant As Decimal
    'Dim ac1 As Integer, ac2 As Integer
    Dim lat, lon 'in degrees
    If Len(n) < 2 Then Exit Function
    i = i + 1: lat = OLCChar_ToNum(Mid(aPlusCode, i, 1)) * 20
    i = i + 1: lon = OLCChar_ToNum(Mid(aPlusCode, i, 1)) * 20
    If Len(n) < 4 Then Exit Function
    i = i + 1: lat = lat + OLCChar_ToNum(Mid(aPlusCode, i, 1)) * 20
    i = i + 1: lon = lon + OLCChar_ToNum(Mid(aPlusCode, i, 1)) * 20
    
    'For i = 1 To n
    '    ac = OLCChar_ToNum(Mid(aPlusCode, i, 1))
    'Next
End Function

Function OLCChar_ToNum(ByVal c As String) As Integer
    Dim ac As Integer: ac = Asc(c) - 50
    If 0 <= ac And ac <= 7 Then
        OLCChar_ToNum = ac: Exit Function
    End If
    Select Case ac
    Case 67: ac = 8
    Case 70: ac = 9
    Case 71: ac = 10
    Case 72: ac = 11
    Case 74: ac = 12
    Case 77: ac = 13
    Case 80: ac = 14
    Case 81: ac = 15
    Case 82: ac = 16
    Case 86: ac = 17
    Case 87: ac = 18
    Case 88: ac = 19
    End Select
    OLCChar_ToNum = ac
End Function

