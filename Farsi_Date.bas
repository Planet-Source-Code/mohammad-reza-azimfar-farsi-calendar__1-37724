Attribute VB_Name = "Farsi_Date"

Public Function Fa_Day(En_Date As String) As String
Select Case Weekday(En_Date)
Case 1
    Fa_Day = "Ìﬂ‘‰»Â"
Case 2
    Fa_Day = "œÊ‘‰»Â"
Case 3
    Fa_Day = "”Â ‘‰»Â"
Case 4
    Fa_Day = "çÂ«—‘‰»Â"
Case 5
    Fa_Day = "Å‰Ã‘‰»Â"
Case 6
    Fa_Day = "Ã„⁄Â"
Case 7
    Fa_Day = "‘‰»Â"
End Select
End Function

Public Function Fa_Date(En_Date As String) As String
Dim The_Select As Integer
Dim The_Year As Integer
Dim The_Month As Integer
Dim The_Day As Integer
The_Year = CInt(Mid(En_Date, 7, 4))
The_Month = CInt(Mid(En_Date, 1, 2))
The_Day = CInt(Mid(En_Date, 4, 2))

If (The_Year Mod 4 = 0) Then
    The_Select = 1
Else
    The_Select = 2
End If

If ((The_Year - 1) Mod 4 = 0) Then
    The_Select = 3
End If

If The_Select = 1 Then
'------------------------------------------------------
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 10
    The_Month = 10
    The_Year = The_Year - 622
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 11
    The_Year = The_Year - 622
    End Select
Case 2: Select Case The_Day
    Case 1 To 19: The_Day = The_Day + 11
    The_Month = 11
    The_Year = The_Year - 622
    Case 20 To 29: The_Day = The_Day - 19
    The_Month = 12
    The_Year = The_Year - 622
    End Select
Case 3: Select Case The_Day
    Case 1 To 19: The_Day = The_Day + 10
    The_Month = 12
    The_Year = The_Year - 622
    Case 20 To 31: The_Day = The_Day - 19
    The_Month = 1
    The_Year = The_Year - 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 19: The_Day = The_Day + 12
    The_Month = 1
    The_Year = The_Year - 621
    Case 20 To 30: The_Day = The_Day - 19
    The_Month = 2
    The_Year = The_Year - 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 11
    The_Month = 2
    The_Year = The_Year - 621
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 3
    The_Year = The_Year - 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 11
    The_Month = 3
    The_Year = The_Year - 621
    Case 21 To 30: The_Day = The_Day - 20
    The_Month = 4
    The_Year = The_Year - 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 4
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 5
    The_Year = The_Year - 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 5
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 6
    The_Year = The_Year - 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 6
    The_Year = The_Year - 621
    Case 22 To 30: The_Day = The_Day - 21
    The_Month = 7
    The_Year = The_Year - 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 9
    The_Month = 7
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 8
    The_Year = The_Year - 621
    End Select
Case 11: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 10
    The_Month = 8
    The_Year = The_Year - 621
    Case 21 To 30: The_Day = The_Day - 20
    The_Month = 9
    The_Year = The_Year - 621
    End Select
Case 12: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 10
    The_Month = 9
    The_Year = The_Year - 621
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 10
    The_Year = The_Year - 621
    End Select
End Select
'------------------------------------------------------
End If

If The_Select = 2 Then
'------------------------------------------------------
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 10
    The_Month = 10
    The_Year = The_Year - 622
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 11
    The_Year = The_Year - 622
    End Select
Case 2: Select Case The_Day
    Case 1 To 19: The_Day = The_Day + 11
    The_Month = 11
    The_Year = The_Year - 622
    Case 19 To 28: The_Day = The_Day - 19
    The_Month = 12
    The_Year = The_Year - 622
    End Select
Case 3: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 9
    The_Month = 12
    The_Year = The_Year - 622
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 1
    The_Year = The_Year - 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 11
    The_Month = 1
    The_Year = The_Year - 621
    Case 21 To 30: The_Day = The_Day - 20
    The_Month = 2
    The_Year = The_Year - 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 2
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 3
    The_Year = The_Year - 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 3
    The_Year = The_Year - 621
    Case 22 To 30: The_Day = The_Day - 21
    The_Month = 4
    The_Year = The_Year - 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 4
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 5
    The_Year = The_Year - 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 5
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 6
    The_Year = The_Year - 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 6
    The_Year = The_Year - 621
    Case 23 To 30: The_Day = The_Day - 22
    The_Month = 7
    The_Year = The_Year - 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 8
    The_Month = 7
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 8
    The_Year = The_Year - 621
    End Select
Case 11: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 9
    The_Month = 8
    The_Year = The_Year - 621
    Case 22 To 30: The_Day = The_Day - 21
    The_Month = 9
    The_Year = The_Year - 621
    End Select
Case 12: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 9
    The_Month = 9
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 10
    The_Year = The_Year - 621
    End Select
End Select
'------------------------------------------------------
End If

If The_Select = 3 Then
'------------------------------------------------------
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 19: The_Day = The_Day + 11
    The_Month = 10
    The_Year = The_Year - 622
    Case 20 To 31: The_Day = The_Day - 19
    The_Month = 11
    The_Year = The_Year - 622
    End Select
Case 2: Select Case The_Day
    Case 1 To 18: The_Day = The_Day + 12
    The_Month = 11
    The_Year = The_Year - 622
    Case 19 To 28: The_Day = The_Day - 18
    The_Month = 12
    The_Year = The_Year - 622
    End Select
Case 3: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 10
    The_Month = 12
    The_Year = The_Year - 622
    Case 21 To 31: The_Day = The_Day - 20
    The_Month = 1
    The_Year = The_Year - 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 20: The_Day = The_Day + 11
    The_Month = 1
    The_Year = The_Year - 621
    Case 21 To 30: The_Day = The_Day - 20
    The_Month = 2
    The_Year = The_Year - 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 2
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 3
    The_Year = The_Year - 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 10
    The_Month = 3
    The_Year = The_Year - 621
    Case 22 To 30: The_Day = The_Day - 21
    The_Month = 4
    The_Year = The_Year - 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 4
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 5
    The_Year = The_Year - 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 5
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 6
    The_Year = The_Year - 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 9
    The_Month = 6
    The_Year = The_Year - 621
    Case 23 To 30: The_Day = The_Day - 22
    The_Month = 7
    The_Year = The_Year - 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 22: The_Day = The_Day + 8
    The_Month = 7
    The_Year = The_Year - 621
    Case 23 To 31: The_Day = The_Day - 22
    The_Month = 8
    The_Year = The_Year - 621
    End Select
Case 11: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 9
    The_Month = 8
    The_Year = The_Year - 621
    Case 22 To 30: The_Day = The_Day - 21
    The_Month = 9
    The_Year = The_Year - 621
    End Select
Case 12: Select Case The_Day
    Case 1 To 21: The_Day = The_Day + 9
    The_Month = 9
    The_Year = The_Year - 621
    Case 22 To 31: The_Day = The_Day - 21
    The_Month = 10
    The_Year = The_Year - 621
    End Select
End Select
'------------------------------------------------------
End If

Fa_Date = Format(CStr(The_Year), "0000") & "/" & _
         Format(CStr(The_Month), "00") & "/" & _
         Format(CStr(The_Day), "00")
End Function


Public Function En_Date(Fa_Date As String) As String
Dim The_Year As Integer
Dim The_Month As Integer
Dim The_Day As Integer
The_Year = CInt(Mid(Fa_Date, 1, 4))
The_Month = CInt(Mid(Fa_Date, 6, 2))
The_Day = CInt(Mid(Fa_Date, 9, 2))

Dim The_Select As Integer
The_Select = The_Year Mod 4

'------------------------------------------------------------------------------------------------------------------------
If The_Select = 0 Then                'Like : 1360, 1364, 1368, 1372, 1376, 1380, 1384, ...
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 3
    The_Year = The_Year + 621
    Case 12 To 31: The_Day = The_Day - 11
    The_Month = 4
    The_Year = The_Year + 621
    End Select
Case 2: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 20
    The_Month = 4
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 5
    The_Year = The_Year + 621
    End Select
Case 3: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 5
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 6
    The_Year = The_Year + 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 6
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 7
    The_Year = The_Year + 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 7
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 8
    The_Year = The_Year + 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 8
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 9
    The_Year = The_Year + 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 8: The_Day = The_Day + 22
    The_Month = 9
    The_Year = The_Year + 621
    Case 9 To 30: The_Day = The_Day - 8
    The_Month = 10
    The_Year = The_Year + 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 10
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 11
    The_Year = The_Year + 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 11
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 12
    The_Year = The_Year + 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 12
    The_Year = The_Year + 621
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 1
    The_Year = The_Year + 622
    End Select
Case 11: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 1
    The_Year = The_Year + 622
    Case 12 To 30: The_Day = The_Day - 11
    The_Month = 2
    The_Year = The_Year + 622
    End Select
Case 12: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 19
    The_Month = 2
    The_Year = The_Year + 622
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 3
    The_Year = The_Year + 622
    End Select
End Select
End If
'------------------------------------------------------------------------------------------------------------------------
If The_Select = 1 Then                'Like : 1361, 1365, 1369, 1373, 1377, 1381, 1385, ...
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 3
    The_Year = The_Year + 621
    Case 12 To 31: The_Day = The_Day - 11
    The_Month = 4
    The_Year = The_Year + 621
    End Select
Case 2: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 20
    The_Month = 4
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 5
    The_Year = The_Year + 621
    End Select
Case 3: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 22
    The_Month = 5
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 6
    The_Year = The_Year + 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 6
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 7
    The_Year = The_Year + 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 7
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 8
    The_Year = The_Year + 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 8
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 9
    The_Year = The_Year + 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 8: The_Day = The_Day + 22
    The_Month = 9
    The_Year = The_Year + 621
    Case 9 To 30: The_Day = The_Day - 8
    The_Month = 10
    The_Year = The_Year + 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 10
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 11
    The_Year = The_Year + 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 11
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 12
    The_Year = The_Year + 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 12
    The_Year = The_Year + 621
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 1
    The_Year = The_Year + 622
    End Select
Case 11: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 1
    The_Year = The_Year + 622
    Case 12 To 30: The_Day = The_Day - 11
    The_Month = 2
    The_Year = The_Year + 622
    End Select
Case 12: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 19
    The_Month = 2
    The_Year = The_Year + 622
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 3
    The_Year = The_Year + 622
    End Select
End Select
End If
'------------------------------------------------------------------------------------------------------------------------
If The_Select = 2 Then                'Like : 1362, 1366, 1370, 1374, 1378, 1382, 1386, ...
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 3
    The_Year = The_Year + 621
    Case 12 To 31: The_Day = The_Day - 11
    The_Month = 4
    The_Year = The_Year + 621
    End Select
Case 2: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 20
    The_Month = 4
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 5
    The_Year = The_Year + 621
    End Select
Case 3: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 5
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 6
    The_Year = The_Year + 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 6
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 7
    The_Year = The_Year + 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 7
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 8
    The_Year = The_Year + 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 8
    The_Year = The_Year + 621
    Case 10 To 31: The_Day = The_Day - 9
    The_Month = 9
    The_Year = The_Year + 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 8: The_Day = The_Day + 22
    The_Month = 9
    The_Year = The_Year + 621
    Case 9 To 30: The_Day = The_Day - 8
    The_Month = 10
    The_Year = The_Year + 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 22
    The_Month = 10
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 11
    The_Year = The_Year + 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 11
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 12
    The_Year = The_Year + 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 12
    The_Year = The_Year + 621
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 1
    The_Year = The_Year + 622
    End Select
Case 11: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 1
    The_Year = The_Year + 622
    Case 12 To 30: The_Day = The_Day - 11
    The_Month = 2
    The_Year = The_Year + 622
    End Select
Case 12: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 19
    The_Month = 2
    The_Year = The_Year + 622
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 3
    The_Year = The_Year + 622
    End Select
End Select
End If
'------------------------------------------------------------------------------------------------------------------------
If The_Select = 3 Then                'Like : 1363, 1367, 1371, 1375, 1379, 1383, 1387, ...
Select Case The_Month
Case 1: Select Case The_Day
    Case 1 To 12: The_Day = The_Day + 19
    The_Month = 3
    The_Year = The_Year + 621
    Case 13 To 31: The_Day = The_Day - 12
    The_Month = 4
    The_Year = The_Year + 621
    End Select
Case 2: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 19
    The_Month = 4
    The_Year = The_Year + 621
    Case 12 To 31: The_Day = The_Day - 11
    The_Month = 5
    The_Year = The_Year + 621
    End Select
Case 3: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 5
    The_Year = The_Year + 621
    Case 12 To 31: The_Day = The_Day - 11
    The_Month = 6
    The_Year = The_Year + 621
    End Select
Case 4: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 20
    The_Month = 6
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 7
    The_Year = The_Year + 621
    End Select
Case 5: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 7
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 8
    The_Year = The_Year + 621
    End Select
Case 6: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 8
    The_Year = The_Year + 621
    Case 11 To 31: The_Day = The_Day - 10
    The_Month = 9
    The_Year = The_Year + 621
    End Select
Case 7: Select Case The_Day
    Case 1 To 9: The_Day = The_Day + 21
    The_Month = 9
    The_Year = The_Year + 621
    Case 10 To 30: The_Day = The_Day - 9
    The_Month = 10
    The_Year = The_Year + 621
    End Select
Case 8: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 21
    The_Month = 10
    The_Year = The_Year + 621
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 11
    The_Year = The_Year + 621
    End Select
Case 9: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 20
    The_Month = 11
    The_Year = The_Year + 621
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 12
    The_Year = The_Year + 621
    End Select
Case 10: Select Case The_Day
    Case 1 To 11: The_Day = The_Day + 20
    The_Month = 12
    The_Year = The_Year + 621
    Case 12 To 30: The_Day = The_Day - 11
    The_Month = 1
    The_Year = The_Year + 622
    End Select
Case 11: Select Case The_Day
    Case 1 To 12: The_Day = The_Day + 19
    The_Month = 1
    The_Year = The_Year + 622
    Case 13 To 30: The_Day = The_Day - 12
    The_Month = 2
    The_Year = The_Year + 622
    End Select
Case 12: Select Case The_Day
    Case 1 To 10: The_Day = The_Day + 18
    The_Month = 2
    The_Year = The_Year + 622
    Case 11 To 30: The_Day = The_Day - 10
    The_Month = 3
    The_Year = The_Year + 622
    End Select
End Select
End If
'------------------------------------------------------------------------------------------------------------------------
En_Date = Format(CStr(The_Month), "00") & "/" & _
                            Format(CStr(The_Day), "00") & "/" & _
                                    Format(CStr(The_Year), "0000")
End Function



