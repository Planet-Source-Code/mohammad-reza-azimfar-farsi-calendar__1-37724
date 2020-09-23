Attribute VB_Name = "Global"
Option Explicit
'----------------------------
'Global Variavles & Functions
'----------------------------
'
Public Cn As New ADODB.Connection
Public Fs As New FileSystemObject
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public S_HTML As String
Public S_Body As String
Public S_Title As String
Public E_Title As String
Public S_Table As String
Public S_Header_TD As String
Public E_Header_TD As String
Public S_Row_TD As String
Public S_Row_TD1 As String
Public S_Row_TD2 As String
Public E_Row_TD As String
'

Public Function Open_Cn()
' ÇíÌÇÏ íß ÇÊÕÇá  ÓÑÊÇÓÑí Èå ÈÇäß ÇØáÇÚÇÊí
On Error GoTo Cn_Err
Cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Calendar.mdb;Persist Security Info=False"
Cn.CursorLocation = adUseClient
Cn.Mode = adModeReadWrite
Cn.Open
Exit Function
Cn_Err:
MsgBox "ÇÑÊÈÇØ ÈÇ ÈÇäß ÇØáÇÚÇÊí ÇíÌÇÏ äÔÏ "
End Function

Public Function Encode(Txt As String) As String
If Txt = "" Then
    Encode = "&nbsp;"
    Exit Function
End If
Dim i As Long
Dim Ret As String
For i = 1 To Len(Txt)
    Select Case Mid(Txt, i, 1)
        Case Chr(13)
        Ret = Ret & "<br>"
        Case Chr(10)
        Ret = Ret & Chr(253)
        Case Else
        Ret = Ret & Mid(Txt, i, 1)
    End Select
Next
Encode = Ret
End Function

Public Sub Set_Report()
S_HTML = "<HTML><HEAD><META http-equiv=""Content-Type"" Content =""Text/html;charset=Windows-1256""></HEAD>" & vbCrLf
S_Body = "<Body dir=RTL TopMargin=0 LeftMargin=0 >" & vbCrLf
S_Title = "<p align=Center><B><U><font size=5 face=Arial >" & vbCrLf
E_Title = "</font></U></B></P>" & vbCrLf
S_Table = "<Table Width=90% Align=Center Border=1 CellPadding=2 CellSpacing=2 >" & vbCrLf
S_Header_TD = "<td align=Center ><B><font size=4 face=Arial>" & vbCrLf
E_Header_TD = "</font></B></td>" & vbCrLf
S_Row_TD = "<td align=Right ><font size=4 face=Arial>" & vbCrLf
S_Row_TD1 = "<td align=Right Width=30% ><font size=3 face=Arial >" & vbCrLf
S_Row_TD2 = "<td align=Right Width=70% ><font size=3 face=Arial >" & vbCrLf
E_Row_TD = "</font></td>" & vbCrLf
End Sub
