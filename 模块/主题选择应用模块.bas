Attribute VB_Name = "����ģ��"
Option Explicit

Public Function applyTheme(ByVal themefile As String) As Boolean
Dim freeFle As Integer
Dim lineString As String
Dim a
Dim i As Integer
Dim myFrm As Form
Dim FontColor As Long '������ɫ
Dim FormColor As Long '���屳��ɫ
'��֤�����ļ��Ƿ����
If exitFile(themefile) = False Then Exit Function '������
freeFle = freeFile()

'�������ļ�
Open themefile For Input As freeFle
    Line Input #freeFle, lineString
    If lineString <> "#�����������ļ�#" Then
        Close #freeFle
        Exit Function
    End If
    Do While EOF(freeFle) = False
        Line Input #freeFle, lineString
        If InStr(lineString, ":") <> 0 Then '��������
            a = Split(lineString, ":")
            Select Case a(0)
                Case "����ɫ"
                    FormColor = chColor(a(1))
                    '���ñ���ɫ
                    For Each myFrm In Forms
                        myFrm.BackColor = FormColor
                    Next
                    SaveSetting App.EXEName, "����", "������ɫ", FormColor
                Case "����ɫ"
                    FontColor = chColor(a(1))
                    
                    For i = 0 To frm������.Label1.UBound
                        frm������.Label1(i).ForeColor = FontColor
                    Next i
                    SaveSetting App.EXEName, "����", "������ɫ", FontColor
                ''''��������
                
            End Select
        End If
    Loop
Close #freeFle

End Function

Private Function chColor(ByVal rgbColor As String) As Long
    chColor = RGB(Left(rgbColor, 3), Mid(rgbColor, 4, 3), Right(rgbColor, 3))
    MsgBox longtoRGB(chColor)
End Function

'���ô�RGB��Long
Private Function longtoRGB(ByVal Color As Long) As String
Dim Red As Integer, Green As Integer, Blue As Integer
Dim Red2 As String, Green2 As String, Blue2 As String

    Red = Color And &HFF '�����ɫ
    Green = (Color And 65280) \ 256
    Blue = (Color And &HFF0000) \ 65536
    
    '��ɫ����
    Select Case Len(CStr(Red))
    Case 3
        Red2 = CStr(Red)
    Case 2
        Red2 = "0" & CStr(Red)
    Case 1
        Red2 = "00" & CStr(Red)
    End Select
    '��ɫ����
    Select Case Len(CStr(Green))
    Case 3
        Green2 = CStr(Green)
    Case 2
        Green2 = "0" & CStr(Green)
    Case 1
        Green2 = "00" & CStr(Green)
    End Select
    '��ɫ����
    Select Case Len(CStr(Blue))
    Case 3
        Blue2 = CStr(Blue)
    Case 2
        Blue2 = "0" & CStr(Blue)
    Case 1
        Blue2 = "00" & CStr(Blue)
    End Select
longtoRGB = Red2 & Green2 & Blue2
End Function
