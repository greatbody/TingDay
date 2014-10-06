Attribute VB_Name = "����ģ��"
Option Explicit
Public Sub listbackcolor(ByVal isBackColor As Boolean)
    If isBackColor = False Then
        frmMenu.ColorList.Checked = False
        SaveSetting App.EXEName, "����", "��ɫ�б�", False
        frmTingDay.List1.ForeColor = vbBlack
        frmTingDay.List2.ForeColor = vbBlack
        frmTingDay.List1.BackColor = vbWhite
        frmTingDay.List2.BackColor = vbWhite
    Else
        frmMenu.ColorList.Checked = True
        SaveSetting App.EXEName, "����", "��ɫ�б�", True
        frmTingDay.List1.ForeColor = vbWhite
        frmTingDay.List2.ForeColor = vbWhite
        frmTingDay.List1.BackColor = frmTingDay.BackColor
        frmTingDay.List2.BackColor = frmTingDay.BackColor
    End If
End Sub
'Ӧ�������ļ�
Public Function applyTheme(ByVal themeFile As String) As Boolean
Dim FREEfle As Integer
Dim lineString As String
Dim a
Dim i As Integer
Dim myFrm As Form
Dim FontColor As Long '������ɫ
Dim FormColor As Long '���屳��ɫ
'��֤�����ļ��Ƿ����
If exitFile(themeFile) = False Then Exit Function '������
FREEfle = FreeFile()
'�ļ�
Open themeFile For Input As FREEfle
    Line Input #FREEfle, lineString
    If lineString <> "#�����������ļ�#" Then
        Close #FREEfle
        Exit Function
    End If
    Do While EOF(FREEfle) = False
        Line Input #FREEfle, lineString
        If InStr(lineString, ":") <> 0 Then '��������
            a = Split(lineString, ":")
            Select Case a(0)
                Case "����ɫ"
                    FormColor = RGBtoLong(a(1))
                    '���ñ���ɫ
                    For Each myFrm In Forms
                        myFrm.BackColor = FormColor
                    Next
                    SaveSetting App.EXEName, "����", "������ɫ", FormColor
                Case "����ɫ"
                    FontColor = RGBtoLong(a(1))
                    
                    SaveSetting App.EXEName, "����", "������ɫ", FontColor
                ''''��������
                
            End Select
        End If
    Loop
Close #FREEfle
End Function

'���浱ǰ���õ������ļ�
Public Sub saveThemetoFile(ByVal themeName As String)
    Dim themeFile As String
    Dim themePath As String '�����ļ�������ļ���
    Dim FREEfle As Integer
    Dim head As String '����ͷ�ļ�
    Dim description As String '����˵���ļ�
    Dim setting(0 To 1) As String '����
    themePath = App.path & "\theme"
    If exitFolder(themePath) = False Then '����ļ��в����ڣ�����һ�������ļ���
        Call buildFolder(themePath)
    End If
    FREEfle = FreeFile()
    themeFile = themePath & "\" & themeName & ".txt" '�����ļ���
    head = "#�����������ļ�#"
    description = "#�������ļ������������ֲ��������ɣ��벻Ҫ�޸�������ݣ�������ɳ������#"
    setting(0) = "����ɫ:" & longtoRGB(GetSetting(App.EXEName, "����", "������ɫ", vbBlack))
    setting(1) = "����ɫ:" & longtoRGB(GetSetting(App.EXEName, "����", "������ɫ", vbRed))
    '���������ļ�
    Open themeFile For Output As #FREEfle
        Print #FREEfle, head & vbCrLf & description & vbCrLf & setting(0) & vbCrLf & setting(1)
    Close #FREEfle
End Sub
'Long��RGB ��ɫת��
Public Function RGBtoLong(ByVal rgbColor As String) As Long
    RGBtoLong = RGB(Left(rgbColor, 3), Mid(rgbColor, 4, 3), Right(rgbColor, 3))
End Function

'��ɫ��RGB��Long ��ת��
Public Function longtoRGB(ByVal Color As Long) As String
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
