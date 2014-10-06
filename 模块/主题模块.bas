Attribute VB_Name = "主题模块"
Option Explicit
Public Sub listbackcolor(ByVal isBackColor As Boolean)
    If isBackColor = False Then
        frmMenu.ColorList.Checked = False
        SaveSetting App.EXEName, "设置", "颜色列表", False
        frmTingDay.List1.ForeColor = vbBlack
        frmTingDay.List2.ForeColor = vbBlack
        frmTingDay.List1.BackColor = vbWhite
        frmTingDay.List2.BackColor = vbWhite
    Else
        frmMenu.ColorList.Checked = True
        SaveSetting App.EXEName, "设置", "颜色列表", True
        frmTingDay.List1.ForeColor = vbWhite
        frmTingDay.List2.ForeColor = vbWhite
        frmTingDay.List1.BackColor = frmTingDay.BackColor
        frmTingDay.List2.BackColor = frmTingDay.BackColor
    End If
End Sub
'应用主题文件
Public Function applyTheme(ByVal themeFile As String) As Boolean
Dim FREEfle As Integer
Dim lineString As String
Dim a
Dim i As Integer
Dim myFrm As Form
Dim FontColor As Long '字体颜色
Dim FormColor As Long '窗体背景色
'验证主题文件是否存在
If exitFile(themeFile) = False Then Exit Function '不存在
FREEfle = FreeFile()
'文件
Open themeFile For Input As FREEfle
    Line Input #FREEfle, lineString
    If lineString <> "#随心听主题文件#" Then
        Close #FREEfle
        Exit Function
    End If
    Do While EOF(FREEfle) = False
        Line Input #FREEfle, lineString
        If InStr(lineString, ":") <> 0 Then '存在配置
            a = Split(lineString, ":")
            Select Case a(0)
                Case "背景色"
                    FormColor = RGBtoLong(a(1))
                    '设置背景色
                    For Each myFrm In Forms
                        myFrm.BackColor = FormColor
                    Next
                    SaveSetting App.EXEName, "设置", "背景颜色", FormColor
                Case "字体色"
                    FontColor = RGBtoLong(a(1))
                    
                    SaveSetting App.EXEName, "设置", "字体颜色", FontColor
                ''''其它配置
                
            End Select
        End If
    Loop
Close #FREEfle
End Function

'保存当前设置到主题文件
Public Sub saveThemetoFile(ByVal themeName As String)
    Dim themeFile As String
    Dim themePath As String '主题文件保存的文件夹
    Dim FREEfle As Integer
    Dim head As String '主题头文件
    Dim description As String '主题说明文件
    Dim setting(0 To 1) As String '配置
    themePath = App.path & "\theme"
    If exitFolder(themePath) = False Then '如果文件夹不存在，则建立一个主题文件夹
        Call buildFolder(themePath)
    End If
    FREEfle = FreeFile()
    themeFile = themePath & "\" & themeName & ".txt" '主题文件名
    head = "#随心听主题文件#"
    description = "#此主题文件由随心听音乐播放器生成，请不要修改相关数据，以免造成程序错误！#"
    setting(0) = "背景色:" & longtoRGB(GetSetting(App.EXEName, "设置", "背景颜色", vbBlack))
    setting(1) = "字体色:" & longtoRGB(GetSetting(App.EXEName, "设置", "字体颜色", vbRed))
    '生成主题文件
    Open themeFile For Output As #FREEfle
        Print #FREEfle, head & vbCrLf & description & vbCrLf & setting(0) & vbCrLf & setting(1)
    Close #FREEfle
End Sub
'Long到RGB 颜色转换
Public Function RGBtoLong(ByVal rgbColor As String) As Long
    RGBtoLong = RGB(Left(rgbColor, 3), Mid(rgbColor, 4, 3), Right(rgbColor, 3))
End Function

'颜色从RGB到Long 的转换
Public Function longtoRGB(ByVal Color As Long) As String
Dim Red As Integer, Green As Integer, Blue As Integer
Dim Red2 As String, Green2 As String, Blue2 As String

    Red = Color And &HFF '拆分颜色
    Green = (Color And 65280) \ 256
    Blue = (Color And &HFF0000) \ 65536
    
    '红色处理
    Select Case Len(CStr(Red))
    Case 3
        Red2 = CStr(Red)
    Case 2
        Red2 = "0" & CStr(Red)
    Case 1
        Red2 = "00" & CStr(Red)
    End Select
    '绿色处理
    Select Case Len(CStr(Green))
    Case 3
        Green2 = CStr(Green)
    Case 2
        Green2 = "0" & CStr(Green)
    Case 1
        Green2 = "00" & CStr(Green)
    End Select
    '蓝色处理
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
