Attribute VB_Name = "主题模块"
Option Explicit

Public Function applyTheme(ByVal themefile As String) As Boolean
Dim freeFle As Integer
Dim lineString As String
Dim a
Dim i As Integer
Dim myFrm As Form
Dim FontColor As Long '字体颜色
Dim FormColor As Long '窗体背景色
'验证主题文件是否存在
If exitFile(themefile) = False Then Exit Function '不存在
freeFle = freeFile()

'打开主题文件
Open themefile For Input As freeFle
    Line Input #freeFle, lineString
    If lineString <> "#随心听主题文件#" Then
        Close #freeFle
        Exit Function
    End If
    Do While EOF(freeFle) = False
        Line Input #freeFle, lineString
        If InStr(lineString, ":") <> 0 Then '存在配置
            a = Split(lineString, ":")
            Select Case a(0)
                Case "背景色"
                    FormColor = chColor(a(1))
                    '设置背景色
                    For Each myFrm In Forms
                        myFrm.BackColor = FormColor
                    Next
                    SaveSetting App.EXEName, "设置", "背景颜色", FormColor
                Case "字体色"
                    FontColor = chColor(a(1))
                    
                    For i = 0 To frm主窗体.Label1.UBound
                        frm主窗体.Label1(i).ForeColor = FontColor
                    Next i
                    SaveSetting App.EXEName, "设置", "字体颜色", FontColor
                ''''其它配置
                
            End Select
        End If
    Loop
Close #freeFle

End Function

Private Function chColor(ByVal rgbColor As String) As Long
    chColor = RGB(Left(rgbColor, 3), Mid(rgbColor, 4, 3), Right(rgbColor, 3))
    MsgBox longtoRGB(chColor)
End Function

'配置从RGB到Long
Private Function longtoRGB(ByVal Color As Long) As String
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
