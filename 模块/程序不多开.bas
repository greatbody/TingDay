Attribute VB_Name = "程序不多开"
Option Explicit
'窗口操作API
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SW_SHOWNORMAL = 1
'设置程序是否多开
Public Sub setOnly(ByVal hwnd As Long, ByVal FileName As String)
    If App.PrevInstance = True Then  '程序多开
            SaveSetting App.EXEName, "程序多开", "歌曲文件", FileName
             hwnd = GetSetting(App.EXEName, "程序多开", "句柄", 0)
             If hwnd <> 0 Then
                ShowWindow hwnd, SW_SHOWNORMAL
            End If
             Call SNIcon_Del
            End
    End If
End Sub

