Attribute VB_Name = "���򲻶࿪"
Option Explicit
'���ڲ���API
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SW_SHOWNORMAL = 1
'���ó����Ƿ�࿪
Public Sub setOnly(ByVal hwnd As Long, ByVal FileName As String)
    If App.PrevInstance = True Then  '����࿪
            SaveSetting App.EXEName, "����࿪", "�����ļ�", FileName
             hwnd = GetSetting(App.EXEName, "����࿪", "���", 0)
             If hwnd <> 0 Then
                ShowWindow hwnd, SW_SHOWNORMAL
            End If
             Call SNIcon_Del
            End
    End If
End Sub

