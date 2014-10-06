Attribute VB_Name = "热键注册"
Option Explicit

Public Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

Public Const WM_HOTKEY = &H312
 Public Const GWL_WNDPROC = (-4)
 
'Public GetProc As Long '自身钩子
Public preProc As Long '系统钩子



Public Sub RegHotKeyPlayNext()
    RegisterHotKey frmMenu.hWnd, 1, MOD_CONTROL, vbKeyLeft
End Sub

Public Sub RegHotKeyPlayPre()
    RegisterHotKey frmMenu.hWnd, 2, MOD_CONTROL, vbKeyRight
End Sub

Public Sub RegHotKeyPlay()
    RegisterHotKey frmMenu.hWnd, 3, MOD_CONTROL, vbKeyQ
End Sub

Public Sub RegHotKeyShow()
    RegisterHotKey frmMenu.hWnd, 4, MOD_CONTROL, vbKeyA
End Sub


Public Sub RegHotKeyAll()
    RegHotKeyPlayNext
    RegHotKeyPlayPre
    RegHotKeyPlay
    RegHotKeyShow
End Sub

Public Sub UnRegHot()
    Dim i As Long
    For i = 1 To 4
        UnregisterHotKey frmMenu.hWnd, i
    Next i
End Sub

Public Function GetProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Debug.Print msg
    If msg = WM_HOTKEY Then
        Select Case wParam
        Case 1 '上一曲
            playPre
        Case 2 '下一曲
            If mPlayStyle = 1 Then
                Rndplay
            Else
                playNext '播放下一曲
            End If
        
        Case 3 '暂停、播放
            If mStatus.fPlay = 1 Then
                zplay_Pause mTingDay
            ElseIf mStatus.fPause = 1 Then
                zplay_Play mTingDay
            End If
        Case 4 '显示、隐藏
            If frmTingDay.Visible = True Then
                frmTingDay.Hide
                frmTingDay.WindowState = 0
            Else
                showMe
            End If
            
        End Select
        updataUI
    End If
    GetProc = CallWindowProc(preProc, hWnd, msg, wParam, lParam)
 End Function
