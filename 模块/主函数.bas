Attribute VB_Name = "������"
Option Explicit
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long    '������������رմ򿪵�ע���
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Type POINTAPI
X As Long
Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public whnd As Long
Public mCaption As String 'ͷ����ʾ
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Private Sub Main()
    Dim info As Boolean
    SaveSetting App.EXEName, "�汾����", "��ǰ�汾", "20141001"
    playList.playWhere = GetSetting(App.EXEName, "����", "playWhere", 0)
    GeJiA.playWhere = GetSetting(App.EXEName, "����", "listIndex", 0)
    playList.Index = playList.playWhere
    GeJiA.Index = GeJiA.playWhere
    GeJiPath = App.path & "\listrecord"
    mVolume = CLng(GetSetting(App.EXEName, "����", "����", "80")) '������Ϣ�Ļ�ȡ
    'mVolume = 90
    'setLink
    info = GetSetting(App.EXEName, "����", "firstRun", True)
    If info = True Then
        MsgBox "�װ���TingDayʹ������ã��������ʹ�õĹ�������ʲô���⣬���Ը��ҷ��ʼ���" & vbCrLf & "woyufan@163.com", vbOKOnly, "��һ������"
        Call SaveSetting(App.EXEName, "����", "firstRun", False)
    End If
    InitmTingDay
    'InitCommonControls
    If exitFile(GeJiPath) = False Then
        saveGeJi GeJiPath
    End If
    openGeJi GeJiPath
    If GeJiA.playWhere > 0 Then openList playList, listA(GeJiA.playWhere).path
    setOnly frmTingDay.hWnd, Command
    showMe
    If Command <> "" Then
        Select Case Command
        
        Case "autorun"
            frmTingDay.Hide
        Case Else
            FileOpen Command
        End Select
    End If
    SNIcon_Add frmMenu.Icon, frmMenu.hWnd, "TingDay"
    '�������UI����
    refreshGeJi
    listShow GeJiA.Index
    updataUI
End Sub


