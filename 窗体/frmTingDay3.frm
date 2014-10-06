VERSION 5.00
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "TingDay"
   ClientHeight    =   2655
   ClientLeft      =   8160
   ClientTop       =   4110
   ClientWidth     =   4455
   Icon            =   "frmTingDay3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   720
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   2040
      Top             =   1200
   End
   Begin VB.Menu m_BackColor 
      Caption         =   "背景颜色"
      Begin VB.Menu ColorList 
         Caption         =   "颜色列表"
         Checked         =   -1  'True
      End
      Begin VB.Menu skyBlue 
         Caption         =   "天空蓝"
      End
      Begin VB.Menu perpol 
         Caption         =   "紫色"
      End
      Begin VB.Menu 分割线_颜色3 
         Caption         =   "-"
      End
      Begin VB.Menu 酷黑 
         Caption         =   "酷黑"
      End
      Begin VB.Menu MoLv 
         Caption         =   "墨绿"
      End
      Begin VB.Menu coffee 
         Caption         =   "咖啡色"
      End
      Begin VB.Menu 分割线_颜色1 
         Caption         =   "-"
      End
      Begin VB.Menu orange 
         Caption         =   "橙色"
      End
      Begin VB.Menu Caolv 
         Caption         =   "草绿"
      End
      Begin VB.Menu Red 
         Caption         =   "玫瑰红"
      End
   End
   Begin VB.Menu 菜单1 
      Caption         =   "菜单1"
      Begin VB.Menu 添加文件夹 
         Caption         =   "添加文件夹"
      End
      Begin VB.Menu 打开歌单 
         Caption         =   "打开歌单"
      End
      Begin VB.Menu 分割线_菜单1 
         Caption         =   "-"
      End
      Begin VB.Menu 重命名 
         Caption         =   "重命名"
      End
      Begin VB.Menu 分割线2_菜单1 
         Caption         =   "-"
      End
      Begin VB.Menu 新建歌单 
         Caption         =   "新建歌单"
      End
      Begin VB.Menu 删除歌单 
         Caption         =   "删除歌单"
      End
      Begin VB.Menu 保存歌单 
         Caption         =   "保存歌单"
      End
   End
   Begin VB.Menu 菜单 
      Caption         =   "菜单"
      Begin VB.Menu 显示 
         Caption         =   "TingDay"
      End
      Begin VB.Menu 分割线3 
         Caption         =   "-"
      End
      Begin VB.Menu 禁用热键 
         Caption         =   "禁用热键"
      End
      Begin VB.Menu TingDay分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 上一曲 
         Caption         =   "上一曲"
      End
      Begin VB.Menu 停播 
         Caption         =   "播放"
      End
      Begin VB.Menu 下一曲 
         Caption         =   "下一曲"
      End
      Begin VB.Menu 分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 随机播放 
         Caption         =   "随机播放"
      End
      Begin VB.Menu 分割线2 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu 菜单2 
      Caption         =   "菜单2"
      Begin VB.Menu 添加文件 
         Caption         =   "添加文件"
      End
      Begin VB.Menu 打开文件位置 
         Caption         =   "打开文件位置"
      End
      Begin VB.Menu 删除歌曲 
         Caption         =   "删除歌曲"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub coffee_Click()
setBackColor "096056017"
allFalse
coffee.Checked = True
End Sub

Private Sub ColorList_Click()
    If ColorList.Checked = True Then
        listbackcolor False
    Else
        listbackcolor True
    End If
End Sub

Private Sub Form_Load()
Me.Visible = False
preProc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf GetProc
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
msg = X / 15
Select Case msg

Case WM_LBUTTONDBLCLK
    frmTingDay.Show
    If frmTingDay.Top < 0 Then frmTingDay.Top = 100
Case WM_RBUTTONDOWN
    Call SetForegroundWindow(Me.hWnd)
    Me.PopupMenu 菜单
    
End Select
End Sub

Private Sub MoLv_Click()
setBackColor "000070048"
allFalse
MoLv.Checked = True
End Sub

Private Sub orange_Click()
setBackColor "255097000"
allFalse
orange.Checked = True
End Sub

Private Sub perpol_Click()
setBackColor "153051225"
allFalse
perpol.Checked = True
End Sub

Private Sub Red_Click()
setBackColor "252033094"
allFalse
Red.Checked = True
End Sub
Private Sub allFalse()
skyBlue.Checked = False
perpol.Checked = False
MoLv.Checked = False
coffee.Checked = False
orange.Checked = False
Caolv.Checked = False
Red.Checked = False
酷黑.Checked = False
End Sub
Private Sub skyBlue_Click()
setBackColor "030144225"
allFalse
skyBlue.Checked = True

End Sub

Private Sub Timer1_Timer()
Dim FileName As String


FileName = GetSetting(App.EXEName, "程序多开", "歌曲文件", "")
If FileName = "" Then Exit Sub
FileName = Mid(FileName, 2, Len(FileName) - 2)
If exitFile(FileName) = True Then
    FileOpen FileName
    SaveSetting App.EXEName, "程序多开", "歌曲文件", ""
End If

End Sub

Private Sub caolv_Click()
setBackColor "000199140"
allFalse
Caolv.Checked = True
End Sub

Private Sub Timer2_Timer()
'注册热键
RegHotKeyAll
End Sub

Private Sub 保存歌单_Click()
Dim FileName As String
Dim msg As Boolean
Dim mList As list
FileName = ShowSave("保存歌单", "歌单文件（*.tdl）" & Chr(0) & "*.tdl", Me.hWnd)
mList = playList
msg = saveList(mList, FileName)
saveGeJi GeJiPath
If msg = vbYes Then MessageBox frmMenu.hWnd, "保存成功", "提示", vbOKOnly
End Sub

Private Sub 打开歌单_Click()
    Dim FileName As String
    Dim mList As list
    FileName = ShowOpen("打开歌单", "歌单文件（*.tdl）" & Chr(0) & "*.tdl", Me.hWnd)
    If FileName = "" Then Exit Sub
    openList mList, FileName
    addList GeJiA, mList
    saveGeJi GeJiPath
    GeJiA.Index = GeJiA.Count
    refreshList listA(GeJiA.Index)
    refreshGeJi
End Sub

Private Sub 打开文件位置_Click()
On Error Resume Next
Dim File As String
Dim FileName As String
FileName = SongA(playList.Index).url
If FileName = "" Then Exit Sub
File = Dir(FileName)
Shell "explorer /select," & FileName, vbNormalFocus
End Sub

Private Sub 歌曲信息_Click()

End Sub

Private Sub 禁用热键_Click()
If 禁用热键.Checked = False Then
    Timer2.Enabled = False
    禁用热键.Checked = True
    UnRegHot
Else
    Timer2.Enabled = True
    禁用热键.Checked = False
End If
End Sub

Private Sub 酷黑_Click()
setBackColor "000000000"
allFalse
酷黑.Checked = True
End Sub

Private Sub 删除歌单_Click()
Dim msg As Integer
If GeJiA.Index > 0 Then
    msg = MessageBox(Me.hWnd, " 删除的歌单将不能恢复,是否继续", "询问", vbYesNo)
    If msg = vbNo Then Exit Sub
    deleteList GeJiA, GeJiA.Index
    If GeJiA.Count > 0 Then
        If GeJiA.Index > GeJiA.Count Then
            GeJiA.Index = GeJiA.Count
            zplay_Stop mTingDay
            playList.isPlay = False
            playList.Count = 0
            updataUI
        End If
        If GeJiA.playWhere > GeJiA.Index Then GeJiA.playWhere = GeJiA.playWhere - 1
        listShow GeJiA.Index
    Else
        
    End If
    refreshList listA(GeJiA.Index)
    refreshGeJi
    saveGeJi GeJiPath
End If
End Sub

Private Sub 删除歌曲_Click()
Dim msg As Integer
Dim i As Integer
i = frmTingDay.List1.listIndex + 1
If i > 0 Then
    msg = MessageBox(Me.hWnd, "这首歌曲只是从列表中移除是否继续", "询问", vbYesNo)
    If msg = vbNo Then Exit Sub
    deleteSong listA(GeJiA.Index), i
    refreshList listA(GeJiA.Index)
    saveList listA(GeJiA.Index), listA(GeJiA.Index).path
    playList = listA(GeJiA.Index)
End If
End Sub

Private Sub 上一曲_Click()
playPre
updataUI
End Sub

Private Sub 随机播放_Click()
Rndplay
updataUI
End Sub

Private Sub 添加文件_Click()
    Dim mSong As Song
    With mSong
        .url = ShowOpen("添加歌曲", "MP3文件（*.mp3）" & Chr(0) & "*.mp3", Me.hWnd)
        If .url = "" Or exitFile(.url) = False Then Exit Sub
        .title = delTuo(.url, True)
        .Islocal = True
    End With
    addsong listA(GeJiA.Index), mSong
    refreshList listA(GeJiA.Index)
    saveList listA(GeJiA.Index), listA(GeJiA.Index).path
End Sub

Private Sub 添加文件夹_Click()
    Dim FileName As String
    Dim Caption As String
    Dim mList As list
    Dim Free As Long
    Free = FreeFile
    FileName = showFolder("添加文件夹", Me.hWnd)
    If FileName = "" Then Exit Sub
    Caption = FolderName(FileName)
    With mList
        .Name = Caption
        .path = FileName
        seachFiles mList, FileName, ".mp3", True
        .Index = 1
        .playWhere = 0
    End With
    addList GeJiA, mList
    saveGeJi GeJiPath
    refreshList mList
    refreshGeJi
    GeJiA.Index = GeJiA.Count
    frmTingDay.List1.listIndex = mList.Index - 1
    frmTingDay.List2.listIndex = GeJiA.Index - 1
End Sub

Private Sub 停播_Click()
    If playList.Count = 0 Then
        FileOpen ShowOpen("打开歌集", "所有文件", Me.hWnd)
        Exit Sub
    End If
    If mStatus.fPlay = 1 Then
        zplay_Pause mTingDay
        frmTingDay.Label4.Caption = "4"
        停播.Caption = "播放"
    ElseIf mStatus.fPause = 1 Then
        zplay_Play mTingDay
        frmTingDay.Label4.Caption = ";"
        停播.Caption = "暂停"
    Else
        playList = playList
        listPlay playList, playList.Index + 1
        frmTingDay.Label4.Caption = ";"
        停播.Caption = "暂停"
    End If
    updataUI
End Sub

Private Sub 退出_Click()
EndmTingDay
End Sub

Private Sub 下一曲_Click()
If mPlayStyle = 1 Then
    Rndplay
Else
    playNext '播放下一曲
End If
updataUI
End Sub

Private Sub 显示_Click()
showMe
End Sub

Private Sub 新建歌单_Click()
Dim Caption As String
Dim FileName As String
Dim mTuo As String
Dim mList As list
FileName = ShowSave("保存歌单", "歌集文件（*.tdl）" & Chr(0) & "*.tdl", Me.hWnd)
If FileName = "" Then Exit Sub
Caption = delTuo(FileName, True)
If Caption = "" Then Exit Sub

With mList
    .path = FileName
    .Name = Caption
End With
saveList mList, FileName
addList GeJiA, mList
saveGeJi GeJiPath
refreshList mList
refreshGeJi
End Sub

Private Sub 重命名_Click()
Dim Name As String
Name = InputBox("请输入新的名称：", "重命名")
If Name = "" Then Exit Sub
listA(GeJiA.Index).Name = Name
saveGeJi GeJiPath
refreshGeJi
End Sub
