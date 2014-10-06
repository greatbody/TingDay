VERSION 5.00
Begin VB.Form frmTingDay 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "TingDay"
   ClientHeight    =   11235
   ClientLeft      =   7110
   ClientTop       =   255
   ClientWidth     =   4485
   FillColor       =   &H80000011&
   Icon            =   "frmTingDay1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11235
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   3360
      ScaleHeight     =   0.8
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   18
      Top             =   2640
      Width           =   735
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   8
         X1              =   0
         X2              =   49.929
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6930
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   6930
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   0
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4920
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   240
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   0.482
         Y1              =   0.4
         Y2              =   0.4
      End
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "词"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3480
      TabIndex        =   20
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3360
      TabIndex        =   19
      ToolTipText     =   " 更换皮肤"
      Top             =   120
      Width           =   315
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      ToolTipText     =   "音量调节/静音"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "TingDay"
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      ToolTipText     =   "搜索网络歌曲"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   14
      ToolTipText     =   "添加文件夹"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   13
      ToolTipText     =   "定位到当前播放的歌曲"
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "顺"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   12
      ToolTipText     =   "当前播放的模式"
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "最新公告：随心听正式改名为：TingDay!"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   360
      TabIndex        =   11
      Top             =   10800
      Width           =   3240
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4320
      Y1              =   10440
      Y2              =   10440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4080
      TabIndex        =   10
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   3720
      TabIndex        =   7
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   42
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1680
      TabIndex        =   6
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TingDay"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "frmTingDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim keyword As String
Dim i As Integer
'搜索音乐
If KeyCode = vbKeyF And Shift = 2 Then
    keyword = InputBox("请输入要搜索的关键词", "歌曲搜索")
    If keyword = "" Then Exit Sub
    For i = 1 To listA(GeJiA.Index).Count
        If InStr(SongA(i).Title, keyword) <> 0 Then
            frmTingDay.List1.listIndex = i - 1
            Exit Sub
        End If
    Next i
    MessageBox Me.hWnd, "没有找到你要的歌曲，非常抱歉！", "歌曲搜索", vbOKOnly
ElseIf KeyCode = 13 Then
        GeJiA.Index = List2.listIndex + 1
        playList.Index = List1.listIndex + 1
        If GeJiA.Index = 0 Then GeJiA.Index = 1
        If playList.Index = 0 Then playList.Index = 1
        If listA(GeJiA.Index).Count = 0 And listA(GeJiA.Index).path <> "" Then
            openList listA(GeJiA.Index), listA(GeJiA.Index).path
            listA(GeJiA.Index).Index = playList.Index
            listA(GeJiA.Index).playWhere = playList.playWhere
            playList = listA(GeJiA.Index)
        Else
            listA(GeJiA.Index).Index = playList.Index
            listA(GeJiA.Index).playWhere = playList.playWhere
            playList = listA(GeJiA.Index)
        End If
        listPlay playList, playList.Index
        Label4.Caption = ";"
        frmMenu.停播.Caption = "暂停"
End If
updataUI
End Sub

Private Sub Form_Load()
Dim lstBcolor As Boolean
'Dim backpic As String
Dim Color As Long
'backpic = App.path & "\backpic.jpg"
Color = GetSetting(App.EXEName, "设置", "背景色", RGBtoLong("000000000"))
lstBcolor = CBool(GetSetting(App.EXEName, "设置", "颜色列表", False))
Me.BackColor = Color
If lstBcolor = True Then
    frmMenu.ColorList.Checked = True
    listbackcolor True
Else
    frmMenu.ColorList.Checked = False
    listbackcolor False
End If
'If exitFile(backpic) = True Then Me.Picture = LoadPicture(backpic)
setTrsp 210, Me.hWnd
Me.KeyPreview = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '移动窗体
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndmTingDay
End Sub

Private Sub Label10_Click()
    Select Case mPlayStyle
    Case s单曲播放
        mPlayStyle = s顺序播放
        Label10 = "顺"
    Case s顺序播放
        mPlayStyle = s随机播放
        Label10 = "随"
    Case s随机播放
        mPlayStyle = s循环播放
        Label10 = "循"
    Case s循环播放
        mPlayStyle = s单曲播放
        Label10 = "单"
    End Select
SaveSetting App.EXEName, "设置", "播放模式", CStr(mPlayStyle)
End Sub

Private Sub Label12_Click()
    定位歌曲
End Sub

Private Sub Label14_Click()
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

Private Sub Label15_Click()
Dim msg As Integer
msg = MessageBox(Me.hWnd, "将会打开百度音乐，是否打开？", "下载音乐", vbYesNo)
If msg = vbYes Then
    ShellExecute Me.hWnd, "Open", "http://music.baidu.com/", "", "", 1
End If
End Sub

Private Sub Label18_Click() '静音
    setVolume 0, 0
    Label18.ForeColor = vbRed
End Sub

Private Sub Label19_Click()
frmMenu.PopupMenu frmMenu.m_BackColor
End Sub

Private Sub Label4_Click()
    If playList.Count = 0 Then
        FileOpen ShowOpen("打开歌集", "所有文件", Me.hWnd)
        Exit Sub
    End If
    If mStatus.fPlay = 1 Then
        zplay_Pause mTingDay
        Label4.Caption = "4"
        frmMenu.停播.Caption = "播放"
    ElseIf mStatus.fPause = 1 Then
        zplay_Play mTingDay
        Label4.Caption = ";"
        frmMenu.停播.Caption = "暂停"
    Else
        GeJiA.Index = List2.listIndex + 1
        playList.Index = List1.listIndex + 1
        If GeJiA.Index = 0 Then GeJiA.Index = 1
        If playList.Index = 0 Then playList.Index = 1
        If listA(GeJiA.Index).Count = 0 And listA(GeJiA.Index).path <> "" Then
            openList listA(GeJiA.Index), listA(GeJiA.Index).path
            listA(GeJiA.Index).Index = playList.Index
            listA(GeJiA.Index).playWhere = playList.playWhere
            playList = listA(GeJiA.Index)
        Else
            listA(GeJiA.Index).Index = playList.Index
            listA(GeJiA.Index).playWhere = playList.playWhere
            playList = listA(GeJiA.Index)
        End If
        listPlay playList, playList.Index
        Label4.Caption = ";"
        frmMenu.停播.Caption = "暂停"
    End If
    updataUI
End Sub

Private Sub Label5_Click()
'Exit Sub
If frmLRC.Visible = True Then
    Unload frmLRC
    Label5.ForeColor = &H808080
    
Else
    With frmLRC
    .Left = Me.Left + Me.Width + 50
    .Top = Me.Top
    .Show
    End With
    Label5.ForeColor = &HFFFFFF
    
End If
End Sub

Private Sub Label6_Click()
    Me.Hide
    Me.WindowState = 0
End Sub

Private Sub Label7_Click()
    playPre '播放上一曲
    updataUI
End Sub

Private Sub Label8_Click()
    If mPlayStyle = 1 Then
        Rndplay
    Else
        playNext '播放下一曲
    End If
    updataUI
End Sub

Private Sub Label9_Click()
EndmTingDay
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If Me.List1.listIndex < 0 Then
        frmMenu.打开文件位置.Enabled = False
        frmMenu.删除歌曲.Enabled = False
    Else
        frmMenu.打开文件位置.Enabled = True
        frmMenu.删除歌曲.Enabled = True
    End If
    frmMenu.PopupMenu frmMenu.菜单2
ElseIf Button = 1 Then
    playList.Index = List1.listIndex + 1
End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    frmMenu.PopupMenu frmMenu.菜单1
ElseIf Button = 1 Then
    GeJiA.Index = List2.listIndex + 1
    listShow GeJiA.Index
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim a As Single
Dim b As Single
Dim c As Single
Dim sTime As TStreamTime
a = X / 100
b = curLength / totalLength
If a > b Then
    c = a - b
    sTime.sec = c * totalLength
    seekTime sTime, smFromCurrentForward
Else
    c = b - a
    sTime.sec = c * totalLength
    seekTime sTime, smFromCurrentBackward
End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    setVolume X, X
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then setVolume X, X
End Sub

Private Sub Timer1_Timer()
    Dim mForm As RECT
    mForm.Left = ScaleX(Me.Left, vbTwips, vbPixels)
    mForm.Right = ScaleX(Me.Width, vbTwips, vbPixels) + ScaleX(frmTingDay.Left, vbTwips, vbPixels)
    mForm.Top = ScaleY(frmTingDay.Top, vbTwips, vbPixels)
    mForm.Bottom = ScaleY(frmTingDay.Height, vbTwips, vbPixels) + ScaleY(frmTingDay.Top, vbTwips, vbPixels)
      '窗体顶部隐藏
    hit mForm, 1
End Sub

'进度条 判断是否已经播放完毕
Private Sub Timer2_Timer()
    Dim t As TStreamTime
    zplay_GetStatus mTingDay, mStatus
    zplay_GetPosition mTingDay, t
    curLength = t.sec
    '改变进度条
    Line1.X2 = CInt((curLength / totalLength) * 100)
    '改变时间
    Label2.Caption = SeccondToMin(curLength)
    '判断是否播放完毕
    If mStatus.fPlay = 0 And mStatus.fPause = 0 And playList.isPlay = True Then  '播放完毕,检查是否需要播放下一首
         playOver '播放完成
         updataUI
    End If
End Sub

Private Sub List1_DblClick()
Dim lrcPath As String
'播放选定歌曲
playList = listA(GeJiA.Index)
playList.playWhere = List1.listIndex + 1
GeJiA.playWhere = List2.listIndex + 1
listPlay playList, playList.playWhere
Label1.Caption = SongA(playList.playWhere).Title

updataUI
End Sub



