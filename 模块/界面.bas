Attribute VB_Name = "界面"
Option Explicit
Public FormA  As RECT

Public Sub showMe()
    frmTingDay.Left = 4080
    frmTingDay.Top = 240
    frmTingDay.Show
    frmTingDay.Timer1.Enabled = True
    frmTingDay.Timer2.Enabled = True
    SaveSetting App.EXEName, "程序多开", "句柄", frmTingDay.hWnd
    setVolume mVolume, mVolume
End Sub

Public Sub setTrsp(ByVal Trsp As Long, ByVal hWnd As Long)
    Dim Trs As Long
    Trs = GetWindowLong(hWnd, GWL_EXSTYLE)
    Trs = Trs Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Trs
    SetLayeredWindowAttributes hWnd, 0, Trsp, LWA_ALPHA
End Sub

Public Sub showVolume()
    Dim leftVolume As Long
    Dim rightVolume As Long
    zplay_GetPlayerVolume mTingDay, leftVolume, rightVolume
    frmTingDay.Line9.X2 = rightVolume
    If rightVolume > 0 Then
        frmTingDay.Label18.ForeColor = vbWhite
    Else
        frmTingDay.Label18.ForeColor = vbRed
    End If
End Sub

Public Sub 定位歌曲()
    '定位列表在歌单的位置
    If playList.isPlay = False Then Exit Sub
    If GeJiA.playWhere = 0 Or playList.playWhere = 0 Then Exit Sub
    listShow GeJiA.playWhere
    frmTingDay.List2.listIndex = GeJiA.playWhere - 1
    frmTingDay.List1.listIndex = playList.playWhere - 1
    GeJiA.Index = GeJiA.playWhere
    playList.Index = playList.playWhere
End Sub

Public Sub listShow(ByVal listIndex As Integer)
    Dim mList As list
    If listIndex = 0 Then Exit Sub
    mList = listA(listIndex)
    If mList.Count = 0 And mList.path <> "" Then
        With mList
            If listIndex > 0 Then openList listA(listIndex), .path
            mList = listA(listIndex)
        End With
    End If
    openList mList, mList.path
    refreshList mList
End Sub

Public Sub updataUI() '更新界面
zplay_GetStatus mTingDay, mStatus
If playList.playWhere = 0 Then Exit Sub
With frmTingDay
    If mStatus.fPlay = 1 Then
        .Label1.Caption = SongA(playList.playWhere).title
        .Label3 = SeccondToMin(totalLength)
        .Label4.Caption = ";"
    Else
        .Label1.Caption = "TingDay"
        .Label3 = SeccondToMin(0)
        .Label4.Caption = "4"
    End If
    mPlayStyle = CInt(GetSetting(App.EXEName, "设置", "播放模式", CStr(s随机播放)))
    .Label10.Caption = NowPlayStyle
End With
If mStatus.fPlay = 1 Then
    frmMenu.停播.Caption = "暂停"
    SNIcon_Refresh frmMenu.Icon, frmMenu.hWnd, "正在播放：" & SongA(playList.playWhere).title
Else
    frmMenu.停播.Caption = "播放"
    SNIcon_Refresh frmMenu.Icon, frmMenu.hWnd, "TingDay"
End If
End Sub

'刷新歌集
Public Sub refreshList(ByRef mList As list)
    Dim i As Integer
    frmTingDay.List1.Clear
    If GeJiA.Count = 0 Then Exit Sub
    If mList.Count = 0 Then Exit Sub
    With mList
        For i = 1 To .Count
             If SongA(i).Singer = "" Then
                    frmTingDay.List1.AddItem SongA(i).title
            Else
                    frmTingDay.List1.AddItem SongA(i).Singer & "-" & SongA(i).title
            End If
        Next i
        If .Index > 0 Then
            If .Count >= .Index Then
                frmTingDay.List1.listIndex = .Index - 1
            End If
        End If
    End With
End Sub

'刷新歌单
Public Sub refreshGeJi()
On Error Resume Next
Dim i As Integer
    With frmTingDay
        .List2.Clear
        If GeJiA.Count = 0 Then Exit Sub
        For i = 1 To GeJiA.Count
            .List2.AddItem listA(i).Name
        Next i
        .List2.listIndex = GeJiA.Index - 1
    End With
End Sub

Public Sub hit(ByRef mForm As RECT, ByVal mForm2 As Long)
    Dim mCursor As POINTAPI '获取鼠标的位置
    '检查是否应该缩到顶部
    GetCursorPos mCursor '获取位置
    If mCursor.X > mForm.Left And mCursor.X < mForm.Right And mCursor.Y > mForm.Top - 5 And mCursor.Y < mForm.Bottom Then
        '鼠标在窗体里面
        Select Case mForm2
        Case 1
            If frmTingDay.Top < 0 Then
                frmTingDay.Move frmTingDay.Left, 0
                setTop frmTingDay.hWnd
            End If
            
        Case 2
            If frmLRC.Top < 0 Then
                frmLRC.Move frmLRC.Left, 0
                setTop frmLRC.hWnd
            End If
        End Select
    Else
        Select Case mForm2
        Case 1
            If frmTingDay.Top = 0 Then frmTingDay.Move frmTingDay.Left, -frmTingDay.Height + 50
        Case 2
            With frmLRC
                If .Top = 0 Then .Move .Left, -.Height + 50
            End With
        End Select
    End If
End Sub

Public Sub setTop(ByVal hWnd As Long)
'窗体置顶
SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
End Sub

Public Sub setBackColor(ByVal RGB As String)
Dim Color As Long
Color = RGBtoLong(RGB)
frmTingDay.BackColor = Color
frmLRC.BackColor = Color
frmLRC.Text1.BackColor = Color
If frmMenu.ColorList.Checked = True Then
        frmTingDay.List1.ForeColor = vbWhite
        frmTingDay.List2.ForeColor = vbWhite
        frmTingDay.List1.BackColor = frmTingDay.BackColor
        frmTingDay.List2.BackColor = frmTingDay.BackColor
End If
SaveSetting App.EXEName, "设置", "背景色", Color

End Sub
