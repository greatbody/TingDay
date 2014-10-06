Attribute VB_Name = "TingDayPlayer"
Option Explicit
Public mTingDay As Long '播放器
Public totalLength As Long '总时长
Public curLength As Long '当前时长
Public mStatus As TStreamStatus '播放状态
Public mPlayStyle As PlayStyle '播放模式
Public mVolume As Long '当前的音量

Public Enum PlayStyle
    s单曲播放 = 0
    s随机播放 = 1
    s顺序播放 = 2
    s循环播放 = 3
End Enum

'装载播放器
Public Sub InitmTingDay()
    mTingDay = zplay_CreateZPlay
    totalLength = 1
    curLength = 0
    setVolume mVolume, mVolume '设置音量
End Sub
'显示播放状态
Public Function NowPlayStyle() As String
    Select Case mPlayStyle
    Case s单曲播放
        NowPlayStyle = "单"
    Case s顺序播放
        NowPlayStyle = "顺"
    Case s随机播放
        NowPlayStyle = "随"
    Case s循环播放
        NowPlayStyle = "循"
    End Select
End Function
'停止播放
Public Sub StopPlay()
    zplay_Stop mTingDay
    zplay_Close mTingDay
    playList.isPlay = False
End Sub

Public Sub setVolume(ByVal leftVolume As Long, ByVal rightVolume As Long)
    zplay_SetPlayerVolume mTingDay, leftVolume, rightVolume
    SaveSetting App.EXEName, "设置", "音量", rightVolume
    mVolume = rightVolume
    showVolume
End Sub
'播放完毕后
Public Sub playOver()
    Select Case mPlayStyle
    Case s单曲播放
        '停止播放
        zplay_Stop mTingDay
    Case s随机播放
        '随机播放
        Rndplay
    Case s顺序播放
        playNext
    Case s循环播放
        If playList.Count = playList.playWhere Then
            playList.playWhere = 1
            listPlay playList, playList.playWhere
        Else
            playNext
        End If
    End Select
End Sub
'播放下一曲
Public Sub playNext()
With playList
    If .playWhere = .Count Then Exit Sub
    .playWhere = .playWhere + 1
    listPlay playList, .playWhere
End With
定位歌曲
End Sub
'播放上一曲
Public Sub playPre()
With playList
    If .playWhere = 1 Then Exit Sub
    .playWhere = .playWhere - 1
    listPlay playList, .playWhere
End With
定位歌曲
End Sub
''随机播放
Public Sub Rndplay()
    Dim i As Integer
    Randomize
    i = Rnd * (playList.Count - 1)
    listPlay playList, i
    定位歌曲
End Sub
'打开mp3/m4a等音频文件
Public Sub mp3Open(ByVal FileName As String)
Dim mSong As Song
Dim mList As list
If GeJiA.Count = 0 Then
    With mList
        .isPlay = True
        .Name = "默认列表"
        .Save = False
        .playWhere = 1
        .Index = 1
    End With
    addList GeJiA, mList
End If

With mSong
    .Islocal = True
    .url = FileName
    .Title = delTuo(.url, True)
End With
addsong listA(GeJiA.Index), mSong
listShow GeJiA.Index
listPlay listA(GeJiA.Index), listA(GeJiA.Index).Count
updataUI
定位歌曲
End Sub
'打开自定义列表yyu /tdl
Public Sub tdlOpen(ByVal FileName As String)
    Dim mList As list
    With mList
        .path = FileName
        .Index = 1
        openList mList, .path
    End With
    addList GeJiA, mList
    saveGeJi GeJiPath
    GeJiA.Index = GeJiA.Count
    playList = mList
End Sub

'播放文件
Public Sub FileOpen(ByVal Value As String)
    Dim strPath As String '文件的路径
    Dim strTuo As String '拓展名
    Dim Temp As String
    Dim a As Variant
    Dim url As String
    Dim mSong As Song
    Dim mList As list
    If Value = "" Then Exit Sub
    If Right$(Value, 1) = Chr(34) Then
        strPath = Mid(Value, 2, Len(Value) - 2)
    Else
        strPath = Value
    End If
    strTuo = getTuo(strPath)
   If exitFile(strPath) = False Then Exit Sub
        Select Case strTuo
        Case ".tdl"
            tdlOpen strPath
        Case ".yyu"
            tdlOpen strPath
        Case ".mp3"
            mp3Open strPath
        Case ".mp2"
            mp3Open strPath
        Case ".mp2"
            mp3Open strPath
        Case ".mp1"
            mp3Open strPath
        Case ".ogg"
            mp3Open strPath
        Case ".flac"
            mp3Open strPath
        Case ".aac"
            mp3Open strPath
        Case ".ac3"
            mp3Open strPath
        Case ".oga"
            mp3Open strPath
        Case ".wav"
            mp3Open strPath
        Case ".pcm"
            mp3Open strPath
        Case Else
            MessageBox frmMenu.hWnd, "还未支持该类型文件", "提示", vbOKOnly
            Exit Sub
        End Select
End Sub

'播放列表中特定的曲目
Public Sub listPlay(ByRef mList As list, ByVal Index As Integer)
    If mList.Count = 0 Then Exit Sub
    If Index = 0 Then Exit Sub
    With mList
            .playWhere = Index
            GeJiA.playWhere = GeJiA.Index
            SaveSetting App.EXEName, "设置", "playWhere", mList.playWhere
            SaveSetting App.EXEName, "设置", "listIndex", GeJiA.playWhere
            If SongA(Index).Islocal = True Then
                playMP3 SongA(Index).url
            Else
                MessageBox frmMenu.hWnd, "我去！这是一首网络歌曲么？", "提示", vbCritical
                MessageBox frmMenu.hWnd, "什么！真的是网络歌曲？难道你不知道你大爷我最近心情不好不播网络歌曲了吗？你还是洗洗睡吧！", "提示", vbCritical
                Exit Sub
                If SongA(Index).Title <> "" And SongA(Index).url = "" Then
                    'SongA(index).url = GetSongforBaidu(SongA(index).title, SongA(index).Singer, True)
                    If SongA(Index).url = "" Then Exit Sub
                End If
                playNET SongA(Index).url
            End If
    .isPlay = True
    End With
    playList = mList
End Sub

Public Sub seekTime(ByRef mTime As TStreamTime, ByVal nMoveMethod As TSeekMethod)
    Select Case nMoveMethod
    Case smFromCurrentForward '前移
        zplay_Seek mTingDay, tfSecond, mTime, smFromCurrentForward
    Case smFromBeginning
        zplay_Seek mTingDay, tfSecond, mTime, smFromBeginning
    Case smFromCurrentBackward
        zplay_Seek mTingDay, tfSecond, mTime, smFromCurrentBackward
    Case smFromEnd
        zplay_Seek mTingDay, tfSecond, mTime, smFromEnd
    End Select
End Sub

Public Sub DestroymTingDay()
    zplay_DestroyZPlay mTingDay
End Sub

Public Sub EndmTingDay()
    zplay_Stop mTingDay '停止播放器
    zplay_Close mTingDay
    DestroymTingDay '销毁播放器
    SNIcon_Del  '销毁托盘图标
    SetWindowLong frmMenu.hWnd, GWL_WNDPROC, preProc '交还全局钩子
    UnRegHot '取消热键注册
    End
End Sub

'播放本地歌曲
Public Sub playMP3(ByVal FileName As String)
    Dim sInfo As TStreamInfo '记录流信息
    Dim info As TID3Info '歌曲信息
    Dim lrcPath As String '歌词地址
    Dim Message As Long '打开文件返回的信息
    Dim i As Integer
    zplay_Stop mTingDay '停止当前歌曲
    zplay_Close mTingDay '关闭文件
    DestroymTingDay '销毁播放器
    InitmTingDay '载入播放器
    If exitFile(FileName) = False Then Exit Sub
    zplay_OpenFile mTingDay, FileName, sfAutodetect
    'If Message <> 1 Then MessageBox frmMenu.hWnd, "不得鸟在飞行的过程中被洲际导弹打到了，但是没事！", "提示", vbCritical
    zplay_Play mTingDay
    zplay_GetStreamInfo mTingDay, sInfo
    totalLength = sInfo.length.sec
    If totalLength = 0 Then totalLength = 1
    '装载歌词
    lrcPath = Left(SongA(playList.playWhere).url, Len(SongA(playList.playWhere).url) - 4) & ".lrc"
    If exitFile(lrcPath) = False Then
        'dowloadlrc SongA(playList.playWhere)
        If exitFile(App.path & "\歌词\" & SongA(playList.playWhere).Title & ".lrc") = True Then
            lrcPath = App.path & "\歌词\" & SongA(playList.playWhere).Title & ".lrc"
            SongA(playList.playWhere).lrcPath = lrcPath
        End If
    End If
    For i = 1 To 10
        frmLRC.Label2(i).Caption = ""
    Next i
        ReDim Lyric(1) As LRC

    LyricsAnalyse lrcPath
End Sub

'播放网络歌曲
Public Sub playNET(ByVal url As String)
    Dim sInfo As TStreamInfo '记录流信息
    Dim info As TID3Info '歌曲信息
    Dim lrcPath As String '歌词地址
    Dim streamSize As Long
    If Not zplay_OpenStream(mTingDay, False, False, url, streamSize, sfMp3) Then
        MessageBox frmMenu.hWnd, "哥伦布在回家的时候摔了一跤，结果死掉了！", "提示", vbCritical
        Exit Sub
    End If
    zplay_Play mTingDay
    zplay_GetStreamInfo mTingDay, sInfo
    totalLength = sInfo.length.sec '获取总时长
    If totalLength = 0 Then totalLength = 1
    zplay_LoadID3 mTingDay, id3Version2, info
    zplay_Play mTingDay
End Sub


