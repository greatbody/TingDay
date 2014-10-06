Attribute VB_Name = "TingDayPlayer"
Option Explicit
Public mTingDay As Long '������
Public totalLength As Long '��ʱ��
Public curLength As Long '��ǰʱ��
Public mStatus As TStreamStatus '����״̬
Public mPlayStyle As PlayStyle '����ģʽ
Public mVolume As Long '��ǰ������

Public Enum PlayStyle
    s�������� = 0
    s������� = 1
    s˳�򲥷� = 2
    sѭ������ = 3
End Enum

'װ�ز�����
Public Sub InitmTingDay()
    mTingDay = zplay_CreateZPlay
    totalLength = 1
    curLength = 0
    setVolume mVolume, mVolume '��������
End Sub
'��ʾ����״̬
Public Function NowPlayStyle() As String
    Select Case mPlayStyle
    Case s��������
        NowPlayStyle = "��"
    Case s˳�򲥷�
        NowPlayStyle = "˳"
    Case s�������
        NowPlayStyle = "��"
    Case sѭ������
        NowPlayStyle = "ѭ"
    End Select
End Function
'ֹͣ����
Public Sub StopPlay()
    zplay_Stop mTingDay
    zplay_Close mTingDay
    playList.isPlay = False
End Sub

Public Sub setVolume(ByVal leftVolume As Long, ByVal rightVolume As Long)
    zplay_SetPlayerVolume mTingDay, leftVolume, rightVolume
    SaveSetting App.EXEName, "����", "����", rightVolume
    mVolume = rightVolume
    showVolume
End Sub
'������Ϻ�
Public Sub playOver()
    Select Case mPlayStyle
    Case s��������
        'ֹͣ����
        zplay_Stop mTingDay
    Case s�������
        '�������
        Rndplay
    Case s˳�򲥷�
        playNext
    Case sѭ������
        If playList.Count = playList.playWhere Then
            playList.playWhere = 1
            listPlay playList, playList.playWhere
        Else
            playNext
        End If
    End Select
End Sub
'������һ��
Public Sub playNext()
With playList
    If .playWhere = .Count Then Exit Sub
    .playWhere = .playWhere + 1
    listPlay playList, .playWhere
End With
��λ����
End Sub
'������һ��
Public Sub playPre()
With playList
    If .playWhere = 1 Then Exit Sub
    .playWhere = .playWhere - 1
    listPlay playList, .playWhere
End With
��λ����
End Sub
''�������
Public Sub Rndplay()
    Dim i As Integer
    Randomize
    i = Rnd * (playList.Count - 1)
    listPlay playList, i
    ��λ����
End Sub
'��mp3/m4a����Ƶ�ļ�
Public Sub mp3Open(ByVal FileName As String)
Dim mSong As Song
Dim mList As list
If GeJiA.Count = 0 Then
    With mList
        .isPlay = True
        .Name = "Ĭ���б�"
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
��λ����
End Sub
'���Զ����б�yyu /tdl
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

'�����ļ�
Public Sub FileOpen(ByVal Value As String)
    Dim strPath As String '�ļ���·��
    Dim strTuo As String '��չ��
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
            MessageBox frmMenu.hWnd, "��δ֧�ָ������ļ�", "��ʾ", vbOKOnly
            Exit Sub
        End Select
End Sub

'�����б����ض�����Ŀ
Public Sub listPlay(ByRef mList As list, ByVal Index As Integer)
    If mList.Count = 0 Then Exit Sub
    If Index = 0 Then Exit Sub
    With mList
            .playWhere = Index
            GeJiA.playWhere = GeJiA.Index
            SaveSetting App.EXEName, "����", "playWhere", mList.playWhere
            SaveSetting App.EXEName, "����", "listIndex", GeJiA.playWhere
            If SongA(Index).Islocal = True Then
                playMP3 SongA(Index).url
            Else
                MessageBox frmMenu.hWnd, "��ȥ������һ���������ô��", "��ʾ", vbCritical
                MessageBox frmMenu.hWnd, "ʲô�����������������ѵ��㲻֪�����ү��������鲻�ò���������������㻹��ϴϴ˯�ɣ�", "��ʾ", vbCritical
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
    Case smFromCurrentForward 'ǰ��
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
    zplay_Stop mTingDay 'ֹͣ������
    zplay_Close mTingDay
    DestroymTingDay '���ٲ�����
    SNIcon_Del  '��������ͼ��
    SetWindowLong frmMenu.hWnd, GWL_WNDPROC, preProc '����ȫ�ֹ���
    UnRegHot 'ȡ���ȼ�ע��
    End
End Sub

'���ű��ظ���
Public Sub playMP3(ByVal FileName As String)
    Dim sInfo As TStreamInfo '��¼����Ϣ
    Dim info As TID3Info '������Ϣ
    Dim lrcPath As String '��ʵ�ַ
    Dim Message As Long '���ļ����ص���Ϣ
    Dim i As Integer
    zplay_Stop mTingDay 'ֹͣ��ǰ����
    zplay_Close mTingDay '�ر��ļ�
    DestroymTingDay '���ٲ�����
    InitmTingDay '���벥����
    If exitFile(FileName) = False Then Exit Sub
    zplay_OpenFile mTingDay, FileName, sfAutodetect
    'If Message <> 1 Then MessageBox frmMenu.hWnd, "�������ڷ��еĹ����б��޼ʵ������ˣ�����û�£�", "��ʾ", vbCritical
    zplay_Play mTingDay
    zplay_GetStreamInfo mTingDay, sInfo
    totalLength = sInfo.length.sec
    If totalLength = 0 Then totalLength = 1
    'װ�ظ��
    lrcPath = Left(SongA(playList.playWhere).url, Len(SongA(playList.playWhere).url) - 4) & ".lrc"
    If exitFile(lrcPath) = False Then
        'dowloadlrc SongA(playList.playWhere)
        If exitFile(App.path & "\���\" & SongA(playList.playWhere).Title & ".lrc") = True Then
            lrcPath = App.path & "\���\" & SongA(playList.playWhere).Title & ".lrc"
            SongA(playList.playWhere).lrcPath = lrcPath
        End If
    End If
    For i = 1 To 10
        frmLRC.Label2(i).Caption = ""
    Next i
        ReDim Lyric(1) As LRC

    LyricsAnalyse lrcPath
End Sub

'�����������
Public Sub playNET(ByVal url As String)
    Dim sInfo As TStreamInfo '��¼����Ϣ
    Dim info As TID3Info '������Ϣ
    Dim lrcPath As String '��ʵ�ַ
    Dim streamSize As Long
    If Not zplay_OpenStream(mTingDay, False, False, url, streamSize, sfMp3) Then
        MessageBox frmMenu.hWnd, "���ײ��ڻؼҵ�ʱ��ˤ��һ�ӣ���������ˣ�", "��ʾ", vbCritical
        Exit Sub
    End If
    zplay_Play mTingDay
    zplay_GetStreamInfo mTingDay, sInfo
    totalLength = sInfo.length.sec '��ȡ��ʱ��
    If totalLength = 0 Then totalLength = 1
    zplay_LoadID3 mTingDay, id3Version2, info
    zplay_Play mTingDay
End Sub


