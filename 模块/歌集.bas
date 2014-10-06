Attribute VB_Name = "歌集"
'定义歌曲
Private Type Song
    mTitle As String '歌曲名
    mURL As String '歌曲地址
    mSinger As String '歌手
    mAblum As String '所在专辑
    mIslocal As Boolean '本地歌曲
End Type
'定义歌集
Public Type GeJi
    Gsong(1 To 800) As Song  '该歌集的所有歌曲 ，默认最大为200条歌曲（可调整）
    i As Integer     '当前正在操作的歌曲，但不一定为播放中
    N As Integer  '该歌集的歌曲数，0表示该歌集不存在
    NameG As String '该歌集的名称,空表示该歌集不存在
    pathG As String  '该歌集所在的地址
    SaveG As Boolean '该歌集是否已经保存
    isPlay As Boolean '该歌集是否在播放
    playWhere As Integer '该歌集现在正播放第几条歌，如果歌集没有被播放，则为0
    hUrl As Boolean '是否已经取得URL
End Type
'公共歌集
Public yGeJi As GeJi
 '为歌集取得所有URL
Public Sub GetUrlforGeJi(ByRef ugeji As GeJi)
With ugeji
    For i = 1 To .N
    DoEvents
        .Gsong(i).mURL = GetSongforBaidu(.Gsong(i).mTitle, .Gsong(i).mSinger, True)
    Next i
.hUrl = True
End With
End Sub
 '播放歌集！
Public Sub playGeJi(ByRef ugeji As GeJi, ByVal where As Integer)
If ugeji.N = 0 Then Exit Sub
With ugeji
        .playWhere = where
        If .Gsong(.playWhere).mURL = "" Then .Gsong(.playWhere).mURL = GetSongforBaidu(.Gsong(.playWhere).mTitle, .Gsong(.playWhere).mSinger, True)
        frm主窗体.WindowsMediaPlayer1.url = .Gsong(.playWhere).mURL
        frm主窗体.WindowsMediaPlayer1.Controls.play
        .isPlay = True
        Int最后的播放位置 = .playWhere
        SaveSetting "随心听", "歌集", "最后的播放位置", Int最后的播放位置
End With
frm歌曲列表.List1.ListIndex = ugeji.playWhere - 1
End Sub
为指定歌集随机播放一首歌
Public Sub Rndplay(ByRef ugeji As GeJi)
    Dim i As Integer
    Randomize
    i = Rnd * (.N - 1) + 1
    Call playGeJi(yGeJi, i)
End Sub

'为歌集添加歌曲
Public Sub addsong(ByRef ugeji As GeJi, ByVal uTitle As String, uSinger As String, uUrl As String)
    With ugeji
        .N = .N + 1
        .Gsong(.N).mSinger = uSinger
        .Gsong(.N).mTitle = uTitle
        .Gsong(.N).mURL = uUrl
        .SaveG = False
    End With
End Sub
'为本地歌集提取歌曲信息
Public Sub getInfoLocal(ByRef ugeji As GeJi)
With ugeji
    For i = 1 To .N
        If .Gsong(i).mIslocal = True Then '确定为本地歌曲
            .Gsong(i).mTitle = Right(.Gsong(i).mURL, Len(.Gsong(i).mURL) - InStrRev(.Gsong(i).mURL, "\"))
        End If
    Next i
End With
End Sub
'歌集正确性与重复性检查与修正
Public Sub checkGeJi(ByRef ugeji As GeJi)
Dim i去括号1 As Integer
Dim i去括号2 As Integer

    With ugeji
        '去掉/n 因为从网页得来的名称有带/n换行符 但VB识别不了是换行符
        For i = 1 To .N
            If Len(.Gsong(i).mTitle) <= 0 Then Call deleteSong(ugeji, i)
            .Gsong(i).mTitle = Replace(.Gsong(i).mTitle, Chr(10), "")
            .Gsong(i).mSinger = Replace(.Gsong(i).mSinger, Chr(10), "")
            .Gsong(i).mTitle = Replace(.Gsong(i).mTitle, Chr(13), "")
            .Gsong(i).mSinger = Replace(.Gsong(i).mSinger, Chr(13), "")
        '去掉歌曲名中的括号
        i去括号 = InStr(.Gsong(i).mTitle, "(")
        If i去括号 <> 0 Then '歌曲名存在括号
            .Gsong(i).mTitle = Left(.Gsong(i).mTitle, i去括号 - 2) '暂时把括号 “（”后面的全部删掉
        End If
        Next i
        '歌曲去重！
        For i = 1 To .N - 1
            For j = i + 1 To .N
                If .Gsong(j).mTitle = .Gsong(i).mTitle Then  '发现有歌名重复的
                    If .Gsong(j).mSinger = .Gsong(i).mSinger Or .Gsong(j).mSinger = "" Then   '歌手重复
                        Call deleteSong(ugeji, j) '删除该歌曲
                    End If
                End If
            Next j
        Next i
    End With
End Sub
'打开歌集文件
Public Sub openGeJi(ByRef ugeji As GeJi, uFilePath As String)
    Dim temp As String
    Dim a As Variant
    Dim Free As Integer
    Dim i As Integer
    Dim tGeJi As GeJi '创建一个缓存歌集来初始化数据
    If exitFile(uFilePath) = False Then Exit Sub
    Free = FreeFile()
    With tGeJi
        .pathG = uFilePath
        Open uFilePath For Input As #Free
            DoEvents '提交给系统，防止假死
            Do While EOF(Free) = False
                Line Input #Free, temp
                If temp = "" Then GoTo Ne
                If InStr(temp, "|") = 0 Then
                    .NameG = temp
                    GoTo Ne
                End If
                .N = .N + 1
                a = Split(temp, "|")
                .Gsong(.N).mSinger = a(0)
                .Gsong(.N).mTitle = a(1)
                If UBound(a) >= 2 Then '为本地歌曲
                    .Gsong(.N).mIslocal = True
                    .Gsong(.N).mURL = a(2)
                End If
Ne:
            Loop
        Close #Free
        .SaveG = True
        .isPlay = False
        .playWhere = 1
        frm主窗体.WindowsMediaPlayer1.url = .Gsong(.playWhere).mURL
    End With
    '把缓存歌集过给存在的歌集
    ugeji = tGeJi
    SaveSetting "随心听", "歌集", "最后的歌集地址", ugeji.pathG
End Sub
'保存歌集
Public Sub saveGeJi(ByRef ugeji As GeJi, ByVal path As String)
Dim Free As Integer
Free = FreeFile
With ugeji
    .NameG = InputBox("请为你这张专辑起一个好听的名字吧！", "专辑名称", .NameG)
    If .NameG = "" Then Exit Sub
    .pathG = path
    If .pathG = "" Then Exit Sub
    Open .pathG For Output As #Free
        Print #Free, .NameG
    Close #Free
    For i = 1 To .N
        Open .pathG For Append As #Free
            If .Gsong(i).mIslocal = True Then
                Print #Free, .Gsong(i).mSinger & "|" & .Gsong(i).mTitle & "|" & .Gsong(i).mURL
            Else
                Print #Free, .Gsong(i).mSinger & "|" & .Gsong(i).mTitle
            End If
        Close #Free
Next i
MsgBox "已经保存！", vbOKOnly, "保存"
.SaveG = True
End With
End Sub
'删除曲目
Public Sub deleteSong(ByRef ugeji As GeJi, ByVal pos_i As Integer)
    With ugeji
        For i = pos_i To .N
            .Gsong(i).mAblum = .Gsong(i + 1).mAblum
            .Gsong(i).mSinger = .Gsong(i + 1).mSinger
            .Gsong(i).mTitle = .Gsong(i + 1).mTitle
            .Gsong(i).mURL = .Gsong(i + 1).mURL
        Next i
        If .playWhere = .N Then .playWhere = .N - 1
        .N = .N - 1
        .SaveG = False
    End With
End Sub
