Attribute VB_Name = "tdl"
Option Explicit
Public GeJiA As GeJi '歌集
Public playList As list  '播放的歌集
Public listA(1 To 20) As list  '歌集列表
Public SongA(1 To 800) As Song '歌曲列表
Public GeJiPath As String 'listrecord 的位置
'定义歌曲
Public Type Song
    title As String '歌曲名
    url As String '歌曲地址
    Singer As String '歌手
    Ablum As String '所在专辑
    Islocal As Boolean '本地歌曲
    lrcPath As String '歌词地址
    love As Long  '喜爱
End Type
'列表
Public Type list
    Count As Integer  '该歌集的歌曲数，0表示该歌集不存在
    Name As String '该歌集的名称,空表示该歌集不存在
    path As String  '该歌集所在的地址
    Save As Boolean '该歌集是否已经保存
    index As Integer
    isPlay As Boolean '该歌集是否在播放
    playWhere As Integer '该歌集现在正播放第几条歌，如果歌集没有被播放，则为0
End Type
'定义歌集
Public Type GeJi
    playWhere As Integer
    index As Integer
    Count As Integer
    FileName As String
    Caption As String
    Save As Boolean
End Type
'添加歌曲
Public Sub addsong(ByRef mList As list, ByRef mSong As Song)
mList.Count = mList.Count + 1
SongA(mList.Count) = mSong
saveList mList, mList.path
End Sub

'添加歌单
Public Sub addList(ByRef mGeJi As GeJi, ByRef mList As list)
mGeJi.Count = mGeJi.Count + 1
listA(mGeJi.Count) = mList
End Sub

'打开歌单
Public Sub openGeJi(ByVal FileName As String)
Dim mName  As String
Dim mFileName As String
Dim Free As Integer
Dim Temp As String
If exitFile(FileName) = False Then Exit Sub
Free = FreeFile()
If FileLen(FileName) = 0 Then Exit Sub
Open FileName For Input As #Free
    Line Input #Free, Temp
    If Temp <> "#EXTGEJI" Then
        Close #Free
        Exit Sub
    End If
    Do While EOF(Free) = False
        Line Input #Free, mName
        Line Input #Free, mFileName
        If InStr(mName, "#EXTINF:") > 0 And exitFile(mFileName) = True Then
            GeJiA.Count = GeJiA.Count + 1
            With listA(GeJiA.Count)
                .Name = Right$(mName, Len(mName) - Len("#EXTINF:"))
                .path = mFileName
                .Save = True
            End With
        End If
    Loop
Close #Free
End Sub


'打开歌集文件
Public Sub openList(ByRef mList As list, ByVal FileName As String)
    Dim Temp As String
    Dim a As Variant
    Dim Free As Integer
    Dim i As Integer
    If exitFile(FileName) = False Then Exit Sub
    Free = FreeFile()
    With mList
        .path = FileName
        .Count = 0
        Close
        Open FileName For Input As #Free
            Do While EOF(Free) = False
                Line Input #Free, Temp
                If Temp = "" Then GoTo Ne
                If InStr(Temp, "|") = 0 Then
                    .Name = Temp
                    GoTo Ne
                End If
               .Count = .Count + 1
                a = Split(Temp, "|")
                SongA(.Count).Singer = a(0)
                SongA(.Count).title = a(1)
                If SongA(.Count).title = "" Then
                    .Count = .Count - 1
                    GoTo Ne
                End If
                If UBound(a) = 2 Then '为本地歌曲
                    SongA(.Count).Islocal = True
                    SongA(.Count).url = a(2)
                    If SongA(.Count).url = "" Then
                        .Count = .Count - 1
                        GoTo Ne
                    End If
                End If
Ne:
            Loop
        Close #Free
        .Save = True
    End With
End Sub
'保存歌单
Public Sub saveGeJi(ByVal FileName As String)
    Dim Free As Integer
    Dim i As Integer
    Free = FreeFile()
    Open FileName For Output As #Free
        Print #Free, "#EXTGEJI"
        For i = 1 To GeJiA.Count
            If listA(i).path <> "" Then
                Print #Free, "#EXTINF:" & listA(i).Name & vbCrLf & listA(i).path
            End If
        Next i
    Close #Free
End Sub
'保存歌集
Public Function saveList(ByRef mList As list, ByVal FileName As String) As Boolean
Dim Free As Integer
Dim i As Integer
Dim Alist As String
If FileName = "" Then Exit Function
Free = FreeFile
With mList
    If .Name = "" Then .Name = InputBox("请为你这张专辑起一个好听的名字吧！", "专辑名称", .Name)
    If .Name = "" Then Exit Function
    .path = FileName
    Open .path For Output As #Free '专辑名称
        Print #Free, .Name
    Close #Free
    For i = 1 To .Count
        Open .path For Append As #Free
            If SongA(i).Islocal = True Then
                If SongA(i).title = "" And TrimA(SongA(i).url) <> "" Then
                    SongA(i).title = delTuo(TrimA(SongA(i).url), True)
                End If
                Print #Free, SongA(i).Singer & "|" & SongA(i).title & "|" & TrimA(SongA(i).url)
            Else
                Print #Free, SongA(i).Singer & "|" & SongA(i).title
            End If
        Close #Free
    Next i
    .Save = True
End With
End Function
'删除曲目
Public Sub deleteSong(ByRef mList As list, ByVal index As Integer)
Dim i As Integer
    With mList
        If index = .Count Then
            .Count = .Count - 1
            Exit Sub
        End If
        For i = index To .Count
            SongA(i) = SongA(i + 1)
        Next i
        .Count = .Count - 1
    End With
    saveList mList, mList.path
End Sub

'删除歌单
Public Sub deleteList(ByRef mGeJi As GeJi, ByVal index As Integer)
Dim i As Integer
    With mGeJi
        If index = .Count Then
            .Count = .Count - 1
            Exit Sub
        End If
        For i = index To .Count
            listA(i) = listA(i + 1)
        Next i
        .Count = .Count - 1
    End With
    saveGeJi GeJiPath
End Sub

