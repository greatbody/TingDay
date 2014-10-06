Attribute VB_Name = "搜狗音乐搜索"
'******************************************************************
'模块用法
'函数为:sogouMusics
'返回值：GeJi
'参数分别为：keyword 歌曲名/歌手 isHide 找不到是否提示
'联系：blog.putaot.cn
'保留我的链接，或许API会更新，我会做相应修改：http://blog.putaot.cn/?post=98
'****************************************************************
Option Explicit
'利用搜狗音乐搜索
'http://wap.sogou.com/music/musicSearchResult.jsp?mode=1&uID=cpmozaddqef-zyUz&v=2&p=1&w=1104&keyword=[keyword]
Public Function sogouMusics(ByVal KeyWord As String, isHide As Boolean) As GeJi
Dim url As String '默认的URL，也就是搜狗的音乐搜索地址 API
Dim WangYe As String '获得的网页源码，已经编码好了
Dim temp_title As String '歌曲名缓存
Dim temp_singer As String '歌手缓存
Dim i去括号 As Integer '去括号部分代码的变量
Dim n_i As Integer
Dim p As Integer '对应的页码
'检查网络是否链接！
If GetNetConnectString = False Then Exit Function

If KeyWord = "" Then Exit Function
p:
p = p + 1
n_i = 0
url = "http://wap.sogou.com/music/musicSearchResult.jsp?mode=1&uID=cpmozaddqef-zyUz&v=2&p=" & p & "&w=1104&keyword=" & KeyWord
WangYe = htmlStr$(url)
'Debug.Print url
If InStr(WangYe, "抱歉，没有找到您搜索") <> 0 And isHide = False Then
    MsgBox "没有找到相关曲目", vbOKOnly, "歌曲没有找到"
    Exit Function
End If
'从获得的源码中得到歌曲信息
With sogouMusics
Do
    .N = .N + 1
    n_i = n_i + 1
    temp_title = cutStr(WangYe, "-->", "</a>", n_i + 1)
    temp_singer = cutStr(WangYe, "singerDetail", "</a>", n_i)
    '歌手信息处理
     temp_singer = Mid(temp_singer, InStr(temp_singer, Chr(34) & ">") + 2)
    If InStr(temp_singer, ".jsp?") <> 0 Then temp_singer = ""
    If InStr(temp_singer, "<span class='keyword'>") <> 0 Then '去红
        temp_singer = Replace(temp_singer, "<span class='keyword'>", "")
        temp_singer = Replace(temp_singer, "</span>", "")
    End If
    temp_singer = Replace(temp_singer, Chr(13), "")
    '歌曲名信息处理
    temp_title = Replace(temp_title, Chr(9), "")
    If InStr(temp_title, "<span class='keyword'>") <> 0 Then '去红
        temp_title = Replace(temp_title, "<span class='keyword'>", "")
        temp_title = Replace(temp_title, "</span>", "")
    End If
    temp_title = Replace(temp_title, Chr(13), "")
    temp_title = Replace(temp_title, Chr(10), "")
    '去掉括号里面的东西
    i去括号 = InStr(temp_title, "(")
    If i去括号 <> 0 Then '歌曲名存在括号
        temp_title = Left(temp_title, i去括号) '暂时把括号 “（”后面的全部删掉
    End If
    '把信息添加到歌集中
    .Gsong(.N).mTitle = temp_title
    .Gsong(.N).mSinger = temp_singer
    If InStr(.Gsong(.N).mTitle, "喜欢这首歌的用户") <> 0 Or InStr(.Gsong(.N).mTitle, "<br>") <> 0 Or n_i > 50 Then
        .N = .N - 1
        Exit Do
    End If
    If Len(.Gsong(.N).mTitle) > 30 Or .Gsong(.N).mTitle = vbCrLf Or Len(.Gsong(.N).mTitle) = 0 Then
        '去掉出错的歌曲
        .N = .N - 1
    End If
Loop
If InStr(WangYe, "下一页") <> 0 Then GoTo p
.NameG = KeyWord
.SaveG = False
.hUrl = False
End With
End Function


