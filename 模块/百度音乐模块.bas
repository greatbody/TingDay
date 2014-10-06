Attribute VB_Name = "百度音乐模块"
'******************************************************************
'模块用法
'函数为:GetSongforBaidu
'返回值：url
'参数分别为：Title 歌曲名 Singer歌手名 isHide 当没有找到音乐时是否提示
'联系：blog.putaot.cn
'****************************************************************
 '从百度音乐中获得音乐播放地址
Public Function GetSongforBaidu(ByVal title As String, Optional Singer As String, Optional isHide As Boolean) As String
Dim HtmlS As String
Dim url_Temp As String
Dim int_url As Integer '“/的位置”
Dim i As Integer
Dim Singer2 As String
'检查网络是否链接！
If GetNetConnectString = False Then Exit Function
Singer2 = Singer
st:
If Singer2 = "" Then
    url_Temp = "http://box.zhangmen.baidu.com/x?op=12&count=1&title=" & title & "$$$$"
Else
    url_Temp = "http://box.zhangmen.baidu.com/x?op=12&count=1&title=" & title & "$$" & Singer2 & "$$$$"
End If
'虽然URL对应的是一个xml文件，但是为了方便所以直接使用字符串的处理方式处理了
HtmlS = htmlStr(url_Temp)
If InStr(HtmlS, "<count>0</count>") Or HtmlS = "" <> 0 Then
    If Singer2 <> "" Then
        '没有找到该歌曲尝试只寻找歌曲名
        Singer2 = ""
        GoTo st
    Else
        If isHide = False Then MsgBox "这首歌没有找到：" & vbCrLf & "歌曲名：" & title & vbCrLf & "歌手：" & Singer
    End If
    Exit Function
End If
url1 = "http" & cutStr(HtmlS, "http", "]]", 1)
'MsgBox URL1
temp = cutStr(HtmlS, "<decode>", "</decode>", 1)
If temp = "" Then
    GetSongforBaidu = ""
    Exit Function
End If
url2 = cutStr(temp, "CDATA[", "]", 1)
'取得ＵＲＬ
i = 2
Do While InStr(url2, ".mp3") = 0 And url2 <> ""
        url1 = "http" & cutStr(HtmlS, "http", "]]", i)
        'MsgBox URL1
        temp = cutStr(HtmlS, "<decode>", "</decode>", i)
        url2 = cutStr(temp, "CDATA[", "]]", 1)
        i = i + 1
Loop
'2014-01-18 发现百度已经修改了ＡＰＩ，其中有可能是改为.mp3+.mp3的那种形式,下面做出相应修改
If InStr(url1, ".mp3") <> 0 Then
    int_url = InStrRev(url1, "/")
    url1 = Mid(url1, 1, int_url)
End If
GetSongforBaidu = url1 & url2
End Function


