Attribute VB_Name = "�ٶ�����ģ��"
'******************************************************************
'ģ���÷�
'����Ϊ:GetSongforBaidu
'����ֵ��url
'�����ֱ�Ϊ��Title ������ Singer������ isHide ��û���ҵ�����ʱ�Ƿ���ʾ
'��ϵ��blog.putaot.cn
'****************************************************************
 '�Ӱٶ������л�����ֲ��ŵ�ַ
Public Function GetSongforBaidu(ByVal title As String, Optional Singer As String, Optional isHide As Boolean) As String
Dim HtmlS As String
Dim url_Temp As String
Dim int_url As Integer '��/��λ�á�
Dim i As Integer
Dim Singer2 As String
'��������Ƿ����ӣ�
If GetNetConnectString = False Then Exit Function
Singer2 = Singer
st:
If Singer2 = "" Then
    url_Temp = "http://box.zhangmen.baidu.com/x?op=12&count=1&title=" & title & "$$$$"
Else
    url_Temp = "http://box.zhangmen.baidu.com/x?op=12&count=1&title=" & title & "$$" & Singer2 & "$$$$"
End If
'��ȻURL��Ӧ����һ��xml�ļ�������Ϊ�˷�������ֱ��ʹ���ַ����Ĵ���ʽ������
HtmlS = htmlStr(url_Temp)
If InStr(HtmlS, "<count>0</count>") Or HtmlS = "" <> 0 Then
    If Singer2 <> "" Then
        'û���ҵ��ø�������ֻѰ�Ҹ�����
        Singer2 = ""
        GoTo st
    Else
        If isHide = False Then MsgBox "���׸�û���ҵ���" & vbCrLf & "��������" & title & vbCrLf & "���֣�" & Singer
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
'ȡ�ãգң�
i = 2
Do While InStr(url2, ".mp3") = 0 And url2 <> ""
        url1 = "http" & cutStr(HtmlS, "http", "]]", i)
        'MsgBox URL1
        temp = cutStr(HtmlS, "<decode>", "</decode>", i)
        url2 = cutStr(temp, "CDATA[", "]]", 1)
        i = i + 1
Loop
'2014-01-18 ���ְٶ��Ѿ��޸��ˣ��Уɣ������п����Ǹ�Ϊ.mp3+.mp3��������ʽ,����������Ӧ�޸�
If InStr(url1, ".mp3") <> 0 Then
    int_url = InStrRev(url1, "/")
    url1 = Mid(url1, 1, int_url)
End If
GetSongforBaidu = url1 & url2
End Function


