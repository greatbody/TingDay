Attribute VB_Name = "�ѹ���������"
'******************************************************************
'ģ���÷�
'����Ϊ:sogouMusics
'����ֵ��GeJi
'�����ֱ�Ϊ��keyword ������/���� isHide �Ҳ����Ƿ���ʾ
'��ϵ��blog.putaot.cn
'�����ҵ����ӣ�����API����£��һ�����Ӧ�޸ģ�http://blog.putaot.cn/?post=98
'****************************************************************
Option Explicit
'�����ѹ���������
'http://wap.sogou.com/music/musicSearchResult.jsp?mode=1&uID=cpmozaddqef-zyUz&v=2&p=1&w=1104&keyword=[keyword]
Public Function sogouMusics(ByVal KeyWord As String, isHide As Boolean) As GeJi
Dim url As String 'Ĭ�ϵ�URL��Ҳ�����ѹ�������������ַ API
Dim WangYe As String '��õ���ҳԴ�룬�Ѿ��������
Dim temp_title As String '����������
Dim temp_singer As String '���ֻ���
Dim iȥ���� As Integer 'ȥ���Ų��ִ���ı���
Dim n_i As Integer
Dim p As Integer '��Ӧ��ҳ��
'��������Ƿ����ӣ�
If GetNetConnectString = False Then Exit Function

If KeyWord = "" Then Exit Function
p:
p = p + 1
n_i = 0
url = "http://wap.sogou.com/music/musicSearchResult.jsp?mode=1&uID=cpmozaddqef-zyUz&v=2&p=" & p & "&w=1104&keyword=" & KeyWord
WangYe = htmlStr$(url)
'Debug.Print url
If InStr(WangYe, "��Ǹ��û���ҵ�������") <> 0 And isHide = False Then
    MsgBox "û���ҵ������Ŀ", vbOKOnly, "����û���ҵ�"
    Exit Function
End If
'�ӻ�õ�Դ���еõ�������Ϣ
With sogouMusics
Do
    .N = .N + 1
    n_i = n_i + 1
    temp_title = cutStr(WangYe, "-->", "</a>", n_i + 1)
    temp_singer = cutStr(WangYe, "singerDetail", "</a>", n_i)
    '������Ϣ����
     temp_singer = Mid(temp_singer, InStr(temp_singer, Chr(34) & ">") + 2)
    If InStr(temp_singer, ".jsp?") <> 0 Then temp_singer = ""
    If InStr(temp_singer, "<span class='keyword'>") <> 0 Then 'ȥ��
        temp_singer = Replace(temp_singer, "<span class='keyword'>", "")
        temp_singer = Replace(temp_singer, "</span>", "")
    End If
    temp_singer = Replace(temp_singer, Chr(13), "")
    '��������Ϣ����
    temp_title = Replace(temp_title, Chr(9), "")
    If InStr(temp_title, "<span class='keyword'>") <> 0 Then 'ȥ��
        temp_title = Replace(temp_title, "<span class='keyword'>", "")
        temp_title = Replace(temp_title, "</span>", "")
    End If
    temp_title = Replace(temp_title, Chr(13), "")
    temp_title = Replace(temp_title, Chr(10), "")
    'ȥ����������Ķ���
    iȥ���� = InStr(temp_title, "(")
    If iȥ���� <> 0 Then '��������������
        temp_title = Left(temp_title, iȥ����) '��ʱ������ �����������ȫ��ɾ��
    End If
    '����Ϣ��ӵ��輯��
    .Gsong(.N).mTitle = temp_title
    .Gsong(.N).mSinger = temp_singer
    If InStr(.Gsong(.N).mTitle, "ϲ�����׸���û�") <> 0 Or InStr(.Gsong(.N).mTitle, "<br>") <> 0 Or n_i > 50 Then
        .N = .N - 1
        Exit Do
    End If
    If Len(.Gsong(.N).mTitle) > 30 Or .Gsong(.N).mTitle = vbCrLf Or Len(.Gsong(.N).mTitle) = 0 Then
        'ȥ������ĸ���
        .N = .N - 1
    End If
Loop
If InStr(WangYe, "��һҳ") <> 0 Then GoTo p
.NameG = KeyWord
.SaveG = False
.hUrl = False
End With
End Function


