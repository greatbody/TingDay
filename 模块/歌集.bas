Attribute VB_Name = "�輯"
'�������
Private Type Song
    mTitle As String '������
    mURL As String '������ַ
    mSinger As String '����
    mAblum As String '����ר��
    mIslocal As Boolean '���ظ���
End Type
'����輯
Public Type GeJi
    Gsong(1 To 800) As Song  '�ø輯�����и��� ��Ĭ�����Ϊ200���������ɵ�����
    i As Integer     '��ǰ���ڲ����ĸ���������һ��Ϊ������
    N As Integer  '�ø輯�ĸ�������0��ʾ�ø輯������
    NameG As String '�ø輯������,�ձ�ʾ�ø輯������
    pathG As String  '�ø輯���ڵĵ�ַ
    SaveG As Boolean '�ø輯�Ƿ��Ѿ�����
    isPlay As Boolean '�ø輯�Ƿ��ڲ���
    playWhere As Integer '�ø輯���������ŵڼ����裬����輯û�б����ţ���Ϊ0
    hUrl As Boolean '�Ƿ��Ѿ�ȡ��URL
End Type
'�����輯
Public yGeJi As GeJi
 'Ϊ�輯ȡ������URL
Public Sub GetUrlforGeJi(ByRef ugeji As GeJi)
With ugeji
    For i = 1 To .N
    DoEvents
        .Gsong(i).mURL = GetSongforBaidu(.Gsong(i).mTitle, .Gsong(i).mSinger, True)
    Next i
.hUrl = True
End With
End Sub
 '���Ÿ輯��
Public Sub playGeJi(ByRef ugeji As GeJi, ByVal where As Integer)
If ugeji.N = 0 Then Exit Sub
With ugeji
        .playWhere = where
        If .Gsong(.playWhere).mURL = "" Then .Gsong(.playWhere).mURL = GetSongforBaidu(.Gsong(.playWhere).mTitle, .Gsong(.playWhere).mSinger, True)
        frm������.WindowsMediaPlayer1.url = .Gsong(.playWhere).mURL
        frm������.WindowsMediaPlayer1.Controls.play
        .isPlay = True
        Int���Ĳ���λ�� = .playWhere
        SaveSetting "������", "�輯", "���Ĳ���λ��", Int���Ĳ���λ��
End With
frm�����б�.List1.ListIndex = ugeji.playWhere - 1
End Sub
Ϊָ���輯�������һ�׸�
Public Sub Rndplay(ByRef ugeji As GeJi)
    Dim i As Integer
    Randomize
    i = Rnd * (.N - 1) + 1
    Call playGeJi(yGeJi, i)
End Sub

'Ϊ�輯��Ӹ���
Public Sub addsong(ByRef ugeji As GeJi, ByVal uTitle As String, uSinger As String, uUrl As String)
    With ugeji
        .N = .N + 1
        .Gsong(.N).mSinger = uSinger
        .Gsong(.N).mTitle = uTitle
        .Gsong(.N).mURL = uUrl
        .SaveG = False
    End With
End Sub
'Ϊ���ظ輯��ȡ������Ϣ
Public Sub getInfoLocal(ByRef ugeji As GeJi)
With ugeji
    For i = 1 To .N
        If .Gsong(i).mIslocal = True Then 'ȷ��Ϊ���ظ���
            .Gsong(i).mTitle = Right(.Gsong(i).mURL, Len(.Gsong(i).mURL) - InStrRev(.Gsong(i).mURL, "\"))
        End If
    Next i
End With
End Sub
'�輯��ȷ�����ظ��Լ��������
Public Sub checkGeJi(ByRef ugeji As GeJi)
Dim iȥ����1 As Integer
Dim iȥ����2 As Integer

    With ugeji
        'ȥ��/n ��Ϊ����ҳ�����������д�/n���з� ��VBʶ�����ǻ��з�
        For i = 1 To .N
            If Len(.Gsong(i).mTitle) <= 0 Then Call deleteSong(ugeji, i)
            .Gsong(i).mTitle = Replace(.Gsong(i).mTitle, Chr(10), "")
            .Gsong(i).mSinger = Replace(.Gsong(i).mSinger, Chr(10), "")
            .Gsong(i).mTitle = Replace(.Gsong(i).mTitle, Chr(13), "")
            .Gsong(i).mSinger = Replace(.Gsong(i).mSinger, Chr(13), "")
        'ȥ���������е�����
        iȥ���� = InStr(.Gsong(i).mTitle, "(")
        If iȥ���� <> 0 Then '��������������
            .Gsong(i).mTitle = Left(.Gsong(i).mTitle, iȥ���� - 2) '��ʱ������ �����������ȫ��ɾ��
        End If
        Next i
        '����ȥ�أ�
        For i = 1 To .N - 1
            For j = i + 1 To .N
                If .Gsong(j).mTitle = .Gsong(i).mTitle Then  '�����и����ظ���
                    If .Gsong(j).mSinger = .Gsong(i).mSinger Or .Gsong(j).mSinger = "" Then   '�����ظ�
                        Call deleteSong(ugeji, j) 'ɾ���ø���
                    End If
                End If
            Next j
        Next i
    End With
End Sub
'�򿪸輯�ļ�
Public Sub openGeJi(ByRef ugeji As GeJi, uFilePath As String)
    Dim temp As String
    Dim a As Variant
    Dim Free As Integer
    Dim i As Integer
    Dim tGeJi As GeJi '����һ������輯����ʼ������
    If exitFile(uFilePath) = False Then Exit Sub
    Free = FreeFile()
    With tGeJi
        .pathG = uFilePath
        Open uFilePath For Input As #Free
            DoEvents '�ύ��ϵͳ����ֹ����
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
                If UBound(a) >= 2 Then 'Ϊ���ظ���
                    .Gsong(.N).mIslocal = True
                    .Gsong(.N).mURL = a(2)
                End If
Ne:
            Loop
        Close #Free
        .SaveG = True
        .isPlay = False
        .playWhere = 1
        frm������.WindowsMediaPlayer1.url = .Gsong(.playWhere).mURL
    End With
    '�ѻ���輯�������ڵĸ輯
    ugeji = tGeJi
    SaveSetting "������", "�輯", "���ĸ輯��ַ", ugeji.pathG
End Sub
'����輯
Public Sub saveGeJi(ByRef ugeji As GeJi, ByVal path As String)
Dim Free As Integer
Free = FreeFile
With ugeji
    .NameG = InputBox("��Ϊ������ר����һ�����������ְɣ�", "ר������", .NameG)
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
MsgBox "�Ѿ����棡", vbOKOnly, "����"
.SaveG = True
End With
End Sub
'ɾ����Ŀ
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
