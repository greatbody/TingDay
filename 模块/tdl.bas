Attribute VB_Name = "tdl"
Option Explicit
Public GeJiA As GeJi '�輯
Public playList As list  '���ŵĸ輯
Public listA(1 To 20) As list  '�輯�б�
Public SongA(1 To 800) As Song '�����б�
Public GeJiPath As String 'listrecord ��λ��
'�������
Public Type Song
    title As String '������
    url As String '������ַ
    Singer As String '����
    Ablum As String '����ר��
    Islocal As Boolean '���ظ���
    lrcPath As String '��ʵ�ַ
    love As Long  'ϲ��
End Type
'�б�
Public Type list
    Count As Integer  '�ø輯�ĸ�������0��ʾ�ø輯������
    Name As String '�ø輯������,�ձ�ʾ�ø輯������
    path As String  '�ø輯���ڵĵ�ַ
    Save As Boolean '�ø輯�Ƿ��Ѿ�����
    index As Integer
    isPlay As Boolean '�ø輯�Ƿ��ڲ���
    playWhere As Integer '�ø輯���������ŵڼ����裬����輯û�б����ţ���Ϊ0
End Type
'����輯
Public Type GeJi
    playWhere As Integer
    index As Integer
    Count As Integer
    FileName As String
    Caption As String
    Save As Boolean
End Type
'��Ӹ���
Public Sub addsong(ByRef mList As list, ByRef mSong As Song)
mList.Count = mList.Count + 1
SongA(mList.Count) = mSong
saveList mList, mList.path
End Sub

'��Ӹ赥
Public Sub addList(ByRef mGeJi As GeJi, ByRef mList As list)
mGeJi.Count = mGeJi.Count + 1
listA(mGeJi.Count) = mList
End Sub

'�򿪸赥
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


'�򿪸輯�ļ�
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
                If UBound(a) = 2 Then 'Ϊ���ظ���
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
'����赥
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
'����輯
Public Function saveList(ByRef mList As list, ByVal FileName As String) As Boolean
Dim Free As Integer
Dim i As Integer
Dim Alist As String
If FileName = "" Then Exit Function
Free = FreeFile
With mList
    If .Name = "" Then .Name = InputBox("��Ϊ������ר����һ�����������ְɣ�", "ר������", .Name)
    If .Name = "" Then Exit Function
    .path = FileName
    Open .path For Output As #Free 'ר������
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
'ɾ����Ŀ
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

'ɾ���赥
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

