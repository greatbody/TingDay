Attribute VB_Name = "����ɨ��"
Option Explicit
Public Function seachFiles(ByRef mList As list, ByVal Folder As String, ByVal Filter As String, ByVal justFirst As Boolean) As Long
On Error Resume Next
Dim fso As New FileSystemObject
Dim mFile As String
Dim objFolder As Object
Dim path As Variant
Dim Files As Variant
Dim saveFolder As String
If Folder = "" Then Exit Function
    '����������Ŀ¼���ļ�
    For Each Files In fso.GetFolder(Folder).Files
            mFile = Files
            If Filter <> ".*" Then
                If Right(mFile, 4) = Filter Then
                    '���뵽�б�
                    mList.Count = mList.Count + 1
                    With SongA(mList.Count)
                        .Islocal = True
                        .url = mFile
                        .title = delTuo(.url, True)
                        If mList.Count >= 800 Then Exit For
                    End With
                End If
            End If
    Next
    If justFirst = True Then
        saveFolder = App.path & "\�б�"
        If exitFolder(saveFolder) = False Then buildFolder saveFolder
        saveList mList, saveFolder & "\" & mList.Name & second(Now) & ".tdl"
        MessageBox frmMenu.hWnd, "����һ��ɨ����" & mList.Count & "������", "ɨ��ɹ�", vbOKOnly
        Set fso = Nothing
        Exit Function
    End If
    '������Ŀ¼�µ������ļ���
    Set objFolder = fso.GetFolder(Folder) '����Ҫ�������ļ�
    For Each path In objFolder.SubFolders
        For Each Files In fso.GetFolder(Folder).Files
                mFile = Files
                If Filter <> ".*" Then
                    If Right(mFile, 4) = Filter Then
                        '���뵽�б�
                        mList.Count = mList.Count + 1
                        With SongA(mList.Count)
                            .Islocal = True
                            .url = mFile
                            .title = delTuo(.url, True)
                            If mList.Count >= 800 Then Exit For
                        End With
                    End If
                End If
        Next
    Next
    Set fso = Nothing
End Function

