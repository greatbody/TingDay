Attribute VB_Name = "本地扫描"
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
    '事先搜索本目录的文件
    For Each Files In fso.GetFolder(Folder).Files
            mFile = Files
            If Filter <> ".*" Then
                If Right(mFile, 4) = Filter Then
                    '加入到列表
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
        saveFolder = App.path & "\列表"
        If exitFolder(saveFolder) = False Then buildFolder saveFolder
        saveList mList, saveFolder & "\" & mList.Name & second(Now) & ".tdl"
        MessageBox frmMenu.hWnd, "本次一共扫描了" & mList.Count & "首音乐", "扫描成功", vbOKOnly
        Set fso = Nothing
        Exit Function
    End If
    '搜索本目录下的其它文件夹
    Set objFolder = fso.GetFolder(Folder) '设置要搜索的文件
    For Each path In objFolder.SubFolders
        For Each Files In fso.GetFolder(Folder).Files
                mFile = Files
                If Filter <> ".*" Then
                    If Right(mFile, 4) = Filter Then
                        '加入到列表
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

