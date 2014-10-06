Attribute VB_Name = "检查文件存在"
Option Explicit

'检查文件是否存在
Public Function exitFile(ByVal FileName As String) As Boolean
Dim fso As New FileSystemObject
exitFile = fso.FileExists(FileName)
End Function

'检查文件夹是否存在
Public Function exitFolder(ByVal Folder As String) As Boolean
    Dim fso As New FileSystemObject
    If fso.FolderExists(Folder) = True Then
        exitFolder = True
    Else
        exitFolder = False
    End If
End Function

'创建文件夹
Public Function buildFolder(ByVal FolderPath As String) As Boolean
    Dim fso As New FileSystemObject
    If exitFolder(FolderPath) = True Then Exit Function '存在则退出函数
    fso.CreateFolder FolderPath
    buildFolder = True
End Function


