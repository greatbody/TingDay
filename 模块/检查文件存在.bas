Attribute VB_Name = "����ļ�����"
Option Explicit

'����ļ��Ƿ����
Public Function exitFile(ByVal FileName As String) As Boolean
Dim fso As New FileSystemObject
exitFile = fso.FileExists(FileName)
End Function

'����ļ����Ƿ����
Public Function exitFolder(ByVal Folder As String) As Boolean
    Dim fso As New FileSystemObject
    If fso.FolderExists(Folder) = True Then
        exitFolder = True
    Else
        exitFolder = False
    End If
End Function

'�����ļ���
Public Function buildFolder(ByVal FolderPath As String) As Boolean
    Dim fso As New FileSystemObject
    If exitFolder(FolderPath) = True Then Exit Function '�������˳�����
    fso.CreateFolder FolderPath
    buildFolder = True
End Function


