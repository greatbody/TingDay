Attribute VB_Name = "�µ�����ѡ����"
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = &H1
Private pidl As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Function showFolder(ByVal yTitle As String, ByVal hWnd As Long) As String     '���ļ��� ��ѡ��һ���ļ���
Dim bi As BROWSEINFO
Dim r As Long
Dim pidl As Long
Dim path As String
Dim pos As Integer
'���
bi.hOwner = hWnd
'չ����Ŀ¼
bi.pidlRoot = 0&
'�б�����
bi.lpszTitle = yTitle
'�涨ֻ��ѡ���ļ��У�������Ч
bi.ulFlags = BIF_RETURNONLYFSDIRS
'����API������ʾ�б��
pidl = SHBrowseForFolder(bi)
'����API������ȡ���ص�·��
path = Space$(512)
r = SHGetPathFromIDList(ByVal pidl&, ByVal path)
If r Then
    pos = InStr(path, Chr$(0))
    showFolder = Left(path, pos - 1)
Else
    showFolder = ""
End If
End Function


