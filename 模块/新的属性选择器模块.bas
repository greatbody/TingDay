Attribute VB_Name = "新的属性选择器"
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

Public Function showFolder(ByVal yTitle As String, ByVal hWnd As Long) As String     '打开文件夹 并选择一个文件夹
Dim bi As BROWSEINFO
Dim r As Long
Dim pidl As Long
Dim path As String
Dim pos As Integer
'句柄
bi.hOwner = hWnd
'展开根目录
bi.pidlRoot = 0&
'列表框标题
bi.lpszTitle = yTitle
'规定只能选择文件夹，其他无效
bi.ulFlags = BIF_RETURNONLYFSDIRS
'调用API函数显示列表框
pidl = SHBrowseForFolder(bi)
'利用API函数获取返回的路径
path = Space$(512)
r = SHGetPathFromIDList(ByVal pidl&, ByVal path)
If r Then
    pos = InStr(path, Chr$(0))
    showFolder = Left(path, pos - 1)
Else
    showFolder = ""
End If
End Function


