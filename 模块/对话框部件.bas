Attribute VB_Name = "�Ի��򲿼�"
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
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type
Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type
'���ļ��Ի���
Public Function ShowOpen(ByVal yTitle As String, ByVal Filter As String, ByVal hwnd As Long) As String
    Dim i As Integer
    Dim OpenFolder As OPENFILENAME
    Dim FileName As String
    With OpenFolder
        .lStructSize = Len(OpenFolder)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        '.lpstrInitialDir = App.path
        .flags = 6148
        '�����ļ�����
        .lpstrFilter = Filter
        '�Ի�������
        .lpstrTitle = yTitle
        '��ʾ�Ի���
        i = GetOpenFileName(OpenFolder)
        If i >= 1 Then
            FileName = .lpstrFile
        End If
    End With
    FileName = TrimA(FileName)
    ShowOpen = FileName
End Function

'�����ļ��Ի���
Public Function ShowSave(ByVal yTitle As String, ByVal Filter As String, ByVal hwnd As Long) As String
    Dim i As Integer
    Dim saveFolder As OPENFILENAME
    Dim FileName As String
    With saveFolder
        .lStructSize = Len(saveFolder)
        .hwndOwner = hwnd
        .hInstance = App.hInstance
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        '.lpstrInitialDir = App.path
        .flags = 6148
        '�����ļ�����
        .lpstrFilter = Filter
        '�Ի�������
        .lpstrTitle = yTitle
        '��ʾ�Ի���
        i = GetSaveFileName(saveFolder)
        If i >= 1 Then
            FileName = .lpstrFile
        End If
    End With
    FileName = TrimA(FileName)
    If FileName = "" Then Exit Function
    If getTuo(FileName, True) <> ".tdl" Then
        FileName = FileName & ".tdl"
    End If
    ShowSave = FileName
End Function

'���ļ���
Public Function showFolder(ByVal yTitle As String, ByVal hwnd As Long) As String     '���ļ��� ��ѡ��һ���ļ���
Dim bi As BROWSEINFO
Dim r As Long
Dim pidl As Long
Dim path As String
Dim pos As Integer
'���
bi.hOwner = hwnd
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


