Attribute VB_Name = "UTF�ı��ļ�����"
Option Explicit

'mTextUTF.bas
'ģ�飺UTF�ı��ļ�����
'���ߣ�zyl910
'�汾��1.0
'���ڣ�2006-1-23


'== ˵�� ===================================================
'֧��Unicode������ı��ļ���д����ʱ֧��ANSI��UTF-8��UTF-16LE��UTF-16BE�⼸�ֱ����ı�


'== ���¼�¼ ===============================================
'[V1.0] 2006-1-23
'1.֧�������ANSI��UTF-8��UTF-16LE��UTF-16BE�⼸�ֱ����ı�


'## ����Ԥ������ #########################################
'== ȫ�ֳ��� ===============================================
'IncludeAPILib��������API�⣬��ʱ����Ҫ�ֶ�дAPI����


'## API ####################################################
#If IncludeAPILib = 0 Then
'== File ===================================================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1

Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const Create_NEW = 1
Private Const Create_ALWAYS = 2
Private Const OPEN_EXISTING = 3
Private Const OPEN_ALWAYS = 4
Private Const TRUNCATE_EXISTING = 5

Private Const FILE_ATTRIBUTE_NORMAL = &H80

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2


'== Unicode ================================================

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpWideCharStr As Any, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByRef lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

#End If


'###########################################################

'Unicode�����ʽ
Public Enum UnicodeEncodeFormat
UEF_ANSI = 0 'ANSI+DBCS
UEF_UTF8 'UTF-8
UEF_UTF16LE 'UTF-16LE
UEF_UTF16BE 'UTF-16BE
UEF_UTF32LE 'UTF-32LE
UEF_UTF32BE 'UTF-32BE

UEF_Auto = -1 '�Զ�ʶ�����

'������Ŀ
[_UEF_Min] = UEF_ANSI
[_UEF_Max] = UEF_UTF32BE

End Enum

'ANSI+DBCS��ʽ���ı���ʹ�õĴ���ҳ��Ĭ��Ϊ0����ʾʹ��ϵͳ��ǰ����ҳ���������øò���ʵ�ֶ�ȡ�������������ı����������� ��������ƽ̨�� ��ȡ ��������ƽ̨���ɵ�txt���ͽ�����Ϊ950
Public UEFCodePage As Long

'�ж�BOM
'����ֵ��BOM��ռ�ֽ�
'dwFirst��[in]�ļ��ʼ��4���ֽ�
'fmt��[out]���ر�������
Public Function UEFCheckBOM(ByVal dwFirst As Long, ByRef fmt As UnicodeEncodeFormat) As Long
If dwFirst = &HFEFF& Then
fmt = UEF_UTF32LE
UEFCheckBOM = 4
ElseIf dwFirst = &HFFFE0000 Then
fmt = UEF_UTF32BE
UEFCheckBOM = 4
ElseIf (dwFirst And &HFFFF&) = &HFEFF& Then
fmt = UEF_UTF16LE
UEFCheckBOM = 2
ElseIf (dwFirst And &HFFFF&) = &HFFFE& Then
fmt = UEF_UTF16BE
UEFCheckBOM = 2
ElseIf (dwFirst And &HFFFFFF) = &HBFBBEF Then
fmt = UEF_UTF8
UEFCheckBOM = 3
Else
fmt = UEF_ANSI
UEFCheckBOM = 0
End If
End Function

'����BOM
'����ֵ��BOM��ռ�ֽ�
'fmt��[in]��������
'dwFirst��[out]�ļ��ʼ��4���ֽ�
Public Function UEFMakeBOM(ByVal fmt As UnicodeEncodeFormat, ByRef dwFirst As Long) As Long
Select Case fmt
Case UEF_UTF8
dwFirst = &HBFBBEF
UEFMakeBOM = 3
Case UEF_UTF16LE
dwFirst = &HFEFF&
UEFMakeBOM = 2
Case UEF_UTF16BE
dwFirst = &HFFFE&
UEFMakeBOM = 2
Case UEF_UTF32LE
dwFirst = &HFEFF&
UEFMakeBOM = 4
Case UEF_UTF32BE
dwFirst = &HFFFE0000
UEFMakeBOM = 4
Case Else
dwFirst = 0
UEFMakeBOM = 0
End Select
End Function

'�ж��ı��ļ��ı�������
'����ֵ���������͡��ļ��޷���ʱ������UEF_Auto
'FileName���ļ���
Public Function UEFCheckTextFileFormat(ByVal FileName As String) As UnicodeEncodeFormat
Dim hFile As Long
Dim dwFirst As Long
Dim nNumRead As Long

'���ļ�
hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
UEFCheckTextFileFormat = UEF_Auto
Exit Function
End If

'�ж�BOM
dwFirst = 0
Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
nNumRead = UEFCheckBOM(dwFirst, UEFCheckTextFileFormat)
'Debug.Print nNumRead

'�ر��ļ�
Call CloseHandle(hFile)

End Function

'��ȡ�ı��ļ�
'����ֵ����ȡ���ı�������vbNullString��ʾ�ļ��޷���
'FileName��[in]�ļ���
'fmt��[in,out]ʹ�ú����ı������ʽ����ȡ�ı���ΪUEF_Autoʱ��ʾ�Զ��жϣ�����fmt���������ı����ñ����ʽ
Public Function UEFLoadTextFile(ByVal FileName As String, Optional ByRef fmt As UnicodeEncodeFormat = UEF_Auto) As String
Dim hFile As Long
Dim nFileSize As Long
Dim nNumRead As Long
Dim dwFirst As Long
Dim CurFmt As UnicodeEncodeFormat
Dim cbBOM As Long
Dim cbTextData As Long
Dim CurCP As Long
Dim byBuf() As Byte
Dim cchStr As Long
Dim I As Long
Dim byTemp As Byte

'�ж�fmt��Χ
If fmt <> UEF_Auto Then
If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
GoTo FunEnd
End If
End If

'���ļ�
hFile = CreateFile(FileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, ByVal 0&)
If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
GoTo FunEnd
End If

'�ж��ļ���С
nFileSize = GetFileSize(hFile, nNumRead)
If nNumRead <> 0 Then '����4GB
GoTo FreeHandle
End If
If nFileSize < 0 Then '����2GB
GoTo FreeHandle
End If

'�ж�BOM
dwFirst = 0
Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
cbBOM = UEFCheckBOM(dwFirst, CurFmt)

'�ָ��ļ�ָ��
If fmt = UEF_Auto Then '�Զ��ж�
fmt = CurFmt
'cbBOM = cbBOM
Else '�ֶ����ñ���
If fmt = CurFmt Then '��������ͬ�������BOM���
'cbBOM = cbBOM
Else '���벻ͬ����ô��������
cbBOM = 0
End If
End If
Call SetFilePointer(hFile, cbBOM, ByVal 0&, FILE_BEGIN)
cbTextData = nFileSize - cbBOM

'��ȡ����
UEFLoadTextFile = ""
Select Case fmt
Case UEF_ANSI, UEF_UTF8
'�ж�Ӧʹ�õ�CodePage
CurCP = IIf(fmt = UEF_UTF8, CP_UTF8, UEFCodePage)

'���仺����
On Error GoTo FreeHandle
ReDim byBuf(0 To cbTextData - 1)
On Error GoTo 0

'��ȡ����
nNumRead = 0
Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)

'ȡ��Unicode�ı�����
cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal 0&, ByVal 0&)
If cchStr > 0 Then
'�����ַ����ռ�
On Error GoTo FreeHandle
UEFLoadTextFile = String$(cchStr, 0)
On Error GoTo 0

'ȡ���ı�
cchStr = MultiByteToWideChar(CurCP, 0, byBuf(0), nNumRead, ByVal StrPtr(UEFLoadTextFile), cchStr + 1)

End If

Case UEF_UTF16LE
cchStr = (cbTextData + 1) \ 2

'�����ַ����ռ�
On Error GoTo FreeHandle
UEFLoadTextFile = String$(cchStr, 0)
On Error GoTo 0

'ȡ���ı�
nNumRead = 0
Call ReadFile(hFile, ByVal StrPtr(UEFLoadTextFile), cbTextData, nNumRead, ByVal 0&)

'�����ı�����
cchStr = (nNumRead + 1) \ 2
If cchStr > 0 Then
If Len(UEFLoadTextFile) > cchStr Then
UEFLoadTextFile = Left$(UEFLoadTextFile, cchStr)
End If
Else
UEFLoadTextFile = ""
End If

Case UEF_UTF16BE
'���仺����
On Error GoTo FreeHandle
ReDim byBuf(0 To cbTextData - 1)
On Error GoTo 0

'��ȡ����
nNumRead = 0
Call ReadFile(hFile, byBuf(0), cbTextData, nNumRead, ByVal 0&)

If nNumRead > 0 Then
'�����ֽڷ�ת�����ֽ�
For I = 0 To nNumRead - 1 - 1 Step 2 '��-1��Ϊ�˱�����������Ǹ��ֽ�
byTemp = byBuf(I)
byBuf(I) = byBuf(I + 1)
byBuf(I + 1) = byTemp
Next I

'ȡ���ı�
UEFLoadTextFile = byBuf 'VB����String�е��ַ���������Byte����ֱ��ת��

End If

Case UEF_UTF32LE
UEFLoadTextFile = vbNullString '��ʱ��֧��
Case UEF_UTF32BE
UEFLoadTextFile = vbNullString '��ʱ��֧��
Case Else
Debug.Assert False
End Select

FreeHandle:
'�ر��ļ�
Call CloseHandle(hFile)

FunEnd:
End Function


'�����ı��ļ�
'����ֵ���Ƿ�ɹ�
'FileName��[in]�ļ���
'sText��[in]��������ı�
'IsAppend��[in]�Ƿ�����ӷ�ʽ
'fmt��[in,out]ʹ�ú����ı������ʽ���洢�ı�����IsAppend=Trueʱ����UEF_Auto�Զ��жϣ�����fmt���������ı����ñ����ʽ
'DefFmt��[in]��ʹ�����ģʽʱ�����ļ���������fmt = UEF_AutoʱӦʹ�õı����ʽ
Public Function UEFSaveTextFile(ByVal FileName As String, _
ByRef sText As String, Optional ByVal IsAppend As Boolean = False, _
Optional ByRef fmt As UnicodeEncodeFormat = UEF_Auto, Optional ByVal DefFmt As UnicodeEncodeFormat = UEF_ANSI) As Boolean
Dim hFile As Long
Dim nFileSize As Long
Dim nNumRead As Long
Dim dwFirst As Long
Dim cbBOM As Long
Dim CurCP As Long
Dim byBuf() As Byte
Dim cbBuf As Long
Dim I As Long
Dim byTemp As Byte

'�ж�fmt��Χ
If IsAppend And (fmt = UEF_Auto) Then
Else
If fmt < [_UEF_Min] Or fmt > [_UEF_Max] Then
GoTo FunEnd
End If
End If

'���ļ�
hFile = CreateFile(FileName, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, IIf(IsAppend, OPEN_ALWAYS, Create_ALWAYS), FILE_ATTRIBUTE_NORMAL, ByVal 0&)
If INVALID_HANDLE_VALUE = hFile Then '�ļ��޷���
GoTo FunEnd
End If

'�ж��ļ���С
nFileSize = GetFileSize(hFile, nNumRead)
If nFileSize = 0 And nNumRead = 0 Then '�ļ���СΪ0�ֽ�
IsAppend = False '��ʱ��ҪдBOM��־
If fmt = UEF_Auto Then fmt = DefFmt
End If

'�ж�BOM
If IsAppend And (fmt = UEF_Auto) Then
dwFirst = 0
Call ReadFile(hFile, dwFirst, 4, nNumRead, ByVal 0&)
cbBOM = UEFCheckBOM(dwFirst, fmt)
ElseIf IsAppend = False Then
cbBOM = UEFMakeBOM(fmt, dwFirst)
End If

'�ļ�ָ�붨λ
Call SetFilePointer(hFile, 0, ByVal 0&, IIf(IsAppend, FILE_END, FILE_BEGIN))

'дBOM
If IsAppend = False Then
If cbBOM > 0 Then
Call WriteFile(hFile, dwFirst, cbBOM, nNumRead, ByVal 0&)
End If
End If

'д�ı�����
If Len(sText) > 0 Then
Select Case fmt
Case UEF_ANSI, UEF_UTF8
'�ж�Ӧʹ�õ�CodePage
CurCP = IIf(fmt = UEF_UTF8, CP_UTF8, UEFCodePage)

'ȡ�û�������С
cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), ByVal 0&, 0, ByVal 0&, ByVal 0&)
If cbBuf > 0 Then
'���仺����
On Error GoTo FreeHandle
ReDim byBuf(0 To cbBuf)
On Error GoTo 0

'ת���ı�
cbBuf = WideCharToMultiByte(CurCP, 0, ByVal StrPtr(sText), Len(sText), byBuf(0), cbBuf + 1, ByVal 0&, ByVal 0&)

'д�ļ�
Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)

UEFSaveTextFile = True

End If

Case UEF_UTF16LE
'д�ļ�
Call WriteFile(hFile, ByVal StrPtr(sText), LenB(sText), nNumRead, ByVal 0&)

UEFSaveTextFile = True

Case UEF_UTF16BE
'���ַ����е����ݸ��Ƶ�byBuf
On Error GoTo FreeHandle
byBuf = sText
On Error GoTo 0
cbBuf = UBound(byBuf) - LBound(byBuf) + 1

'�����ֽڷ�ת�����ֽ�
For I = 0 To cbBuf - 1 - 1 Step 2 '��-1��Ϊ�˱�����������Ǹ��ֽ�
byTemp = byBuf(I)
byBuf(I) = byBuf(I + 1)
byBuf(I + 1) = byTemp
Next I

'д�ļ�
Call WriteFile(hFile, byBuf(0), cbBuf, nNumRead, ByVal 0&)

UEFSaveTextFile = True

Case UEF_UTF32LE
UEFSaveTextFile = False '��ʱ��֧��
Case UEF_UTF32BE
UEFSaveTextFile = False '��ʱ��֧��
Case Else
Debug.Assert False
End Select
Else
UEFSaveTextFile = True
End If

FreeHandle:
'�ر��ļ�
Call CloseHandle(hFile)

FunEnd:
End Function

