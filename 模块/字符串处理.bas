Attribute VB_Name = "�ַ�������"
Option Explicit
'��ȡ��ǰ�����������ַ���
Public Function cutStr(ByVal youString As String, ByVal strStr As String, ByVal Lstr As String, ByVal N As Integer) As String
Dim intStr As Long
Dim intL As Long
Dim i As Integer
'��ʼ��
If youString = "" Then Exit Function
intStr = 1
For i = 1 To N
    intStr = InStr(intStr + 1, youString, strStr)
Next i
If intStr = 0 Then
    cutStr = ""
    Exit Function
End If
intL = InStr(intStr + 1, youString, Lstr)
If intL <> 0 Then cutStr = Mid$(youString, intStr + Len(strStr), intL - intStr - Len(strStr))
End Function
'��ȡ·�����ļ����ڵ��ļ�������
Public Function FolderName(ByVal FileName As String) As String
    Dim Caption As String
    If exitFile(FileName) = True Then
        Caption = Dir(FileName)
        Caption = Left(FileName, Len(FileName) - Len(Caption) - 1)
    Else
        Caption = FileName
    End If
    FolderName = Right(Caption, Len(Caption) - InStrRev(Caption, "\"))
End Function
'��ȡ�ļ����ڵ��ļ���·��
Public Function getPathName(ByVal FileName As String) As String
Dim i As Integer
i = InStrRev(FileName, "\")
FileName = Left(FileName, i)
getPathName = FileName
End Function
'ȥ����չ��
Public Function delTuo(ByVal FileName As String, Optional hvePath As Boolean)
Dim i As Integer
If FileName = "" Then Exit Function
If hvePath = True Then
    If exitFile(FileName) = False Then
        i = InStrRev(FileName, "\")
        FileName = Right(FileName, Len(FileName) - i)
    Else
        FileName = Dir(FileName)
    End If
End If
delTuo = Left$(FileName, Len(FileName) - 4)
End Function
Public Sub rSong()

End Sub

'ȥ���ո�
Public Function TrimA(ByVal FileName As String) As String
    FileName = Trim(FileName)
    TrimA = Replace(FileName, Chr(0), "")
End Function
'��ȡ��չ��
Public Function getTuo(ByVal FileName As String, Optional hvePath As Boolean)
Dim mFileName As String
If Len(FileName) <= 4 Then
    getTuo = ""
    Exit Function
End If
If InStr(FileName, ".") = 0 Then
    getTuo = ""
    Exit Function
End If
If hvePath = True Then
    mFileName = Dir(FileName)
    FileName = mFileName
End If
getTuo = Right$(FileName, Len(FileName) - InStr(FileName, ".") + 1)
End Function
'����
Public Function SeccondToMin(ByVal Seccond As Integer) As String
    Dim Minte As String  '��
    Dim mSecond As String  '��
    Minte = CStr(Seccond \ 60)
    mSecond = CStr(Seccond Mod 60)
    If CInt(Minte) < 10 Then
        Minte = "0" & Minte
    End If
    If CInt(mSecond) < 10 Then
        mSecond = "0" & mSecond
    End If
    SeccondToMin = Minte & ":" & mSecond
End Function
