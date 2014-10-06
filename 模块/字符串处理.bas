Attribute VB_Name = "字符串处理"
Option Explicit
'截取最前符合条件的字符串
Public Function cutStr(ByVal youString As String, ByVal strStr As String, ByVal Lstr As String, ByVal N As Integer) As String
Dim intStr As Long
Dim intL As Long
Dim i As Integer
'初始化
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
'获取路径、文件所在的文件夹名称
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
'获取文件所在的文件夹路径
Public Function getPathName(ByVal FileName As String) As String
Dim i As Integer
i = InStrRev(FileName, "\")
FileName = Left(FileName, i)
getPathName = FileName
End Function
'去掉拓展名
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

'去除空格
Public Function TrimA(ByVal FileName As String) As String
    FileName = Trim(FileName)
    TrimA = Replace(FileName, Chr(0), "")
End Function
'获取拓展名
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
'秒变分
Public Function SeccondToMin(ByVal Seccond As Integer) As String
    Dim Minte As String  '分
    Dim mSecond As String  '秒
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
