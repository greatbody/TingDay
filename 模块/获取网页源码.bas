Attribute VB_Name = "获取网页源码"
Option Explicit
Public Function htmlStr(url$) As String
'-----通用变量声明---
 Dim xmlHttp
 Dim content As Variant
 '----声明结束------------------------------
 
 '检查网络是否链接
 If GetNetConnectString = False Then Exit Function
 '地址转换
 url = GBtoUTF8(url)
 '获取源码
 Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
 xmlHttp.Open "GET", url, False
 DoEvents
  xmlHttp.send '为什么有时候要很久才能有回应？
 content = xmlHttp.responsebody
 If xmlHttp.readyState = 4 And CStr(content) <> "" Then
    htmlStr = EncodingConvertor(content, "utf-8")
 End If
 Set xmlHttp = Nothing
End Function

'把内容转为ansi格式 或者是UTF-8格式！
Public Function EncodingConvertor(ByVal content As Variant, ByVal encoding As String) As String '转码
    Dim objStream As Object
    Set objStream = CreateObject("Adodb.Stream")
    With objStream
        .Type = 1
        .mode = 3
        .Open
        .Write content
        .Position = 0
        .Type = 2
        .Charset = encoding
        EncodingConvertor = .ReadText
        .Close
    End With
    Set objStream = Nothing
    If Err.Number <> 0 Then
        EncodingConvertor = ""
    End If
End Function

'转换为UTF-8编码的URL编码 GBtoUTF8
Public Function GBtoUTF8(ByVal szInput As String) As String
Dim wch, uch, szRet
Dim X
Dim nAsc, nAsc2, nAsc3
'如果输入参数为空，则退出函数
If szInput = "" Then
    GBtoUTF8 = szInput
    Exit Function
End If
'开始转换
For X = 1 To Len(szInput)
    wch = Mid(szInput, X, 1)
    nAsc = AscW(wch)
    If nAsc < 0 Then nAsc = nAsc + 65536
        If (nAsc And &HFF80) = 0 Then
            szRet = szRet & wch
        Else
            If (nAsc And &HF000) = 0 Then
                uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            Else
                uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
                Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
                Hex(nAsc And &H3F Or &H80)
                szRet = szRet & uch
            End If
        End If
Next
GBtoUTF8 = szRet
End Function
