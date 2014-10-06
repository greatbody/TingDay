Attribute VB_Name = "��ȡ��ҳԴ��"
Option Explicit
Public Function htmlStr(url$) As String
'-----ͨ�ñ�������---
 Dim xmlHttp
 Dim content As Variant
 '----��������------------------------------
 
 '��������Ƿ�����
 If GetNetConnectString = False Then Exit Function
 '��ַת��
 url = GBtoUTF8(url)
 '��ȡԴ��
 Set xmlHttp = CreateObject("Microsoft.XMLHTTP")
 xmlHttp.Open "GET", url, False
 DoEvents
  xmlHttp.send 'Ϊʲô��ʱ��Ҫ�ܾò����л�Ӧ��
 content = xmlHttp.responsebody
 If xmlHttp.readyState = 4 And CStr(content) <> "" Then
    htmlStr = EncodingConvertor(content, "utf-8")
 End If
 Set xmlHttp = Nothing
End Function

'������תΪansi��ʽ ������UTF-8��ʽ��
Public Function EncodingConvertor(ByVal content As Variant, ByVal encoding As String) As String 'ת��
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

'ת��ΪUTF-8�����URL���� GBtoUTF8
Public Function GBtoUTF8(ByVal szInput As String) As String
Dim wch, uch, szRet
Dim X
Dim nAsc, nAsc2, nAsc3
'����������Ϊ�գ����˳�����
If szInput = "" Then
    GBtoUTF8 = szInput
    Exit Function
End If
'��ʼת��
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
