Attribute VB_Name = "�ַ�������ת��"
Option Explicit
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
        '.Charset = encoding
        EncodingConvertor = .ReadText
        .Close
    End With
    Set objStream = Nothing
    If Err.Number <> 0 Then
        EncodingConvertor = ""
    End If
End Function

