Attribute VB_Name = "字符串编码转换"
Option Explicit
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
        '.Charset = encoding
        EncodingConvertor = .ReadText
        .Close
    End With
    Set objStream = Nothing
    If Err.Number <> 0 Then
        EncodingConvertor = ""
    End If
End Function

