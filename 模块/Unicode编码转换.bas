Attribute VB_Name = "Unicode编码转换"
Option Explicit

Public Function Unicode8Decode(bTemp() As Byte) As String
  '解码UNICODE UTF-8
  Dim i As Long
  Dim k As Long
 
  Dim strReturn As String
  Dim strTmp() As Byte
  Dim Code As Long
  Dim Code1 As Long
  Dim Code2 As Long
  Dim Code3 As Long
  Dim Code4 As Long
  Dim bNo As Long
 
  k = UBound(bTemp)
  ReDim strTmp(k * 2)
  bNo = 0
  For i = 0 To k
    If (bTemp(i) And 128) = 0 Then
      strTmp(bNo) = bTemp(i)
      bNo = bNo + 1
    ElseIf (bTemp(i) And 252) = 252 Then '11111100
      Code1 = (bTemp(i) And 1) * 64 + bTemp(i + 1) And 63
      Code2 = (bTemp(i + 2) And 63) * 4 + (bTemp(i + 3) And 48) \ 16
      Code3 = (bTemp(i + 3) And 15) * 16 + (bTemp(i + 4) And 60) \ 4
      Code4 = (bTemp(i + 4) And 3) * 64 + (bTemp(i + 5) And 63)
      Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
      Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
      i = i + 5
      strTmp(bNo) = Code And 255
      strTmp(bNo + 1) = Code \ 256
      strTmp(bNo + 1) = Code \ 65536
      strTmp(bNo + 1) = Code \ 16777216
      bNo = bNo + 4

    ElseIf (bTemp(i) And 248) = 248 Then '11111000
      Code1 = (bTemp(i) And 3)
      Code2 = (bTemp(i + 1) And 63) * 4 + (bTemp(i + 2) And 48) \ 16
      Code3 = (bTemp(i + 2) And 15) * 16 + (bTemp(i + 3) And 60) \ 4
      Code4 = (bTemp(i + 3) And 3) * 64 + (bTemp(i + 4) And 63)
      Code = ((Code1 * 256 + Code2) * 256 + Code3) * 256 + Code4
      Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
      i = i + 4
      strTmp(bNo) = Code And 255
      strTmp(bNo + 1) = Code \ 256
      strTmp(bNo + 1) = Code \ 65536
      strTmp(bNo + 1) = Code \ 16777216
      bNo = bNo + 4

    ElseIf (bTemp(i) And 240) = 240 Then '11110000
      Code1 = (bTemp(i) And 7) * 8 + (bTemp(i + 1) And 48) \ 16
      Code2 = (bTemp(i + 1) And 15) * 16 + (bTemp(i + 2) And 60) \ 4
      Code3 = (bTemp(i + 2) And 3) * 64 + (bTemp(i + 3) And 63)
      Code = (Code1 * 256 + Code2) * 256 + Code3
      Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
      i = i + 3
      strTmp(bNo) = Code And 255
      strTmp(bNo + 1) = Code \ 256
      strTmp(bNo + 1) = Code \ 65536
      strTmp(bNo + 1) = Code \ 16777216
      bNo = bNo + 4

    ElseIf (bTemp(i) And 224) = 224 Then '11100000
      Code1 = (bTemp(i) And 15) * 16 + (bTemp(i + 1) And 60) \ 4
      Code2 = (bTemp(i + 1) And 3) * 64 + (bTemp(i + 2) And 63)
      Code = Code1 * 256 + Code2
      Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
      i = i + 2
      strTmp(bNo) = Code And 255
      strTmp(bNo + 1) = Code \ 256
      bNo = bNo + 2

    ElseIf (bTemp(i) And 192) = 192 Then '11000000
      Code1 = (bTemp(i) And 28) \ 4
      Code2 = (bTemp(i) And 3) * 64 + (bTemp(i + 1) And 63)
      Code = Code1 * 256 + Code2
      Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
      i = i + 1
      strTmp(bNo) = Code And 255
      strTmp(bNo + 1) = Code \ 256
      bNo = bNo + 2
     
    End If
  Next
  ReDim Preserve strTmp(bNo - 1)
  strReturn = StrConv(strTmp, vbUnicode)
  Unicode8Decode = strReturn
End Function

Public Function Unicode8Encode(bTemp As String) As Byte()
  '编码UNICODE UTF-8
  Dim i As Long
  Dim j As Long
  Dim k As Long
 
  Dim strTotal() As Byte
  Dim strTmp As String
 
  Dim Code As Long
  Dim Code1 As Long
  Dim Code2 As Long
  Dim Code3 As Long
  Dim Code4 As Long
  Dim Code5 As Long
  Dim Code6 As Long
  '已生成的字节数
  Dim bNo As Long
 
  k = Len(bTemp)
  bNo = 0
  ReDim strTotal(k * 3)
  For i = 1 To k
    Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
    If Code < 128& Then
      strTotal(bNo) = Code
      bNo = bNo + 1
      If bNo > 422386 Then
        Debug.Print Code
      End If
    ElseIf Code < 2048& Then
      Code1 = ((Code And 1984&) \ 32&) + 192&
      Code2 = (Code And 63&) + 128&
      strTotal(bNo) = Code1
      strTotal(bNo + 1) = Code2
      bNo = bNo + 2
     
    ElseIf Code < 65536 Then
      Code1 = ((Code And 61440) \ 4096&) + 224&
      Code2 = ((Code And 4032&) \ 64&) + 128&
      Code3 = (Code And 63&) + 128&
      strTotal(bNo) = Code1
      strTotal(bNo + 1) = Code2
      strTotal(bNo + 2) = Code3
      bNo = bNo + 3
     
    ElseIf Code < 2097152 Then
      Code1 = ((Code And 1835008) \ 262144) + 240&
      Code2 = ((Code And 258048) \ 4096&) + 128&
      Code3 = ((Code And 4032&) \ 64&) + 128&
      Code4 = (Code And 63&) + 128&
      strTotal(bNo) = Code1
      strTotal(bNo + 1) = Code2
      strTotal(bNo + 2) = Code3
      strTotal(bNo + 3) = Code4
      bNo = bNo + 4
     
    ElseIf Code < 67108864 Then
      Code1 = ((Code And 50331648) \ 16777216) + 248&
      Code2 = ((Code And 16515072) \ 262144) + 128&
      Code3 = ((Code And 258048) \ 4096&) + 128&
      Code4 = ((Code And 4032&) \ 64&) + 128&
      Code5 = (Code And 63&) + 128&
      strTotal(bNo) = Code1
      strTotal(bNo + 1) = Code2
      strTotal(bNo + 2) = Code3
      strTotal(bNo + 3) = Code4
      strTotal(bNo + 4) = Code5
      bNo = bNo + 5
     
    Else
      Code1 = IIf(Code And 1073741824 = 1073741824, 253&, 252&)
      Code2 = ((Code And 1056964608) \ 16777216) + 128&
      Code3 = ((Code And 16515072) \ 262144) + 128&
      Code4 = ((Code And 258048) \ 4096&) + 128&
      Code5 = ((Code And 4032&) \ 64&) + 128&
      Code6 = (Code And 63&) + 128&
      strTotal(bNo) = Code1
      strTotal(bNo + 1) = Code2
      strTotal(bNo + 2) = Code3
      strTotal(bNo + 3) = Code4
      strTotal(bNo + 4) = Code5
      strTotal(bNo + 5) = Code6
      bNo = bNo + 6
    End If
  Next
  ReDim Preserve strTotal(bNo - 1)
  Unicode8Encode = strTotal
End Function

Public Function UnicodeDecode(bTemp() As Byte, Optional BigEndian As Boolean = False) As String
  '解码UNICODE UTF-16
  Dim i As Long
 
  Dim strTotal() As Byte
  Dim strReturn As String
  Dim Code As Long
  Dim Code1 As Long
  Dim Code2 As Long
  Dim bNo As Long
 
  bNo = 0
  ReDim strTotal(UBound(bTemp))
  If BigEndian Then
    For i = LBound(bTemp) To UBound(bTemp) Step 2
      Code1 = bTemp(i)
      Code2 = bTemp(i + 1)
      Code = Code1 * 256 + Code2
      If Code > 255 Then
        Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
        strTotal(bNo) = Code And 255
        strTotal(bNo + 1) = Code \ 256
        bNo = bNo + 2
      Else
        strTotal(bNo) = Code
        bNo = bNo + 1
      End If
    Next
  Else
    For i = LBound(bTemp) To UBound(bTemp) Step 2
      Code1 = bTemp(i)
      Code2 = bTemp(i + 1)
      Code = Code2 * 256 + Code1
      If Code > 255 Then
        Code = CLng("&H" + Hex(AscW(StrConv(ChrW(Code), vbFromUnicode))))
        strTotal(bNo) = Code And 255
        strTotal(bNo + 1) = Code \ 256
        bNo = bNo + 2
      Else
        strTotal(bNo) = Code
        bNo = bNo + 1
      End If
    Next
  End If
  ReDim Preserve strTotal(bNo - 1)
  strReturn = StrConv(strTotal, vbUnicode)
  UnicodeDecode = strReturn
End Function

Public Function UnicodeEncode(bTemp As String, Optional BigEndian As Boolean = False) As Byte()
  '编码UNICODE UTF-16
  Dim i As Long
  Dim k As Long
 
  Dim strTotal() As Byte
  Dim Code As Long
  Dim bNo As Long
 
  k = Len(bTemp)
  ReDim strTotal(k * 2)
  bNo = 0
  If BigEndian Then
    For i = 1 To k
      Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
      strTotal(bNo) = Code \ 256
      strTotal(bNo + 1) = Code And 255
      bNo = bNo + 2
    Next
  Else
    For i = 1 To k
      Code = CLng("&H" + Hex(AscW(Mid(bTemp, i, 1))))
      strTotal(bNo) = Code And 255
      strTotal(bNo + 1) = Code \ 256
      bNo = bNo + 2
    Next
  End If
  ReDim Preserve strTotal(bNo - 1)
  UnicodeEncode = strTotal
End Function

