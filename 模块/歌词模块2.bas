Attribute VB_Name = "歌词模块2"
Option Explicit
'处理LRC歌词
Public Lyric() As String 'LRC歌词按行存放在数组中
'Public iDimention As Integer
Public sGM As String    '存储歌曲名称
Public sGS As String    '存储歌手名字
Public sZJ As String    '存储歌曲专辑
Public sGCLY As String   '存储歌词来源
Public LrcOffset As String 'LRC歌词时间偏移量
Public Function LyricsAnalyse(LRCName As String) As Boolean 'LRC歌词分析函数
    Dim FileNumber As Integer, FileCount As Integer, i As Integer, j As Integer, k As Long, L As Integer
    Dim TimeLabelLength As Integer '时间标签的长度
    Dim bFlag As Boolean
    Dim MyValue As String
    Dim sRow As String
    Dim sTime As String
    Dim miniute As String    '定义分
    Dim second As String     '定义秒
    Dim msecond As String    '定义毫秒
    Dim s1 As String, s2 As String
    '       On Error Resume Next
    sGM = "未知"  '存放歌曲名
    sGS = "未知" '存放歌手名字
    sZJ = "未知"  '存放专辑名字
    sGCLY = "未知"  '存放歌词来自何处，即编者
    '打开LRC歌词文件LyricsName
    FileCount = 0: j = 0: TimeLabelLength = 0: bFlag = True
    If Dir(LRCName) <> "" Then
        LyricsAnalyse = True
        LrcOffset = "0"  '存放歌词偏移量
        FileNumber = FreeFile
        Open (LRCName) For Input As #FileNumber '打开文件，从中读取LRC歌词
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, MyValue '读取一行歌词到变量MyValue
            MyValue = Trim(MyValue)
            For i = 1 To Len(MyValue)
                sRow = Mid(MyValue, i, 1)
                If sRow = "[" Then
                    FileCount = FileCount + 1 'FileCount计数用来确定歌词有多少行
                    L = i + 1
                    sRow = Mid(MyValue, L, 1)
                    If IsNumeric(sRow) Then
                        Do While bFlag
                            sRow = Mid(MyValue, L + 1, 1)
                            If IsNumeric(sRow) Or sRow = ":" Or sRow = "." Then
                                TimeLabelLength = TimeLabelLength + 1
                            End If
                            If sRow = "]" Then TimeLabelLength = TimeLabelLength + 3: bFlag = False: Exit Do
                            L = L + 1
                        Loop
                    End If
                End If
            Next
            '处理LRC歌词头
            If InStr(LCase(MyValue), "[ti:") > 0 Then
                i = InStr(LCase(MyValue), "[ti:")
                Mid(MyValue, i, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGM = IIf(Trim(MyValue) <> "", Trim(MyValue), "未知") '存放歌曲名
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[ar:") > 0 Then
                i = InStr(LCase(MyValue), "[ar:")
                Mid(MyValue, i, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGS = IIf(Trim(MyValue) <> "", Trim(MyValue), "未知") '存放歌手名字
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[al:") > 0 Then
                i = InStr(LCase(MyValue), "[al:")
                Mid(MyValue, i, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sZJ = IIf(Trim(MyValue) <> "", Trim(MyValue), "未知") '存放专辑名字
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[by:") > 0 Then
                i = InStr(LCase(MyValue), "[by:")
                Mid(MyValue, i, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGCLY = IIf(Trim(MyValue) <> "", Trim(MyValue), "未知") '存放歌词来自何处
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[offset:") > 0 Then
                i = InStr(LCase(MyValue), "[offset:")
                Mid(MyValue, i, 8) = Space(8)
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                LrcOffset = IIf(Trim(MyValue) <> "", Trim(MyValue), "0") '存放歌词偏移量
                j = j + 1
            End If
        Loop
        Close #FileNumber
        ReDim Lyric(Trim(Str(FileCount - j - 1))) '重新定义实际歌词的行数组，有FileCount - j个元素
        '处理LRC歌词文本
        FileCount = 0
        FileNumber = FreeFile
        Open (LRCName) For Input As #FileNumber '打开文件，从中读取LRC歌词
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, MyValue '读取一行歌词到变量MyValue
            MyValue = Trim(MyValue)
            sRow = Mid(MyValue, 2, 2)
            j = 0: i = 0
            If IsNumeric(sRow) Then '判断是否是时间标签行
                i = i + 1
                Do While True
                    j = j + TimeLabelLength
                    sRow = Mid(MyValue, j + 2, 2)
                    If IsNumeric(sRow) Then  '判断是否是时间标签行
                        i = i + 1 'i累加后也就是每行时间标签的个数
                    Else
                        Exit Do
                    End If
                Loop
                For j = 1 To i 'Step 1
                    sTime = Mid(MyValue, (j - 1) * TimeLabelLength + 1, TimeLabelLength)
                    If TimeLabelLength > 2 Then miniute = Mid(sTime, 2, 2)
                    If TimeLabelLength > 5 Then second = Mid(sTime, 5, 2)
                    If TimeLabelLength > 8 Then msecond = Mid(sTime, 8, 2)
                    k = Val(miniute) * 60000 + Val(second) * 1000 + Val(msecond) - Val(LrcOffset)
                    s1 = Trim(Str(k))
                    If Len(s1) = 1 Then s2 = "[0000000" & s1 & "]"
                    If Len(s1) = 2 Then s2 = "[000000" & s1 & "]"
                    If Len(s1) = 3 Then s2 = "[00000" & s1 & "]"
                    If Len(s1) = 4 Then s2 = "[0000" & s1 & "]"
                    If Len(s1) = 5 Then s2 = "[000" & s1 & "]"
                    If Len(s1) = 6 Then s2 = "[00" & s1 & "]"
                    If Len(s1) = 7 Then s2 = "[0" & s1 & "]"
                    If Len(s1) >= 8 Then s2 = "[" & s1 & "]"
                    Lyric(Trim(Str(FileCount))) = Trim(s2 & Mid(MyValue, i * TimeLabelLength + 1, Len(MyValue) - (i * TimeLabelLength)))  'i是时间标签个数
                    FileCount = FileCount + 1
                Next

            End If
        Loop
'        iDimention = FileCount - 1
        Close #FileNumber
        '冒泡排序
        For i = 0 To FileCount - 2
            For j = 0 To FileCount - 2 - i
                If Val(Mid(Lyric(j), 2, 8)) > Val(Mid(Lyric(j + 1), 2, 8)) Then
                    sRow = Lyric(j)
                    Lyric(j) = Lyric(j + 1)
                    Lyric(j + 1) = sRow
                End If
            Next j
        Next i
    Else
        LyricsAnalyse = False
    End If
End Function
