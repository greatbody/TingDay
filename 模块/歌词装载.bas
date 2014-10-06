Attribute VB_Name = "���װ��ģ��"
Option Explicit
'LRC�����Ϣ
Public Type LRC
Time As String  '��λΪ����
Caption As String
End Type
Public Lyric() As LRC  'LRC��ʰ��д����������
'Public iDimention As Integer
Public sGM As String    '�洢��������
Public sGS As String    '�洢��������
Public sZJ As String    '�洢����ר��
Public sGCLY As String   '�洢�����Դ
Public LrcOffset As String 'LRC���ʱ��ƫ����
'LRC��ʷ�������
Public Function LyricsAnalyse(LRCName As String) As Boolean
On Error Resume Next
    Dim FileNumber As Integer, FileCount As Integer, I As Integer, j As Integer, k As Long, L As Integer
    Dim TimeLabelLength As Integer 'ʱ���ǩ�ĳ���
    Dim bFlag As Boolean
    Dim MyValue As String
    Dim sRow As String
    Dim sRow2 As LRC
    Dim sTime As String
    Dim miniute As String    '�����
    Dim second As String     '������
    Dim mSecond As String    '�������
    Dim msg As Integer
    Dim s1 As String, s2 As String
    '       On Error Resume Next
    sGM = "δ֪"  '��Ÿ�����
    sGS = "δ֪" '��Ÿ�������
    sZJ = "δ֪"  '���ר������
    sGCLY = "δ֪"  '��Ÿ�����Ժδ���������
    If exitFile(LRCName) = False Then Exit Function
    FileCount = 0: j = 0: TimeLabelLength = 0: bFlag = True
    If Dir(LRCName) <> "" Then
        LyricsAnalyse = True
        LrcOffset = "0"  '��Ÿ��ƫ����
        FileNumber = FreeFile
        Open (LRCName) For Input As #FileNumber '���ļ������ж�ȡLRC���
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, MyValue '��ȡһ�и�ʵ�����MyValue
            MyValue = Trim(MyValue)
            For I = 1 To Len(MyValue)
                sRow = Mid(MyValue, I, 1)
                If sRow = "[" Then
                    FileCount = FileCount + 1 'FileCount��������ȷ������ж�����
                    L = I + 1
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
            
            '����LRC���ͷ
            If InStr(LCase(MyValue), "[ti:") > 0 Then
                I = InStr(LCase(MyValue), "[ti:")
                Mid(MyValue, I, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGM = IIf(Trim(MyValue) <> "", Trim(MyValue), "δ֪") '��Ÿ�����
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[ar:") > 0 Then
                I = InStr(LCase(MyValue), "[ar:")
                Mid(MyValue, I, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGS = IIf(Trim(MyValue) <> "", Trim(MyValue), "δ֪") '��Ÿ�������
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[al:") > 0 Then
                I = InStr(LCase(MyValue), "[al:")
                Mid(MyValue, I, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sZJ = IIf(Trim(MyValue) <> "", Trim(MyValue), "δ֪") '���ר������
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[by:") > 0 Then
                I = InStr(LCase(MyValue), "[by:")
                Mid(MyValue, I, 4) = "    "
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                sGCLY = IIf(Trim(MyValue) <> "", Trim(MyValue), "δ֪") '��Ÿ�����Ժδ�
                j = j + 1
            End If
            If InStr(LCase(MyValue), "[offset:") > 0 Then
                I = InStr(LCase(MyValue), "[offset:")
                Mid(MyValue, I, 8) = Space(8)
                MyValue = Mid(MyValue, 2, Len(MyValue) - 2)
                LrcOffset = IIf(Trim(MyValue) <> "", Trim(MyValue), "0") '��Ÿ��ƫ����
                j = j + 1
            End If
        Loop
        Close #FileNumber
        
        
        ReDim Lyric(Trim(Str(FileCount - j - 1))) As LRC  '���¶���ʵ�ʸ�ʵ������飬��FileCount - j��Ԫ��
        '����LRC����ı�
        FileCount = 0
        FileNumber = FreeFile
        Open (LRCName) For Input As #FileNumber '���ļ������ж�ȡLRC���
        Do While Not EOF(FileNumber)
            Line Input #FileNumber, MyValue '��ȡһ�и�ʵ�����MyValue
            sRow = Mid(MyValue, 2, 2)
            j = 0: I = 0
            If IsNumeric(sRow) Then '�ж��Ƿ���ʱ���ǩ��
                I = I + 1
                Do While True
                    j = j + TimeLabelLength
                    sRow = Mid(MyValue, j + 2, 2)
                    If IsNumeric(sRow) Then  '�ж��Ƿ���ʱ���ǩ��
                        I = I + 1 'i�ۼӺ�Ҳ����ÿ��ʱ���ǩ�ĸ���
                    Else
                        Exit Do
                    End If
                Loop
                For j = 1 To I 'Step 1
                    sTime = Mid(MyValue, (j - 1) * TimeLabelLength + 1, TimeLabelLength)
                    If TimeLabelLength > 2 Then miniute = Mid(sTime, 2, 2)
                    If TimeLabelLength > 5 Then second = Mid(sTime, 5, 2)
                    If TimeLabelLength > 8 Then mSecond = Mid(sTime, 8, 2)
                    k = Val(miniute) * 60000 + Val(second) * 1000 + Val(mSecond) - Val(LrcOffset)
                    s1 = Trim(Str(k))
                    If Len(s1) = 1 Then s2 = "0000000" & s1
                    If Len(s1) = 2 Then s2 = "000000" & s1
                    If Len(s1) = 3 Then s2 = "00000" & s1
                    If Len(s1) = 4 Then s2 = "0000" & s1
                    If Len(s1) = 5 Then s2 = "000" & s1
                    If Len(s1) = 6 Then s2 = "00" & s1
                    If Len(s1) = 7 Then s2 = "0" & s1
                    If Len(s1) >= 8 Then s2 = s1
                    Lyric(Trim(Str(FileCount))).Caption = Trim(Mid(MyValue, I * TimeLabelLength + 1, Len(MyValue) - (I * TimeLabelLength))) '���
                    Lyric(Trim(Str(FileCount))).Time = Trim(s2) '���ʱ��
                    FileCount = FileCount + 1
                Next

            End If
        Loop
'        iDimention = FileCount - 1
        Close #FileNumber
        'ð������
        For I = 0 To FileCount - 2
            For j = 0 To FileCount - 2 - I
                If Val(Lyric(j).Time) > Val(Lyric(j + 1).Time) Then
                    sRow2 = Lyric(j)
                    Lyric(j) = Lyric(j + 1)
                    Lyric(j + 1) = sRow2
                End If
            Next j
        Next I
    Else
        LyricsAnalyse = False
    End If
End Function
