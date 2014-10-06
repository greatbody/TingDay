Attribute VB_Name = "歌词窗体显示模块"
Option Explicit
Public Sub drawLRC(ByVal curtime As Long)
'获取歌词位置
Dim mSeccond As Long
Dim I As Integer
Dim lrc_i As Integer '歌词位置
mSeccond = curtime * 1000
For I = 1 To UBound(Lyric)
    If Lyric(I - 1).Time <= mSeccond And Lyric(I + 1).Time > mSeccond Then
        lrc_i = I
        Exit For
    ElseIf Lyric(0).Time > mSeccond Then
        lrc_i = 1
        Exit For
    End If
Next I
If lrc_i = 0 Then lrc_i = 1
With frmLRC
    For I = 1 To 10
        If lrc_i - 5 + I > 0 Then
            .Label2(I).Caption = Lyric(lrc_i - 5 + I).Caption
            .Label2(I).Left = (frmLRC.Width - .Label2(I).Width) / 2
        Else
            .Label2(I).Caption = ""
        End If
    Next I
End With
End Sub

