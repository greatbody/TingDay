Attribute VB_Name = "��ʴ�����ʾģ��"
Option Explicit
Public Sub drawLRC(ByVal curtime As Long)
'��ȡ���λ��
Dim mSeccond As Long
Dim I As Integer
Dim lrc_i As Integer '���λ��
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

