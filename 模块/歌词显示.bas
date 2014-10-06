Attribute VB_Name = "¸è´Ê´°ÌåÏÔÊ¾Ä£¿é"
Option Explicit

Public Sub drawLRC()
    Dim i As Integer
    With frmLRC
        For i = 1 To 10
            If lrc_i - 5 + i > 0 Then
                .Label2(i).Caption = mlrc(lrc_i - 5 + i).Caption
                .Label2(i).Left = (frmLRC.Width - .Label2(i).Width) / 2
            Else
                .Label2(i).Caption = ""
            End If
        Next i
    End With
    
End Sub

Public Sub getcurLRC(ByVal curTime As Long)
Dim i As Integer
For i = 1 To 499
    If mlrc(i).Time = curTime Then
        lrc_i = i
        Exit Sub
    End If
Next i
End Sub
